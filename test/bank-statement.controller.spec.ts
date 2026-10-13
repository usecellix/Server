import * as fs from 'fs';
import * as path from 'path';
import { Body, Controller, ExecutionContext, Post, ValidationPipe } from '@nestjs/common';
import { GUARDS_METADATA } from '@nestjs/common/constants';
import { Reflector } from '@nestjs/core';
import { FastifyAdapter, NestFastifyApplication } from '@nestjs/platform-fastify';
import { Test } from '@nestjs/testing';
import { of } from 'rxjs';
import { AuthGuard } from '../src/auth/auth.guard';
import { BankStatementController } from '../src/bank-statement/bank-statement.controller';
import { BankStatementService } from '../src/bank-statement/bank-statement.service';
import { HttpExceptionFilter } from '../src/common/filters/http-exception.filter';
import { applyRouteBodyLimits, ROUTE_BODY_LIMITS } from '../src/common/http/route-body-limits';
import { RequestResponseCaptureInterceptor } from '../src/common/interceptors/request-response-capture.interceptor';
import { ResponseEnvelopeInterceptor } from '../src/common/interceptors/response-envelope.interceptor';
import { sanitizeLogBody } from '../src/common/logging/log-body.util';
import { getCapturedResponse } from '../src/common/logging/request-response-capture.util';
import { RawTable } from '../src/domain-tools/ingestion/raw-table.types';

const fixture = (): RawTable =>
  JSON.parse(
    fs.readFileSync(
      path.join(
        __dirname,
        '..',
        'src',
        'domain-tools',
        'ingestion',
        'fixtures',
        'synthetic-federal-bank-statement.rawtable.json',
      ),
      'utf8',
    ),
  ) as RawTable;

/** A route with no entry in ROUTE_BODY_LIMITS, to show the default still applies elsewhere. */
@Controller('probe')
class ProbeController {
  @Post('echo')
  echo(@Body() body: { padding?: string }) {
    return { received: body.padding?.length ?? 0 };
  }
}

/** Past Fastify's 1 MiB default. */
const OVER_DEFAULT_LIMIT = 'x'.repeat(1_200_000);

describe('POST /ingest/bank-statement', () => {
  let app: NestFastifyApplication;

  beforeAll(async () => {
    const moduleRef = await Test.createTestingModule({
      controllers: [BankStatementController, ProbeController],
      providers: [BankStatementService],
    })
      .overrideGuard(AuthGuard)
      .useValue({ canActivate: () => true })
      .compile();

    app = moduleRef.createNestApplication<NestFastifyApplication>(new FastifyAdapter());
    // Same order as main.ts: hook first, then the global pipe, filter and interceptors.
    applyRouteBodyLimits(app.getHttpAdapter().getInstance());
    app.useGlobalPipes(new ValidationPipe({ whitelist: true, forbidNonWhitelisted: true, transform: true }));
    app.useGlobalFilters(new HttpExceptionFilter());
    app.useGlobalInterceptors(
      new RequestResponseCaptureInterceptor(),
      new ResponseEnvelopeInterceptor(app.get(Reflector)),
    );
    await app.init();
    await app.getHttpAdapter().getInstance().ready();
  });

  afterAll(async () => {
    await app.close();
  });

  it('is behind the auth guard', () => {
    const guards = Reflect.getMetadata(GUARDS_METADATA, BankStatementController) as unknown[];
    expect(guards).toContain(AuthGuard);
  });

  it('returns a verified import in the standard envelope', async () => {
    const response = await app.inject({
      method: 'POST',
      url: '/ingest/bank-statement',
      payload: { rawTable: fixture(), existingSheetNames: ['Sheet1'], activeSheetName: 'Sheet1' },
    });

    expect(response.statusCode).toBe(200);
    const body = response.json();
    expect(body.success).toBe(true);
    expect(body.data.outputSheetName).toBe('Bank Statement');
    expect(body.data.statement.transactionCount).toBe(54);
    expect(body.data.verification.state).toBe('verified');
    expect(body.data.actions[0]).toMatchObject({ type: 'CREATE_SHEET', relativeTo: 'Sheet1' });
  });

  it('accepts a statement larger than the 1 MiB default body limit', async () => {
    // Leading whitespace is valid JSON, so this adds bytes without changing the statement.
    const big = JSON.stringify({ rawTable: fixture() }).replace('{', `{${' '.repeat(1_200_000)}`);
    expect(Buffer.byteLength(big)).toBeGreaterThan(1024 * 1024);
    expect(Buffer.byteLength(big)).toBeLessThan(ROUTE_BODY_LIMITS['/ingest/bank-statement']);

    const response = await app.inject({
      method: 'POST',
      url: '/ingest/bank-statement',
      headers: { 'content-type': 'application/json' },
      payload: big,
    });
    expect(response.statusCode).toBe(200);
    expect(response.json().data.statement.transactionCount).toBe(54);
  });

  it('leaves the 1 MiB default in place for every other route', async () => {
    const response = await app.inject({
      method: 'POST',
      url: '/probe/echo',
      payload: { padding: OVER_DEFAULT_LIMIT },
    });
    expect(response.statusCode).toBe(413);
  });

  it('answers 422 with a reason when the file is not a statement', async () => {
    const response = await app.inject({
      method: 'POST',
      url: '/ingest/bank-statement',
      payload: {
        rawTable: {
          source: 'csv',
          fileName: 'invoices.csv',
          layout: 'grid',
          rows: [{ ref: 'row 1', cells: [{ t: 'Invoice No' }, { t: 'Total' }] }],
        },
      },
    });
    expect(response.statusCode).toBe(422);
    const body = response.json();
    expect(body.success).toBe(false);
    expect(body.error.details.code).toBe('no_header');
    expect(body.error.message).toMatch(/Could not find the transaction table/);
  });

  it('answers 400 for a body that is not a decoded file', async () => {
    const missing = await app.inject({ method: 'POST', url: '/ingest/bank-statement', payload: {} });
    expect(missing.statusCode).toBe(400);

    const malformed = await app.inject({
      method: 'POST',
      url: '/ingest/bank-statement',
      payload: { rawTable: { source: 'pdf', layout: 'positioned', fileName: 'a.pdf', rows: 'nope' } },
    });
    expect(malformed.statusCode).toBe(400);

    const extraField = await app.inject({
      method: 'POST',
      url: '/ingest/bank-statement',
      payload: { rawTable: fixture(), userId: 'someone-else' },
    });
    expect(extraField.statusCode).toBe(400);
  });
});

describe('bank statement contents stay out of the request log', () => {
  function runInterceptor(handler: (...args: never[]) => unknown) {
    const reply = { raw: {} };
    const context = {
      getType: () => 'http',
      getHandler: () => handler,
      switchToHttp: () => ({ getResponse: () => reply }),
    } as unknown as ExecutionContext;
    new RequestResponseCaptureInterceptor()
      .intercept(context, { handle: () => of({ secret: 'transactions' }) })
      .subscribe();
    return getCapturedResponse(reply);
  }

  it('does not capture the response of the import route', () => {
    const handler = BankStatementController.prototype.importBankStatement;
    expect(runInterceptor(handler)).toBeUndefined();
  });

  it('still captures the response of an ordinary route', () => {
    expect(runInterceptor(ProbeController.prototype.echo)).toEqual({
      kind: 'json',
      body: { secret: 'transactions' },
    });
  });

  it('logs only the size of the decoded file from the request body', () => {
    const logged = sanitizeLogBody({ rawTable: fixture(), existingSheetNames: ['Sheet1'] });
    expect(logged).toEqual({
      rawTable: { source: 'pdf', layout: 'positioned', rows: 145 },
      existingSheetNames: ['Sheet1'],
    });
    expect(JSON.stringify(logged)).not.toMatch(/SAMPLE|99990100001234|testupi/);
  });
});
