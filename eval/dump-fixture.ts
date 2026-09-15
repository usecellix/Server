/** Writes the audit fixture to JSON so external tools (e.g. an Excel COM setup script) can seed a real workbook with it. */
import * as fs from 'fs';
import * as path from 'path';
import { buildFixture } from './usecase-fixture';

const out = process.env.OUT ?? path.join(__dirname, 'usecase-fixture.json');
fs.writeFileSync(out, JSON.stringify(buildFixture(), null, 1));
console.log(`fixture written to ${out}`);
