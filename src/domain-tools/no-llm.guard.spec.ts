import * as fs from 'fs';
import * as path from 'path';

/**
 * Enforces: no domain tool may import or call an LLM client.
 * Spec 06 — "No domain tool contains an LLM call anywhere in its call graph".
 */
const DOMAIN_TOOLS_ROOT = path.join(__dirname);
const FORBIDDEN_IMPORT_PATTERNS = [
  /from\s+['"][^'"]*openrouter[^'"]*['"]/i,
  /from\s+['"][^'"]*openai[^'"]*['"]/i,
  /from\s+['"][^'"]*llm[^'"]*['"]/i,
  /OpenRouterService/,
  /ConversationEngineService/,
  /LlmRouterService/,
  /@anthropic/,
  /chat\.completions/,
  /\.complete\s*\(/,
  /\.stream\s*\(/,
];

function collectSourceFiles(dir: string): string[] {
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  const files: string[] = [];
  for (const entry of entries) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      if (entry.name === 'fixtures' || entry.name === 'node_modules') continue;
      files.push(...collectSourceFiles(full));
      continue;
    }
    if (!entry.name.endsWith('.ts')) continue;
    if (entry.name.endsWith('.spec.ts')) continue;
    if (entry.name.includes('test-utils')) continue;
    files.push(full);
  }
  return files;
}

describe('domain-tools no-LLM guard', () => {
  it('source files under domain-tools must not reference LLM clients', () => {
    const sources = collectSourceFiles(DOMAIN_TOOLS_ROOT);
    expect(sources.length).toBeGreaterThan(5);

    const violations: string[] = [];
    for (const file of sources) {
      const content = fs.readFileSync(file, 'utf8');
      for (const pattern of FORBIDDEN_IMPORT_PATTERNS) {
        if (pattern.test(content)) {
          violations.push(`${path.relative(DOMAIN_TOOLS_ROOT, file)} matches ${pattern}`);
        }
      }
    }

    expect(violations).toEqual([]);
  });
});
