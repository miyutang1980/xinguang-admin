import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

// Requires Playwright/Chromium. Every HTTP request is intercepted with fake data.
const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '../..') + '/';
const output = await fs.mkdtemp(path.join(os.tmpdir(), 'xg-browser-qa-')) + '/';
const AsyncFunction = Object.getPrototypeOf(async function(){}).constructor;
for (const files of [
  ['login-regression.js'],
  ['pickup-setup.js', 'pickup-regression.js'],
  ['pickup-recovery.js'],
  ['details-regression.js'],
  ['pickup-setup.js','linkage-regression.js'],
  ['pickup-latency.js']
]) {
  const fixture = await fs.readFile(path.join(here, 'bootstrap.js'), 'utf8');
  const checks = (await Promise.all(files.map(f => fs.readFile(path.join(here, f), 'utf8')))).join('\n');
  const code = (fixture + '\ntry {\n' + checks + '\n} finally { await ssQABrowser.close(); }')
    .replaceAll('/home/user/workspace/xinguang-admin/', root)
    .replaceAll('/home/user/workspace/frontend_qa/', output);
  await new AsyncFunction(code)();
}
console.log('Screenshots:', output);
