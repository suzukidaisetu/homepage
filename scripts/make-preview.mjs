// Builds a preview folder for viewing inside Claude: the current site/index.html + the latest built content
// from the snapshot branch. Usage (in a clone):
//   git fetch --depth 1 origin snapshot:refs/remotes/origin/snapshot
//   node scripts/make-preview.mjs           -> writes preview/
import fs from 'node:fs/promises';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
const ROOT = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
const OUT = path.join(ROOT, 'preview');
await fs.rm(OUT, { recursive: true, force: true });
await fs.mkdir(OUT, { recursive: true });
// content from the last published build
const files = execFileSync('git', ['ls-tree', '-r', '--name-only', 'origin/snapshot'], { cwd: ROOT }).toString().trim().split('\n');
for (const f of files) {
  if (f === 'index.html' || f === '.nojekyll') continue;
  const buf = execFileSync('git', ['show', `origin/snapshot:${f}`], { cwd: ROOT, maxBuffer: 1 << 28 });
  await fs.mkdir(path.join(OUT, path.dirname(f)), { recursive: true });
  await fs.writeFile(path.join(OUT, f), buf);
}
// the page itself, from the working copy, without the outer document tags (the preview host adds its own)
let html = await fs.readFile(path.join(ROOT, 'site', 'index.html'), 'utf8');
html = html.replace(/<!doctype html>\s*/i, '').replace(/<\/?html[^>]*>\s*/gi, '').replace(/<\/?head>\s*/gi, '').replace(/<\/?body>\s*/gi, '')
  .replace(/<meta charset[^>]*>\s*/i, '').replace(/<meta name="viewport"[^>]*>\s*/i, '');
await fs.writeFile(path.join(OUT, 'index.html'), html);
console.log(`preview/ ready: ${files.length - 2} content files`);
