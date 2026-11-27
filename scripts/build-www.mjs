// Copies the web app into www/ for the Windows app (Tauri bundles www/). www/ is generated and gitignored.
import { cpSync, rmSync, readdirSync, mkdirSync, writeFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const out = join(root, 'www');
const skip = new Set(['.git', '.github', '.gitignore', '.vercel', 'node_modules', 'www', 'scripts', 'src-tauri', 'package.json', 'package-lock.json']);
rmSync(out, { recursive: true, force: true });
mkdirSync(out);
for (const name of readdirSync(root)) {
  if (skip.has(name) || name.endsWith('.md')) continue;
  cpSync(join(root, name), join(out, name), { recursive: true });
}
// the app carries its own files: no offline cache needed
writeFileSync(join(out, 'sw.js'), 'self.addEventListener("install",()=>self.skipWaiting());self.addEventListener("activate",e=>e.waitUntil(self.registration.unregister()));\n');
console.log('Built www/');
