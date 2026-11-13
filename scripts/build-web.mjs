// Bundles web/src into one self-contained HTML file (no network requests, no CDN).
//   dist/BGI-Underbilling.html           - standalone copy that opens by double-click
//   cmd/bgi-underbilling/ui/index.html   - copy the Go launcher embeds and serves
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const rootDir = join(dirname(fileURLToPath(import.meta.url)), '..');
const src = (f) => readFileSync(join(rootDir, 'web/src', f), 'utf8');
const version = readFileSync(join(rootDir, 'VERSION'), 'utf8').trim();

const scripts = [
  `globalThis.BGI = globalThis.BGI || {}; globalThis.BGI.version = ${JSON.stringify(version)};`,
  src('xlsx.js'), src('engine.js'), src('app.js')
].join('\n');
if (/<\/script/i.test(scripts)) throw new Error('script source contains a closing script tag');

// Window and taskbar icon: Edge app windows take it from the page, so it is inlined at full size.
const b64 = (f) => readFileSync(join(rootDir, 'assets', f)).toString('base64');
const icon = `<link rel="icon" type="image/svg+xml" href="data:image/svg+xml;base64,${b64('icon.svg')}">\n` +
  `<link rel="icon" type="image/png" sizes="256x256" href="data:image/png;base64,${b64('icon-256.png')}">`;
const template = src('index.template.html').replace('<!--ICON-->', icon);
const fill = (html) => html.replace('/*__STYLES__*/', () => src('styles.css')).replace('/*__SCRIPTS__*/', () => scripts);

// Served copy: the launcher swaps in a per-response nonce and sends the CSP as a header.
const served = fill(template).replace('<!--CSP-->', '');
// Standalone copy: no server, so the policy travels in the file and the nonce attributes are dropped.
const csp = "default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src data:; base-uri 'none'; form-action 'none'";
const standalone = fill(template)
  .replace('<!--CSP-->', `<meta http-equiv="Content-Security-Policy" content="${csp}">`)
  .replaceAll(' nonce="__CSP_NONCE__"', '');

mkdirSync(join(rootDir, 'dist'), { recursive: true });
mkdirSync(join(rootDir, 'cmd/bgi-underbilling/ui'), { recursive: true });
writeFileSync(join(rootDir, 'dist/BGI-Underbilling.html'), standalone);
writeFileSync(join(rootDir, 'cmd/bgi-underbilling/ui/index.html'), served);
console.log(`built v${version}: standalone ${(standalone.length / 1024).toFixed(0)} KB`);
