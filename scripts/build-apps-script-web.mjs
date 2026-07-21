import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';

const SCRIPT_DIR = dirname(fileURLToPath(import.meta.url));
const ROOT_DIR = join(SCRIPT_DIR, '..');
const OUTPUT_PATH = join(ROOT_DIR, 'apps-script', 'WebApp.html');
const SHEETJS_OUTPUT_PATH = join(ROOT_DIR, 'apps-script', 'SheetJS.html');

const SOURCE_PATHS = Object.freeze({
  html: join(ROOT_DIR, 'index.html'),
  styles: join(ROOT_DIR, 'assets', 'styles.css'),
  core: join(ROOT_DIR, 'assets', 'core.js'),
  app: join(ROOT_DIR, 'assets', 'app.js'),
  qrcode: join(ROOT_DIR, 'vendor', 'qrcode.js'),
  xlsx: join(ROOT_DIR, 'vendor', 'xlsx.full.min.js'),
  favicon: join(ROOT_DIR, 'favicon.svg'),
  notices: join(ROOT_DIR, 'THIRD_PARTY_NOTICES.md'),
  sheetJsLicense: join(ROOT_DIR, 'vendor', 'LICENSE-SheetJS.txt')
});

const EXTERNAL_ASSET_TAGS = Object.freeze([
  /\s*<link\b[^>]*\bhref=["']assets\/styles\.css(?:\?[^"']*)?["'][^>]*>\s*/i,
  /\s*<script\b[^>]*\bsrc=["']assets\/config\.js(?:\?[^"']*)?["'][^>]*>\s*<\/script>\s*/i,
  /\s*<script\b[^>]*\bsrc=["']vendor\/qrcode\.js(?:\?[^"']*)?["'][^>]*>\s*<\/script>\s*/i,
  /\s*<script\b[^>]*\bsrc=["']vendor\/xlsx\.full\.min\.js(?:\?[^"']*)?["'][^>]*>\s*<\/script>\s*/i,
  /\s*<script\b[^>]*\bsrc=["']assets\/app\.js(?:\?[^"']*)?["'][^>]*>\s*<\/script>\s*/i
]);

const APP_IMPORT_PATTERN = /^\s*import\s*\{[\s\S]*?\}\s*from\s*["']\.\/core\.js(?:\?[^"']*)?["'];\s*/;
const LOCAL_ASSET_REFERENCE_PATTERN = /(?:src|href)=["'](?:assets|vendor)\/|favicon\.svg(?:\?[^"'\s)]*)?|\.\/core\.js/;

function normalizeText(value) {
  return value.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n');
}

async function readText(path) {
  return normalizeText(await readFile(path, 'utf8'));
}

function replaceExactlyOnce(source, pattern, replacement, description) {
  const matches = source.match(new RegExp(pattern.source, pattern.flags.includes('g') ? pattern.flags : `${pattern.flags}g`));
  if ((matches || []).length !== 1) {
    throw new Error(`${description} should occur exactly once; found ${(matches || []).length}.`);
  }
  // A function replacement prevents JavaScript source sequences such as `$'`
  // from being interpreted as String.prototype.replace substitution tokens.
  return source.replace(pattern, () => replacement);
}

function removeOptionalOnce(source, pattern, description) {
  const flags = pattern.flags.includes('g') ? pattern.flags : `${pattern.flags}g`;
  const matches = source.match(new RegExp(pattern.source, flags)) || [];
  if (matches.length > 1) throw new Error(`${description} should occur at most once; found ${matches.length}.`);
  return source.replace(pattern, '');
}

function escapeAppsScriptScriptlets(source) {
  // Apps Script evaluates <? ... ?> before returning the page. SheetJS contains
  // literal "<?xml" strings, so encode the '<' without changing runtime text.
  return source.replaceAll('<?', '\\x3C?');
}

function assertInlineSafe(source, kind) {
  const closingTag = kind === 'style' ? '</style' : '</script';
  if (source.toLowerCase().includes(closingTag)) {
    throw new Error(`${kind} source contains ${closingTag}, which would terminate its inline element.`);
  }
}

function buildApplicationScript(coreSource, appSource, faviconDataUrl) {
  const importMatch = appSource.match(APP_IMPORT_PATTERN);
  if (!importMatch) throw new Error('assets/app.js must import named exports from ./core.js.');

  const core = coreSource.replace(/^export\s+(?=(?:const|function)\b)/gm, '');
  if (/^\s*(?:import|export)\b/m.test(core)) {
    throw new Error('assets/core.js contains an unsupported import or export declaration.');
  }

  let app = appSource.replace(APP_IMPORT_PATTERN, '');
  if (/^\s*(?:import|export)\b/m.test(app)) {
    throw new Error('assets/app.js contains an unsupported import or export declaration.');
  }

  app = app.replace(/(["'])favicon\.svg(?:\?[^"']*)?\1/g, JSON.stringify(faviconDataUrl));

  const script = [
    '/* First-party application bundle: assets/core.js followed by assets/app.js. */',
    '(() => {',
    "'use strict';",
    core.trim(),
    app.trim(),
    '})();'
  ].join('\n\n');

  assertInlineSafe(script, 'script');
  return escapeAppsScriptScriptlets(script);
}

function inlineScript(label, source) {
  assertInlineSafe(source, 'script');
  return `<script>\n/* ${label} */\n${escapeAppsScriptScriptlets(source.trim())}\n</script>`;
}

async function build() {
  const [htmlSource, styles, core, app, qrcode, xlsx, favicon, notices, sheetJsLicense] = await Promise.all([
    readText(SOURCE_PATHS.html),
    readText(SOURCE_PATHS.styles),
    readText(SOURCE_PATHS.core),
    readText(SOURCE_PATHS.app),
    readText(SOURCE_PATHS.qrcode),
    readText(SOURCE_PATHS.xlsx),
    readText(SOURCE_PATHS.favicon),
    readText(SOURCE_PATHS.notices),
    readText(SOURCE_PATHS.sheetJsLicense)
  ]);

  assertInlineSafe(styles, 'style');
  if (notices.includes('--') || sheetJsLicense.includes('--')) {
    throw new Error('Third-party notice text cannot be embedded safely in an HTML comment.');
  }
  const faviconDataUrl = `data:image/svg+xml;base64,${Buffer.from(favicon, 'utf8').toString('base64')}`;
  const applicationScript = buildApplicationScript(core, app, faviconDataUrl);

  let html = htmlSource;
  html = replaceExactlyOnce(
    html,
    /<meta\b(?=[^>]*\bhttp-equiv=["']Content-Security-Policy["'])[^>]*>/i,
    '',
    'Content-Security-Policy meta tag'
  );

  html = replaceExactlyOnce(
    html,
    EXTERNAL_ASSET_TAGS[0],
    `\n  <style>\n${styles.trim()}\n  </style>\n`,
    'local stylesheet tag'
  );
  for (const [index, pattern] of EXTERNAL_ASSET_TAGS.slice(1).entries()) {
    html = removeOptionalOnce(html, pattern, `local script tag ${index + 1}`);
  }

  html = html.replaceAll(/favicon\.svg(?:\?[^"'\s)]*)?/g, faviconDataUrl);

  const runtimeConfig = [
    '<script>',
    '/* The bound Apps Script doGet() assigns template.WEB_APP_URL before evaluation. */',
    'window.TRAINING_SIGN_WEB_APP_URL = <?!= JSON.stringify(WEB_APP_URL) ?>;',
    'window.TRAINING_SIGN_CONFIG = Object.freeze({',
    '  API_URL: window.TRAINING_SIGN_WEB_APP_URL,',
    "  APP_NAME: '학교 연수 전자서명'",
    '});',
    '</script>'
  ].join('\n');

  const thirdPartyNotice = [
    '<!--',
    'BEGIN THIRD_PARTY_NOTICES.md',
    notices.trim(),
    'END THIRD_PARTY_NOTICES.md',
    '',
    'BEGIN vendor/LICENSE-SheetJS.txt',
    sheetJsLicense.trim(),
    'END vendor/LICENSE-SheetJS.txt',
    '-->'
  ].join('\n');

  const scripts = [
    runtimeConfig,
    thirdPartyNotice,
    inlineScript('qrcode-generator 2.0.4 — MIT; original copyright/license header follows.', qrcode),
    `<script>\n${applicationScript}\n</script>`
  ].join('\n\n');

  html = replaceExactlyOnce(html, /\s*<\/body>/i, `\n\n${scripts}\n</body>`, 'closing body tag');
  html = `${normalizeText(html).trimEnd()}\n`;

  if (!html.includes('window.TRAINING_SIGN_WEB_APP_URL = <?!= JSON.stringify(WEB_APP_URL) ?>;')) {
    throw new Error('WEB_APP_URL template expression is missing from the generated HTML.');
  }
  if (LOCAL_ASSET_REFERENCE_PATTERN.test(html)) {
    throw new Error('A local asset reference remains in the generated HTML.');
  }
  if (/https?:\/\/(?:cdn|fonts\.)/i.test(html)) {
    throw new Error('An external CDN or font URL remains in the generated HTML.');
  }
  const scriptletCount = (html.match(/<\?/g) || []).length;
  if (scriptletCount !== 1) {
    throw new Error(`Generated HTML must contain only the WEB_APP_URL scriptlet; found ${scriptletCount}.`);
  }

  await mkdir(dirname(OUTPUT_PATH), { recursive: true });
  await writeFile(OUTPUT_PATH, html, 'utf8');
  const sheetJsAsset = `${escapeAppsScriptScriptlets(xlsx.trim())}\n`;
  if ((sheetJsAsset.match(/<\?/g) || []).length !== 0) {
    throw new Error('Generated SheetJS asset contains an Apps Script scriptlet sequence.');
  }
  await writeFile(SHEETJS_OUTPUT_PATH, sheetJsAsset, 'utf8');

  const bytes = Buffer.byteLength(html, 'utf8');
  const sha256 = createHash('sha256').update(html, 'utf8').digest('hex');
  const sheetJsBytes = Buffer.byteLength(sheetJsAsset, 'utf8');
  const sheetJsSha256 = createHash('sha256').update(sheetJsAsset, 'utf8').digest('hex');
  console.log(`Built ${relative(ROOT_DIR, OUTPUT_PATH).replaceAll('\\', '/')}`);
  console.log(`Bytes: ${bytes}`);
  console.log(`SHA-256: ${sha256}`);
  console.log(`Built ${relative(ROOT_DIR, SHEETJS_OUTPUT_PATH).replaceAll('\\', '/')}`);
  console.log(`Bytes: ${sheetJsBytes}`);
  console.log(`SHA-256: ${sheetJsSha256}`);
}

await build();
