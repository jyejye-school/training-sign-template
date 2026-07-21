import test from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';

const root = path.resolve(import.meta.dirname, '..');
const outputPath = path.join(root, 'apps-script', 'WebApp.html');
const sheetJsOutputPath = path.join(root, 'apps-script', 'SheetJS.html');
const readOutput = () => fs.readFileSync(outputPath, 'utf8');
const readSheetJsOutput = () => fs.readFileSync(sheetJsOutputPath, 'utf8');
const digest = value => createHash('sha256').update(value, 'utf8').digest('hex');

test('Apps Script 화면 번들은 외부 정적 호스트 없이 한 파일로 실행된다', () => {
  const html = readOutput();
  assert.match(html, /window\.TRAINING_SIGN_WEB_APP_URL = <\?!= JSON\.stringify\(WEB_APP_URL\) \?>/);
  assert.equal((html.match(/<\?/g) || []).length, 1);
  assert.doesNotMatch(html, /Content-Security-Policy/);
  assert.doesNotMatch(html, /(?:src|href)=["'](?:assets|vendor)\//);
  assert.doesNotMatch(html, /\.\/core\.js|https?:\/\/(?:cdn|fonts\.)/i);
  assert.match(html, /qrcode-generator 2\.0\.4/);
  assert.match(html, /SheetJS Community Edition 0\.20\.3/);
  assert.doesNotMatch(html, /function make_xlsx_lib|window\.XLSX\s*=\s*XLSX/);
  assert.ok(Buffer.byteLength(html, 'utf8') < 500_000, 'Participant HTML should not embed the SheetJS runtime.');

  const sheetJs = readSheetJsOutput();
  assert.match(sheetJs, /xlsx\.js \(C\) 2013-present SheetJS/);
  assert.match(sheetJs, /0\.20\.3/);
  assert.equal((sheetJs.match(/<\?/g) || []).length, 0);
  new vm.Script(sheetJs);

  const executable = html.replace(
    '<?!= JSON.stringify(WEB_APP_URL) ?>',
    JSON.stringify('https://script.google.com/macros/s/AKfycb_bundle-test/exec')
  );
  const scripts = [...executable.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi)].map(match => match[1]);
  assert.equal(scripts.length, 3);
  scripts.forEach(source => new vm.Script(source));
});

test('SheetJS는 관리자 엑셀 기능에서만 웹앱 자체 주소로 지연 로드된다', () => {
  const app = fs.readFileSync(path.join(root, 'assets', 'app.js'), 'utf8');
  const backend = fs.readFileSync(path.join(root, 'apps-script', 'Code.gs'), 'utf8');
  const index = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
  assert.doesNotMatch(index, /<script[^>]+xlsx\.full\.min\.js/);
  assert.match(app, /function loadXlsxLibrary\(\)/);
  assert.match(app, /function xlsxLibraryUrl\(\)[\s\S]*hasAppsScriptLocationBridge\(\)[\s\S]*vendor\/xlsx\.full\.min\.js/);
  assert.match(app, /url\.searchParams\.set\('asset', 'sheetjs'\)/);
  assert.match(app, /await loadXlsxLibrary\(\)/);
  assert.match(backend, /event\.parameter\.asset[\s\S]*HtmlService\.createHtmlOutputFromFile\('SheetJS'\)[\s\S]*MimeType\.JAVASCRIPT/);
});

test('Apps Script 화면 번들은 같은 소스에서 동일하게 재생성된다', () => {
  const before = digest(readOutput());
  const sheetJsBefore = digest(readSheetJsOutput());
  execFileSync(process.execPath, [path.join(root, 'scripts', 'build-apps-script-web.mjs')], {
    cwd: root,
    stdio: 'pipe'
  });
  assert.equal(digest(readOutput()), before);
  assert.equal(digest(readSheetJsOutput()), sheetJsBefore);
});
