import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve(import.meta.dirname, '..');
const read = file => fs.readFileSync(path.join(root, file), 'utf8');
const backend = read('apps-script/Code.gs');
const app = read('assets/app.js');
const index = read('index.html');
const config = read('assets/config.js');

const extractFunction = (source, name, nextName) => source.slice(
  source.indexOf(`function ${name}`),
  source.indexOf(`function ${nextName}`)
);

test('기본 GET은 번들 웹앱을 열고 format=json은 기존 진단 응답을 유지한다', () => {
  const body = backend.slice(backend.indexOf('function doGet'), backend.indexOf('function doPost'));
  assert.match(body, /event\.parameter\.format/);
  assert.match(body, /jsonOutput_\(\{ ok: true/);
  assert.match(body, /HtmlService\.createTemplateFromFile\('WebApp'\)/);
  assert.match(body, /const webAppUrl = currentWebAppUrl_\(\)[\s\S]*template\.WEB_APP_URL = webAppUrl/);
  assert.match(body, /template\.evaluate\(\)/);
});

test('Apps Script 바깥 탭 파비콘은 공개 파일 없이 데이터 URL로 기관 설정을 제공한다', () => {
  const body = backend.slice(backend.indexOf('function doGet'), backend.indexOf('function doPost'));
  assert.match(body, /\.setFaviconUrl\(currentFaviconDataUrl_\(\)\)/);
  assert.match(body, /function currentFaviconDataUrl_\(\)/);
  assert.match(body, /CacheService\.getScriptCache\(\)/);
  assert.match(body, /readSettings_\(\)\.faviconData/);
  assert.match(body, /data:image\/svg\+xml;base64,/);
  assert.match(body, /function defaultFaviconSvg_\(\)[\s\S]*#2563eb[\s\S]*>서명<\/text>/);
  assert.doesNotMatch(body, /MimeType\.XML|DriveApp|setSharing|ANYONE|imageFileId/);
});

test('브라우저는 서버가 주입한 Apps Script exec 주소만 POST 대상으로 사용한다', () => {
  assert.doesNotMatch(app, /\bconfig\./);
  assert.doesNotMatch(config, /__APPS_SCRIPT_WEB_APP_URL__/);
  assert.match(app, /window\.TRAINING_SIGN_WEB_APP_URL \|\| window\.TRAINING_SIGN_CONFIG\?\.API_URL/);
  assert.match(config, /TRAINING_SIGN_WEB_APP_URL[\s\S]*TRAINING_SIGN_CONFIG[\s\S]*API_URL: window\.TRAINING_SIGN_WEB_APP_URL/);
  assert.match(app, /url\.origin === 'https:\/\/script\.google\.com'/);
  assert.match(app, /\^\\\/macros\\\/s\\\//);
  assert.doesNotMatch(app + backend, /frontendUrl:\s*baseUrl|request\.frontendUrl|normalizeFrontendUrl_/);
  assert.doesNotMatch(backend, /getProperty\('FRONTEND_URL'\)|setProperty\('FRONTEND_URL'\)/);
});

test('주입 URL 검증은 정확한 script.google.com exec 주소만 허용한다', () => {
  const source = extractFunction(app, 'trustedWebAppUrl', 'hasAppsScriptLocationBridge');
  const trustedWebAppUrl = new Function(`${source}; return trustedWebAppUrl;`)();
  const valid = 'https://script.google.com/macros/s/AKfycb_safe-123/exec';
  assert.equal(trustedWebAppUrl(valid), valid);
  assert.equal(trustedWebAppUrl('https://evil.example/macros/s/AKfycb_safe-123/exec'), '');
  assert.equal(trustedWebAppUrl('https://script.google.com.evil.example/macros/s/AKfycb_safe-123/exec'), '');
  assert.equal(trustedWebAppUrl('https://script.google.com/macros/s/AKfycb_safe-123/dev'), '');
  assert.equal(trustedWebAppUrl(`${valid}?backend=https://evil.example`), '');
  assert.equal(trustedWebAppUrl(`${valid}#k=token`), '');
});

test('Apps Script IFRAME은 바깥 해시를 읽고 갱신하며 정적 데모는 History API를 쓴다', () => {
  assert.match(app, /window\.google\.script\.url\.getLocation\(finish\)/);
  assert.match(app, /parseShareToken\(info\?\.hash \|\| ''\)/);
  assert.doesNotMatch(app, /setTimeout\(\(\) => finish\(null\)/);
  assert.match(app, /window\.google\.script\.history\.replace\(\{\}, \{\}, hash\)/);
  assert.match(app, /history\.replaceState\(null, '', token \? buildShareUrl\(staticBaseUrl, token\) : staticBaseUrl\)/);
  assert.match(app, /function appShareBaseUrl\(\)[\s\S]*!DEMO && API_URL \? API_URL : staticBaseUrl/);
  assert.match(app, /function replaceAppHash\(token\)[\s\S]*history\.replaceState\(null, '', token \? buildShareUrl\(staticBaseUrl, token\) : staticBaseUrl\)/);
  assert.match(app, /initializeAppLocation\(\)\.then\(initializePublicApp\)/);
});

test('공유 주소는 저장된 외부 주소가 아니라 현재 배포 URL에서 생성된다', () => {
  assert.match(backend, /shareUrl: buildShareUrl_\(currentWebAppUrl_\(\), shareToken\)/);
  assert.match(backend, /function rotateShareToken_\(\)[\s\S]*const url = currentWebAppUrl_\(\)/);
  assert.match(app, /function openShareDialog\(\)[\s\S]*buildShareUrl\(appShareBaseUrl\(\), shareToken\)/);
});
