/**
 * 학교 연수 전자서명 API
 * - Google 시트에 연결된 Apps Script 웹앱으로 배포합니다.
 * - 데이터와 파일은 이 스크립트를 소유한 Google 계정에만 저장합니다.
 */

const APP = Object.freeze({
  VERSION: '1.9.1',
  TIME_ZONE: 'Asia/Seoul',
  DATA_FILE: '학교 연수 전자서명 데이터',
  GUIDE_SHEET: '사용설명서',
  ROOT_FOLDER: '학교 연수 전자서명',
  SIGNATURE_FOLDER: '서명 원본',
  EXPORT_FOLDER: '출력 보관',
  SESSION_SECONDS: 1800,
  MAX_SIGNATURE_BYTES: 400 * 1024,
  MAX_STAFF: 500,
  MAX_EXPORT_ROWS: 200,
  EXPORT_BATCH_SIZE: 30,
  DOWNLOAD_CHUNK_SIZE: 1024 * 1024,
  EXPORT_LEASE_MS: 7 * 60 * 1000,
  ONSITE_CODE_MAX_FAILURES: 5,
  ONSITE_CODE_MAX_SESSION_FAILURES: 50
});

const ONSITE_RECEPTION_PROPERTY_PREFIX_ = 'ONSITE_RECEPTION_';

const SHEETS = Object.freeze({
  SETTINGS: { name: '설정', headers: ['key', 'value'] },
  STAFF: { name: '구성원', headers: ['id', 'department', 'name', 'active', 'sortOrder', 'createdAt'] },
  TRAININGS: { name: '연수', headers: ['id', 'title', 'target', 'date', 'daily', 'startTime', 'endTime', 'active', 'sortOrder', 'createdAt', 'updatedAt', 'audienceMode', 'audienceDepartments', 'onsiteCodeRequired'] },
  SIGNATURES: { name: '서명', headers: ['id', 'trainingId', 'staffId', 'signDate', 'signTime', 'department', 'name', 'imageFileId', 'createdAt', 'scopeDate', 'verificationMethod', 'onsiteSessionId'] },
  EXPORTS: { name: '출력 작업', headers: ['jobId', 'trainingId', 'trainingTitle', 'date', 'sort', 'columns', 'showRate', 'status', 'progress', 'total', 'tempSpreadsheetId', 'pdfFileId', 'xlsxFileId', 'createdAt', 'updatedAt', 'error', 'purgedAt', 'outputType', 'previewFileId', 'printOpenedAt', 'signatureSnapshot'] },
  AUDIT: { name: '감사 기록', headers: ['timestamp', 'action', 'target', 'count', 'detail'] }
});

const SETTING_KEYS = Object.freeze([
  'schoolName', 'subtitle', 'notice', 'brandColor',
  'privacyPurpose', 'privacyItems', 'privacyRetention', 'faviconData'
]);

const INSTANCE_PROPERTIES = Object.freeze([
  'SPREADSHEET_ID', 'INSTANCE_ID', 'ROOT_FOLDER_ID', 'SIGNATURE_FOLDER_ID', 'EXPORT_FOLDER_ID',
  'SHARE_TOKEN', 'SETUP_CODE', 'ADMIN_PEPPER', 'ADMIN_EPOCH', 'ADMIN_SALT', 'ADMIN_HASH', 'FRONTEND_URL',
  'ONSITE_CODE_SECRET', 'FAVICON_REV'
]);

let REQUEST_CONTEXT_ = null;

function resetRequestContext_() {
  REQUEST_CONTEXT_ = { spreadsheet: null, sheets: {}, rows: {}, webAppUrl: '' };
}

function requestContext_() {
  if (!REQUEST_CONTEXT_) resetRequestContext_();
  return REQUEST_CONTEXT_;
}

/** 시트를 열 때 관리용 메뉴를 표시합니다. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🖊️ 전자서명 관리')
    .addItem('초기 설정 실행', 'initializeSystemFromMenu')
    .addItem('초기 설정 코드 보기', 'showSetupCode')
    .addItem('웹앱 주소 확인', 'showWebAppUrl')
    .addSeparator()
    .addItem('데이터 탭 표시·숨기기', 'toggleDataSheets')
    .addItem('관리자 비밀번호 복구', 'resetAdminPasswordFromMenu')
    .addItem('사용설명서 다시 만들기', 'rebuildGuideSheetFromMenu')
    .addToUi();
}

function onInstall() {
  onOpen();
}

function initializeSystemFromMenu() {
  try {
    const result = initializeSystem();
    SpreadsheetApp.getUi().alert(
      '초기 설정 완료',
      '현재 학교용 시트를 안전한 데이터 파일로 초기화했습니다.\n\n초기 설정 코드: ' + result.setupCode +
        '\n\n이 코드는 관리자 첫 비밀번호를 설정한 뒤 자동으로 폐기됩니다.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('초기 설정을 완료하지 못했습니다', String(error && error.message || error), SpreadsheetApp.getUi().ButtonSet.OK);
    throw error;
  }
}

function showSetupCode() {
  const properties = PropertiesService.getScriptProperties();
  const code = properties.getProperty('SETUP_CODE');
  const message = !properties.getProperty('SPREADSHEET_ID')
    ? '먼저 ‘초기 설정 실행’을 선택해 주세요.'
    : code
      ? '초기 설정 코드: ' + code + '\n\n관리자 첫 비밀번호 설정이 끝나면 이 코드는 자동 폐기됩니다.'
      : '관리자 첫 설정이 이미 완료되어 초기 설정 코드가 폐기되었습니다.';
  SpreadsheetApp.getUi().alert('초기 설정 코드', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function showWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  const message = url
    ? '현재 웹앱 주소:\n\n' + url + '\n\n주소가 /exec로 끝나는지 확인하세요.'
    : '아직 웹앱으로 배포되지 않았습니다. Apps Script에서 ‘배포 → 새 배포 → 웹 앱’을 실행해 주세요.';
  SpreadsheetApp.getUi().alert('웹앱 주소', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function toggleDataSheets() {
  const spreadsheet = boundSpreadsheet_();
  ensureGuideSheet_(spreadsheet, false);
  const sheets = dataSheetDefinitions_().map(function(definition) { return spreadsheet.getSheetByName(definition.name); }).filter(Boolean);
  const shouldShow = sheets.some(function(sheet) { return sheet.isSheetHidden(); });
  sheets.forEach(function(sheet) { if (shouldShow) sheet.showSheet(); else sheet.hideSheet(); });
  SpreadsheetApp.getUi().alert('데이터 탭', shouldShow ? '데이터 탭을 표시했습니다.' : '데이터 탭을 숨겼습니다.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function resetAdminPasswordFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('관리자 비밀번호 복구', '기존 관리자 세션을 모두 끝내고 임시 비밀번호를 발급할까요?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;
  const temporaryPassword = resetAdminPasswordFromEditor();
  ui.alert('임시 관리자 비밀번호', temporaryPassword + '\n\n관리자 화면에 로그인한 뒤 즉시 새 비밀번호로 변경해 주세요.', ui.ButtonSet.OK);
}

function rebuildGuideSheetFromMenu() {
  const spreadsheet = boundSpreadsheet_();
  ensureGuideSheet_(spreadsheet, true);
  SpreadsheetApp.getUi().alert('사용설명서', '사용설명서 탭을 다시 만들었습니다. 기존 데이터 탭은 변경하지 않았습니다.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function doGet(event) {
  resetRequestContext_();
  const pathInfo = String(event && event.pathInfo || '').replace(/^\/+|\/+$/g, '');
  if (pathInfo === 'favicon.svg') return faviconOutput_();
  if (event && event.parameter && String(event.parameter.asset || '').toLowerCase() === 'sheetjs') {
    return ContentService
      .createTextOutput(HtmlService.createHtmlOutputFromFile('SheetJS').getContent())
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  if (event && event.parameter && String(event.parameter.format || '').toLowerCase() === 'json') {
    return jsonOutput_({ ok: true, data: { service: '학교 연수 전자서명 API', version: APP.VERSION } });
  }
  const webAppUrl = currentWebAppUrl_();
  const template = HtmlService.createTemplateFromFile('WebApp');
  template.WEB_APP_URL = webAppUrl;
  const faviconRevision = PropertiesService.getScriptProperties().getProperty('FAVICON_REV') || APP.VERSION;
  return template.evaluate()
    .setTitle('학교 연수 전자서명')
    .setFaviconUrl(webAppUrl + '/favicon.svg?v=' + encodeURIComponent(faviconRevision))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** HtmlService 바깥 브라우저 탭에서도 기관 파비콘을 표시합니다. */
function faviconOutput_() {
  return ContentService.createTextOutput(faviconSvg_())
    .setMimeType(ContentService.MimeType.XML);
}

function faviconSvg_() {
  try {
    const properties = PropertiesService.getScriptProperties();
    if (properties.getProperty('SPREADSHEET_ID')) {
      const faviconData = validateFaviconData_(readSettings_().faviconData);
      if (faviconData) {
        return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64">' +
          '<image width="64" height="64" href="' + faviconData + '"/>' +
          '</svg>';
      }
    }
  } catch (error) {
    console.warn('사용자 지정 파비콘을 읽지 못해 기본 아이콘을 사용합니다: ' + String(error && error.message || error));
  }
  return defaultFaviconSvg_();
}

function defaultFaviconSvg_() {
  return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64">' +
    '<rect width="64" height="64" rx="14" fill="#2563eb"/>' +
    '<rect x="6" y="8" width="43" height="14" rx="7" fill="#fff"/>' +
    '<path d="M47 8 60 15 47 22Z" fill="#bfdbfe"/>' +
    '<path d="m57 13.4 3 1.6-3 1.6Z" fill="#172554"/>' +
    '<path d="M43 9v12" fill="none" stroke="#2563eb" stroke-width="3" stroke-linecap="round"/>' +
    '<text x="32" y="56" fill="#fff" font-size="27" font-weight="900" letter-spacing="-1.5" text-anchor="middle" font-family="Malgun Gothic,Apple SD Gothic Neo,sans-serif">서명</text>' +
    '</svg>';
}

function doPost(event) {
  resetRequestContext_();
  try {
    if (!event || !event.postData || !event.postData.contents) apiError_('BAD_REQUEST', '요청 본문이 없습니다.');
    const request = JSON.parse(event.postData.contents);
    const data = dispatch_(request || {});
    return jsonOutput_({ ok: true, data: data === undefined ? null : data });
  } catch (error) {
    console.error(error && error.stack ? error.stack : error);
    return jsonOutput_({
      ok: false,
      error: {
        code: error.apiCode || 'SERVER_ERROR',
        message: error.apiCode ? error.message : '서버에서 요청을 처리하지 못했습니다.',
        details: error.details || null
      }
    });
  }
}

function dispatch_(request) {
  const action = string_(request.action, 80);
  if (!action) apiError_('BAD_REQUEST', '작업 이름이 없습니다.');
  if (action === 'get_setup_status') return getSetupStatus_();
  if (action === 'complete_setup') return completeSetup_(request);
  if (action === 'admin_login') return adminLogin_(request.password, request.view);
  if (action === 'get_public_data') return getPublicData_(request.shareToken);
  if (action === 'submit_signature') return submitSignature_(request);

  const sessionToken = requireAdminSession_(request.sessionToken);
  if (action === 'logout') return logout_(sessionToken);
  if (action === 'get_admin_data') return getAdminData_();
  if (action === 'get_admin_section') return getAdminSection_(request.section);
  if (action === 'get_training_signature_status') return getTrainingSignatureStatus_(request.trainingId, request.date);
  if (action === 'get_training_reception_status') return getTrainingReception_(request.trainingId);
  if (action === 'start_training_reception') return withAdminMutationLock_(function() { return startTrainingReception_(request.trainingId, request.durationMinutes); });
  if (action === 'close_training_reception') return withAdminMutationLock_(function() { return closeTrainingReception_(request.trainingId); });
  if (action === 'save_settings') return withAdminMutationLock_(function() { return saveSettings_(request.settings); });
  if (action === 'save_training') return withAdminMutationLock_(function() { return saveTraining_(request.training); });
  if (action === 'delete_training') return withAdminMutationLock_(function() { return deleteTraining_(request.trainingId); });
  if (action === 'move_training') return withAdminMutationLock_(function() { return moveTraining_(request.trainingId, request.direction); });
  if (action === 'save_staff_batch') return withAdminMutationLock_(function() { return saveStaffBatch_(request.people); });
  if (action === 'update_staff') return withAdminMutationLock_(function() { return updateStaff_(request.person); });
  if (action === 'delete_staff') return withAdminMutationLock_(function() { return deleteStaff_(request.staffId); });
  if (action === 'rename_department') return withAdminMutationLock_(function() { return renameDepartment_(request.oldDepartment, request.newDepartment); });
  if (action === 'list_records') return listRecords_(request.trainingId, request.date);
  if (action === 'delete_record') return withAdminMutationLock_(function() { return deleteRecord_(request.recordId); });
  if (action === 'rotate_share_token') return withAdminMutationLock_(function() { return rotateShareToken_(); });
  if (action === 'change_password') return withAdminMutationLock_(function() { return changePassword_(request.currentPassword, request.newPassword); });
  if (action === 'start_export') return startExport_(request);
  if (action === 'continue_export') return continueExport_(request.jobId);
  if (action === 'finalize_export') return finalizeExport_(request.jobId);
  if (action === 'record_print_opened') return recordPrintOpened_(request.jobId);
  if (action === 'download_export_chunk') return downloadExportChunk_(request.jobId, request.format, request.offset, request.chunkSize);
  if (action === 'purge_originals') return withAdminMutationLock_(function() { return purgeOriginals_(request.jobId, request.confirmation); });
  apiError_('UNKNOWN_ACTION', '지원하지 않는 작업입니다.');
}

/**
 * 학교용 시트의 메뉴 또는 연결형 Apps Script 편집기에서 실행합니다.
 * 웹앱 요청에서는 호출하지 않으며, 반복 실행해도 기존 데이터와 비밀값을 보존합니다.
 */
function initializeSystem() {
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty('TEMPLATE_LOCK') === '1') {
    apiError_('TEMPLATE_LOCKED', '이 파일은 비어 있는 배포용 원본입니다. 학교용 사본을 만든 뒤 사본에서 초기화해 주세요.');
  }

  const spreadsheet = boundSpreadsheet_();
  const spreadsheetId = spreadsheet.getId();
  const storedId = properties.getProperty('SPREADSHEET_ID');
  if (storedId && storedId !== spreadsheetId) {
    clearCopiedInstanceProperties_(properties);
  }

  DriveApp.getFileById(spreadsheetId).setName(APP.DATA_FILE);
  spreadsheet.setSpreadsheetTimeZone(APP.TIME_ZONE);
  ensureGuideSheet_(spreadsheet, false);
  dataSheetDefinitions_().forEach(function(definition) { ensureSheet_(spreadsheet, definition); });

  const rootFolder = getOrRepairFolder_(properties, 'ROOT_FOLDER_ID', null, APP.ROOT_FOLDER);
  const signatureFolder = getOrRepairFolder_(properties, 'SIGNATURE_FOLDER_ID', rootFolder, APP.SIGNATURE_FOLDER);
  const exportFolder = getOrRepairFolder_(properties, 'EXPORT_FOLDER_ID', rootFolder, APP.EXPORT_FOLDER);

  const secrets = {
    SPREADSHEET_ID: spreadsheetId,
    INSTANCE_ID: properties.getProperty('INSTANCE_ID') || randomToken_(24),
    ROOT_FOLDER_ID: rootFolder.getId(),
    SIGNATURE_FOLDER_ID: signatureFolder.getId(),
    EXPORT_FOLDER_ID: exportFolder.getId(),
    SHARE_TOKEN: properties.getProperty('SHARE_TOKEN') || randomToken_(24),
    ADMIN_PEPPER: properties.getProperty('ADMIN_PEPPER') || randomToken_(32),
    ADMIN_EPOCH: properties.getProperty('ADMIN_EPOCH') || '1',
    ONSITE_CODE_SECRET: properties.getProperty('ONSITE_CODE_SECRET') || randomToken_(32)
  };
  if (!properties.getProperty('ADMIN_HASH')) secrets.SETUP_CODE = properties.getProperty('SETUP_CODE') || randomToken_(24);
  properties.setProperties(secrets, false);

  if (!readRows_(SHEETS.SETTINGS).length) {
    writeSettings_({
      schoolName: '학교 연수 전자서명',
      subtitle: '연수 참여 확인',
      notice: '',
      brandColor: '#315c54',
      privacyPurpose: '',
      privacyItems: '',
      privacyRetention: '',
      faviconData: ''
    });
  }
  ensureCleanupTrigger_();
  hideDataSheets_(spreadsheet);
  audit_('initialize', spreadsheetId, 1, storedId ? '시스템 구성 복구' : '학교용 사본 초기화');
  const setupCode = properties.getProperty('SETUP_CODE') || '(이미 관리자 설정 완료)';
  console.log('초기 설정 코드: ' + setupCode);
  console.log('학교용 데이터 시트: ' + spreadsheet.getUrl());
  return { spreadsheetId: spreadsheetId, setupCode: setupCode };
}

function getSetupStatus_() {
  const properties = PropertiesService.getScriptProperties();
  return {
    initialized: Boolean(properties.getProperty('SPREADSHEET_ID')),
    adminConfigured: Boolean(properties.getProperty('ADMIN_HASH'))
  };
}

function completeSetup_(request) {
  requireInitialized_();
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty('ADMIN_HASH')) apiError_('ALREADY_CONFIGURED', '관리자 설정이 이미 완료되었습니다.');
  if (!safeEqual_(string_(request.setupCode, 100), properties.getProperty('SETUP_CODE') || '')) apiError_('BAD_SETUP_CODE', '초기 설정 코드가 올바르지 않습니다.');
  const password = validatePassword_(request.password);
  const salt = randomToken_(18);
  properties.setProperties({
    ADMIN_SALT: salt,
    ADMIN_HASH: passwordHash_(password, salt)
  });
  properties.deleteProperty('FRONTEND_URL');
  properties.deleteProperty('SETUP_CODE');
  audit_('complete_setup', 'admin', 1, '관리자 비밀번호 최초 설정');
  return createAdminLoginResult_(request.view);
}

function adminLogin_(password, view) {
  requireInitialized_();
  const properties = PropertiesService.getScriptProperties();
  if (!properties.getProperty('ADMIN_HASH')) apiError_('SETUP_REQUIRED', '관리자 첫 설정이 필요합니다.');
  const cache = CacheService.getScriptCache();
  const lockedUntil = Number(cache.get('admin-login-locked-until') || 0);
  if (lockedUntil > Date.now()) apiError_('LOGIN_LOCKED', '로그인 시도가 잠시 제한되었습니다. 5분 뒤 다시 시도해 주세요.');
  const valid = verifyPassword_(String(password || ''));
  if (!valid) {
    const failures = Number(cache.get('admin-login-failures') || 0) + 1;
    cache.put('admin-login-failures', String(failures), 300);
    if (failures >= 5) cache.put('admin-login-locked-until', String(Date.now() + 300000), 300);
    apiError_('BAD_PASSWORD', failures >= 5 ? '로그인 시도가 잠시 제한되었습니다. 5분 뒤 다시 시도해 주세요.' : '관리자 비밀번호가 올바르지 않습니다.');
  }
  cache.remove('admin-login-failures');
  cache.remove('admin-login-locked-until');
  audit_('admin_login', 'admin', 1, '관리자 로그인');
  return createAdminLoginResult_(view);
}

function createAdminLoginResult_(view) {
  const token = randomToken_(32);
  const epoch = PropertiesService.getScriptProperties().getProperty('ADMIN_EPOCH') || '1';
  CacheService.getScriptCache().put('admin-session:' + token, epoch, APP.SESSION_SECONDS);
  return {
    sessionToken: token,
    expiresIn: APP.SESSION_SECONDS,
    adminData: view === 'bootstrap' ? getAdminBootstrap_() : getAdminData_()
  };
}

function withAdminMutationLock_(callback) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    return callback();
  } finally {
    try { lock.releaseLock(); } catch (ignore) { /* Lock may not have been acquired. */ }
  }
}

function requireAdminSession_(token) {
  const value = string_(token, 100);
  const epoch = PropertiesService.getScriptProperties().getProperty('ADMIN_EPOCH') || '1';
  if (!value || CacheService.getScriptCache().get('admin-session:' + value) !== epoch) apiError_('SESSION_EXPIRED', '관리자 로그인이 만료되었습니다. 다시 로그인해 주세요.');
  CacheService.getScriptCache().put('admin-session:' + value, epoch, APP.SESSION_SECONDS);
  return value;
}

function logout_(token) {
  CacheService.getScriptCache().remove('admin-session:' + token);
  return { loggedOut: true };
}

function changePassword_(currentPassword, newPassword) {
  if (!verifyPassword_(String(currentPassword || ''))) apiError_('BAD_PASSWORD', '현재 비밀번호가 올바르지 않습니다.');
  const password = validatePassword_(newPassword);
  const salt = randomToken_(18);
  PropertiesService.getScriptProperties().setProperties({ ADMIN_SALT: salt, ADMIN_HASH: passwordHash_(password, salt) });
  audit_('change_password', 'admin', 1, '관리자 비밀번호 변경');
  return { changed: true };
}

/**
 * 관리자 비밀번호를 잊었을 때 관리용 Google 계정으로 편집기를 열어 직접 실행합니다.
 * 모든 기존 관리자 세션을 무효화하고 임시 비밀번호를 실행 로그에 표시합니다.
 */
function resetAdminPasswordFromEditor() {
  requireInitialized_();
  const temporaryPassword = 'R' + randomToken_(18) + '9';
  const salt = randomToken_(18);
  const properties = PropertiesService.getScriptProperties();
  const nextEpoch = number_(properties.getProperty('ADMIN_EPOCH')) + 1;
  properties.setProperties({ ADMIN_SALT: salt, ADMIN_HASH: passwordHash_(temporaryPassword, salt), ADMIN_EPOCH: String(nextEpoch) });
  audit_('reset_admin_password', 'admin', 1, '편집기에서 임시 비밀번호 발급');
  console.log('임시 관리자 비밀번호: ' + temporaryPassword);
  console.log('로그인 후 공유·보안 메뉴에서 즉시 새 비밀번호로 변경하세요.');
  return temporaryPassword;
}

function passwordHash_(password, salt) {
  const pepper = PropertiesService.getScriptProperties().getProperty('ADMIN_PEPPER') || '';
  const bytes = Utilities.computeHmacSha256Signature(salt + ':' + password, pepper, Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

function verifyPassword_(password) {
  const properties = PropertiesService.getScriptProperties();
  const salt = properties.getProperty('ADMIN_SALT') || '';
  const expected = properties.getProperty('ADMIN_HASH') || '';
  return expected && safeEqual_(passwordHash_(password, salt), expected);
}

function validatePassword_(password) {
  const value = String(password || '');
  if (/^\d{4}$/.test(value)) return value;
  if (value.length < 10 || value.length > 100 || !/[A-Za-z가-힣]/.test(value) || !/\d/.test(value)) {
    apiError_('WEAK_PASSWORD', '관리자 비밀번호는 숫자 4자리 또는 문자와 숫자를 포함한 10자 이상 100자 이하로 설정해 주세요.');
  }
  return value;
}

function getPublicData_(shareToken) {
  requireInitialized_();
  requireShareToken_(shareToken);
  const settings = readSettings_();
  const privacyReady = privacyReady_(settings);
  if (!privacyReady) apiError_('PRIVACY_NOT_READY', '관리자가 개인정보 처리 안내를 완료하지 않았습니다.');
  const today = today_();
  const activeStaff = readRows_(SHEETS.STAFF)
    .filter(row => bool_(row.active))
    .sort(staffSort_);
  const trainingRows = readRows_(SHEETS.TRAININGS)
    .filter(function(row) { return isTrainingPublicOnDate_(row, today); })
    .sort(orderSort_);
  const staff = activeStaff
    .filter(function(person) {
      return trainingRows.some(function(training) { return trainingTargetsStaff_(training, person); });
    })
    .map(publicStaff_);
  const trainings = trainingRows.map(publicTraining_);
  return { settings: settings, staff: staff, trainings: trainings, privacyReady: true, serverDate: today };
}

function submitSignature_(request) {
  requireInitialized_();
  requireShareToken_(request.shareToken);
  const trainingId = id_(request.trainingId, '연수');
  const staffId = id_(request.staffId, '구성원');
  const onsiteCode = request.onsiteCode;
  const signatureData = String(request.signatureData || '');
  if (signatureData.length > APP.MAX_SIGNATURE_BYTES * 1.5) apiError_('SIGNATURE_TOO_LARGE', '서명 이미지가 너무 큽니다. 다시 작성해 주세요.');
  const match = signatureData.match(/^data:image\/png;base64,([A-Za-z0-9+/=]+)$/);
  if (!match) apiError_('BAD_SIGNATURE', '서명 이미지 형식이 올바르지 않습니다.');
  const bytes = Utilities.base64Decode(match[1]);
  if (bytes.length < 100 || bytes.length > APP.MAX_SIGNATURE_BYTES) apiError_('BAD_SIGNATURE', '서명 이미지 크기가 올바르지 않습니다.');
  if ((bytes[0] & 255) !== 137 || (bytes[1] & 255) !== 80 || (bytes[2] & 255) !== 78 || (bytes[3] & 255) !== 71) apiError_('BAD_SIGNATURE', 'PNG 서명 이미지만 등록할 수 있습니다.');

  const training = findRow_(SHEETS.TRAININGS, 'id', trainingId);
  const person = findRow_(SHEETS.STAFF, 'id', staffId);
  validateSigningWindow_(training, person, onsiteCode, true);
  const fileCreatedAt = new Date();
  const fileDate = formatDate_(fileCreatedAt, 'yyyy-MM-dd');
  const folder = getOrCreateTrainingFolder_(trainingId, training.title);
  const fileName = safeFileName_(fileDate + '_' + person.department + '_' + person.name) + '.png';
  const file = folder.createFile(Utilities.newBlob(bytes, 'image/png', fileName));

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    invalidateRows_(SHEETS.TRAININGS);
    invalidateRows_(SHEETS.STAFF);
    invalidateRows_(SHEETS.SIGNATURES);
    const freshTraining = findRow_(SHEETS.TRAININGS, 'id', trainingId);
    const freshPerson = findRow_(SHEETS.STAFF, 'id', staffId);
    const verification = validateSigningWindow_(freshTraining, freshPerson, onsiteCode, false);
    const now = new Date();
    const date = formatDate_(now, 'yyyy-MM-dd');
    const time = formatDate_(now, 'HH:mm:ss');
    const scopeDate = trainingScopeDate_(freshTraining, date);
    const duplicate = readRows_(SHEETS.SIGNATURES).some(function(row) {
      return row.trainingId === trainingId &&
        row.staffId === staffId &&
        signatureMatchesTrainingDate_(row, freshTraining, date);
    });
    if (duplicate) apiError_('DUPLICATE', '[' + freshTraining.title + '] ' + freshPerson.name + '님은 이미 서명을 완료했습니다.');
    appendObject_(SHEETS.SIGNATURES, {
      id: Utilities.getUuid(), trainingId: trainingId, staffId: staffId,
      signDate: date, signTime: time, department: freshPerson.department, name: freshPerson.name,
      imageFileId: file.getId(), createdAt: now.toISOString(), scopeDate: scopeDate,
      verificationMethod: verification.verificationMethod,
      onsiteSessionId: verification.onsiteSessionId
    });
    return { registeredAt: now.toISOString(), signDate: date, signTime: time };
  } catch (error) {
    try { file.setTrashed(true); } catch (ignore) { /* Best effort orphan cleanup. */ }
    throw error;
  } finally {
    try { lock.releaseLock(); } catch (ignore) { /* Lock may not have been acquired. */ }
  }
}

function validateSigningWindow_(training, person, onsiteCode, recordOnsiteFailure) {
  if (!training || !bool_(training.active)) apiError_('TRAINING_CLOSED', '현재 서명할 수 없는 연수입니다.');
  if (!person || !bool_(person.active)) apiError_('STAFF_NOT_FOUND', '구성원 명단에서 확인할 수 없습니다.');
  if (!trainingTargetsStaff_(training, person)) apiError_('STAFF_NOT_TARGET', '이 연수의 서명 대상이 아닙니다.');
  const now = new Date();
  const today = formatDate_(now, 'yyyy-MM-dd');
  const trainingDate = sheetDateText_(training.date);
  const startTime = sheetTimeText_(training.startTime, false);
  const endTime = sheetTimeText_(training.endTime, false);
  const daily = bool_(training.daily);
  if (!daily) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(trainingDate)) apiError_('TRAINING_DATE', '연수 날짜가 올바르지 않습니다.');
    if (trainingDate > today) apiError_('TRAINING_DATE', trainingDate + '부터 서명할 수 있습니다.');
  }
  const verification = validateTrainingReception_(training, person, onsiteCode, now, recordOnsiteFailure !== false);
  // 과거 고정 연수의 시각은 연수 당일 일정입니다. 재수합 중에는 활성 상태를 접수 스위치로 사용합니다.
  if (!daily && trainingDate < today) return verification;
  const nowTime = formatDate_(now, 'HH:mm');
  if (startTime && nowTime < startTime) apiError_('TOO_EARLY', '아직 서명 가능 시간이 아닙니다. ' + startTime + '부터 서명할 수 있습니다.');
  if (endTime && nowTime > endTime) apiError_('TOO_LATE', '서명 가능 시간이 종료되었습니다.');
  return verification;
}

function isTrainingPublicOnDate_(training, today) {
  if (!training || !bool_(training.active)) return false;
  if (trainingAudienceMode_(training) === 'invalid') return false;
  if (bool_(training.daily)) return true;
  const trainingDate = sheetDateText_(training.date);
  return /^\d{4}-\d{2}-\d{2}$/.test(trainingDate) && trainingDate <= today;
}

function signatureMatchesTrainingDate_(signature, training, date) {
  if (!signature || !training) return false;
  const signatureScopeDate = sheetDateText_(signature.scopeDate || signature.signDate);
  return signatureScopeDate === trainingScopeDate_(training, date);
}

function trainingScopeDate_(training, date) {
  const scopeDate = bool_(training && training.daily)
    ? sheetDateText_(date)
    : sheetDateText_(training && training.date);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(scopeDate)) apiError_('TRAINING_DATE', '연수 기준 날짜가 올바르지 않습니다.');
  return scopeDate;
}

function getTrainingReception_(trainingId) {
  const id = id_(trainingId, '연수');
  const training = findRow_(SHEETS.TRAININGS, 'id', id);
  if (!training) apiError_('NOT_FOUND', '연수를 찾을 수 없습니다.');
  return publicTrainingReception_(training, readTrainingReceptionState_(id), new Date());
}

function startTrainingReception_(trainingId, durationMinutes) {
  const id = id_(trainingId, '연수');
  const duration = Number(durationMinutes);
  if (!Number.isInteger(duration) || [5, 10, 15].indexOf(duration) < 0) {
    apiError_('RECEPTION_DURATION', '현장 접수 시간은 5분·10분·15분 중에서 선택해 주세요.');
  }
  const training = findRow_(SHEETS.TRAININGS, 'id', id);
  if (!training) apiError_('NOT_FOUND', '연수를 찾을 수 없습니다.');
  if (!bool_(training.onsiteCodeRequired)) apiError_('RECEPTION_NOT_ENABLED', '이 연수는 현장 접수 코드를 사용하지 않습니다.');
  if (!bool_(training.active) || trainingAudienceMode_(training) === 'invalid') {
    apiError_('TRAINING_CLOSED', '현재 서명할 수 없는 연수입니다.');
  }

  const now = new Date();
  const serverDate = formatDate_(now, 'yyyy-MM-dd');
  const trainingDate = sheetDateText_(training.date);
  if (!bool_(training.daily)) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(trainingDate)) apiError_('TRAINING_DATE', '연수 날짜가 올바르지 않습니다.');
    if (trainingDate > serverDate) apiError_('TRAINING_DATE', trainingDate + '부터 현장 접수를 시작할 수 있습니다.');
  }
  const schedule = trainingReceptionSchedule_(training, now);
  if (schedule.beforeStart) apiError_('TOO_EARLY', '아직 서명 가능 시간이 아닙니다. ' + schedule.startTime + '부터 현장 접수를 시작할 수 있습니다.');
  if (schedule.afterEnd) apiError_('TOO_LATE', '서명 가능 시간이 종료되어 현장 접수를 시작할 수 없습니다.');

  const code = generateOnsiteCode_();
  const properties = PropertiesService.getScriptProperties();
  const propertyKey = trainingReceptionPropertyKey_(id);
  const previousRawState = properties.getProperty(propertyKey);
  const previousState = readTrainingReceptionState_(id);
  const requestedExpiresAt = now.getTime() + duration * 60 * 1000;
  const expiresAt = schedule.maxExpiresAt
    ? Math.min(requestedExpiresAt, schedule.maxExpiresAt)
    : requestedExpiresAt;
  if (expiresAt <= now.getTime()) apiError_('TOO_LATE', '서명 가능 시간이 종료되어 현장 접수를 시작할 수 없습니다.');
  const state = {
    v: 1,
    sessionId: randomToken_(18),
    trainingId: id,
    serverDate: serverDate,
    scopeDate: trainingScopeDate_(training, serverDate),
    salt: randomToken_(18),
    codeHash: '',
    startedAt: now.toISOString(),
    expiresAt: new Date(expiresAt).toISOString(),
    durationMinutes: duration,
    failureTotal: 0,
    failureCounts: {}
  };
  state.codeHash = onsiteCodeHash_(state, code);
  properties.setProperty(propertyKey, JSON.stringify(state));
  try {
    audit_(
      'start_training_reception',
      id,
      1,
      training.title + ' · ' + duration + '분 · ' + state.scopeDate + ' · 세션 ' + state.sessionId +
        (previousState ? ' · 기존 접수 교체' : '')
    );
  } catch (error) {
    if (previousRawState === null) properties.deleteProperty(propertyKey);
    else properties.setProperty(propertyKey, previousRawState);
    throw error;
  }
  return Object.assign(publicTrainingReception_(training, state, now), { code: code });
}

function closeTrainingReception_(trainingId) {
  const id = id_(trainingId, '연수');
  const training = findRow_(SHEETS.TRAININGS, 'id', id);
  if (!training) apiError_('NOT_FOUND', '연수를 찾을 수 없습니다.');
  const state = removeTrainingReceptionState_(id);
  if (state) {
    auditReceptionBestEffort_('close_training_reception', id, 1, training.title + ' · 세션 ' + state.sessionId + ' · 관리자 종료');
  }
  return publicTrainingReception_(training, null, new Date());
}

function publicTrainingReception_(training, state, now) {
  const required = bool_(training && training.onsiteCodeRequired);
  const current = now instanceof Date ? now : new Date();
  const open = required && isTrainingReceptionOpen_(training, state, current);
  const status = open ? 'open' : required && state ? 'expired' : 'closed';
  const expiresAt = state ? String(state.expiresAt || '') : '';
  const scopeDate = state && /^\d{4}-\d{2}-\d{2}$/.test(String(state.scopeDate || ''))
    ? String(state.scopeDate)
    : bool_(training && training.daily)
      ? formatDate_(current, 'yyyy-MM-dd')
      : sheetDateText_(training && training.date);
  const remainingSeconds = open
    ? Math.max(0, Math.ceil((new Date(expiresAt).getTime() - current.getTime()) / 1000))
    : 0;
  return {
    trainingId: String(training && training.id || ''),
    date: scopeDate,
    required: required,
    status: status,
    sessionId: state ? String(state.sessionId || '') : '',
    startedAt: state ? String(state.startedAt || '') : '',
    expiresAt: expiresAt,
    durationMinutes: state ? number_(state.durationMinutes) : 0,
    remainingSeconds: remainingSeconds
  };
}

function validateTrainingReception_(training, person, onsiteCode, now, recordFailure) {
  if (!bool_(training && training.onsiteCodeRequired)) {
    return { verificationMethod: 'share_link', onsiteSessionId: '' };
  }
  if (recordFailure) {
    return withAdminMutationLock_(function() {
      return validateTrainingReceptionState_(training, person, onsiteCode, new Date(), true);
    });
  }
  return validateTrainingReceptionState_(training, person, onsiteCode, now, false);
}

function validateTrainingReceptionState_(training, person, onsiteCode, now, recordFailure) {
  const state = readTrainingReceptionState_(String(training.id || ''));
  if (!isTrainingReceptionOpen_(training, state, now)) {
    apiError_('TRAINING_RECEPTION_CLOSED', '현장 접수가 시작되지 않았거나 접수 시간이 종료되었습니다.');
  }
  assertOnsiteCodeAttemptsAvailable_(state, person);
  const code = typeof onsiteCode === 'string' ? onsiteCode.trim() : '';
  if (!code) apiError_('ONSITE_CODE_REQUIRED', '현장에서 안내된 6자리 접수 코드를 입력해 주세요.');
  const valid = /^\d{6}$/.test(code) && safeEqual_(onsiteCodeHash_(state, code), state.codeHash);
  if (!valid) {
    const failures = recordFailure ? recordOnsiteCodeFailure_(state, person) : onsiteCodeFailureCount_(state, person);
    if (failures >= APP.ONSITE_CODE_MAX_FAILURES || number_(state.failureTotal) >= APP.ONSITE_CODE_MAX_SESSION_FAILURES) {
      apiError_('ONSITE_CODE_LOCKED', '현장 접수 코드를 여러 번 잘못 입력했습니다. 관리자에게 확인해 주세요.');
    }
    apiError_('ONSITE_CODE_INVALID', '현장 접수 코드가 올바르지 않습니다.');
  }
  if (recordFailure) clearOnsiteCodeFailures_(state, person);
  return { verificationMethod: 'onsite_code', onsiteSessionId: state.sessionId };
}

function isTrainingReceptionOpen_(training, state, now) {
  if (!training || !bool_(training.active) || !bool_(training.onsiteCodeRequired) || trainingAudienceMode_(training) === 'invalid') return false;
  if (!state || String(state.trainingId || '') !== String(training.id || '')) return false;
  const current = now instanceof Date ? now : new Date();
  const startedAt = new Date(String(state.startedAt || '')).getTime();
  const expiresAt = new Date(String(state.expiresAt || '')).getTime();
  if (!startedAt || !expiresAt || startedAt > current.getTime() || expiresAt <= current.getTime()) return false;
  const serverDate = formatDate_(current, 'yyyy-MM-dd');
  if (String(state.serverDate || '') !== serverDate) return false;
  const schedule = trainingReceptionSchedule_(training, current);
  if (schedule.beforeStart || schedule.afterEnd) return false;
  try {
    return String(state.scopeDate || '') === trainingScopeDate_(training, serverDate);
  } catch (error) {
    return false;
  }
}

function trainingReceptionSchedule_(training, now) {
  const current = now instanceof Date ? now : new Date();
  const today = formatDate_(current, 'yyyy-MM-dd');
  const daily = bool_(training && training.daily);
  const trainingDate = sheetDateText_(training && training.date);
  const pastFixed = !daily && /^\d{4}-\d{2}-\d{2}$/.test(trainingDate) && trainingDate < today;
  const startTime = sheetTimeText_(training && training.startTime, false);
  const endTime = sheetTimeText_(training && training.endTime, false);
  if (pastFixed) return { beforeStart: false, afterEnd: false, startTime: startTime, endTime: endTime, maxExpiresAt: 0 };
  const nowTime = formatDate_(current, 'HH:mm');
  const beforeStart = Boolean(startTime && nowTime < startTime);
  const afterEnd = Boolean(endTime && nowTime > endTime);
  let maxExpiresAt = 0;
  if (endTime && !afterEnd) {
    const nowParts = formatDate_(current, 'HH:mm:ss').split(':').map(Number);
    const endParts = endTime.split(':').map(Number);
    const nowOffset = ((nowParts[0] * 60 + nowParts[1]) * 60 + nowParts[2]) * 1000 + current.getMilliseconds();
    const endOffset = ((endParts[0] * 60 + endParts[1] + 1) * 60) * 1000;
    maxExpiresAt = current.getTime() + Math.max(0, endOffset - nowOffset);
  }
  return { beforeStart: beforeStart, afterEnd: afterEnd, startTime: startTime, endTime: endTime, maxExpiresAt: maxExpiresAt };
}

function readTrainingReceptionState_(trainingId) {
  const id = String(trainingId || '');
  if (!id) return null;
  const raw = PropertiesService.getScriptProperties().getProperty(trainingReceptionPropertyKey_(id));
  if (!raw) return null;
  try {
    const state = JSON.parse(raw);
    if (!state || typeof state !== 'object' || Number(state.v) !== 1) return null;
    if (String(state.trainingId || '') !== id || !/^[A-Za-z0-9_-]{8,100}$/.test(String(state.sessionId || ''))) return null;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(state.serverDate || '')) || !/^\d{4}-\d{2}-\d{2}$/.test(String(state.scopeDate || ''))) return null;
    if (!String(state.salt || '') || !String(state.codeHash || '') || [5, 10, 15].indexOf(Number(state.durationMinutes)) < 0) return null;
    if (!new Date(String(state.startedAt || '')).getTime() || !new Date(String(state.expiresAt || '')).getTime()) return null;
    state.failureTotal = Math.max(0, Math.floor(number_(state.failureTotal)));
    const failureCounts = state.failureCounts && typeof state.failureCounts === 'object' && !Array.isArray(state.failureCounts)
      ? state.failureCounts
      : {};
    state.failureCounts = {};
    Object.keys(failureCounts).slice(0, APP.MAX_STAFF).forEach(function(key) {
      if (/^[A-Za-z0-9_-]{8,24}$/.test(key)) state.failureCounts[key] = Math.max(0, Math.floor(number_(failureCounts[key])));
    });
    return state;
  } catch (error) {
    return null;
  }
}

function removeTrainingReceptionState_(trainingId) {
  const id = String(trainingId || '');
  const state = readTrainingReceptionState_(id);
  PropertiesService.getScriptProperties().deleteProperty(trainingReceptionPropertyKey_(id));
  return state;
}

function trainingReceptionPropertyKey_(trainingId) {
  return ONSITE_RECEPTION_PROPERTY_PREFIX_ + String(trainingId || '');
}

function onsiteCodeHash_(state, code) {
  const properties = PropertiesService.getScriptProperties();
  let secret = properties.getProperty('ONSITE_CODE_SECRET');
  if (!secret) {
    secret = randomToken_(32);
    properties.setProperty('ONSITE_CODE_SECRET', secret);
  }
  const message = [
    String(state.v || 1),
    properties.getProperty('INSTANCE_ID') || '',
    String(state.trainingId || ''),
    String(state.sessionId || ''),
    String(state.serverDate || ''),
    String(state.scopeDate || ''),
    String(state.salt || ''),
    String(state.startedAt || ''),
    String(state.expiresAt || ''),
    String(code || '')
  ].join('\n');
  const bytes = Utilities.computeHmacSha256Signature(message, secret, Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

function generateOnsiteCode_() {
  const seed = Utilities.getUuid() + ':' + Utilities.getUuid() + ':' + new Date().getTime() + ':' + Math.random();
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed, Utilities.Charset.UTF_8);
  const value = (((bytes[0] & 255) * 0x1000000) + ((bytes[1] & 255) << 16) + ((bytes[2] & 255) << 8) + (bytes[3] & 255)) >>> 0;
  return String(100000 + (value % 900000));
}

function assertOnsiteCodeAttemptsAvailable_(state, person) {
  const count = onsiteCodeFailureCount_(state, person);
  if (count >= APP.ONSITE_CODE_MAX_FAILURES || number_(state.failureTotal) >= APP.ONSITE_CODE_MAX_SESSION_FAILURES) {
    apiError_('ONSITE_CODE_LOCKED', '현장 접수 코드를 여러 번 잘못 입력했습니다. 관리자에게 확인해 주세요.');
  }
}

function recordOnsiteCodeFailure_(state, person) {
  const key = onsiteCodeFailureKey_(person);
  const count = onsiteCodeFailureCount_(state, person) + 1;
  state.failureCounts = state.failureCounts || {};
  state.failureCounts[key] = count;
  state.failureTotal = number_(state.failureTotal) + 1;
  PropertiesService.getScriptProperties().setProperty(trainingReceptionPropertyKey_(state.trainingId), JSON.stringify(state));
  return count;
}

function clearOnsiteCodeFailures_(state, person) {
  const key = onsiteCodeFailureKey_(person);
  if (!state.failureCounts || !Object.prototype.hasOwnProperty.call(state.failureCounts, key)) return;
  delete state.failureCounts[key];
  PropertiesService.getScriptProperties().setProperty(trainingReceptionPropertyKey_(state.trainingId), JSON.stringify(state));
}

function onsiteCodeFailureCount_(state, person) {
  const counts = state && state.failureCounts && typeof state.failureCounts === 'object' ? state.failureCounts : {};
  return Math.max(0, Math.floor(number_(counts[onsiteCodeFailureKey_(person)])));
}

function onsiteCodeFailureKey_(person) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(person && person.id || ''), Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '').slice(0, 16);
}

function getAdminData_() {
  requireInitialized_();
  const properties = PropertiesService.getScriptProperties();
  const shareToken = properties.getProperty('SHARE_TOKEN') || '';
  const staff = readRows_(SHEETS.STAFF).sort(staffSort_).map(publicStaff_);
  const trainings = readRows_(SHEETS.TRAININGS).sort(orderSort_).map(publicTraining_);
  const signatures = readRows_(SHEETS.SIGNATURES);
  const exports = readRows_(SHEETS.EXPORTS)
    .sort(function(a, b) { return String(b.createdAt).localeCompare(String(a.createdAt)); })
    .map(publicJob_);
  return {
    settings: readSettings_(), staff: staff, departments: listActiveDepartments_(), trainings: trainings, exports: exports,
    shareToken: shareToken, shareUrl: buildShareUrl_(currentWebAppUrl_(), shareToken),
    stats: { staff: staff.length, trainings: trainings.length, signatures: signatures.length }
  };
}

function getAdminBootstrap_() {
  requireInitialized_();
  const properties = PropertiesService.getScriptProperties();
  const shareToken = properties.getProperty('SHARE_TOKEN') || '';
  return {
    settings: readSettings_(),
    trainings: readRows_(SHEETS.TRAININGS).sort(orderSort_).map(publicTraining_),
    staff: [],
    exports: [],
    shareToken: shareToken,
    shareUrl: buildShareUrl_(currentWebAppUrl_(), shareToken),
    loadedSections: ['settings', 'trainings', 'share']
  };
}

function getAdminSection_(section) {
  const name = string_(section, 30);
  if (['staff', 'exports', 'settings', 'share', 'trainings', 'training_workspace'].indexOf(name) < 0) {
    apiError_('VALIDATION', '불러올 관리자 화면이 올바르지 않습니다.');
  }
  if (name === 'training_workspace') {
    return {
      section: name,
      trainings: readRows_(SHEETS.TRAININGS).sort(orderSort_).map(publicTraining_),
      departments: listActiveDepartments_(),
      exports: readRows_(SHEETS.EXPORTS)
        .sort(function(a, b) { return String(b.createdAt).localeCompare(String(a.createdAt)); })
        .map(publicJob_)
    };
  }
  if (name === 'staff') {
    return { section: name, staff: readRows_(SHEETS.STAFF).sort(staffSort_).map(publicStaff_), departments: listActiveDepartments_() };
  }
  if (name === 'exports') {
    return {
      section: name,
      exports: readRows_(SHEETS.EXPORTS)
        .sort(function(a, b) { return String(b.createdAt).localeCompare(String(a.createdAt)); })
        .map(publicJob_)
    };
  }
  if (name === 'settings') return { section: name, settings: readSettings_() };
  if (name === 'trainings') {
    return { section: name, trainings: readRows_(SHEETS.TRAININGS).sort(orderSort_).map(publicTraining_) };
  }
  const properties = PropertiesService.getScriptProperties();
  const shareToken = properties.getProperty('SHARE_TOKEN') || '';
  return { section: name, shareToken: shareToken, shareUrl: buildShareUrl_(currentWebAppUrl_(), shareToken) };
}

function saveSettings_(input) {
  const current = readSettings_();
  const settings = {};
  SETTING_KEYS.forEach(function(key) {
    if (key === 'faviconData') return;
    settings[key] = string_(input && input[key], key === 'notice' ? 1000 : 500);
  });
  settings.faviconData = input && Object.prototype.hasOwnProperty.call(input, 'faviconData')
    ? validateFaviconData_(input.faviconData)
    : String(current.faviconData || '');
  if (!privacyReady_(settings)) apiError_('PRIVACY_REQUIRED', '학교명과 개인정보 처리 안내를 모두 입력해 주세요.');
  if (!/^#[0-9a-f]{6}$/i.test(settings.brandColor)) settings.brandColor = '#315c54';
  writeSettings_(settings);
  if (settings.faviconData !== String(current.faviconData || '')) {
    PropertiesService.getScriptProperties().setProperty('FAVICON_REV', randomToken_(10));
  }
  audit_('save_settings', 'settings', 1, '기관 설정 변경');
  return { settings: settings };
}

function saveTraining_(input) {
  const training = normalizeTraining_(input);
  if (training.active && !privacyReady_(readSettings_())) apiError_('PRIVACY_REQUIRED', '개인정보 처리 안내를 모두 입력해야 연수를 활성화할 수 있습니다.');
  const sheet = sheet_(SHEETS.TRAININGS);
  const rows = readRowsWithNumbers_(SHEETS.TRAININGS);
  const existing = training.id ? rows.find(function(item) { return item.data.id === training.id; }) : null;
  const now = new Date().toISOString();
  let stored;
  if (existing) {
    stored = Object.assign({}, existing.data, training, { updatedAt: now });
    writeObjectRow_(sheet, SHEETS.TRAININGS.headers, existing.rowNumber, stored, SHEETS.TRAININGS);
  } else {
    training.id = Utilities.getUuid();
    training.sortOrder = rows.length ? Math.max.apply(null, rows.map(function(item) { return number_(item.data.sortOrder); })) + 1 : 1;
    training.createdAt = now;
    training.updatedAt = now;
    appendObject_(SHEETS.TRAININGS, training);
    stored = training;
  }
  if (existing) {
    const receptionMustClose = !bool_(stored.active) || !bool_(stored.onsiteCodeRequired) ||
      bool_(existing.data.daily) !== bool_(stored.daily) ||
      sheetDateText_(existing.data.date) !== sheetDateText_(stored.date) ||
      sheetTimeText_(existing.data.startTime, false) !== sheetTimeText_(stored.startTime, false) ||
      sheetTimeText_(existing.data.endTime, false) !== sheetTimeText_(stored.endTime, false);
    if (receptionMustClose) {
      const closedState = removeTrainingReceptionState_(stored.id);
      if (closedState) {
        auditReceptionBestEffort_('close_training_reception', stored.id, 1, stored.title + ' · 세션 ' + closedState.sessionId + ' · 연수 설정 변경으로 자동 종료');
      }
    }
  }
  audit_('save_training', stored.id, 1, stored.title);
  return { training: publicTraining_(stored) };
}

function normalizeTraining_(input) {
  const title = string_(input && input.title, 100);
  const daily = bool_(input && input.daily);
  const date = string_(input && input.date, 10);
  const startTime = string_(input && input.startTime, 5);
  const endTime = string_(input && input.endTime, 5);
  if (!title) apiError_('VALIDATION', '연수명을 입력해 주세요.');
  if (!daily && !/^\d{4}-\d{2}-\d{2}$/.test(date)) apiError_('VALIDATION', '연수 날짜가 올바르지 않습니다.');
  if (startTime && !/^\d{2}:\d{2}$/.test(startTime)) apiError_('VALIDATION', '시작 시각이 올바르지 않습니다.');
  if (endTime && !/^\d{2}:\d{2}$/.test(endTime)) apiError_('VALIDATION', '종료 시각이 올바르지 않습니다.');
  if (startTime && endTime && startTime >= endTime) apiError_('VALIDATION', '종료 시각은 시작 시각보다 늦어야 합니다.');
  const normalized = { id: input && input.id ? id_(input.id, '연수') : '', title: title, target: '', date: daily ? date || today_() : date, daily: daily, startTime: startTime, endTime: endTime, active: bool_(input && input.active) };
  const hasAudienceMode = input && Object.prototype.hasOwnProperty.call(input, 'audienceMode');
  const hasAudienceDepartments = input && Object.prototype.hasOwnProperty.call(input, 'audienceDepartments');
  const hasAudienceInput = hasAudienceMode || hasAudienceDepartments;
  if (hasAudienceInput) {
    if (!hasAudienceMode || !hasAudienceDepartments) {
      apiError_('VALIDATION', '서명 대상 설정 형식이 올바르지 않습니다.');
    }
    const audienceMode = String(input.audienceMode || '');
    if (audienceMode !== 'all' && audienceMode !== 'departments') {
      apiError_('VALIDATION', '서명 대상 방식이 올바르지 않습니다.');
    }
    if (!Array.isArray(input.audienceDepartments)) {
      apiError_('VALIDATION', '서명 대상 부서 형식이 올바르지 않습니다.');
    }
    if (
      input.audienceDepartments.length > 200 ||
      input.audienceDepartments.some(function(value) {
        if (typeof value !== 'string') return true;
        const department = value.trim();
        return !department || department.length > 50;
      })
    ) {
      apiError_('VALIDATION', '서명 대상 부서 이름을 확인해 주세요.');
    }
    const departments = normalizeAudienceDepartments_(input.audienceDepartments);
    if (audienceMode === 'departments' && !departments.length) {
      apiError_('VALIDATION', '서명 대상 부서를 하나 이상 선택해 주세요.');
    }
    normalized.audienceMode = audienceMode;
    normalized.audienceDepartments = audienceMode === 'departments' ? JSON.stringify(departments) : '';
  }
  if (input && Object.prototype.hasOwnProperty.call(input, 'onsiteCodeRequired')) {
    if (typeof input.onsiteCodeRequired !== 'boolean') {
      apiError_('VALIDATION', '현장 접수 코드 사용 여부가 올바르지 않습니다.');
    }
    normalized.onsiteCodeRequired = input.onsiteCodeRequired;
  }
  return normalized;
}

function deleteTraining_(trainingId) {
  const id = id_(trainingId, '연수');
  const training = findRow_(SHEETS.TRAININGS, 'id', id);
  if (!training) apiError_('NOT_FOUND', '연수를 찾을 수 없습니다.');
  const closedState = removeTrainingReceptionState_(id);
  deleteRowById_(SHEETS.TRAININGS, id);
  if (closedState) {
    auditReceptionBestEffort_('close_training_reception', id, 1, training.title + ' · 세션 ' + closedState.sessionId + ' · 연수 삭제로 자동 종료');
  }
  audit_('delete_training', id, 1, '서명 기록은 유지');
  return { deleted: true, deletedId: id };
}

function moveTraining_(trainingId, direction) {
  const id = id_(trainingId, '연수');
  if (direction !== 'up' && direction !== 'down') apiError_('VALIDATION', '이동 방향이 올바르지 않습니다.');
  const rows = readRowsWithNumbers_(SHEETS.TRAININGS).sort(function(a, b) { return orderSort_(a.data, b.data); });
  const index = rows.findIndex(function(item) { return item.data.id === id; });
  const target = direction === 'up' ? index - 1 : index + 1;
  if (index < 0) apiError_('NOT_FOUND', '연수를 찾을 수 없습니다.');
  if (target < 0 || target >= rows.length) {
    return { moved: false, trainings: rows.map(function(item) { return publicTraining_(item.data); }) };
  }
  const firstOrder = number_(rows[index].data.sortOrder) || index + 1;
  const secondOrder = number_(rows[target].data.sortOrder) || target + 1;
  const sheet = sheet_(SHEETS.TRAININGS);
  const sortColumn = SHEETS.TRAININGS.headers.indexOf('sortOrder') + 1;
  sheet.getRange(rows[index].rowNumber, sortColumn).setValue(secondOrder);
  sheet.getRange(rows[target].rowNumber, sortColumn).setValue(firstOrder);
  invalidateRows_(SHEETS.TRAININGS);
  return {
    moved: true,
    trainings: readRows_(SHEETS.TRAININGS).sort(orderSort_).map(publicTraining_)
  };
}

function saveStaffBatch_(people) {
  if (!Array.isArray(people) || !people.length) apiError_('VALIDATION', '등록할 구성원이 없습니다.');
  if (people.length > APP.MAX_STAFF) apiError_('VALIDATION', '한 번에 등록할 수 있는 인원은 ' + APP.MAX_STAFF + '명입니다.');
  const existing = readRows_(SHEETS.STAFF);
  if (existing.length >= APP.MAX_STAFF) apiError_('LIMIT', '구성원은 최대 ' + APP.MAX_STAFF + '명까지 등록할 수 있습니다.');
  const seen = {};
  existing.forEach(function(person) { seen[staffKey_(person.department, person.name)] = true; });
  let skipped = 0;
  let nextOrder = existing.length ? Math.max.apply(null, existing.map(function(person) { return number_(person.sortOrder); })) + 1 : 1;
  const created = [];
  const now = new Date().toISOString();
  people.forEach(function(person) {
    const department = string_(person && person.department, 50);
    const name = string_(person && person.name, 50);
    const key = staffKey_(department, name);
    if (!department || !name || seen[key] || existing.length + created.length >= APP.MAX_STAFF) { skipped += 1; return; }
    created.push({ id: Utilities.getUuid(), department: department, name: name, active: true, sortOrder: nextOrder++, createdAt: now });
    seen[key] = true;
  });
  if (created.length) {
    const sheet = sheet_(SHEETS.STAFF);
    sheet.getRange(sheet.getLastRow() + 1, 1, created.length, SHEETS.STAFF.headers.length)
      .setValues(created.map(function(person) { return objectValues_(SHEETS.STAFF.headers, person); }));
    invalidateRows_(SHEETS.STAFF);
  }
  audit_('save_staff_batch', 'staff', created.length, '건너뜀 ' + skipped);
  return { added: created.length, skipped: skipped, people: created.map(publicStaff_) };
}

function updateStaff_(input) {
  const id = id_(input && input.id, '구성원');
  const department = string_(input && input.department, 50);
  const name = string_(input && input.name, 50);
  if (!department || !name) apiError_('VALIDATION', '부서와 성명을 입력해 주세요.');
  const rows = readRowsWithNumbers_(SHEETS.STAFF);
  const current = rows.find(function(item) { return item.data.id === id; });
  if (!current) apiError_('NOT_FOUND', '구성원을 찾을 수 없습니다.');
  const duplicate = rows.some(function(item) { return item.data.id !== id && staffKey_(item.data.department, item.data.name) === staffKey_(department, name); });
  if (duplicate) apiError_('DUPLICATE_STAFF', '같은 부서와 성명의 구성원이 이미 있습니다.');
  const stored = Object.assign({}, current.data, { department: department, name: name });
  writeObjectRow_(sheet_(SHEETS.STAFF), SHEETS.STAFF.headers, current.rowNumber, stored, SHEETS.STAFF);
  audit_('update_staff', id, 1, department + ' ' + name);
  return { updated: true, person: publicStaff_(stored) };
}

function deleteStaff_(staffId) {
  const id = id_(staffId, '구성원');
  deleteRowById_(SHEETS.STAFF, id);
  audit_('delete_staff', id, 1, '기존 서명 기록 유지');
  return { deleted: true, deletedId: id };
}

function renameDepartment_(oldDepartment, newDepartment) {
  const oldName = string_(oldDepartment, 50);
  const newName = string_(newDepartment, 50);
  if (!oldName || !newName) apiError_('VALIDATION', '기존 부서와 새 부서명을 입력해 주세요.');
  const sheet = sheet_(SHEETS.STAFF);
  const rows = readRowsWithNumbers_(SHEETS.STAFF);
  let updated = 0;
  rows.forEach(function(item) {
    if (item.data.department === oldName) {
      item.data.department = newName;
      updated += 1;
    }
  });
  if (!updated) apiError_('NOT_FOUND', '변경할 부서를 찾지 못했습니다.');
  sheet.getRange(2, 1, rows.length, SHEETS.STAFF.headers.length)
    .setValues(rows.map(function(item) { return objectValues_(SHEETS.STAFF.headers, item.data); }));
  invalidateRows_(SHEETS.STAFF);
  const trainingRows = readRowsWithNumbers_(SHEETS.TRAININGS);
  const changedTrainings = [];
  trainingRows.forEach(function(item) {
    if (trainingAudienceMode_(item.data) !== 'departments') return;
    const before = trainingAudienceDepartments_(item.data);
    if (before.indexOf(oldName) < 0) return;
    const after = normalizeAudienceDepartments_(before.map(function(department) {
      return department === oldName ? newName : department;
    }));
    item.data.audienceDepartments = JSON.stringify(after);
    changedTrainings.push(publicTraining_(item.data));
  });
  if (changedTrainings.length) {
    sheet_(SHEETS.TRAININGS).getRange(2, 1, trainingRows.length, SHEETS.TRAININGS.headers.length)
      .setValues(trainingRows.map(function(item) { return objectValues_(SHEETS.TRAININGS.headers, item.data); }));
    invalidateRows_(SHEETS.TRAININGS);
  }
  audit_('rename_department', oldName, updated, newName + ' · 연수 대상 ' + changedTrainings.length + '건');
  return {
    updated: updated,
    oldDepartment: oldName,
    newDepartment: newName,
    people: rows.filter(function(item) { return item.data.department === newName; }).map(function(item) { return publicStaff_(item.data); }),
    trainings: changedTrainings
  };
}

function listRecords_(trainingId, date) {
  const id = id_(trainingId, '연수');
  const signDate = validDate_(date);
  const training = findRow_(SHEETS.TRAININGS, 'id', id);
  if (!training) apiError_('NOT_FOUND', '연수를 찾을 수 없습니다.');
  assertTrainingRequestDate_(training, signDate, '기록');
  const records = readRows_(SHEETS.SIGNATURES)
    .filter(function(row) {
      return row.trainingId === id && signatureMatchesTrainingDate_(row, training, signDate);
    })
    .sort(function(a, b) { return String(a.createdAt).localeCompare(String(b.createdAt)); })
    .map(function(row) {
      return {
        id: row.id,
        trainingId: row.trainingId,
        signDate: sheetDateText_(row.signDate),
        signTime: sheetTimeText_(row.signTime, true),
        department: row.department,
        name: row.name,
        verificationMethod: String(row.verificationMethod || 'legacy'),
        onsiteSessionId: String(row.onsiteSessionId || '')
      };
    });
  return { records: records };
}

function getTrainingSignatureStatus_(trainingId, date) {
  const id = id_(trainingId, '연수');
  const signDate = validDate_(date);
  const training = findRow_(SHEETS.TRAININGS, 'id', id);
  if (!training) apiError_('NOT_FOUND', '연수를 찾을 수 없습니다.');
  assertTrainingRequestDate_(training, signDate, '현황');

  const activeStaff = readRows_(SHEETS.STAFF).filter(function(person) {
    return bool_(person.active) && trainingTargetsStaff_(training, person);
  });
  const signatures = readRows_(SHEETS.SIGNATURES)
    .filter(function(record) {
      return String(record.trainingId) === id && signatureMatchesTrainingDate_(record, training, signDate);
    })
    .sort(function(a, b) { return String(a.createdAt).localeCompare(String(b.createdAt)); });
  return buildTrainingSignatureStatus_(id, signDate, activeStaff, signatures);
}

function buildTrainingSignatureStatus_(trainingId, signDate, activeStaff, signatures) {
  const sortedStaff = activeStaff.slice().sort(staffSort_);
  const activeIds = new Set(sortedStaff.map(function(person) { return String(person.id); }));
  const signedByStaff = new Map();
  let outsideRosterSignedCount = 0;
  signatures.forEach(function(record) {
    const staffId = String(record.staffId || '');
    if (!activeIds.has(staffId)) {
      outsideRosterSignedCount += 1;
      return;
    }
    if (!signedByStaff.has(staffId)) signedByStaff.set(staffId, record);
  });

  let signedCount = 0;
  const people = sortedStaff.map(function(person) {
    const signature = signedByStaff.get(String(person.id)) || null;
    if (signature) signedCount += 1;
    return {
      staffId: String(person.id),
      department: String(person.department || ''),
      name: String(person.name || ''),
      sortOrder: number_(person.sortOrder),
      status: signature ? 'signed' : 'unsigned',
      signDate: signature ? sheetDateText_(signature.signDate) : '',
      signTime: signature ? sheetTimeText_(signature.signTime, true).slice(0, 5) : ''
    };
  });
  const targetCount = people.length;
  const unsignedCount = targetCount - signedCount;
  const rate = targetCount ? Math.round(signedCount / targetCount * 1000) / 10 : 0;
  return {
    trainingId: trainingId,
    date: signDate,
    summary: {
      targetCount: targetCount,
      signedCount: signedCount,
      unsignedCount: unsignedCount,
      rate: rate,
      outsideRosterSignedCount: outsideRosterSignedCount
    },
    people: people
  };
}

function assertTrainingRequestDate_(training, date, purpose) {
  const trainingDate = sheetDateText_(training && training.date);
  if (!bool_(training && training.daily) && trainingDate !== date) {
    apiError_('TRAINING_DATE', '해당 연수는 ' + trainingDate + '의 ' + (purpose || '자료') + '만 확인할 수 있습니다.');
  }
}

function deleteRecord_(recordId) {
  const id = id_(recordId, '서명 기록');
  const rows = readRowsWithNumbers_(SHEETS.SIGNATURES);
  const record = rows.find(function(item) { return item.data.id === id; });
  if (!record) apiError_('NOT_FOUND', '서명 기록을 찾을 수 없습니다.');
  trashFileIfExists_(record.data.imageFileId);
  sheet_(SHEETS.SIGNATURES).deleteRow(record.rowNumber);
  invalidateRows_(SHEETS.SIGNATURES);
  audit_('delete_record', id, 1, record.data.department + ' ' + record.data.name);
  return { deleted: true, deletedId: id };
}

function rotateShareToken_() {
  const properties = PropertiesService.getScriptProperties();
  const token = randomToken_(24);
  const url = currentWebAppUrl_();
  properties.setProperty('SHARE_TOKEN', token);
  properties.deleteProperty('FRONTEND_URL');
  audit_('rotate_share_token', 'share', 1, '기존 공유 링크 무효화');
  return { shareToken: token, shareUrl: buildShareUrl_(url, token) };
}

function requireShareToken_(token) {
  const expected = PropertiesService.getScriptProperties().getProperty('SHARE_TOKEN') || '';
  if (!expected || !safeEqual_(string_(token, 100), expected)) apiError_('INVALID_LINK', '공유 링크가 올바르지 않거나 교체되었습니다. 담당자에게 새 링크를 받아 주세요.');
}

function startExport_(request) {
  const trainingId = id_(request.trainingId, '연수');
  const date = validDate_(request.date);
  const columns = Math.max(1, Math.min(3, number_(request.columns) || 2));
  const sort = ['registration', 'department', 'name'].indexOf(request.sort) >= 0 ? request.sort : 'registration';
  const outputType = ['pdf', 'xlsx', 'print'].indexOf(request.outputType) >= 0 ? request.outputType : 'pdf';
  const showRate = bool_(request.showRate);
  const training = findRow_(SHEETS.TRAININGS, 'id', trainingId);
  if (!training) apiError_('NOT_FOUND', '연수를 찾을 수 없습니다.');
  const trainingDate = sheetDateText_(training.date);
  if (!bool_(training.daily) && trainingDate !== date) {
    apiError_('TRAINING_DATE', '해당 연수는 ' + trainingDate + '만 출력할 수 있습니다.');
  }
  const scopedSignatures = readScopedSignatures_(trainingId, training, date);
  const roster = buildExportRoster_(trainingId, date, sort, training, scopedSignatures);
  if (roster.length > APP.MAX_EXPORT_ROWS) apiError_('EXPORT_LIMIT', '한 번에 출력할 수 있는 인원은 ' + APP.MAX_EXPORT_ROWS + '명입니다.');

  const exportFolder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('EXPORT_FOLDER_ID'));
  const temporary = SpreadsheetApp.create('임시_' + safeFileName_(training.title) + '_' + date + '_' + Date.now());
  DriveApp.getFileById(temporary.getId()).moveTo(exportFolder);
  const output = temporary.getSheets()[0];
  output.setName('서명등록부');
  prepareExportSheet_(output, training, date, roster, columns, showRate, sort, readSettings_());
  const imageRows = [];
  roster.forEach(function(row, index) {
    if (row.fileId) imageRows.push([index, String(row.fileId)]);
  });
  const imageSheet = temporary.insertSheet('_IMAGES');
  imageSheet.getRange(1, 1, 1, 2).setValues([['layoutIndex', 'fileId']]);
  if (imageRows.length) imageSheet.getRange(2, 1, imageRows.length, 2).setValues(imageRows);
  imageSheet.hideSheet();

  const totalImages = imageRows.length;
  const now = new Date().toISOString();
  const job = {
    jobId: Utilities.getUuid(), trainingId: trainingId, trainingTitle: training.title, date: date,
    sort: sort, columns: columns, showRate: showRate, status: 'queued', progress: 0, total: totalImages,
    tempSpreadsheetId: temporary.getId(), pdfFileId: '', xlsxFileId: '', createdAt: now, updatedAt: now, error: '', purgedAt: '',
    outputType: outputType, previewFileId: '', printOpenedAt: '',
    signatureSnapshot: JSON.stringify(roster.map(function(row) { return String(row.signatureId || ''); }).filter(Boolean))
  };
  appendObject_(SHEETS.EXPORTS, job);
  audit_('start_export', job.jobId, roster.length, training.title + ' ' + date);
  return publicJob_(job);
}

function readScopedSignatures_(trainingId, training, date) {
  return readRows_(SHEETS.SIGNATURES)
    .filter(function(row) {
      return row.trainingId === trainingId && signatureMatchesTrainingDate_(row, training, date);
    })
    .sort(function(a, b) { return String(a.createdAt).localeCompare(String(b.createdAt)); });
}

function buildExportRoster_(trainingId, date, sort, training, scopedSignatures) {
  training = training || findRow_(SHEETS.TRAININGS, 'id', trainingId);
  if (!training) apiError_('NOT_FOUND', '연수를 찾을 수 없습니다.');
  const signatures = scopedSignatures || readScopedSignatures_(trainingId, training, date);
  const signedByStaff = new Map();
  signatures.forEach(function(row) {
    const staffId = String(row.staffId || '');
    if (!signedByStaff.has(staffId)) signedByStaff.set(staffId, row);
  });
  const allStaff = readRows_(SHEETS.STAFF);
  const staffById = new Map();
  allStaff.forEach(function(person) { staffById.set(String(person.id || ''), person); });
  const includedStaff = new Set();
  const roster = allStaff.filter(function(person) {
    return bool_(person.active) && trainingTargetsStaff_(training, person);
  }).map(function(person) {
    const staffId = String(person.id || '');
    const signature = signedByStaff.get(staffId);
    includedStaff.add(staffId);
    return {
      staffId: person.id, department: person.department, name: person.name,
      time: signature ? exportSignatureTime_(signature, training, date) : '', fileId: signature ? signature.imageFileId : '',
      signatureId: signature ? signature.id : '', sortOrder: number_(person.sortOrder), createdAt: person.createdAt || ''
    };
  });
  signatures.forEach(function(signature, index) {
    const staffId = String(signature.staffId || '');
    if (includedStaff.has(staffId)) return;
    const currentPerson = staffById.get(staffId);
    if (currentPerson && bool_(currentPerson.active)) return;
    roster.push({
      staffId: signature.staffId, department: signature.department, name: signature.name,
      time: exportSignatureTime_(signature, training, date), fileId: signature.imageFileId,
      signatureId: signature.id, sortOrder: 1000000 + index, createdAt: signature.createdAt || ''
    });
  });
  roster.sort(function(a, b) {
    if (sort === 'name') return compareKo_(a.name, b.name) || compareKo_(a.department, b.department);
    if (sort === 'department') return compareKo_(a.department, b.department) || compareKo_(a.name, b.name) || number_(a.sortOrder) - number_(b.sortOrder);
    return number_(a.sortOrder) - number_(b.sortOrder) || String(a.createdAt).localeCompare(String(b.createdAt));
  });
  return roster;
}

function exportSignatureTime_(signature, training, outputDate) {
  const time = sheetTimeText_(signature && signature.signTime, true).slice(0, 5);
  if (!time) return '';
  const signDate = sheetDateText_(signature && signature.signDate);
  if (!bool_(training && training.daily) && signDate && signDate !== sheetDateText_(outputDate)) {
    return signDate.slice(5).replace('-', '.') + ' ' + time;
  }
  return time;
}

function prepareExportSheet_(sheet, training, date, roster, columns, showRate, sort, settings) {
  const totalColumns = columns * 4;
  sheet.clear();
  sheet.setHiddenGridlines(true);
  const schoolName = String(settings && settings.schoolName || '학교 연수 전자서명');
  const firstHalf = Math.max(1, Math.floor(totalColumns / 2));
  const signedCount = roster.filter(function(row) { return Boolean(row.fileId); }).length;
  const rate = roster.length ? Math.round(signedCount / roster.length * 1000) / 10 : 0;
  const rowsPerBlock = Math.max(1, Math.ceil(roster.length / columns));
  let footerRow = 5 + rowsPerBlock;
  const summaryRow = showRate ? footerRow : 0;
  if (showRate) footerRow += 1;
  const totalRows = footerRow;
  const values = [];
  for (let rowIndex = 0; rowIndex < totalRows; rowIndex += 1) {
    values.push(new Array(totalColumns).fill(''));
  }
  values[0][0] = schoolName;
  values[0][firstHalf] = '연수일: ' + formatKoreanDate_(date);
  values[1][0] = training.title + ' 서명등록부';
  values[2][0] = '교직원 연수 참여 확인 기록';
  for (let block = 0; block < columns; block += 1) {
    const valueBase = block * 4;
    values[3][valueBase] = '번호';
    values[3][valueBase + 1] = '부서';
    values[3][valueBase + 2] = '성명';
    values[3][valueBase + 3] = '서명';
  }
  roster.forEach(function(row, index) {
    const position = exportPosition_(index, rowsPerBlock);
    const valueBase = position.block * 4;
    const valueRow = position.row - 1;
    values[valueRow][valueBase] = index + 1;
    values[valueRow][valueBase + 1] = row.department;
    values[valueRow][valueBase + 2] = row.name + (row.time ? '\n' + String(row.time) : '');
    values[valueRow][valueBase + 3] = row.fileId ? '' : '미서명';
  });
  if (showRate) {
    values[summaryRow - 1][0] = '대상 ' + roster.length + '명 · 서명 ' + signedCount + '명 · 미서명 ' + (roster.length - signedCount) + '명 · 서명률 ' + rate + '%';
  }
  values[footerRow - 1][0] = '연수 참여 확인용 자동 생성 문서 · 생성 시각 ' + formatDate_(new Date(), 'yyyy-MM-dd HH:mm');
  sheet.getRange(1, 1, totalRows, totalColumns).setValues(values);

  sheet.getRange(1, 1, 1, firstHalf).merge()
    .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('left');
  sheet.getRange(1, firstHalf + 1, 1, totalColumns - firstHalf).merge()
    .setFontSize(10).setHorizontalAlignment('right');
  sheet.getRange(1, 1, 1, totalColumns)
    .setBorder(false, false, true, false, false, false, '#315c54', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(2, 1, 1, totalColumns).merge()
    .setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(3, 1, 1, totalColumns).merge()
    .setHorizontalAlignment('center').setFontSize(9).setFontColor('#60706b');
  sheet.getRange(4, 1, 1, totalColumns)
    .setFontWeight('bold').setBackground('#dfece8').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true);

  const body = sheet.getRange(5, 1, rowsPerBlock, totalColumns);
  body.setBorder(true, true, true, true, true, true).setVerticalAlignment('middle');
  for (let block = 0; block < columns; block += 1) {
    const base = 1 + block * 4;
    sheet.setColumnWidth(base, 38);
    sheet.setColumnWidth(base + 1, columns === 3 ? 72 : 90);
    sheet.setColumnWidth(base + 2, columns === 3 ? 78 : 92);
    sheet.setColumnWidth(base + 3, columns === 1 ? 230 : columns === 3 ? 112 : 150);
    sheet.getRange(5, base, rowsPerBlock, 1).setHorizontalAlignment('center');
    sheet.getRange(5, base + 1, rowsPerBlock, 1)
      .setWrap(true).setHorizontalAlignment('center').setFontSize(columns === 3 ? 7 : 8);
    sheet.getRange(5, base + 2, rowsPerBlock, 1)
      .setWrap(true).setHorizontalAlignment('center');
    sheet.getRange(5, base + 3, rowsPerBlock, 1)
      .setHorizontalAlignment('center').setFontColor('#b4473d').setFontSize(8);

    const richTextValues = [];
    for (let offset = 0; offset < rowsPerBlock; offset += 1) {
      const rosterIndex = block * rowsPerBlock + offset;
      const row = rosterIndex < roster.length ? roster[rosterIndex] : null;
      const nameText = row ? row.name + (row.time ? '\n' + String(row.time) : '') : '';
      const richText = SpreadsheetApp.newRichTextValue().setText(nameText);
      if (row && row.name) {
        richText.setTextStyle(0, row.name.length, SpreadsheetApp.newTextStyle()
          .setBold(true).setFontSize(columns === 3 ? 8 : 9).build());
        if (row.time) {
          richText.setTextStyle(row.name.length + 1, nameText.length, SpreadsheetApp.newTextStyle()
            .setFontSize(7).setForegroundColor('#66736f').build());
        }
      }
      richTextValues.push([richText.build()]);
    }
    sheet.getRange(5, base + 2, rowsPerBlock, 1).setRichTextValues(richTextValues);
  }
  sheet.setRowHeights(5, rowsPerBlock, columns === 3 ? 52 : 58);
  if (sort === 'department') mergeExportDepartments_(sheet, roster, columns, rowsPerBlock);
  if (showRate) {
    sheet.getRange(summaryRow, 1, 1, totalColumns).merge()
      .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center').setBackground('#f1f6f4')
      .setBorder(true, true, true, true, false, false, '#9fb8b1', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(summaryRow, 32);
  }
  sheet.getRange(footerRow, 1, 1, totalColumns).merge()
    .setFontSize(7).setFontColor('#7a8783').setHorizontalAlignment('center');
  sheet.setRowHeight(1, 28);
  sheet.setRowHeight(2, 34);
  sheet.setRowHeight(3, 22);
  sheet.setRowHeight(4, 26);
  sheet.setFrozenRows(4);
}

function mergeExportDepartments_(sheet, roster, columns, rowsPerBlock) {
  for (let block = 0; block < columns; block += 1) {
    const startIndex = block * rowsPerBlock;
    const endIndex = Math.min(roster.length, startIndex + rowsPerBlock);
    let cursor = startIndex;
    while (cursor < endIndex) {
      let next = cursor + 1;
      while (next < endIndex && roster[next].department === roster[cursor].department) next += 1;
      if (next - cursor > 1) {
        const row = 5 + (cursor - startIndex);
        const column = 2 + block * 4;
        sheet.getRange(row, column, next - cursor, 1).merge().setVerticalAlignment('middle').setHorizontalAlignment('center');
      }
      cursor = next;
    }
  }
}

function continueExport_(jobId) {
  const id = id_(jobId, '출력 작업');
  const initialJob = findRow_(SHEETS.EXPORTS, 'jobId', id);
  if (!initialJob) apiError_('NOT_FOUND', '출력 작업을 찾을 수 없습니다.');
  if (exportJobIsTerminal_(initialJob)) return publicJob_(initialJob);
  const leaseToken = acquireExportLease_(id);
  if (!leaseToken) {
    invalidateRows_(SHEETS.EXPORTS);
    const current = findRow_(SHEETS.EXPORTS, 'jobId', id);
    if (!current) apiError_('NOT_FOUND', '출력 작업을 찾을 수 없습니다.');
    if (exportJobIsTerminal_(current)) return publicJob_(current);
    const busyJob = publicJob_(current);
    busyJob.busy = true;
    return busyJob;
  }
  let entry = null;
  let job = null;
  try {
    invalidateRows_(SHEETS.EXPORTS);
    entry = findRowWithNumber_(SHEETS.EXPORTS, 'jobId', id);
    if (!entry) apiError_('NOT_FOUND', '출력 작업을 찾을 수 없습니다.');
    job = entry.data;
    if (exportJobIsTerminal_(job)) return publicJob_(job);
    const spreadsheet = SpreadsheetApp.openById(job.tempSpreadsheetId);
    const output = spreadsheet.getSheetByName('서명등록부');
    if (!output) throw new Error('임시 출력 문서를 찾을 수 없습니다.');
    const start = number_(job.progress);
    const imageData = readExportImageBatch_(spreadsheet, start);
    const fetchedImages = fetchPrivateDriveImages_(imageData.batch);
    const rowsPerBlock = exportRowsPerBlock_(output, job);
    insertExportImageBatch_(output, fetchedImages, job, rowsPerBlock);
    const nextProgress = start + imageData.processedCount;
    const nextJob = Object.assign({}, job, {
      progress: nextProgress,
      total: imageData.total,
      status: 'processing',
      error: ''
    });
    if (nextProgress >= imageData.total) {
      job = nextJob;
      job = createExportPreview_(entry.rowNumber, nextJob, spreadsheet);
    } else {
      job = writeExportJobChanges_(entry.rowNumber, job, {
        progress: nextProgress,
        total: imageData.total,
        status: 'processing',
        error: ''
      });
    }
    return publicJob_(job);
  } catch (error) {
    if (!entry || !job) throw error;
    invalidateRows_(SHEETS.EXPORTS);
    const latest = findRow_(SHEETS.EXPORTS, 'jobId', id);
    if (latest && exportJobIsTerminal_(latest)) return publicJob_(latest);
    job = writeExportJobChanges_(entry.rowNumber, job, {
      status: 'failed',
      error: String(error && error.message || error).slice(0, 500)
    });
    return publicJob_(job);
  } finally {
    releaseExportLease_(id, leaseToken);
  }
}

function insertExportImageBatch_(output, fetchedImages, job, rowsPerBlock) {
  if (fetchedImages.some(function(item) { return !item.blob; })) {
    throw new Error('일부 서명 이미지를 불러오지 못했습니다. 원본 삭제를 막기 위해 출력을 중단했습니다. 다시 시도해 주세요.');
  }
  const inserted = [];
  try {
    fetchedImages.forEach(function(item) {
      const layoutIndex = number_(item.layoutIndex);
      const position = exportPosition_(layoutIndex, rowsPerBlock);
      const column = 4 + position.block * 4;
      const image = output.insertImage(item.blob, column, position.row);
      inserted.push(image);
      const imageWidth = number_(job.columns) === 1 ? 215 : number_(job.columns) === 3 ? 104 : 140;
      image.setWidth(imageWidth).setHeight(number_(job.columns) === 3 ? 44 : 50);
    });
  } catch (error) {
    inserted.forEach(function(image) {
      try { image.remove(); } catch (ignore) { /* Best effort rollback before retry. */ }
    });
    throw new Error('서명 이미지를 배치하지 못했습니다. 원본 삭제를 막기 위해 출력을 중단했습니다. 다시 시도해 주세요.');
  }
}

function exportJobIsTerminal_(job) {
  if (!job) return false;
  if (['preview_ready', 'complete', 'expired'].indexOf(String(job.status)) >= 0) return true;
  return job.status === 'failed' && !job.tempSpreadsheetId;
}

function acquireExportLease_(jobId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return '';
  try {
    const properties = PropertiesService.getScriptProperties();
    const key = 'EXPORT_LEASE_' + jobId;
    const current = String(properties.getProperty(key) || '');
    const currentExpiresAt = number_(current.split('|')[1]);
    if (currentExpiresAt > Date.now()) return '';
    const token = Utilities.getUuid() + '|' + String(Date.now() + APP.EXPORT_LEASE_MS);
    properties.setProperty(key, token);
    return token;
  } finally {
    lock.releaseLock();
  }
}

function releaseExportLease_(jobId, leaseToken) {
  if (!leaseToken) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const properties = PropertiesService.getScriptProperties();
    const key = 'EXPORT_LEASE_' + jobId;
    if (properties.getProperty(key) === leaseToken) properties.deleteProperty(key);
  } finally {
    lock.releaseLock();
  }
}

function readExportImageBatch_(spreadsheet, start) {
  const imageSheet = spreadsheet.getSheetByName('_IMAGES');
  if (imageSheet) {
    const total = Math.max(0, imageSheet.getLastRow() - 1);
    const count = Math.max(0, Math.min(APP.EXPORT_BATCH_SIZE, total - start));
    const rows = count ? imageSheet.getRange(start + 2, 1, count, 2).getValues() : [];
    return {
      total: total,
      processedCount: count,
      batch: rows.map(function(row) {
        return { layoutIndex: number_(row[0]), fileId: String(row[1] || '') };
      }).filter(function(row) { return Boolean(row.fileId); })
    };
  }

  // Jobs created before v1.6.0 keep the full roster in _DATA.
  const legacySheet = spreadsheet.getSheetByName('_DATA');
  if (!legacySheet) throw new Error('임시 출력 데이터를 찾을 수 없습니다.');
  const rowCount = Math.max(0, legacySheet.getLastRow() - 1);
  const data = rowCount ? legacySheet.getRange(2, 1, rowCount, 6).getValues() : [];
  const withImages = data.filter(function(row) { return Boolean(row[5]); });
  return {
    total: withImages.length,
    processedCount: Math.min(APP.EXPORT_BATCH_SIZE, Math.max(0, withImages.length - start)),
    batch: withImages.slice(start, start + APP.EXPORT_BATCH_SIZE).map(function(row) {
      return { layoutIndex: number_(row[0]), fileId: String(row[5] || '') };
    })
  };
}

function fetchPrivateDriveImages_(batch) {
  if (!batch.length) return [];
  const token = ScriptApp.getOAuthToken();
  const requests = batch.map(function(item) {
    return {
      url: 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(item.fileId) + '?alt=media&supportsAllDrives=true',
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    };
  });
  let responses = null;
  try {
    responses = UrlFetchApp.fetchAll(requests);
  } catch (error) {
    responses = [];
  }
  return batch.map(function(item, index) {
    const response = responses[index];
    const code = response ? response.getResponseCode() : 0;
    let blob = code >= 200 && code < 300 ? response.getBlob().setName('signature.png') : null;
    if (!blob) {
      try {
        blob = DriveApp.getFileById(item.fileId).getBlob().setName('signature.png');
      } catch (ignore) {
        blob = null;
      }
    }
    return {
      layoutIndex: item.layoutIndex,
      blob: blob
    };
  });
}

function exportRowsPerBlock_(output, job) {
  const footerRows = bool_(job.showRate) ? 6 : 5;
  return Math.max(1, output.getLastRow() - footerRows);
}

function createExportPreview_(rowNumber, job, spreadsheet) {
  spreadsheet = spreadsheet || SpreadsheetApp.openById(job.tempSpreadsheetId);
  SpreadsheetApp.flush();
  const exportFolder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('EXPORT_FOLDER_ID'));
  const safeName = safeFileName_(job.trainingTitle + '_' + sheetDateText_(job.date) + '_서명등록부');
  const previewFile = exportFolder.createFile(exportSpreadsheetBlob_(spreadsheet, 'pdf').setName('미리보기_' + safeName + '.pdf'));
  const stored = writeExportJobChanges_(rowNumber, job, {
    status: 'preview_ready',
    previewFileId: previewFile.getId(),
    error: ''
  });
  try {
    removeExportHelperSheets_(spreadsheet);
  } catch (cleanupError) {
    console.warn('출력 보조 시트 정리 실패: ' + String(cleanupError && cleanupError.message || cleanupError));
  }
  try {
    audit_('preview_export', job.jobId, number_(job.total), safeName);
  } catch (auditError) {
    console.warn('출력 미리보기 감사 기록 실패: ' + String(auditError && auditError.message || auditError));
  }
  return stored;
}

function removeExportHelperSheets_(spreadsheet) {
  const imageSheet = spreadsheet.getSheetByName('_IMAGES');
  if (imageSheet) spreadsheet.deleteSheet(imageSheet);
  const legacySheet = spreadsheet.getSheetByName('_DATA');
  if (legacySheet) spreadsheet.deleteSheet(legacySheet);
}

function exportSpreadsheetBlob_(spreadsheet, format) {
  const base = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() + '/export';
  let url = base + '?format=xlsx';
  if (format === 'pdf') {
    const output = spreadsheet.getSheetByName('서명등록부');
    url = base + '?format=pdf&gid=' + output.getSheetId() + '&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false&pagenum=CENTER&gridlines=false&fzr=true&top_margin=0.35&bottom_margin=0.4&left_margin=0.3&right_margin=0.3';
  }
  const options = {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  let lastError = '';
  for (let attempt = 0; attempt < 3; attempt += 1) {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    const bytes = response.getContent();
    const validHeader = format === 'pdf'
      ? bytes.length > 4 && String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]) === '%PDF'
      : bytes.length > 2 && String.fromCharCode(bytes[0], bytes[1]) === 'PK';
    if (code >= 200 && code < 300 && bytes.length > 100 && validHeader) return response.getBlob();
    lastError = 'HTTP ' + code + ', ' + bytes.length + ' bytes';
    if (attempt < 2) Utilities.sleep(250 * (attempt + 1));
  }
  throw new Error('출력 파일 변환에 실패했습니다. ' + lastError);
}

function finalizeExport_(jobId) {
  const entry = findRowWithNumber_(SHEETS.EXPORTS, 'jobId', id_(jobId, '출력 작업'));
  if (!entry || entry.data.status !== 'preview_ready') apiError_('EXPORT_NOT_READY', '미리보기가 준비된 작업만 파일로 만들 수 있습니다.');
  const job = entry.data;
  const outputType = String(job.outputType || '');
  if (outputType !== 'pdf' && outputType !== 'xlsx') apiError_('EXPORT_FORMAT', '인쇄 작업은 파일 생성 완료로 처리할 수 없습니다.');
  try {
    const safeName = safeFileName_(job.trainingTitle + '_' + sheetDateText_(job.date) + '_서명등록부');
    const changes = { status: 'complete', tempSpreadsheetId: '', previewFileId: '', error: '' };
    if (outputType === 'pdf') {
      const previewFile = DriveApp.getFileById(String(job.previewFileId));
      previewFile.setName(safeName + '.pdf');
      changes.pdfFileId = previewFile.getId();
    } else {
      const spreadsheet = SpreadsheetApp.openById(job.tempSpreadsheetId);
      removeExportHelperSheets_(spreadsheet);
      const exportFolder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('EXPORT_FOLDER_ID'));
      const xlsxFile = exportFolder.createFile(exportSpreadsheetBlob_(spreadsheet, 'xlsx').setName(safeName + '.xlsx'));
      changes.xlsxFileId = xlsxFile.getId();
      trashFileIfExists_(job.previewFileId);
    }
    trashFileIfExists_(job.tempSpreadsheetId);
    updateExportJob_(entry.rowNumber, changes);
    audit_('complete_export', job.jobId, number_(job.total), outputType + ' ' + safeName);
    return publicJob_(findRowWithNumber_(SHEETS.EXPORTS, 'jobId', job.jobId).data);
  } catch (error) {
    updateExportJob_(entry.rowNumber, { status: 'preview_ready', error: String(error && error.message || error).slice(0, 500) });
    apiError_('EXPORT_FAILED', '선택한 출력 파일을 만들지 못했습니다. 잠시 후 다시 시도해 주세요.');
  }
}

function recordPrintOpened_(jobId) {
  const entry = findRowWithNumber_(SHEETS.EXPORTS, 'jobId', id_(jobId, '출력 작업'));
  if (!entry || entry.data.status !== 'preview_ready' || entry.data.outputType !== 'print' || !entry.data.previewFileId) {
    apiError_('EXPORT_NOT_READY', '인쇄 미리보기가 준비되지 않았습니다.');
  }
  const timestamp = new Date().toISOString();
  updateExportJob_(entry.rowNumber, { printOpenedAt: timestamp });
  audit_('print_opened', entry.data.jobId, number_(entry.data.total), entry.data.trainingTitle + ' ' + sheetDateText_(entry.data.date));
  return publicJob_(findRowWithNumber_(SHEETS.EXPORTS, 'jobId', entry.data.jobId).data);
}

function downloadExportChunk_(jobId, format, offset, chunkSize) {
  const job = findRow_(SHEETS.EXPORTS, 'jobId', id_(jobId, '출력 작업'));
  const normalizedFormat = format === 'preview' ? 'preview' : format === 'pdf' ? 'pdf' : format === 'xlsx' ? 'xlsx' : '';
  if (!normalizedFormat) apiError_('VALIDATION', '파일 형식이 올바르지 않습니다.');
  if (!job) apiError_('NOT_FOUND', '출력 작업을 찾을 수 없습니다.');
  if (normalizedFormat === 'preview' && job.status !== 'preview_ready') apiError_('EXPORT_NOT_READY', '미리보기 파일이 아직 준비되지 않았습니다.');
  if (normalizedFormat !== 'preview' && job.status !== 'complete') apiError_('EXPORT_NOT_READY', '출력 파일이 아직 준비되지 않았습니다.');
  const fileId = normalizedFormat === 'preview' ? job.previewFileId : normalizedFormat === 'pdf' ? job.pdfFileId : job.xlsxFileId;
  if (!fileId) apiError_('NOT_FOUND', '출력 파일을 찾을 수 없습니다.');
  const token = ScriptApp.getOAuthToken();
  const driveUrl = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(String(fileId));
  const metadata = getDriveDownloadMetadata_(String(fileId), driveUrl, token);
  const totalBytes = Math.max(0, number_(metadata.size));
  const start = Math.max(0, number_(offset));
  const size = Math.max(32768, Math.min(APP.DOWNLOAD_CHUNK_SIZE, number_(chunkSize) || APP.DOWNLOAD_CHUNK_SIZE));
  if (start >= totalBytes) {
    return {
      base64: '', nextOffset: totalBytes, totalBytes: totalBytes,
      fileName: String(metadata.name || 'download'), mimeType: String(metadata.mimeType || 'application/octet-stream')
    };
  }
  const end = Math.min(totalBytes, start + size);
  const contentResponse = UrlFetchApp.fetch(
    driveUrl + '?alt=media&supportsAllDrives=true',
    {
      headers: {
        Authorization: 'Bearer ' + token,
        Range: 'bytes=' + start + '-' + (end - 1)
      },
      muteHttpExceptions: true
    }
  );
  const responseCode = contentResponse.getResponseCode();
  if (responseCode !== 200 && responseCode !== 206) apiError_('DOWNLOAD_FAILED', '출력 파일을 내려받지 못했습니다.');
  const normalized = normalizeDriveDownloadResponse_(
    responseCode,
    contentResponse.getHeaders(),
    contentResponse.getContent(),
    start,
    end,
    totalBytes
  );
  return {
    base64: Utilities.base64Encode(normalized.bytes), nextOffset: normalized.nextOffset, totalBytes: totalBytes,
    fileName: String(metadata.name || 'download'), mimeType: String(metadata.mimeType || 'application/octet-stream')
  };
}

function normalizeDriveDownloadResponse_(responseCode, headers, bytes, start, end, totalBytes) {
  const normalizedHeaders = headers || {};
  const contentRange = String(
    normalizedHeaders['Content-Range'] ||
    normalizedHeaders['content-range'] ||
    ''
  );
  const expectedLength = Math.max(0, end - start);
  if (responseCode === 206) {
    const match = /^bytes\s+(\d+)-(\d+)\/(\d+|\*)$/i.exec(contentRange);
    if (match && number_(match[1]) !== start) {
      apiError_('DOWNLOAD_FAILED', '출력 파일의 다운로드 위치가 일치하지 않습니다.');
    }
    if (!bytes.length || bytes.length > expectedLength) bytes = bytes.slice(0, expectedLength);
    if (!bytes.length) apiError_('DOWNLOAD_FAILED', '출력 파일 조각이 비어 있습니다.');
    return { bytes: bytes, nextOffset: Math.min(totalBytes, start + bytes.length) };
  }

  // Drive가 Range를 무시한 경우 전체 파일을 이번 응답으로 끝내 재다운로드를 막습니다.
  if (bytes.length === totalBytes) {
    const remaining = start ? bytes.slice(start) : bytes;
    if (!remaining.length) apiError_('DOWNLOAD_FAILED', '출력 파일 조각이 비어 있습니다.');
    return { bytes: remaining, nextOffset: totalBytes };
  }
  if (start === 0 && bytes.length) {
    return { bytes: bytes, nextOffset: Math.min(totalBytes, bytes.length) };
  }
  apiError_('DOWNLOAD_FAILED', '출력 파일을 이어받지 못했습니다. 다시 시도해 주세요.');
}

function getDriveDownloadMetadata_(fileId, driveUrl, token) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'EXPORT_FILE_META_' + fileId;
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      const metadata = JSON.parse(cached);
      if (number_(metadata.size) > 0 && metadata.name && metadata.mimeType) return metadata;
    } catch (ignore) {
      // Invalid cache entries are replaced from Drive below.
    }
  }
  const response = UrlFetchApp.fetch(
    driveUrl + '?fields=name%2CmimeType%2Csize&supportsAllDrives=true',
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    apiError_('NOT_FOUND', '출력 파일을 찾을 수 없습니다.');
  }
  const metadata = JSON.parse(response.getContentText());
  if (number_(metadata.size) <= 0) apiError_('DOWNLOAD_FAILED', '출력 파일 크기를 확인하지 못했습니다.');
  cache.put(cacheKey, JSON.stringify(metadata), 600);
  return metadata;
}

function purgeOriginals_(jobId, confirmation) {
  const entry = findRowWithNumber_(SHEETS.EXPORTS, 'jobId', id_(jobId, '출력 작업'));
  if (!entry || !canPurgeExport_(entry.data)) apiError_('EXPORT_REQUIRED', 'PDF 또는 엑셀 파일이 정상 보관된 작업만 원본을 삭제할 수 있습니다. 인쇄 작업만으로는 삭제할 수 없습니다.');
  const job = entry.data;
  if (!safeEqual_(String(confirmation || ''), String(job.trainingTitle || ''))) apiError_('CONFIRMATION_MISMATCH', '연수명이 일치하지 않습니다.');
  if (job.pdfFileId) assertFileExists_(job.pdfFileId);
  if (job.xlsxFileId) assertFileExists_(job.xlsxFileId);
  const training = findRow_(SHEETS.TRAININGS, 'id', String(job.trainingId || ''));
  const snapshotIds = parseSignatureSnapshot_(job.signatureSnapshot);
  const sheet = sheet_(SHEETS.SIGNATURES);
  const records = readRowsWithNumbers_(SHEETS.SIGNATURES)
    .filter(function(item) {
      return signatureBelongsToExportPurge_(item.data, job, training, snapshotIds);
    })
    .sort(function(a, b) { return b.rowNumber - a.rowNumber; });
  let deleted = 0;
  let failed = 0;
  records.forEach(function(item) {
    try {
      trashFileIfExists_(item.data.imageFileId);
      sheet.deleteRow(item.rowNumber);
      deleted += 1;
    } catch (error) {
      failed += 1;
    }
  });
  invalidateRows_(SHEETS.SIGNATURES);
  if (!failed) updateExportJob_(entry.rowNumber, { purgedAt: new Date().toISOString() });
  audit_('purge_originals', job.jobId, deleted, '실패 ' + failed);
  invalidateRows_(SHEETS.EXPORTS);
  const stored = findRow_(SHEETS.EXPORTS, 'jobId', job.jobId);
  return { deleted: deleted, failed: failed, job: publicJob_(stored || job) };
}

function parseSignatureSnapshot_(value) {
  const raw = String(value || '').trim();
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) throw new Error('not an array');
    return new Set(parsed.map(function(id) { return String(id || ''); }).filter(Boolean));
  } catch (error) {
    apiError_('EXPORT_DATA', '출력 당시 서명 목록을 확인하지 못했습니다. 원본을 삭제하지 않았습니다.');
  }
}

function signatureBelongsToExportPurge_(signature, job, training, snapshotIds) {
  if (!signature || !job || signature.trainingId !== job.trainingId) return false;
  if (snapshotIds) return snapshotIds.has(String(signature.id || ''));
  const signatureCreatedAt = new Date(String(signature.createdAt || '')).getTime();
  const exportStartedAt = new Date(String(job.createdAt || '')).getTime();
  if (!signatureCreatedAt || !exportStartedAt || signatureCreatedAt > exportStartedAt) return false;
  return sheetDateText_(signature.scopeDate || signature.signDate) === sheetDateText_(job.date);
}

function cleanupStaleExportJobs() {
  const cutoff = Date.now() - 24 * 60 * 60 * 1000;
  const rows = readRowsWithNumbers_(SHEETS.EXPORTS);
  rows.forEach(function(item) {
    const status = item.data.status;
    const created = new Date(item.data.createdAt).getTime();
    if (status !== 'complete' && status !== 'expired' && created && created < cutoff) {
      if (item.data.tempSpreadsheetId) trashFileIfExists_(item.data.tempSpreadsheetId);
      if (item.data.previewFileId) trashFileIfExists_(item.data.previewFileId);
      updateExportJob_(item.rowNumber, { status: 'expired', tempSpreadsheetId: '', previewFileId: '', error: '24시간이 지나 자동 정리됨' });
    }
  });
  cleanupExpiredTrainingReceptions_();
}

function cleanupExpiredTrainingReceptions_() {
  return withAdminMutationLock_(function() {
    const properties = PropertiesService.getScriptProperties();
    const all = properties.getProperties();
    const now = new Date();
    const today = formatDate_(now, 'yyyy-MM-dd');
    let cleaned = 0;
    Object.keys(all).forEach(function(key) {
      if (key.indexOf(ONSITE_RECEPTION_PROPERTY_PREFIX_) !== 0) return;
      const trainingId = key.slice(ONSITE_RECEPTION_PROPERTY_PREFIX_.length);
      const state = readTrainingReceptionState_(trainingId);
      const expiresAt = state ? new Date(String(state.expiresAt || '')).getTime() : 0;
      if (!state || !expiresAt || expiresAt <= now.getTime() || String(state.serverDate || '') !== today) {
        properties.deleteProperty(key);
        cleaned += 1;
      }
    });
    if (cleaned) auditReceptionBestEffort_('cleanup_training_reception', 'reception', cleaned, '만료되거나 손상된 현장 접수 상태 정리');
    return cleaned;
  });
}

function updateExportJob_(rowNumber, changes) {
  const sheet = sheet_(SHEETS.EXPORTS);
  const current = rowObject_(SHEETS.EXPORTS.headers, sheet.getRange(rowNumber, 1, 1, SHEETS.EXPORTS.headers.length).getValues()[0]);
  return writeExportJobChanges_(rowNumber, current, changes);
}

function writeExportJobChanges_(rowNumber, current, changes) {
  const stored = Object.assign({}, current, changes, { updatedAt: new Date().toISOString() });
  writeObjectRow_(sheet_(SHEETS.EXPORTS), SHEETS.EXPORTS.headers, rowNumber, stored, SHEETS.EXPORTS);
  return stored;
}

function publicJob_(job) {
  const outputType = job.outputType || (job.pdfFileId && job.xlsxFileId ? 'legacy_both' : job.xlsxFileId ? 'xlsx' : job.pdfFileId ? 'pdf' : 'pdf');
  const canResume = Boolean(job.tempSpreadsheetId && ['queued', 'processing', 'failed'].indexOf(String(job.status)) >= 0);
  return {
    jobId: job.jobId, trainingId: job.trainingId, trainingTitle: job.trainingTitle, date: sheetDateText_(job.date),
    sort: job.sort, columns: number_(job.columns), showRate: bool_(job.showRate), status: job.status,
    progress: number_(job.progress), total: number_(job.total), createdAt: job.createdAt, updatedAt: job.updatedAt,
    outputType: outputType, hasPreview: Boolean(job.previewFileId), hasPdf: Boolean(job.pdfFileId), hasXlsx: Boolean(job.xlsxFileId),
    canPurge: canPurgeExport_(job), canResume: canResume, printOpenedAt: job.printOpenedAt || '', error: job.error || '', purgedAt: job.purgedAt || ''
  };
}

function canPurgeExport_(job) {
  if (!job || job.status !== 'complete') return false;
  if (!job.outputType) return Boolean(job.pdfFileId && job.xlsxFileId);
  if (job.outputType === 'pdf') return Boolean(job.pdfFileId);
  if (job.outputType === 'xlsx') return Boolean(job.xlsxFileId);
  return false;
}

function exportPosition_(index, rowsPerBlock) {
  const block = Math.floor(index / rowsPerBlock);
  return { block: block, row: 5 + (index % rowsPerBlock) };
}

function spreadsheet_() {
  requireInitialized_();
  const context = requestContext_();
  if (!context.spreadsheet) {
    context.spreadsheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
  }
  return context.spreadsheet;
}

/** 초기화·시트 메뉴에서만 사용합니다. 웹앱 요청 경로에서는 호출하지 않습니다. */
function boundSpreadsheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) apiError_('BOUND_SHEET_REQUIRED', '이 함수는 학교용 Google 시트에 연결된 Apps Script에서 실행해야 합니다.');
  return spreadsheet;
}

function dataSheetDefinitions_() {
  return [SHEETS.SETTINGS, SHEETS.STAFF, SHEETS.TRAININGS, SHEETS.SIGNATURES, SHEETS.EXPORTS, SHEETS.AUDIT];
}

function clearCopiedInstanceProperties_(properties) {
  INSTANCE_PROPERTIES.forEach(function(key) { properties.deleteProperty(key); });
  Object.keys(properties.getProperties()).forEach(function(key) {
    if (key.indexOf(ONSITE_RECEPTION_PROPERTY_PREFIX_) === 0) properties.deleteProperty(key);
  });
}

function getOrRepairFolder_(properties, propertyKey, parentFolder, name) {
  const existingId = properties.getProperty(propertyKey);
  if (existingId) {
    try {
      const existing = DriveApp.getFolderById(existingId);
      existing.getName();
      return existing;
    } catch (error) {
      console.warn(propertyKey + ' 폴더를 찾을 수 없어 다시 만듭니다: ' + String(error && error.message || error));
    }
  }
  return parentFolder ? parentFolder.createFolder(name) : DriveApp.createFolder(name);
}

function ensureGuideSheet_(spreadsheet, rebuild) {
  let sheet = spreadsheet.getSheetByName(APP.GUIDE_SHEET);
  if (!sheet) {
    const reusable = spreadsheet.getSheets().find(function(candidate) {
      return dataSheetDefinitions_().every(function(definition) { return candidate.getName() !== definition.name; });
    });
    sheet = reusable || spreadsheet.insertSheet(APP.GUIDE_SHEET, 0);
    sheet.setName(APP.GUIDE_SHEET);
    rebuild = true;
  }
  if (!rebuild && String(sheet.getRange('B2').getValue()).indexOf('학교 연수 전자서명') >= 0) {
    sheet.showSheet();
    spreadsheet.setActiveSheet(sheet);
    return sheet;
  }

  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(4);
  sheet.setTabColor('#165DFF');
  sheet.setColumnWidth(1, 22);
  sheet.setColumnWidth(2, 54);
  sheet.setColumnWidth(3, 178);
  sheet.setColumnWidth(4, 270);
  sheet.setColumnWidth(5, 170);
  sheet.setColumnWidth(6, 22);
  for (let row = 1; row <= 66; row += 1) sheet.setRowHeight(row, 27);
  sheet.setRowHeight(1, 14);
  sheet.setRowHeight(2, 44);
  sheet.setRowHeight(3, 26);
  sheet.setRowHeight(4, 44);

  sheet.getRange('A1:F66')
    .setBackground('#F7F9FC')
    .setFontFamily('Arial')
    .setFontColor('#334155')
    .setFontSize(10)
    .setVerticalAlignment('middle');
  sheet.getRange('B2:E2').merge()
    .setValue('학교 연수 전자서명 시스템 사용설명서')
    .setBackground('#123A73')
    .setFontColor('#FFFFFF')
    .setFontSize(18)
    .setFontWeight('bold');
  sheet.getRange('B3:E3').merge()
    .setValue('다른 학교에서 사본을 만들어 자기 계정으로 독립 운영하는 방법')
    .setBackground('#123A73')
    .setFontColor('#DCE8FF')
    .setFontSize(10);
  const isDistributionTemplate = PropertiesService.getScriptProperties().getProperty('TEMPLATE_LOCK') === '1';
  sheet.getRange('B4:E4').merge()
    .setValue(isDistributionTemplate
      ? '중요  이 파일은 배포용 원본입니다. 원본에서 초기화하지 말고 반드시 파일 → 사본 만들기로 시작하세요.'
      : '중요  현재 운영 시트에는 학교 데이터가 있으므로 다른 학교에 공유하지 마세요. 비어 있는 배포용 원본의 사본을 안내하세요.')
    .setBackground('#FFF4DF')
    .setFontColor('#A05A00')
    .setFontWeight('bold')
    .setWrap(true);

  const guideSections = [
    {
      title: '1. 학교용 사본과 웹앱 만들기',
      color: '#165DFF',
      pale: '#EAF1FF',
      steps: [
        ['1', '학교용 사본 만들기', '상단 메뉴에서 파일 → 사본 만들기를 누릅니다.', '반드시 새로 만든 사본에서 다음 단계를 진행합니다.'],
        ['2', '초기 설정과 권한 승인', '새 사본에서 🖊️ 전자서명 관리 → 초기 설정 실행을 누르고, 완료 창의 초기 설정 코드를 보관합니다.', '‘Google에서 확인하지 않은 앱’이 나오면 개발자 이메일이 본인 계정인지 확인한 뒤 고급 → 학교 연수 전자서명 시스템으로 이동 → 허용을 누릅니다. 모르는 이메일이거나 Google Workspace 정책이 차단하면 중단하고 관리자에게 문의하세요. 소유자만 처음 한 번 승인하며 서명 참여자에게는 이 화면이 나타나지 않습니다.', 134],
        ['3', '웹앱 배포와 화면 확인', '확장 프로그램 → Apps Script에서 배포 → 새 배포 → 유형 선택 → 웹 앱을 고릅니다.', '실행 계정은 나, 액세스 권한은 모든 사용자(로그인 없이 접근 가능)로 설정합니다. 배포 뒤 /exec로 끝나는 주소를 열면 전자서명 화면과 관리자 버튼이 함께 보이며 별도 홈페이지를 만들 필요가 없습니다.', 120]
      ]
    },
    {
      title: '2. 관리자 첫 설정',
      color: '#6F42C1',
      pale: '#F3EDFF',
      steps: [
        ['1', '관리자 비밀번호 만들기', '초기 설정 코드와 관리자 비밀번호를 입력합니다.', '숫자 4자리 또는 문자·숫자를 포함한 10자 이상 비밀번호를 사용할 수 있습니다.'],
        ['2', '기관 설정 입력', '기관 설정에서 학교명·부제목·안내문·개인정보 처리 안내·대표 색상을 저장합니다.', '개인정보 안내가 비어 있으면 연수를 활성화할 수 없습니다.'],
        ['3', '구성원 등록', '구성원에서 부서와 성명을 등록합니다.', '이름은 띄어쓰기·쉼표·줄바꿈으로 여러 명을 한 번에 입력하거나 엑셀·CSV를 가져올 수 있습니다.'],
        ['4', '연수 등록', '연수에서 날짜·시간·활성 상태와 서명 대상(전체 구성원 또는 부서 선택)을 정합니다.', '현장 확인이 필요한 연수만 현장 접수 코드 사용을 켭니다. 기존 연수와 이 설정을 끈 연수는 코드 없이 운영됩니다.', 76],
        ['5', '공유 링크 배포', '공유·보안에서 참여 링크와 QR을 복사해 학교 내부에 안내합니다.', '링크를 받은 사람은 별도 Google 계정 없이 접속합니다.']
      ]
    },
    {
      title: '3. 연수 운영과 출력',
      color: '#165DFF',
      pale: '#F4F8FF',
      steps: [
        ['1', '현장 접수와 참여', '연수 카드에서 5분·10분·15분 중 접수 시간을 고르고 현장 접수를 시작한 뒤 화면의 6자리 코드를 참석자에게 안내합니다.', '코드 재발급·접수 종료 시 이전 코드는 즉시 무효화되며, 오늘·매일 연수는 등록된 종료 시각에 맞춰 접수도 끝납니다.', 76],
        ['2', '미서명 현황 확인', '연수 목록에서 미서명 현황을 눌러 미서명만 또는 전체 현황을 확인합니다.', '매일 연수는 날짜를 바꿀 수 있으며 대상·서명·미서명·서명률과 서명 시각을 볼 수 있습니다.', 76],
        ['3', '서명 기록 확인', '서명 기록에서 연수를 선택하면 해당 연수 날짜가 자동으로 설정됩니다.', '잘못 제출된 기록은 개별 삭제할 수 있습니다.'],
        ['4', '출력 미리보기', '연수 목록의 출력 버튼에서 날짜·열 수·정렬·서명률·PDF/엑셀/인쇄를 고릅니다.', '실제 서명 이미지가 들어간 A4 미리보기를 먼저 확인합니다.', 76],
        ['5', '파일 보관', 'PDF 또는 엑셀을 내려받아 학교가 정한 보관 위치에 저장합니다.', '인쇄하기만 실행한 작업은 원본 삭제 조건을 충족하지 않습니다.'],
        ['6', '원본 삭제', '보관 파일을 확인한 뒤 연수명을 다시 입력해 해당 날짜의 원본 서명과 PNG를 삭제합니다.', '출력 파일은 비공개 Drive의 출력 보관 폴더에 남습니다.']
      ]
    },
    {
      title: '4. 보안·복구',
      color: '#B42318',
      pale: '#FFF0EE',
      steps: [
        ['•', '공유 링크 보호', '공유 링크에는 본인 인증이 없습니다. 학교 내부에서만 안내하세요.', '유출되면 공유·보안 → 공유 키 교체 후 새 링크와 QR을 배포합니다.'],
        ['•', '비밀번호 분실', '학교용 시트에서 🖊️ 전자서명 관리 → 관리자 비밀번호 복구를 누릅니다.', '기존 관리자 세션은 모두 무효화됩니다.'],
        ['•', '데이터 탭 점검', '🖊️ 전자서명 관리 → 데이터 탭 표시·숨기기를 사용합니다.', '관리용 Google 계정 외에는 시트 편집 권한을 주지 않습니다.'],
        ['•', '자료 최소 보관', '학교가 정한 보관기간이 끝나면 출력 파일도 관리용 Drive에서 삭제합니다.', '이 시스템은 연수 참여 확인용이며 공식 동의·결재·법적 전자서명용이 아닙니다.', 76]
      ]
    }
  ];

  let guideRow = 6;
  guideSections.forEach(function(section) {
    sheet.getRange(guideRow, 2, 1, 4).merge()
      .setValue(section.title)
      .setBackground(section.color)
      .setFontColor('#FFFFFF')
      .setFontSize(12)
      .setFontWeight('bold');
    sheet.setRowHeight(guideRow, 32);
    guideRow += 1;
    section.steps.forEach(function(step) {
      sheet.getRange(guideRow, 2)
        .setValue(step[0])
        .setBackground(section.color)
        .setFontColor('#FFFFFF')
        .setFontSize(11)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      sheet.getRange(guideRow, 3)
        .setValue(step[1])
        .setBackground(section.pale)
        .setFontColor('#14213D')
        .setFontWeight('bold')
        .setWrap(true);
      sheet.getRange(guideRow, 4, 1, 2).merge()
        .setValue(step[2] + '\n참고  ' + step[3])
        .setBackground('#FFFFFF')
        .setFontColor('#334155')
        .setWrap(true);
      sheet.getRange(guideRow, 2, 1, 4)
        .setBorder(false, false, true, false, false, false, '#D8E2F0', SpreadsheetApp.BorderStyle.SOLID);
      sheet.setRowHeight(guideRow, step[4] || 58);
      guideRow += 1;
    });
    guideRow += 1;
  });

  sheet.getRange(guideRow, 2, 1, 4).merge()
    .setValue('공식 참고 문서')
    .setBackground('#E9EEF5')
    .setFontColor('#14213D')
    .setFontSize(11)
    .setFontWeight('bold');
  guideRow += 1;
  sheet.getRange(guideRow, 2, 2, 4).merge()
    .setValue('Google Apps Script 연결형 스크립트  https://developers.google.com/apps-script/guides/bound\nGoogle Apps Script 웹앱 배포  https://developers.google.com/apps-script/guides/web')
    .setBackground('#FFFFFF')
    .setFontColor('#64748B')
    .setFontSize(9)
    .setWrap(true);
  sheet.setRowHeights(guideRow, 2, 32);
  sheet.showSheet();
  spreadsheet.setActiveSheet(sheet);
  return sheet;
}

function hideDataSheets_(spreadsheet) {
  ensureGuideSheet_(spreadsheet, false).showSheet();
  dataSheetDefinitions_().forEach(function(definition) {
    const sheet = spreadsheet.getSheetByName(definition.name);
    if (sheet) sheet.hideSheet();
  });
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(APP.GUIDE_SHEET));
}

function sheet_(definition) {
  const context = requestContext_();
  if (!context.sheets[definition.name]) context.sheets[definition.name] = ensureSheet_(spreadsheet_(), definition);
  return context.sheets[definition.name];
}

function ensureSheet_(spreadsheet, definition) {
  let sheet = spreadsheet.getSheetByName(definition.name);
  if (!sheet) sheet = spreadsheet.insertSheet(definition.name);
  const width = definition.headers.length;
  const current = sheet.getRange(1, 1, 1, width).getValues()[0];
  if (current.join('\u0000') !== definition.headers.join('\u0000')) {
    sheet.getRange(1, 1, 1, width).setValues([definition.headers]).setFontWeight('bold').setBackground('#dfece8');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readRows_(definition) {
  return readRowsWithNumbers_(definition).map(function(item) { return item.data; });
}

function readRowsWithNumbers_(definition) {
  const context = requestContext_();
  const cached = context.rows[definition.name];
  if (cached) return cloneRowEntries_(cached);
  const sheet = sheet_(definition);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    context.rows[definition.name] = [];
    return [];
  }
  const rows = sheet.getRange(2, 1, lastRow - 1, definition.headers.length).getValues()
    .map(function(values, index) { return { rowNumber: index + 2, data: rowObject_(definition.headers, values) }; })
    .filter(function(item) { return Object.keys(item.data).some(function(key) { return item.data[key] !== ''; }); });
  context.rows[definition.name] = rows;
  return cloneRowEntries_(rows);
}

function cloneRowEntries_(rows) {
  return rows.map(function(item) { return { rowNumber: item.rowNumber, data: Object.assign({}, item.data) }; });
}

function invalidateRows_(definition) {
  if (REQUEST_CONTEXT_) delete REQUEST_CONTEXT_.rows[definition.name];
}

function rowObject_(headers, values) {
  const result = {};
  headers.forEach(function(header, index) { result[header] = values[index]; });
  return result;
}

function objectValues_(headers, object) {
  return headers.map(function(header) { return object[header] === undefined || object[header] === null ? '' : object[header]; });
}

function appendObject_(definition, object) {
  const sheet = sheet_(definition);
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, definition.headers.length)
    .setValues([objectValues_(definition.headers, object)]);
  invalidateRows_(definition);
}

function writeObjectRow_(sheet, headers, rowNumber, object, definition) {
  sheet.getRange(rowNumber, 1, 1, headers.length).setValues([objectValues_(headers, object)]);
  if (definition) invalidateRows_(definition);
}

function findRow_(definition, key, value) {
  const entry = findRowWithNumber_(definition, key, value);
  return entry ? entry.data : null;
}

function findRowWithNumber_(definition, key, value) {
  return readRowsWithNumbers_(definition).find(function(item) { return String(item.data[key]) === String(value); }) || null;
}

function deleteRowById_(definition, id) {
  const entry = findRowWithNumber_(definition, 'id', id);
  if (!entry) apiError_('NOT_FOUND', definition.name + ' 항목을 찾을 수 없습니다.');
  sheet_(definition).deleteRow(entry.rowNumber);
  invalidateRows_(definition);
}

function readSettings_() {
  const values = {};
  SETTING_KEYS.forEach(function(key) { values[key] = ''; });
  readRows_(SHEETS.SETTINGS).forEach(function(row) {
    if (SETTING_KEYS.indexOf(String(row.key)) >= 0) values[String(row.key)] = String(row.value || '');
  });
  return values;
}

function writeSettings_(settings) {
  const sheet = sheet_(SHEETS.SETTINGS);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, SHEETS.SETTINGS.headers.length).clearContent();
  const values = SETTING_KEYS.map(function(key) {
    return [key, settings[key] === undefined ? '' : String(settings[key])];
  });
  sheet.getRange(2, 1, values.length, SHEETS.SETTINGS.headers.length).setValues(values);
  invalidateRows_(SHEETS.SETTINGS);
}

function privacyReady_(settings) {
  return ['schoolName', 'subtitle', 'privacyPurpose', 'privacyItems', 'privacyRetention']
    .every(function(key) { return Boolean(String(settings[key] || '').trim()); });
}

function validateFaviconData_(value) {
  const data = String(value === undefined || value === null ? '' : value).trim();
  if (!data) return '';
  if (data.length > 45000) apiError_('FAVICON_TOO_LARGE', '파비콘 PNG는 32KB 이하여야 합니다.');
  const match = data.match(/^data:image\/png;base64,([A-Za-z0-9+/]+={0,2})$/);
  if (!match || match[1].length % 4 !== 0) apiError_('BAD_FAVICON', '64×64 PNG 파비콘만 저장할 수 있습니다.');
  let bytes;
  try {
    bytes = Utilities.base64Decode(match[1]);
  } catch (error) {
    apiError_('BAD_FAVICON', '파비콘 PNG 데이터를 읽을 수 없습니다.');
  }
  if (!bytes || bytes.length < 33 || bytes.length > 32 * 1024) apiError_('FAVICON_TOO_LARGE', '파비콘 PNG는 32KB 이하여야 합니다.');
  const expectedSignature = [137, 80, 78, 71, 13, 10, 26, 10];
  const validSignature = expectedSignature.every(function(byte, index) { return (bytes[index] & 255) === byte; });
  const hasIhdr = String.fromCharCode(bytes[12] & 255, bytes[13] & 255, bytes[14] & 255, bytes[15] & 255) === 'IHDR';
  if (!validSignature || !hasIhdr || pngUint32_(bytes, 16) !== 64 || pngUint32_(bytes, 20) !== 64) {
    apiError_('BAD_FAVICON', '파비콘은 가로·세로 64픽셀 PNG여야 합니다.');
  }
  return 'data:image/png;base64,' + match[1];
}

function pngUint32_(bytes, offset) {
  return (((bytes[offset] & 255) * 0x1000000)
    + ((bytes[offset + 1] & 255) << 16)
    + ((bytes[offset + 2] & 255) << 8)
    + (bytes[offset + 3] & 255)) >>> 0;
}

function publicStaff_(row) {
  return { id: String(row.id), department: String(row.department), name: String(row.name), active: bool_(row.active), sortOrder: number_(row.sortOrder) };
}

function normalizeAudienceDepartments_(value) {
  let source = value;
  if (!Array.isArray(source)) {
    try { source = JSON.parse(String(source || '[]')); }
    catch (error) { source = []; }
  }
  if (!Array.isArray(source)) return [];
  const seen = new Set();
  const departments = [];
  source.forEach(function(item) {
    const department = String(item || '').trim();
    const key = department.toLowerCase();
    if (!department || department.length > 50 || seen.has(key)) return;
    seen.add(key);
    departments.push(department);
  });
  return departments.slice(0, 200);
}

function trainingAudienceMode_(training) {
  const mode = String(training && training.audienceMode || '').trim();
  if (!mode || mode === 'all') return 'all';
  if (mode === 'departments') return 'departments';
  return 'invalid';
}

function trainingAudienceDepartments_(training) {
  return trainingAudienceMode_(training) === 'departments'
    ? normalizeAudienceDepartments_(training && training.audienceDepartments)
    : [];
}

function trainingTargetsStaff_(training, person) {
  const mode = trainingAudienceMode_(training);
  if (mode === 'all') return true;
  if (mode !== 'departments') return false;
  const department = String(person && person.department || '').trim();
  const key = department.toLowerCase();
  return trainingAudienceDepartments_(training).some(function(item) {
    return String(item || '').trim().toLowerCase() === key;
  });
}

function listActiveDepartments_() {
  const seen = new Set();
  const departments = [];
  readRows_(SHEETS.STAFF)
    .filter(function(person) { return bool_(person.active); })
    .sort(staffSort_)
    .forEach(function(person) {
      const department = String(person.department || '').trim();
      const key = department.toLowerCase();
      if (!department || seen.has(key)) return;
      seen.add(key);
      departments.push(department);
    });
  return departments;
}

function publicTraining_(row) {
  const storedAudienceMode = trainingAudienceMode_(row);
  const audienceMode = storedAudienceMode === 'all' ? 'all' : 'departments';
  return {
    id: String(row.id), title: String(row.title), date: sheetDateText_(row.date), daily: bool_(row.daily),
    startTime: sheetTimeText_(row.startTime, false), endTime: sheetTimeText_(row.endTime, false), active: bool_(row.active), sortOrder: number_(row.sortOrder),
    audienceMode: audienceMode, audienceDepartments: trainingAudienceDepartments_(row),
    onsiteCodeRequired: bool_(row.onsiteCodeRequired)
  };
}

function staffSort_(a, b) {
  return number_(a.sortOrder) - number_(b.sortOrder) || compareKo_(a.department, b.department) || compareKo_(a.name, b.name);
}

function orderSort_(a, b) {
  return number_(a.sortOrder) - number_(b.sortOrder) || compareKo_(a.title, b.title);
}

function getOrCreateTrainingFolder_(trainingId, title) {
  const root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('SIGNATURE_FOLDER_ID'));
  const folderName = safeFileName_(trainingId.slice(0, 8) + '_' + title);
  const iterator = root.getFoldersByName(folderName);
  return iterator.hasNext() ? iterator.next() : root.createFolder(folderName);
}

function trashFileIfExists_(fileId) {
  if (!fileId) return;
  try {
    const file = DriveApp.getFileById(String(fileId));
    if (!file.isTrashed()) file.setTrashed(true);
  } catch (error) {
    const message = String(error && error.message || error);
    if (/not found|does not exist|찾을 수|유효하지/i.test(message)) return;
    throw error;
  }
}

function assertFileExists_(fileId) {
  try {
    DriveApp.getFileById(String(fileId)).getName();
  } catch (error) {
    apiError_('EXPORT_MISSING', '보관된 출력 파일을 찾을 수 없어 원본을 삭제하지 않았습니다.');
  }
}

function audit_(action, target, count, detail) {
  appendObject_(SHEETS.AUDIT, { timestamp: new Date().toISOString(), action: action, target: String(target || ''), count: number_(count), detail: String(detail || '').slice(0, 500) });
}

function auditReceptionBestEffort_(action, target, count, detail) {
  try {
    audit_(action, target, count, detail);
  } catch (error) {
    console.error('현장 접수 감사 기록 저장 실패: ' + String(error && error.message || error));
  }
}

function ensureCleanupTrigger_() {
  const exists = ScriptApp.getProjectTriggers().some(function(trigger) { return trigger.getHandlerFunction() === 'cleanupStaleExportJobs'; });
  if (!exists) ScriptApp.newTrigger('cleanupStaleExportJobs').timeBased().everyHours(6).create();
}

function requireInitialized_() {
  if (!PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')) apiError_('NOT_INITIALIZED', 'Apps Script 편집기에서 initializeSystem 함수를 먼저 실행해 주세요.');
}

function jsonOutput_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

function apiError_(code, message, details) {
  const error = new Error(message);
  error.apiCode = code;
  error.details = details || null;
  throw error;
}

function string_(value, maxLength) {
  return String(value === undefined || value === null ? '' : value).replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, maxLength || 500);
}

function bool_(value) {
  return value === true || value === 1 || String(value).toLowerCase() === 'true';
}

function number_(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function id_(value, label) {
  const result = string_(value, 100);
  if (!/^[A-Za-z0-9_-]{8,100}$/.test(result)) apiError_('VALIDATION', (label || 'ID') + ' 값이 올바르지 않습니다.');
  return result;
}

function validDate_(value) {
  const date = string_(value, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) apiError_('VALIDATION', '날짜가 올바르지 않습니다.');
  return date;
}

function today_() {
  return formatDate_(new Date(), 'yyyy-MM-dd');
}

function formatDate_(date, pattern) {
  return Utilities.formatDate(date, APP.TIME_ZONE, pattern);
}

function sheetDateText_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return formatDate_(value, 'yyyy-MM-dd');
  return String(value || '').trim();
}

function sheetTimeText_(value, includeSeconds) {
  if (value instanceof Date && !isNaN(value.getTime())) return formatDate_(value, includeSeconds ? 'HH:mm:ss' : 'HH:mm');
  const text = String(value || '').trim();
  if (!text) return '';
  const match = text.match(/^(\d{2}):(\d{2})(?::(\d{2}))?$/);
  if (!match) return text;
  return match[1] + ':' + match[2] + (includeSeconds ? ':' + (match[3] || '00') : '');
}

function formatKoreanDate_(date) {
  const parts = String(date).split('-');
  return parts.length === 3 ? Number(parts[0]) + '년 ' + Number(parts[1]) + '월 ' + Number(parts[2]) + '일' : String(date);
}

function compareKo_(a, b) {
  return String(a || '').localeCompare(String(b || ''), 'ko');
}

function staffKey_(department, name) {
  return String(department || '').trim().toLowerCase() + '\u0000' + String(name || '').trim().toLowerCase();
}

function safeFileName_(value) {
  return String(value || '').replace(/[\\/:*?"<>|\r\n]+/g, '_').replace(/\s+/g, ' ').trim().slice(0, 100) || '연수';
}

function randomToken_(length) {
  const seed = Utilities.getUuid() + Utilities.getUuid() + new Date().getTime() + Math.random();
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed, Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '').slice(0, length || 24);
}

function safeEqual_(left, right) {
  const a = String(left || '');
  const b = String(right || '');
  let mismatch = a.length ^ b.length;
  const length = Math.max(a.length, b.length);
  for (let index = 0; index < length; index += 1) mismatch |= (a.charCodeAt(index % Math.max(1, a.length)) || 0) ^ (b.charCodeAt(index % Math.max(1, b.length)) || 0);
  return mismatch === 0;
}

function currentWebAppUrl_() {
  const context = requestContext_();
  if (context.webAppUrl) return context.webAppUrl;
  const value = String(ScriptApp.getService().getUrl() || '').trim();
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(value)) {
    apiError_('WEB_APP_NOT_DEPLOYED', '현재 Apps Script 웹앱 실행 주소를 확인할 수 없습니다. /exec 주소로 배포한 뒤 다시 시도해 주세요.');
  }
  context.webAppUrl = value;
  return value;
}

function buildShareUrl_(baseUrl, token) {
  return baseUrl && token ? baseUrl + '#k=' + encodeURIComponent(token) : '';
}
