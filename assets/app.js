import {
  buildShareUrl,
  formatKoreanDate,
  formatKoreanHeaderDate,
  groupStaffByDepartment,
  isPrivacyReady,
  isValidAdminPassword,
  localDuplicateKey,
  normalizeNameEntryText,
  normalizeRosterRows,
  parseShareToken,
  splitNames,
  todaySeoul,
  trainingTimeLabel,
  validateTraining
} from './core.js?v=20260715.2';

const $ = id => document.getElementById(id);
const config = window.TRAINING_SIGN_CONFIG || {};
const DEMO = new URLSearchParams(location.search).get('demo') === '1';
const API_URL = String(config.API_URL || '');
let shareToken = DEMO ? 'DEMO_TOKEN_1234567890123456' : parseShareToken(location.hash);
const baseUrl = `${location.origin}${location.pathname}`;
const EXPORT_SETTINGS_KEY = 'training-sign:export-settings';
const ADMIN_SYNC_MS = 30000;
const DEFAULT_FAVICON_URL = 'favicon.svg?v=20260715.2';
const FAVICON_MAX_SOURCE_BYTES = 2 * 1024 * 1024;
const FAVICON_MAX_PNG_BYTES = 32 * 1024;
const ADMIN_SECTION_FOR_TAB = Object.freeze({
  trainings: 'training_workspace',
  staff: 'staff',
  settings: 'settings',
  share: 'share'
});

const state = {
  publicData: null,
  selectedTraining: null,
  selectedStaff: null,
  adminSession: '',
  adminData: null,
  records: [],
  strokes: [],
  drawing: false,
  currentStroke: null,
  demoAdminData: null,
  activePreview: null,
  adminActiveTab: 'trainings',
  adminLoadedAt: {},
  adminSectionPromises: {},
  adminSyncTimer: null,
  adminAuthenticating: false,
  settingsFaviconData: '',
  staffNamesComposing: false,
  activeExportTrainingId: '',
  activeUnsignedTrainingId: '',
  unsignedStatusDate: '',
  unsignedStatusMode: 'unsigned',
  unsignedStatusCache: new Map(),
  unsignedStatusRequest: null,
  unsignedRefreshTimer: null,
  pendingPreviewToken: ''
};

const demoData = {
  settings: {
    schoolName: '한빛고등학교',
    subtitle: '교직원 연수 참여 확인',
    notice: '연수 내용을 확인한 뒤 본인의 부서와 성명을 선택해 서명해 주세요.',
    brandColor: '#315c54',
    privacyPurpose: '교직원 연수 참여 여부 확인 및 서명등록부 작성',
    privacyItems: '부서, 성명, 서명 이미지, 서명 날짜와 시각',
    privacyRetention: '선택한 출력 파일을 보관한 뒤 시스템 원본을 삭제합니다.'
  },
  trainings: [
    { id: 'demo-training-1', title: '2026 개인정보 보호 연수', date: todaySeoul(), daily: false, startTime: '', endTime: '', active: true, sortOrder: 1 },
    { id: 'demo-training-2', title: '학교 안전교육', date: todaySeoul(), daily: true, startTime: '09:00', endTime: '18:00', active: true, sortOrder: 2 }
  ],
  staff: [
    { id: 'staff-1', department: '교무기획부', name: '김하늘', active: true, sortOrder: 1 },
    { id: 'staff-2', department: '교무기획부', name: '박서준', active: true, sortOrder: 2 },
    { id: 'staff-3', department: '교육연구부', name: '이도윤', active: true, sortOrder: 3 },
    { id: 'staff-4', department: '교육연구부', name: '최지우', active: true, sortOrder: 4 },
    { id: 'staff-5', department: '행정실', name: '정민서', active: true, sortOrder: 5 }
  ],
  privacyReady: true,
  serverDate: todaySeoul()
};

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function setHidden(element, hidden) {
  element?.classList.toggle('hidden', Boolean(hidden));
}

function showToast(message, timeout = 2600) {
  const toast = $('toast');
  toast.textContent = message;
  setHidden(toast, false);
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => setHidden(toast, true), timeout);
}

function showStatus(message = '', isError = true) {
  const banner = $('statusBanner');
  banner.textContent = message;
  banner.style.background = isError ? '' : 'var(--brand-soft)';
  banner.style.color = isError ? '' : 'var(--brand-dark)';
  setHidden(banner, !message);
}

function requestActionDialog({ title, message = '', confirmLabel = '확인', danger = false, fields = [] }) {
  return new Promise(resolve => {
    document.querySelector('.action-dialog-overlay')?.remove();
    const overlay = document.createElement('div');
    overlay.className = 'action-dialog-overlay';
    const fieldMarkup = fields.map(field => `
      <label>${escapeHtml(field.label)}
        <input name="${escapeHtml(field.name)}" value="${escapeHtml(field.value || '')}" maxlength="${Number(field.maxLength || 100)}" required>
      </label>`).join('');
    overlay.innerHTML = `
      <div class="action-dialog-card" role="dialog" aria-modal="true" aria-labelledby="actionDialogTitle">
        <form>
          <h3 id="actionDialogTitle">${escapeHtml(title)}</h3>
          ${message ? `<p>${escapeHtml(message)}</p>` : ''}
          ${fieldMarkup}
          <div class="button-row">
            <button class="button secondary" type="button" data-action="cancel">취소</button>
            <button class="button ${danger ? 'danger' : 'primary'}" type="submit">${escapeHtml(confirmLabel)}</button>
          </div>
        </form>
      </div>`;
    const form = overlay.querySelector('form');
    const finish = value => {
      overlay.remove();
      resolve(value);
    };
    form.addEventListener('submit', event => {
      event.preventDefault();
      const values = {};
      fields.forEach(field => { values[field.name] = form.elements[field.name].value.trim(); });
      finish(fields.length ? values : true);
    });
    overlay.querySelector('[data-action="cancel"]').addEventListener('click', () => finish(null));
    overlay.addEventListener('click', event => {
      if (event.target === overlay) finish(null);
    });
    overlay.addEventListener('keydown', event => {
      if (event.key === 'Escape') {
        event.preventDefault();
        event.stopPropagation();
        finish(null);
      }
    });
    ($('adminDialog').open ? $('adminDialog') : document.body).append(overlay);
    (overlay.querySelector('input') || overlay.querySelector('button[type="submit"]')).focus();
  });
}

async function requestConfirmation(options) {
  return Boolean(await requestActionDialog(options));
}

async function rpc(action, payload = {}, options = {}) {
  if (DEMO) return demoRpc(action, payload);
  if (!API_URL || API_URL.includes('__APPS_SCRIPT_WEB_APP_URL__')) {
    throw new Error('아직 Apps Script 주소가 연결되지 않았습니다. 관리자에게 알려 주세요.');
  }
  const body = { action, ...payload };
  if (options.admin !== false && state.adminSession && !body.sessionToken) body.sessionToken = state.adminSession;
  const response = await fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(body),
    redirect: 'follow',
    cache: 'no-store'
  });
  if (!response.ok) throw new Error(`서버 응답 오류 (${response.status})`);
  const result = await response.json();
  if (!result.ok) {
    const error = new Error(result.error?.message || '요청을 처리하지 못했습니다.');
    error.code = result.error?.code || 'UNKNOWN';
    error.details = result.error?.details;
    if (error.code === 'SESSION_EXPIRED' && options.admin !== false) queueMicrotask(() => handleExpiredAdminSession(error.message));
    throw error;
  }
  return result.data;
}

function demoRpc(action, payload) {
  if (!state.demoAdminData) {
    state.demoAdminData = {
      settings: { ...demoData.settings },
      trainings: demoData.trainings.map(item => ({ ...item })),
      staff: demoData.staff.map(item => ({ ...item })),
      exports: [],
      shareToken: shareToken,
      shareUrl: buildShareUrl(baseUrl, shareToken),
      stats: { staff: demoData.staff.length, trainings: demoData.trainings.length, signatures: 2 }
    };
  }
  if (action === 'get_public_data') return Promise.resolve(demoData);
  if (action === 'submit_signature') return Promise.resolve({ registeredAt: new Date().toISOString(), demo: true });
  if (action === 'get_setup_status') return Promise.resolve({ initialized: true, adminConfigured: true });
  if (action === 'admin_login') {
    if (payload.password !== 'demo-admin') return Promise.reject(Object.assign(new Error('데모 비밀번호는 demo-admin입니다.'), { code: 'BAD_PASSWORD' }));
    const adminData = payload.view === 'bootstrap'
      ? { settings: state.demoAdminData.settings, trainings: state.demoAdminData.trainings, staff: [], exports: [], shareToken: state.demoAdminData.shareToken, shareUrl: state.demoAdminData.shareUrl, loadedSections: ['settings', 'trainings', 'share'] }
      : state.demoAdminData;
    return Promise.resolve({ sessionToken: 'demo-session', expiresIn: 1800, adminData });
  }
  if (action === 'get_admin_section') {
    if (payload.section === 'training_workspace') return Promise.resolve({ section: 'training_workspace', trainings: state.demoAdminData.trainings, exports: state.demoAdminData.exports });
    if (payload.section === 'staff') return Promise.resolve({ section: 'staff', staff: state.demoAdminData.staff });
    if (payload.section === 'exports') return Promise.resolve({ section: 'exports', exports: state.demoAdminData.exports });
    if (payload.section === 'settings') return Promise.resolve({ section: 'settings', settings: state.demoAdminData.settings });
    if (payload.section === 'trainings') return Promise.resolve({ section: 'trainings', trainings: state.demoAdminData.trainings });
    if (payload.section === 'share') return Promise.resolve({ section: 'share', shareToken: state.demoAdminData.shareToken, shareUrl: state.demoAdminData.shareUrl });
  }
  if (action === 'get_admin_data') return Promise.resolve(state.demoAdminData);
  if (action === 'get_training_signature_status') {
    const training = state.demoAdminData.trainings.find(item => item.id === payload.trainingId);
    if (!training) return Promise.reject(Object.assign(new Error('연수를 찾을 수 없습니다.'), { code: 'NOT_FOUND' }));
    if (!training.daily && payload.date !== training.date) {
      return Promise.reject(Object.assign(new Error('이 연수는 등록된 날짜의 현황만 볼 수 있습니다.'), { code: 'INVALID_DATE' }));
    }
    const people = sortByRegistration(state.demoAdminData.staff.filter(person => person.active !== false)).map((person, index) => ({
      staffId: person.id,
      department: person.department,
      name: person.name,
      sortOrder: person.sortOrder,
      status: index < 2 ? 'signed' : 'unsigned',
      signTime: index < 2 ? (index ? '10:20' : '10:12') : ''
    }));
    const signedCount = people.filter(person => person.status === 'signed').length;
    return Promise.resolve({
      trainingId: training.id,
      date: payload.date,
      summary: {
        targetCount: people.length,
        signedCount,
        unsignedCount: people.length - signedCount,
        rate: people.length ? Math.round(signedCount / people.length * 1000) / 10 : 0,
        outsideRosterSignedCount: 0
      },
      people
    });
  }
  if (action === 'list_records') return Promise.resolve({ records: [
    { id: 'record-1', department: '교무기획부', name: '김하늘', signDate: todaySeoul(), signTime: '10:12:03', trainingId: 'demo-training-1' },
    { id: 'record-2', department: '교육연구부', name: '이도윤', signDate: todaySeoul(), signTime: '10:20:14', trainingId: 'demo-training-1' }
  ] });
  if (action === 'start_export') {
    const demoTraining = state.demoAdminData.trainings.find(item => item.id === payload.trainingId);
    const job = {
      jobId: `demo-export-${Date.now()}`, trainingId: payload.trainingId, trainingTitle: demoTraining?.title || '데모 연수', date: payload.date,
      sort: payload.sort, columns: payload.columns, showRate: payload.showRate, outputType: payload.outputType,
      status: 'queued', progress: 0, total: 2, hasPreview: false, hasPdf: false, hasXlsx: false, canPurge: false
    };
    state.demoAdminData.exports.unshift(job);
    return Promise.resolve(job);
  }
  if (action === 'continue_export') {
    const job = state.demoAdminData.exports.find(item => item.jobId === payload.jobId);
    if (!job) return Promise.reject(new Error('데모 출력 작업을 찾을 수 없습니다.'));
    Object.assign(job, { status: 'preview_ready', progress: 2, hasPreview: true });
    return Promise.resolve(job);
  }
  if (action === 'finalize_export' || action === 'record_print_opened') {
    return Promise.reject(Object.assign(new Error('데모에서는 파일 생성과 인쇄 기록을 저장하지 않습니다.'), { code: 'DEMO_READ_ONLY' }));
  }
  if (['logout', 'download_export_chunk'].includes(action)) return Promise.resolve({});
  return Promise.reject(Object.assign(new Error('데모에서는 변경 내용을 저장하지 않습니다.'), { code: 'DEMO_READ_ONLY' }));
}

function applySettings(settings) {
  const color = /^#[0-9a-f]{6}$/i.test(settings.brandColor || '') ? settings.brandColor : '#315c54';
  document.documentElement.style.setProperty('--brand', color);
  $('schoolName').textContent = settings.schoolName || config.APP_NAME || '학교 연수 전자서명';
  $('schoolSubtitle').textContent = settings.subtitle || '연수 참여 확인';
  $('schoolDate').textContent = formatKoreanHeaderDate(state.publicData?.serverDate || todaySeoul());
  document.title = `${settings.schoolName || '학교'} 연수 전자서명`;
  $('noticeText').textContent = settings.notice || '';
  setHidden($('noticePanel'), !settings.notice);
  const favicon = String(settings.faviconData || '');
  const faviconLink = $('faviconLink');
  faviconLink.type = favicon ? 'image/png' : 'image/svg+xml';
  faviconLink.href = favicon || DEFAULT_FAVICON_URL;
}

function showPanel(panelId) {
  document.body.dataset.panel = panelId;
  ['loadingPanel', 'invalidPanel', 'trainingPanel', 'personPanel', 'signaturePanel', 'successPanel']
    .forEach(id => setHidden($(id), id !== panelId));
  const step = panelId === 'personPanel' ? 2 : panelId === 'signaturePanel' ? 3 : 1;
  document.querySelectorAll('#stepper li').forEach(item => {
    const value = Number(item.dataset.step);
    item.classList.toggle('active', value === step);
    item.classList.toggle('done', value < step);
  });
}

function renderTrainings() {
  const list = $('trainingList');
  const trainings = state.publicData?.trainings || [];
  list.replaceChildren();
  setHidden($('noTraining'), trainings.length > 0);
  trainings.forEach(training => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'choice-card';
    button.dataset.trainingId = training.id;
    button.innerHTML = `<span><strong>${escapeHtml(training.title)}</strong><span class="choice-meta"><small>${escapeHtml(trainingTimeLabel(training))}</small></span></span><b aria-hidden="true">›</b>`;
    button.addEventListener('click', () => selectTraining(training.id));
    list.append(button);
  });
}

function selectTraining(trainingId) {
  state.selectedTraining = state.publicData.trainings.find(item => item.id === trainingId) || null;
  state.selectedStaff = null;
  if (!state.selectedTraining) return;
  $('selectedTrainingLabel').textContent = `${state.selectedTraining.title} · ${trainingTimeLabel(state.selectedTraining)}`;
  renderDepartments();
  showPanel('personPanel');
}

function renderDepartments() {
  const groups = groupStaffByDepartment(state.publicData?.staff || []);
  const select = $('departmentSelect');
  select.innerHTML = '<option value="">부서를 선택하세요</option>';
  for (const department of groups.keys()) {
    const option = document.createElement('option');
    option.value = department;
    option.textContent = department;
    select.append(option);
  }
  $('staffSelect').innerHTML = '<option value="">먼저 부서를 선택하세요</option>';
  $('staffSelect').disabled = true;
  $('goToSignature').disabled = true;
}

function renderStaffForDepartment() {
  const department = $('departmentSelect').value;
  const people = (state.publicData?.staff || []).filter(person => person.department === department && person.active !== false);
  const select = $('staffSelect');
  select.innerHTML = '<option value="">성명을 선택하세요</option>';
  people.forEach(person => {
    const option = document.createElement('option');
    option.value = person.id;
    option.textContent = person.name;
    select.append(option);
  });
  select.disabled = !department;
  state.selectedStaff = null;
  $('goToSignature').disabled = true;
}

function goToSignature() {
  const person = state.publicData.staff.find(item => item.id === $('staffSelect').value);
  if (!person) return;
  state.selectedStaff = person;
  $('signerSummary').textContent = `${state.selectedTraining.title} · ${person.department} ${person.name}`;
  clearSignature();
  showPanel('signaturePanel');
  requestAnimationFrame(resizeCanvas);
}

function clearSignature() {
  state.strokes = [];
  state.currentStroke = null;
  drawSignature();
}

function undoSignature() {
  state.strokes.pop();
  drawSignature();
}

function resizeCanvas() {
  const canvas = $('signatureCanvas');
  if (!canvas) return;
  const rect = canvas.getBoundingClientRect();
  const ratio = Math.min(window.devicePixelRatio || 1, 2);
  canvas.width = Math.max(1, Math.round(rect.width * ratio));
  canvas.height = Math.max(1, Math.round(rect.height * ratio));
  drawSignature();
}

function drawSignature(targetCanvas = $('signatureCanvas'), width = targetCanvas.width, height = targetCanvas.height) {
  if (!targetCanvas) return;
  const context = targetCanvas.getContext('2d');
  context.clearRect(0, 0, width, height);
  context.lineCap = 'round';
  context.lineJoin = 'round';
  context.strokeStyle = '#13201d';
  context.lineWidth = Math.max(2, width / 250);
  for (const stroke of state.strokes) {
    if (!stroke.length) continue;
    context.beginPath();
    context.moveTo(stroke[0].x * width, stroke[0].y * height);
    for (const point of stroke.slice(1)) context.lineTo(point.x * width, point.y * height);
    if (stroke.length === 1) context.lineTo(stroke[0].x * width + .01, stroke[0].y * height + .01);
    context.stroke();
  }
}

function canvasPoint(event) {
  const rect = $('signatureCanvas').getBoundingClientRect();
  return {
    x: Math.min(1, Math.max(0, (event.clientX - rect.left) / rect.width)),
    y: Math.min(1, Math.max(0, (event.clientY - rect.top) / rect.height))
  };
}

function startDrawing(event) {
  event.preventDefault();
  state.drawing = true;
  state.currentStroke = [canvasPoint(event)];
  state.strokes.push(state.currentStroke);
  $('signatureCanvas').setPointerCapture?.(event.pointerId);
  drawSignature();
}

function continueDrawing(event) {
  if (!state.drawing || !state.currentStroke) return;
  event.preventDefault();
  state.currentStroke.push(canvasPoint(event));
  drawSignature();
}

function stopDrawing() {
  state.drawing = false;
  state.currentStroke = null;
}

function signatureDataUrl() {
  const canvas = document.createElement('canvas');
  canvas.width = 600;
  canvas.height = 220;
  const context = canvas.getContext('2d');
  context.fillStyle = '#ffffff';
  context.fillRect(0, 0, canvas.width, canvas.height);
  drawSignature(canvas, canvas.width, canvas.height);
  return canvas.toDataURL('image/png');
}

async function submitSignature() {
  if (!state.selectedTraining || !state.selectedStaff) return;
  if (state.strokes.length === 0 || !state.strokes.some(stroke => stroke.length > 1)) {
    showToast('서명을 먼저 작성해 주세요.');
    return;
  }
  const date = state.publicData.serverDate || todaySeoul();
  const duplicateKey = localDuplicateKey(state.selectedTraining.id, state.selectedStaff.id, date);
  if (localStorage.getItem(duplicateKey)) {
    const proceed = await requestConfirmation({
      title: '이미 서명한 기록이 있습니다',
      message: '이 기기에서 같은 연수에 서명한 기록이 있습니다. 서버에서 중복 여부를 다시 확인할까요?',
      confirmLabel: '서버에서 확인'
    });
    if (!proceed) return;
  }
  const button = $('submitSignature');
  button.disabled = true;
  button.textContent = '등록 중…';
  try {
    const result = await rpc('submit_signature', {
      shareToken,
      trainingId: state.selectedTraining.id,
      staffId: state.selectedStaff.id,
      signatureData: signatureDataUrl()
    }, { admin: false });
    localStorage.setItem(duplicateKey, result.registeredAt || new Date().toISOString());
    $('successMessage').textContent = `${state.selectedStaff.department} ${state.selectedStaff.name}님의 참여 확인이 완료되었습니다.${result.demo ? ' 데모이므로 실제 저장되지는 않았습니다.' : ''}`;
    showPanel('successPanel');
  } catch (error) {
    showToast(error.message, 4200);
  } finally {
    button.disabled = false;
    button.textContent = '서명 제출';
  }
}

function renderPrivacy() {
  const settings = state.publicData?.settings || state.adminData?.settings || {};
  const fields = [
    ['수집 목적', settings.privacyPurpose],
    ['수집 항목', settings.privacyItems],
    ['보관·삭제', settings.privacyRetention],
    ['이용 제한', '연수 참여 확인용이며, 본인 인증이 필요한 법적 전자서명에는 사용할 수 없습니다.']
  ];
  $('privacyDetails').innerHTML = fields.map(([title, value]) => `<dt>${escapeHtml(title)}</dt><dd>${escapeHtml(value || '관리자가 아직 입력하지 않았습니다.')}</dd>`).join('');
  $('privacyDialog').showModal();
}

function renderQr(container, url) {
  container.replaceChildren();
  if (!url || typeof window.qrcode !== 'function') {
    container.textContent = 'QR 생성 기능을 불러오지 못했습니다.';
    return;
  }
  const qr = window.qrcode(0, 'M');
  qr.addData(url, 'Byte');
  qr.make();
  container.innerHTML = qr.createSvgTag({ cellSize: 6, margin: 4, scalable: true });
}

function openShareDialog() {
  const url = DEMO ? buildShareUrl(baseUrl, shareToken) : location.href;
  $('shareUrl').value = url;
  renderQr($('qrCode'), url);
  $('shareDialog').showModal();
}

async function copyText(value, message = '복사했습니다.') {
  try {
    await navigator.clipboard.writeText(value);
    showToast(message);
  } catch {
    const input = document.createElement('textarea');
    input.value = value;
    document.body.append(input);
    input.select();
    document.execCommand('copy');
    input.remove();
    showToast(message);
  }
}

function emptyAdminData() {
  return {
    settings: {}, trainings: [], staff: [], exports: [], shareToken: '', shareUrl: '', loadedSections: []
  };
}

function setAdminSectionLoading(loading, message = '관리자 정보를 불러오는 중입니다.') {
  const box = $('adminSectionLoading');
  box.querySelector('p').textContent = message;
  setHidden(box, !loading);
}

function setAdminAuthenticating(authenticating) {
  state.adminAuthenticating = authenticating;
  $('adminDialog').classList.toggle('admin-authenticating', authenticating);
  setAdminSectionLoading(authenticating, '비밀번호를 확인하고 연수 목록을 불러오는 중입니다.');
}

function markAdminSectionLoaded(section, timestamp = Date.now()) {
  state.adminLoadedAt[section] = timestamp;
  if (!state.adminData.loadedSections.includes(section)) state.adminData.loadedSections.push(section);
}

function mergeAdminSection(data) {
  if (!state.adminData || !data) return;
  ['settings', 'trainings', 'staff', 'exports', 'shareToken', 'shareUrl'].forEach(key => {
    if (Object.hasOwn(data, key)) state.adminData[key] = data[key];
  });
  markAdminSectionLoaded(data.section);
}

function renderAdminSection(section) {
  if (section === 'trainings' || section === 'training_workspace') {
    renderTrainingAdmin();
    populateTrainingSelects();
  }
  if (section === 'staff') renderStaffAdmin();
  if (section === 'settings') fillSettingsForm();
  if (section === 'share') renderShareAdmin();
}

async function loadAdminSection(section, { force = false, background = false } = {}) {
  if (!section || !state.adminSession || state.adminAuthenticating) return;
  const loadedAt = state.adminLoadedAt[section] || 0;
  if (!force && loadedAt && Date.now() - loadedAt < ADMIN_SYNC_MS) return;
  if (state.adminSectionPromises[section]) return state.adminSectionPromises[section];
  if (!background) setAdminSectionLoading(true, '현재 화면을 불러오는 중입니다.');
  const request = rpc('get_admin_section', { section })
    .then(data => {
      mergeAdminSection(data);
      renderAdminSection(section);
      return data;
    })
    .catch(error => {
      if (error.code === 'SESSION_EXPIRED') handleExpiredAdminSession(error.message);
      else if (!background) showToast(error.message, 4200);
      throw error;
    })
    .finally(() => {
      delete state.adminSectionPromises[section];
      if (!background && !state.adminAuthenticating) setAdminSectionLoading(false);
    });
  state.adminSectionPromises[section] = request;
  return request;
}

function startAdminBackgroundSync() {
  clearInterval(state.adminSyncTimer);
  state.adminSyncTimer = setInterval(() => {
    if (!$('adminDialog').open || !state.adminSession || state.adminAuthenticating) return;
    const section = ADMIN_SECTION_FOR_TAB[state.adminActiveTab];
    if (section === 'settings' && $('settingsForm').contains(document.activeElement)) return;
    if (section) loadAdminSection(section, { force: true, background: true }).catch(() => {});
  }, ADMIN_SYNC_MS);
}

function handleExpiredAdminSession(message = '관리자 로그인이 만료되었습니다. 다시 로그인해 주세요.') {
  clearInterval(state.adminSyncTimer);
  stopUnsignedStatusRefresh();
  state.adminSession = '';
  state.adminData = null;
  state.adminLoadedAt = {};
  state.activeExportTrainingId = '';
  state.activeUnsignedTrainingId = '';
  state.unsignedStatusDate = '';
  state.unsignedStatusCache.clear();
  state.unsignedStatusRequest = null;
  if ($('adminDialog').open) $('adminDialog').close();
  showToast(message, 4200);
}

function setAdminLoginMode(setupRequired) {
  $('adminLoginTitle').textContent = setupRequired ? '관리자 첫 설정' : '관리자 로그인';
  $('adminLoginSubmit').textContent = setupRequired ? '비밀번호 설정' : '로그인';
  setHidden($('setupCodeLabel'), !setupRequired);
  setHidden($('adminPasswordConfirmLabel'), !setupRequired);
  $('setupCode').required = setupRequired;
  $('adminPasswordConfirm').required = setupRequired;
  $('adminPassword').autocomplete = setupRequired ? 'new-password' : 'current-password';
  $('adminLoginForm').dataset.setupRequired = setupRequired ? '1' : '0';
  $('adminLoginSubmit').disabled = false;
}

function openAdminLogin() {
  const errorBox = $('adminLoginError');
  $('adminLoginError').textContent = '';
  setHidden($('adminLoginError'), true);
  $('adminPassword').value = '';
  $('adminPasswordConfirm').value = '';
  $('setupCode').value = '';
  setAdminLoginMode(false);
  if (!$('adminLoginDialog').open) $('adminLoginDialog').showModal();
  $('adminPassword').focus();
  rpc('get_setup_status', {}, { admin: false }).then(status => {
    if (!state.adminSession) setAdminLoginMode(!status.adminConfigured);
  }).catch(error => {
    if ($('adminLoginDialog').open) {
      errorBox.textContent = error.message;
      setHidden(errorBox, false);
    }
  });
}

async function handleAdminLogin(event) {
  event.preventDefault();
  const setupRequired = event.currentTarget.dataset.setupRequired === '1';
  const password = $('adminPassword').value;
  const errorBox = $('adminLoginError');
  if (setupRequired) {
    if (!isValidAdminPassword(password)) {
      errorBox.textContent = '비밀번호는 숫자 4자리 또는 문자와 숫자를 포함한 10자 이상으로 설정해 주세요.';
      setHidden(errorBox, false);
      return;
    }
    if (password !== $('adminPasswordConfirm').value) {
      errorBox.textContent = '비밀번호 확인이 일치하지 않습니다.';
      setHidden(errorBox, false);
      return;
    }
  }
  state.adminSession = '';
  state.adminData = emptyAdminData();
  state.adminLoadedAt = {};
  state.adminActiveTab = 'trainings';
  switchAdminTab('trainings', { sync: false });
  renderTrainingAdmin();
  populateTrainingSelects();
  $('adminLoginDialog').close();
  if (!$('adminDialog').open) $('adminDialog').showModal();
  setAdminAuthenticating(true);
  try {
    let data;
    if (setupRequired) {
      data = await rpc('complete_setup', {
        setupCode: $('setupCode').value.trim(),
        password,
        frontendUrl: baseUrl,
        view: 'bootstrap'
      }, { admin: false });
    } else {
      data = await rpc('admin_login', { password, view: 'bootstrap' }, { admin: false });
    }
    state.adminSession = data.sessionToken;
    state.adminData = Object.assign(emptyAdminData(), data.adminData || {});
    state.adminData.loadedSections = Array.isArray(data.adminData?.loadedSections) ? [...data.adminData.loadedSections] : [];
    state.adminData.loadedSections.forEach(section => { state.adminLoadedAt[section] = Date.now(); });
    renderAdmin();
    setAdminAuthenticating(false);
    startAdminBackgroundSync();
    loadAdminSection('training_workspace', { force: true, background: true }).catch(() => {});
  } catch (error) {
    setAdminAuthenticating(false);
    state.adminSession = '';
    state.adminData = null;
    state.adminLoadedAt = {};
    if ($('adminDialog').open) $('adminDialog').close();
    errorBox.textContent = error.message;
    setHidden(errorBox, false);
    if (!$('adminLoginDialog').open) $('adminLoginDialog').showModal();
    $('adminPassword').select();
  }
}

function switchAdminTab(tab, { sync = true } = {}) {
  state.adminActiveTab = tab;
  if (tab === 'trainings' && state.activeUnsignedTrainingId) startUnsignedStatusRefresh();
  else if (tab !== 'trainings') stopUnsignedStatusRefresh();
  document.querySelectorAll('#adminTabs button').forEach(button => button.classList.toggle('active', button.dataset.adminTab === tab));
  document.querySelectorAll('[data-admin-panel]').forEach(panel => panel.classList.toggle('active', panel.dataset.adminPanel === tab));
  if (!sync) return;
  const section = ADMIN_SECTION_FOR_TAB[tab];
  if (!section) return;
  const firstLoad = !state.adminLoadedAt[section];
  loadAdminSection(section, { background: !firstLoad }).catch(() => {});
}

function trainingMeta(training) {
  if (training.pending) return '서버에 저장하는 중…';
  return [trainingTimeLabel(training), training.active ? '활성' : '비활성'].join(' · ');
}

function upsertAdminItem(list, item, key = 'id') {
  const next = [...(list || [])];
  const index = next.findIndex(current => current[key] === item[key]);
  if (index >= 0) next[index] = item;
  else next.push(item);
  return next;
}

function sortByRegistration(items) {
  return [...(items || [])].sort((a, b) => Number(a.sortOrder || 0) - Number(b.sortOrder || 0));
}

function upsertExportJob(job) {
  if (!job || !state.adminData) return;
  state.adminData.exports = upsertAdminItem(state.adminData.exports, job, 'jobId')
    .sort((a, b) => String(b.createdAt || b.updatedAt || '').localeCompare(String(a.createdAt || a.updatedAt || '')));
  markAdminSectionLoaded('training_workspace');
  renderExportJobs();
  renderOrphanExportJobs();
}

function renderTrainingAdmin() {
  const container = $('trainingAdminList');
  const exportPanel = $('trainingExportPanel');
  const exportPanelHome = $('trainingExportPanelHome');
  const unsignedPanel = $('trainingUnsignedPanel');
  const unsignedPanelHome = $('trainingUnsignedPanelHome');
  if (exportPanel && exportPanelHome && exportPanel.parentElement !== exportPanelHome) exportPanelHome.append(exportPanel);
  if (unsignedPanel && unsignedPanelHome && unsignedPanel.parentElement !== unsignedPanelHome) unsignedPanelHome.append(unsignedPanel);
  const trainings = state.adminData?.trainings || [];
  container.innerHTML = trainings.length ? trainings.map((training, index) => `
    <div class="training-admin-item" data-training-id="${escapeHtml(training.id)}">
      <div class="admin-row${training.pending ? ' pending-row' : ''}">
        <div class="admin-row-main"><strong>${escapeHtml(training.title)}</strong><small>${escapeHtml(trainingMeta(training))}</small></div>
        <div class="row-actions">
          <button data-action="move-up" ${training.pending || index === 0 ? 'disabled' : ''}>위</button>
          <button data-action="move-down" ${training.pending || index === trainings.length - 1 ? 'disabled' : ''}>아래</button>
          <button data-action="edit-training" ${training.pending ? 'disabled' : ''}>수정</button>
          <button data-action="toggle-unsigned" aria-controls="trainingUnsignedPanel" aria-expanded="${state.activeUnsignedTrainingId === training.id ? 'true' : 'false'}" ${training.pending ? 'disabled' : ''}>${state.activeUnsignedTrainingId === training.id ? '현황 접기' : '미서명 현황'}</button>
          <button data-action="toggle-export" aria-controls="trainingExportPanel" aria-expanded="${state.activeExportTrainingId === training.id ? 'true' : 'false'}" ${training.pending ? 'disabled' : ''}>${state.activeExportTrainingId === training.id ? '출력 접기' : '출력'}</button>
          <button data-action="delete-training" class="danger" ${training.pending ? 'disabled' : ''}>삭제</button>
        </div>
      </div>
      <div class="training-tool-slot"></div>
    </div>`).join('') : '<div class="empty-state">등록된 연수가 없습니다.</div>';
  const selectedExport = trainings.find(training => training.id === state.activeExportTrainingId && !training.pending);
  if (selectedExport && exportPanel) {
    const preserveForm = exportPanel.dataset.trainingId === selectedExport.id;
    container.querySelector(`[data-training-id="${CSS.escape(selectedExport.id)}"] .training-tool-slot`)?.append(exportPanel);
    exportPanel.dataset.trainingId = selectedExport.id;
    setHidden(exportPanel, false);
    renderTrainingExportPanel(selectedExport, { preserveForm });
  } else if (exportPanel) {
    state.activeExportTrainingId = '';
    exportPanel.dataset.trainingId = '';
    setHidden(exportPanel, true);
  }
  const selectedUnsigned = trainings.find(training => training.id === state.activeUnsignedTrainingId && !training.pending);
  if (selectedUnsigned && unsignedPanel) {
    const preserveDate = unsignedPanel.dataset.trainingId === selectedUnsigned.id;
    container.querySelector(`[data-training-id="${CSS.escape(selectedUnsigned.id)}"] .training-tool-slot`)?.append(unsignedPanel);
    unsignedPanel.dataset.trainingId = selectedUnsigned.id;
    setHidden(unsignedPanel, false);
    renderTrainingUnsignedPanel(selectedUnsigned, { preserveDate });
  } else if (unsignedPanel) {
    state.activeUnsignedTrainingId = '';
    unsignedPanel.dataset.trainingId = '';
    setHidden(unsignedPanel, true);
    stopUnsignedStatusRefresh();
  }
  renderOrphanExportJobs();
}

function renderTrainingExportPanel(training, { preserveForm = false } = {}) {
  $('trainingExportTitle').textContent = training.title;
  $('trainingExportDateHint').textContent = training.daily
    ? '매일 연수는 출력할 날짜를 선택할 수 있습니다.'
    : `${formatKoreanDate(training.date)} 서명 기록을 출력합니다.`;
  if (!preserveForm) restoreExportSettings(training);
  if (!training.daily) {
    $('exportDate').value = training.date;
    $('exportDate').disabled = true;
  } else {
    $('exportDate').disabled = false;
  }
  renderExportJobs();
}

function collapseTrainingExport() {
  state.activeExportTrainingId = '';
  renderTrainingAdmin();
}

function toggleTrainingExport(training) {
  if (state.activeExportTrainingId === training.id) return collapseTrainingExport();
  state.activeUnsignedTrainingId = '';
  stopUnsignedStatusRefresh();
  state.activeExportTrainingId = training.id;
  renderTrainingAdmin();
  if (!state.adminLoadedAt.training_workspace) {
    loadAdminSection('training_workspace', { force: true, background: true }).catch(() => {});
  }
  $('trainingExportPanel').scrollIntoView({ block: 'nearest', behavior: 'smooth' });
}

function unsignedStatusCacheKey(trainingId, date) {
  return `${trainingId}:${date}`;
}

function isSignedStatus(person) {
  return ['signed', 'complete', 'completed'].includes(String(person?.status || '').toLowerCase());
}

function formatStatusTime(value) {
  const match = String(value || '').match(/(?:^|T|\s)(\d{2}):(\d{2})/);
  return match ? `${match[1]}:${match[2]}` : '';
}

function setUnsignedStatusLoading(loading, message = '미서명 현황을 확인하는 중입니다.') {
  const box = $('unsignedStatusLoading');
  box.querySelector('p').textContent = message;
  setHidden(box, !loading);
  $('refreshUnsignedStatus').disabled = loading;
}

function renderUnsignedStatusContent(data = null) {
  const summary = data?.summary || {};
  $('unsignedTargetCount').textContent = Number(summary.targetCount || 0).toLocaleString('ko-KR');
  $('unsignedSignedCount').textContent = Number(summary.signedCount || 0).toLocaleString('ko-KR');
  $('unsignedMissingCount').textContent = Number(summary.unsignedCount || 0).toLocaleString('ko-KR');
  const rate = Number(summary.rate || 0);
  $('unsignedRate').textContent = `${Number.isFinite(rate) ? rate.toLocaleString('ko-KR', { maximumFractionDigits: 1 }) : 0}%`;
  const outsideCount = Number(summary.outsideRosterSignedCount || 0);
  $('unsignedOutsideRoster').textContent = outsideCount ? `현재 명단 외 서명 ${outsideCount}건` : '';
  setHidden($('unsignedOutsideRoster'), !outsideCount);

  document.querySelectorAll('[data-unsigned-view]').forEach(button => {
    const active = button.dataset.unsignedView === state.unsignedStatusMode;
    button.classList.toggle('active', active);
    button.setAttribute('aria-pressed', active ? 'true' : 'false');
  });

  const people = sortByRegistration(data?.people || []);
  const visiblePeople = state.unsignedStatusMode === 'unsigned'
    ? people.filter(person => !isSignedStatus(person))
    : people;
  const container = $('unsignedStatusList');
  if (!people.length) {
    container.innerHTML = '<div class="empty-state">현재 등록된 구성원이 없습니다.</div>';
    return;
  }
  if (!visiblePeople.length) {
    container.innerHTML = '<div class="unsigned-complete-state"><strong>현재 구성원이 모두 서명했습니다.</strong><p>전체 현황으로 바꾸면 서명 시각을 확인할 수 있습니다.</p></div>';
    return;
  }
  const rows = visiblePeople.map(person => {
    const signed = isSignedStatus(person);
    const time = signed ? formatStatusTime(person.signTime) : '';
    return `<tr>
      <td>${escapeHtml(person.department || '')}</td>
      <td>${escapeHtml(person.name || '')}</td>
      <td><span class="signature-status ${signed ? 'signed' : 'unsigned'}">${signed ? '완료' : '미서명'}</span>${time ? `<time datetime="${escapeHtml(person.signTime)}">${escapeHtml(time)}</time>` : ''}</td>
    </tr>`;
  }).join('');
  container.innerHTML = `<div class="unsigned-status-table-wrap">
    <table class="unsigned-status-table">
      <thead><tr><th scope="col">부서</th><th scope="col">성명</th><th scope="col">상태</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>
  </div>`;
}

function renderTrainingUnsignedPanel(training, { preserveDate = false } = {}) {
  $('trainingUnsignedTitle').textContent = training.title;
  $('trainingUnsignedDateHint').textContent = training.daily
    ? '매일 연수는 확인할 날짜를 선택할 수 있습니다.'
    : `${formatKoreanDate(training.date)} 서명 현황입니다.`;
  if (!preserveDate || !/^\d{4}-\d{2}-\d{2}$/.test(state.unsignedStatusDate)) {
    state.unsignedStatusDate = training.daily ? todaySeoul() : training.date;
  }
  if (!training.daily) state.unsignedStatusDate = training.date;
  $('unsignedStatusDate').value = state.unsignedStatusDate;
  $('unsignedStatusDate').disabled = !training.daily;
  const entry = state.unsignedStatusCache.get(unsignedStatusCacheKey(training.id, state.unsignedStatusDate));
  if (entry?.data) {
    renderUnsignedStatusContent(entry.data);
    setUnsignedStatusLoading(false);
  } else {
    renderUnsignedStatusContent();
    setUnsignedStatusLoading(true);
  }
}

async function fetchUnsignedStatus({ force = false } = {}) {
  const training = state.adminData?.trainings.find(item => item.id === state.activeUnsignedTrainingId);
  if (!training) return;
  const date = training.daily ? $('unsignedStatusDate').value : training.date;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return;
  state.unsignedStatusDate = date;
  const key = unsignedStatusCacheKey(training.id, date);
  const cached = state.unsignedStatusCache.get(key);
  if (!force && cached && Date.now() - cached.fetchedAt < ADMIN_SYNC_MS) {
    renderUnsignedStatusContent(cached.data);
    setUnsignedStatusLoading(false);
    return cached.data;
  }
  if (state.unsignedStatusRequest?.key === key) return state.unsignedStatusRequest.promise;
  const requestSession = state.adminSession;
  setUnsignedStatusLoading(true, force ? '최신 서명 현황을 확인하는 중입니다.' : '미서명 현황을 확인하는 중입니다.');
  const promise = rpc('get_training_signature_status', { trainingId: training.id, date })
    .then(data => {
      if (!requestSession || state.adminSession !== requestSession) return null;
      state.unsignedStatusCache.set(key, { data, fetchedAt: Date.now() });
      if (state.activeUnsignedTrainingId === training.id && state.unsignedStatusDate === date) {
        renderUnsignedStatusContent(data);
        setUnsignedStatusLoading(false);
      }
      return data;
    })
    .catch(error => {
      if (state.activeUnsignedTrainingId === training.id && state.unsignedStatusDate === date) {
        setUnsignedStatusLoading(false);
        $('unsignedStatusList').innerHTML = `<div class="empty-state error-state">${escapeHtml(error.message)}</div>`;
      }
      throw error;
    })
    .finally(() => {
      if (state.unsignedStatusRequest?.promise === promise) state.unsignedStatusRequest = null;
    });
  state.unsignedStatusRequest = { key, promise };
  return promise;
}

function stopUnsignedStatusRefresh() {
  clearInterval(state.unsignedRefreshTimer);
  state.unsignedRefreshTimer = null;
}

function startUnsignedStatusRefresh() {
  stopUnsignedStatusRefresh();
  state.unsignedRefreshTimer = setInterval(() => {
    if (!$('adminDialog').open || !state.adminSession || !state.activeUnsignedTrainingId) return;
    fetchUnsignedStatus({ force: true }).catch(() => {});
  }, ADMIN_SYNC_MS);
}

function collapseTrainingUnsigned() {
  state.activeUnsignedTrainingId = '';
  stopUnsignedStatusRefresh();
  renderTrainingAdmin();
}

function toggleTrainingUnsigned(training) {
  if (state.activeUnsignedTrainingId === training.id) return collapseTrainingUnsigned();
  state.activeExportTrainingId = '';
  state.activeUnsignedTrainingId = training.id;
  state.unsignedStatusDate = training.daily ? todaySeoul() : training.date;
  state.unsignedStatusMode = 'unsigned';
  renderTrainingAdmin();
  fetchUnsignedStatus().catch(() => {});
  startUnsignedStatusRefresh();
  $('trainingUnsignedPanel').scrollIntoView({ block: 'nearest', behavior: 'smooth' });
}

function openTrainingForm(training = null) {
  $('trainingId').value = training?.id || '';
  $('trainingTitle').value = training?.title || '';
  $('trainingDate').value = training?.date && training.date !== '매일' ? training.date : todaySeoul();
  $('trainingDaily').checked = Boolean(training?.daily);
  $('trainingStart').value = training?.startTime || '';
  $('trainingEnd').value = training?.endTime || '';
  $('trainingActive').checked = training ? training.active !== false : true;
  setHidden($('trainingFormError'), true);
  setHidden($('trainingForm'), false);
  $('trainingTitle').focus();
}

async function saveTraining(event) {
  event.preventDefault();
  const training = {
    id: $('trainingId').value || undefined,
    title: $('trainingTitle').value.trim(),
    date: $('trainingDate').value,
    daily: $('trainingDaily').checked,
    startTime: $('trainingStart').value,
    endTime: $('trainingEnd').value,
    active: $('trainingActive').checked
  };
  const errors = validateTraining(training);
  if (training.active && !isPrivacyReady(state.adminData.settings)) errors.push('기관 설정에서 개인정보 안내를 모두 입력해야 활성화할 수 있습니다.');
  if (errors.length) {
    $('trainingFormError').textContent = errors.join(' ');
    setHidden($('trainingFormError'), false);
    return;
  }
  const previousTrainings = state.adminData.trainings.map(item => ({ ...item }));
  const existing = training.id ? state.adminData.trainings.find(item => item.id === training.id) : null;
  const pendingId = training.id || `pending-${Date.now()}`;
  const pendingTraining = {
    ...(existing || {}), ...training, id: pendingId,
    sortOrder: existing?.sortOrder || Math.max(0, ...state.adminData.trainings.map(item => Number(item.sortOrder || 0))) + 1,
    pending: true
  };
  state.adminData.trainings = sortByRegistration(upsertAdminItem(state.adminData.trainings, pendingTraining));
  renderAdminSection('trainings');
  setHidden($('trainingForm'), true);
  showToast('연수를 저장하는 중입니다…', 5000);
  try {
    const result = await rpc('save_training', { training });
    state.adminData.trainings = state.adminData.trainings.filter(item => item.id !== pendingId);
    state.adminData.trainings = sortByRegistration(upsertAdminItem(state.adminData.trainings, result.training));
    markAdminSectionLoaded('trainings');
    renderAdminSection('trainings');
    showToast('연수를 저장했습니다.');
  } catch (error) {
    if (state.adminData) {
      state.adminData.trainings = previousTrainings;
      renderAdminSection('trainings');
      openTrainingForm(existing || training);
    }
    showToast(error.message, 4200);
  }
}

async function handleTrainingListClick(event) {
  const button = event.target.closest('button[data-action]');
  const row = event.target.closest('[data-training-id]');
  if (!button || !row) return;
  const training = state.adminData.trainings.find(item => item.id === row.dataset.trainingId);
  if (!training || training.pending) return;
  const previousTrainings = state.adminData.trainings.map(item => ({ ...item }));
  try {
    if (button.dataset.action === 'edit-training') return openTrainingForm(training);
    if (button.dataset.action === 'toggle-unsigned') return toggleTrainingUnsigned(training);
    if (button.dataset.action === 'toggle-export') return toggleTrainingExport(training);
    if (button.dataset.action === 'delete-training') {
      const confirmed = await requestConfirmation({
        title: '연수를 삭제할까요?',
        message: `'${training.title}' 연수를 삭제합니다. 기존 서명 기록은 남습니다.`,
        confirmLabel: '연수 삭제',
        danger: true
      });
      if (!confirmed) return;
      state.adminData.trainings = state.adminData.trainings.filter(item => item.id !== training.id);
      renderAdminSection('trainings');
      const result = await rpc('delete_training', { trainingId: training.id });
      state.adminData.trainings = state.adminData.trainings.filter(item => item.id !== (result.deletedId || training.id));
    } else {
      const direction = button.dataset.action === 'move-up' ? -1 : 1;
      const currentIndex = state.adminData.trainings.findIndex(item => item.id === training.id);
      const targetIndex = currentIndex + direction;
      if (targetIndex < 0 || targetIndex >= state.adminData.trainings.length) return;
      const optimistic = [...state.adminData.trainings];
      [optimistic[currentIndex], optimistic[targetIndex]] = [optimistic[targetIndex], optimistic[currentIndex]];
      state.adminData.trainings = optimistic.map((item, index) => ({ ...item, sortOrder: index + 1 }));
      renderAdminSection('trainings');
      const result = await rpc('move_training', { trainingId: training.id, direction: button.dataset.action === 'move-up' ? 'up' : 'down' });
      if (Array.isArray(result.trainings)) state.adminData.trainings = result.trainings;
    }
    markAdminSectionLoaded('trainings');
    renderAdminSection('trainings');
  } catch (error) {
    if (state.adminData) {
      state.adminData.trainings = previousTrainings;
      renderAdminSection('trainings');
    }
    showToast(error.message, 4200);
  }
}

function renderStaffAdmin() {
  const staff = state.adminData?.staff || [];
  const groups = groupStaffByDepartment(staff);
  const container = $('staffAdminList');
  const departmentOptions = [...groups.keys()];
  $('oldDepartment').innerHTML = '<option value="">기존 부서 선택</option>' + departmentOptions.map(value => `<option>${escapeHtml(value)}</option>`).join('');
  container.innerHTML = departmentOptions.length ? departmentOptions.map(department => `
    <div class="subcard"><h4>${escapeHtml(department)} <small>${groups.get(department).length}명</small></h4>
    ${groups.get(department).map(person => `<div class="admin-row" data-staff-id="${escapeHtml(person.id)}"><div class="admin-row-main"><strong>${escapeHtml(person.name)}</strong></div><div class="row-actions"><button data-action="edit-staff">수정</button><button data-action="delete-staff" class="danger">삭제</button></div></div>`).join('')}</div>`).join('') : '<div class="empty-state">등록된 구성원이 없습니다.</div>';
}

function normalizeStaffNamesField() {
  const input = $('staffNames');
  const value = input.value;
  const start = input.selectionStart ?? value.length;
  const end = input.selectionEnd ?? start;
  const normalized = normalizeNameEntryText(value);
  if (normalized === value) return;
  const nextStart = normalizeNameEntryText(value.slice(0, start)).length;
  const nextEnd = normalizeNameEntryText(value.slice(0, end)).length;
  input.value = normalized;
  input.setSelectionRange(nextStart, nextEnd);
}

async function addStaff(event) {
  event.preventDefault();
  const department = $('staffDepartment').value.trim();
  const names = splitNames($('staffNames').value);
  if (!department || !names.length) return;
  try {
    const result = await rpc('save_staff_batch', { people: names.map(name => ({ department, name })) });
    $('staffNames').value = '';
    state.adminData.staff = sortByRegistration([...(state.adminData.staff || []), ...(result.people || [])]);
    markAdminSectionLoaded('staff');
    renderStaffAdmin();
    showToast(`${result.added}명 등록, ${result.skipped}명 건너뜀`);
  } catch (error) { showToast(error.message, 4200); }
}

async function handleStaffListClick(event) {
  const button = event.target.closest('button[data-action]');
  const row = event.target.closest('[data-staff-id]');
  if (!button || !row) return;
  const person = state.adminData.staff.find(item => item.id === row.dataset.staffId);
  if (!person) return;
  try {
    if (button.dataset.action === 'edit-staff') {
      const edited = await requestActionDialog({
        title: '구성원 정보 수정',
        confirmLabel: '수정 저장',
        fields: [
          { name: 'department', label: '부서', value: person.department, maxLength: 50 },
          { name: 'name', label: '성명', value: person.name, maxLength: 50 }
        ]
      });
      if (!edited) return;
      const result = await rpc('update_staff', { person: { id: person.id, department: edited.department, name: edited.name } });
      state.adminData.staff = sortByRegistration(upsertAdminItem(state.adminData.staff, result.person || { ...person, ...edited }));
    } else {
      const confirmed = await requestConfirmation({
        title: '구성원을 삭제할까요?',
        message: `${person.department} ${person.name} 구성원을 삭제합니다. 기존 서명 기록은 남습니다.`,
        confirmLabel: '구성원 삭제',
        danger: true
      });
      if (!confirmed) return;
      const result = await rpc('delete_staff', { staffId: person.id });
      state.adminData.staff = state.adminData.staff.filter(item => item.id !== (result.deletedId || person.id));
    }
    markAdminSectionLoaded('staff');
    renderStaffAdmin();
  } catch (error) { showToast(error.message, 4200); }
}

async function renameDepartment(event) {
  event.preventDefault();
  const oldDepartment = $('oldDepartment').value;
  const newDepartment = $('newDepartment').value.trim();
  if (!oldDepartment || !newDepartment) return;
  try {
    const result = await rpc('rename_department', { oldDepartment, newDepartment });
    $('newDepartment').value = '';
    state.adminData.staff = state.adminData.staff.map(person => person.department === oldDepartment ? { ...person, department: newDepartment } : person);
    markAdminSectionLoaded('staff');
    renderStaffAdmin();
    showToast(`${result.updated}명의 부서명을 변경했습니다.`);
  } catch (error) { showToast(error.message, 4200); }
}

function downloadRosterTemplate() {
  if (!window.XLSX) return showToast('엑셀 기능을 불러오지 못했습니다.');
  const sheet = XLSX.utils.aoa_to_sheet([['부서', '성명'], ['교무기획부', '홍길동']]);
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, '구성원명단');
  XLSX.writeFile(book, '구성원_등록_양식.xlsx');
}

async function importRosterFile(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  const status = $('rosterImportStatus');
  try {
    if (!window.XLSX) throw new Error('엑셀 기능을 불러오지 못했습니다.');
    status.textContent = '파일을 읽는 중…';
    const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const people = normalizeRosterRows(rows);
    if (!people.length) throw new Error('부서와 성명 열을 찾지 못했습니다. 양식을 확인해 주세요.');
    const confirmed = await requestConfirmation({
      title: `${people.length}명을 가져올까요?`,
      message: '이미 등록된 같은 부서·성명은 건너뜁니다.',
      confirmLabel: '명단 가져오기'
    });
    if (!confirmed) return;
    const result = await rpc('save_staff_batch', { people });
    state.adminData.staff = sortByRegistration([...(state.adminData.staff || []), ...(result.people || [])]);
    markAdminSectionLoaded('staff');
    renderStaffAdmin();
    status.textContent = `${result.added}명 등록, ${result.skipped}명 건너뜀`;
  } catch (error) {
    status.textContent = error.message;
  } finally {
    event.target.value = '';
  }
}

function fillSettingsForm() {
  const settings = state.adminData?.settings || {};
  $('settingsSchoolName').value = settings.schoolName || '';
  $('settingsSubtitle').value = settings.subtitle || '';
  $('settingsNotice').value = settings.notice || '';
  $('settingsBrandColor').value = /^#[0-9a-f]{6}$/i.test(settings.brandColor || '') ? settings.brandColor : '#315c54';
  $('settingsPrivacyPurpose').value = settings.privacyPurpose || '';
  $('settingsPrivacyItems').value = settings.privacyItems || '';
  $('settingsPrivacyRetention').value = settings.privacyRetention || '';
  state.settingsFaviconData = String(settings.faviconData || '');
  renderFaviconSetting();
}

function renderFaviconSetting(message = '') {
  const favicon = state.settingsFaviconData;
  $('settingsFaviconPreview').src = favicon || DEFAULT_FAVICON_URL;
  $('settingsFaviconStatus').textContent = message || (favicon ? '사용자 지정 아이콘이 선택되어 있습니다.' : '현재 파란색 기본 아이콘을 사용합니다.');
}

function loadImageFromFile(file) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const image = new Image();
    image.onload = () => {
      URL.revokeObjectURL(url);
      resolve(image);
    };
    image.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error('이미지 파일을 읽을 수 없습니다. 손상되지 않은 파일인지 확인해 주세요.'));
    };
    image.src = url;
  });
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ''));
    reader.onerror = () => reject(new Error('파비콘 이미지를 변환하지 못했습니다.'));
    reader.readAsDataURL(blob);
  });
}

async function convertFaviconFile(file) {
  const allowedTypes = new Set(['image/png', 'image/jpeg', 'image/webp']);
  if (!allowedTypes.has(String(file?.type || '').toLowerCase())) throw new Error('PNG·JPG·WebP 이미지만 선택할 수 있습니다.');
  if (!file.size || file.size > FAVICON_MAX_SOURCE_BYTES) throw new Error('파비콘 이미지는 2MB 이하여야 합니다.');
  const image = await loadImageFromFile(file);
  if (!image.naturalWidth || !image.naturalHeight) throw new Error('이미지 크기를 확인할 수 없습니다.');
  const canvas = document.createElement('canvas');
  canvas.width = 64;
  canvas.height = 64;
  const context = canvas.getContext('2d');
  context.clearRect(0, 0, 64, 64);
  const scale = Math.min(64 / image.naturalWidth, 64 / image.naturalHeight);
  const width = Math.max(1, Math.round(image.naturalWidth * scale));
  const height = Math.max(1, Math.round(image.naturalHeight * scale));
  context.drawImage(image, Math.round((64 - width) / 2), Math.round((64 - height) / 2), width, height);
  const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
  if (!blob || blob.size > FAVICON_MAX_PNG_BYTES) throw new Error('변환된 파비콘이 32KB를 넘습니다. 더 단순한 이미지를 선택해 주세요.');
  return blobToDataUrl(blob);
}

async function handleFaviconFile(event) {
  const input = event.currentTarget;
  const file = input.files?.[0];
  if (!file) return;
  try {
    state.settingsFaviconData = await convertFaviconFile(file);
    renderFaviconSetting('64×64 PNG로 맞췄습니다. 설정 저장을 눌러 적용하세요.');
  } catch (error) {
    showToast(error.message, 4200);
  } finally {
    input.value = '';
  }
}

async function saveSettings(event) {
  event.preventDefault();
  const settings = {
    schoolName: $('settingsSchoolName').value.trim(),
    subtitle: $('settingsSubtitle').value.trim(),
    notice: $('settingsNotice').value.trim(),
    brandColor: $('settingsBrandColor').value,
    privacyPurpose: $('settingsPrivacyPurpose').value.trim(),
    privacyItems: $('settingsPrivacyItems').value.trim(),
    privacyRetention: $('settingsPrivacyRetention').value.trim(),
    faviconData: state.settingsFaviconData
  };
  if (!isPrivacyReady(settings)) return showToast('개인정보 안내 항목을 모두 입력해 주세요.');
  try {
    const result = await rpc('save_settings', { settings, frontendUrl: baseUrl });
    state.adminData.settings = result.settings || settings;
    markAdminSectionLoaded('settings');
    fillSettingsForm();
    applySettings(state.adminData.settings);
    showToast('기관 설정을 저장했습니다.');
  } catch (error) { showToast(error.message, 4200); }
}

function populateTrainingSelects() {
  const previousTrainingId = $('recordTraining').value;
  const previousDate = $('recordDate').value;
  const options = (state.adminData?.trainings || []).filter(training => !training.pending).map(training => `<option value="${escapeHtml(training.id)}">${escapeHtml(training.title)}</option>`).join('');
  $('recordTraining').innerHTML = `<option value="">연수 선택</option>${options}`;
  const previousTraining = (state.adminData?.trainings || []).find(training => training.id === previousTrainingId && !training.pending);
  if (previousTraining) {
    $('recordTraining').value = previousTraining.id;
    $('recordDate').value = previousTraining.daily ? (previousDate || todaySeoul()) : previousTraining.date;
    $('recordDate').disabled = !previousTraining.daily;
  } else {
    $('recordDate').value = todaySeoul();
    $('recordDate').disabled = false;
  }
}

function syncRecordDateToTraining() {
  const training = (state.adminData?.trainings || []).find(item => item.id === $('recordTraining').value && !item.pending);
  $('recordDate').value = training && !training.daily ? training.date : todaySeoul();
  $('recordDate').disabled = Boolean(training && !training.daily);
  state.records = [];
  $('recordSummary').replaceChildren();
  $('recordList').replaceChildren();
}

async function loadRecords(event) {
  event?.preventDefault();
  const trainingId = $('recordTraining').value;
  const date = $('recordDate').value;
  if (!trainingId || !date) return;
  try {
    const data = await rpc('list_records', { trainingId, date });
    state.records = data.records || [];
    $('recordSummary').innerHTML = `<p class="selection-summary">서명 ${state.records.length}건</p>`;
    $('recordList').innerHTML = state.records.length ? state.records.map(record => `
      <div class="admin-row" data-record-id="${escapeHtml(record.id)}"><div class="admin-row-main"><strong>${escapeHtml(record.department)} ${escapeHtml(record.name)}</strong><small>${escapeHtml(record.signDate)} ${escapeHtml(record.signTime)}</small></div><div class="row-actions"><button class="danger" data-action="delete-record">기록 삭제</button></div></div>`).join('') : '<div class="empty-state">서명 기록이 없습니다.</div>';
  } catch (error) { showToast(error.message, 4200); }
}

async function handleRecordClick(event) {
  const button = event.target.closest('[data-action="delete-record"]');
  const row = event.target.closest('[data-record-id]');
  if (!button || !row) return;
  const record = state.records.find(item => item.id === row.dataset.recordId);
  if (!record) return;
  const confirmed = await requestConfirmation({
    title: '서명 기록을 삭제할까요?',
    message: `${record.department} ${record.name}의 서명 기록과 이미지 파일을 함께 삭제합니다.`,
    confirmLabel: '기록 삭제',
    danger: true
  });
  if (!confirmed) return;
  try {
    const result = await rpc('delete_record', { recordId: record.id });
    state.records = state.records.filter(item => item.id !== (result.deletedId || record.id));
    $('recordSummary').innerHTML = `<p class="selection-summary">서명 ${state.records.length}건</p>`;
    row.remove();
    if (!state.records.length) $('recordList').innerHTML = '<div class="empty-state">서명 기록이 없습니다.</div>';
    showToast('서명 기록을 삭제했습니다.');
  } catch (error) { showToast(error.message, 4200); }
}

function renderShareAdmin() {
  const url = state.adminData?.shareUrl || buildShareUrl(baseUrl, state.adminData?.shareToken || '');
  $('adminShareUrl').value = url;
  renderQr($('adminQrCode'), url);
}

async function rotateShareToken() {
  const confirmed = await requestConfirmation({
    title: '공유 키를 교체할까요?',
    message: '기존 링크와 QR은 즉시 사용할 수 없게 됩니다.',
    confirmLabel: '공유 키 교체',
    danger: true
  });
  if (!confirmed) return;
  try {
    const data = await rpc('rotate_share_token', { frontendUrl: baseUrl });
    state.adminData.shareToken = data.shareToken;
    state.adminData.shareUrl = data.shareUrl;
    renderShareAdmin();
    showToast('공유 키를 교체했습니다. 새 링크를 안내해 주세요.', 4200);
  } catch (error) { showToast(error.message, 4200); }
}

async function changePassword(event) {
  event.preventDefault();
  try {
    if (!isValidAdminPassword($('newPassword').value)) throw new Error('새 비밀번호는 숫자 4자리 또는 문자와 숫자를 포함한 10자 이상으로 설정해 주세요.');
    await rpc('change_password', { currentPassword: $('currentPassword').value, newPassword: $('newPassword').value });
    $('currentPassword').value = '';
    $('newPassword').value = '';
    showToast('관리자 비밀번호를 변경했습니다.');
  } catch (error) { showToast(error.message, 4200); }
}

function exportStatusLabel(job) {
  if (job.status === 'complete') return '완료';
  if (job.status === 'preview_ready') return job.printOpenedAt ? '인쇄창 열림' : '미리보기 준비됨';
  if (job.status === 'failed') return `실패: ${job.error || '알 수 없는 오류'}`;
  if (job.status === 'expired') return '만료됨';
  return `생성 중 ${job.progress || 0}/${job.total || 0}`;
}

function exportOutputLabel(outputType) {
  if (outputType === 'xlsx') return '엑셀';
  if (outputType === 'print') return '인쇄';
  if (outputType === 'legacy_both') return '기존 PDF·엑셀';
  return 'PDF';
}

function exportPrimaryActionLabel(outputType) {
  if (outputType === 'xlsx') return '엑셀 만들기';
  if (outputType === 'print') return '인쇄하기';
  return 'PDF 내려받기';
}

function exportJobHtml(job, { showTrainingTitle = false } = {}) {
  const title = showTrainingTitle ? (job.trainingTitle || '삭제된 연수') : `${job.date} 출력`;
  const meta = showTrainingTitle
    ? `${job.date} · ${exportOutputLabel(job.outputType)} · ${exportStatusLabel(job)}`
    : `${exportOutputLabel(job.outputType)} · ${exportStatusLabel(job)}`;
  return `<div class="admin-row" data-job-id="${escapeHtml(job.jobId)}">
    <div class="admin-row-main"><strong>${escapeHtml(title)}</strong><small>${escapeHtml(meta)}</small></div>
    <div class="row-actions">
      ${job.hasPreview && job.status === 'preview_ready' ? '<button data-action="open-preview">미리보기</button>' : ''}
      ${job.hasPdf ? '<button data-action="download-pdf">PDF</button>' : ''}
      ${job.hasXlsx ? '<button data-action="download-xlsx">엑셀</button>' : ''}
      ${job.canPurge && !job.purgedAt ? '<button class="danger" data-action="purge-originals">원본 삭제</button>' : ''}
      ${job.canResume || job.status === 'processing' || job.status === 'queued' ? `<button data-action="resume-export">${job.status === 'failed' ? '다시 시도' : '계속 만들기'}</button>` : ''}
    </div>
  </div>`;
}

function renderExportJobs() {
  const container = $('exportJobList');
  if (!state.activeExportTrainingId) {
    container.replaceChildren();
    return;
  }
  if (!state.adminLoadedAt.training_workspace) {
    container.innerHTML = '<div class="empty-state">출력 내역을 불러오는 중입니다.</div>';
    return;
  }
  const jobs = (state.adminData?.exports || []).filter(job => job.trainingId === state.activeExportTrainingId);
  container.innerHTML = jobs.length
    ? jobs.map(job => exportJobHtml(job)).join('')
    : '<div class="empty-state">이 연수에서 생성한 출력 파일이 없습니다.</div>';
}

function renderOrphanExportJobs() {
  const section = $('orphanExportSection');
  if (!state.adminLoadedAt.training_workspace) {
    setHidden(section, true);
    return;
  }
  const trainingIds = new Set((state.adminData?.trainings || []).map(training => training.id));
  const jobs = (state.adminData?.exports || []).filter(job => !trainingIds.has(job.trainingId));
  $('orphanExportCount').textContent = jobs.length ? `(${jobs.length}건)` : '';
  $('orphanExportJobList').innerHTML = jobs.map(job => exportJobHtml(job, { showTrainingTitle: true })).join('');
  setHidden(section, !jobs.length);
}

const EXPORT_PREVIEW_PHASES = ['table', 'images', 'pdf', 'loading'];

function setExportPreviewPhase(phase, message = '') {
  const phaseIndex = EXPORT_PREVIEW_PHASES.indexOf(phase);
  document.querySelectorAll('#exportPreviewSteps [data-preview-phase]').forEach(item => {
    const index = EXPORT_PREVIEW_PHASES.indexOf(item.dataset.previewPhase);
    item.classList.toggle('active', index === phaseIndex);
    item.classList.toggle('done', phaseIndex >= 0 && index < phaseIndex);
  });
  $('exportPreviewLoading').classList.remove('error');
  setHidden($('exportPreviewSpinner'), false);
  $('exportPreviewLoadingText').textContent = message || '출력 미리보기를 준비하고 있습니다.';
}

function setExportPreviewFailure(message) {
  document.querySelectorAll('#exportPreviewSteps [data-preview-phase]').forEach(item => item.classList.remove('active'));
  $('exportPreviewLoading').classList.add('error');
  setHidden($('exportPreviewSpinner'), true);
  $('exportPreviewLoadingText').textContent = message;
  setHidden($('exportPreviewLoading'), false);
  $('confirmExportPreview').disabled = true;
}

function prepareExportPreviewLoading(training, payload) {
  releaseExportPreview();
  const token = `preview-${Date.now()}-${Math.random().toString(36).slice(2)}`;
  state.pendingPreviewToken = token;
  const dialog = $('exportPreviewDialog');
  $('exportPreviewMeta').textContent = `${training?.title || '연수'} · ${payload.date} · ${exportOutputLabel(payload.outputType)}`;
  $('confirmExportPreview').textContent = exportPrimaryActionLabel(payload.outputType);
  $('confirmExportPreview').disabled = true;
  setHidden($('exportPreviewLoading'), false);
  setExportPreviewPhase('table', '출력할 명단과 표를 준비하고 있습니다.');
  if (!dialog.open) dialog.showModal();
  return token;
}

function updateExportPreviewPhase(job) {
  const serverPhase = String(job?.phase || '').toLowerCase();
  if (['table', 'layout', 'prepare'].includes(serverPhase)) {
    setExportPreviewPhase('table', '출력할 명단과 표를 준비하고 있습니다.');
    return;
  }
  if (['pdf', 'render', 'converting'].includes(serverPhase) || (job?.total && Number(job.progress) >= Number(job.total) && job.status !== 'preview_ready')) {
    setExportPreviewPhase('pdf', 'A4 PDF로 변환하고 있습니다.');
    return;
  }
  if (job?.status === 'preview_ready') {
    setExportPreviewPhase('loading', '완성된 미리보기를 불러오고 있습니다.');
    return;
  }
  const count = job?.total ? ` ${Number(job.progress || 0)}/${Number(job.total)}` : '';
  setExportPreviewPhase('images', `실제 서명 이미지를 배치하고 있습니다.${count}`);
}

async function startExport(event) {
  event.preventDefault();
  const training = state.adminData?.trainings.find(item => item.id === state.activeExportTrainingId);
  const payload = {
    trainingId: training?.id || '',
    date: $('exportDate').value,
    columns: Number($('exportColumns').value),
    sort: $('exportSort').value,
    showRate: $('exportShowRate').checked,
    outputType: $('exportOutputType').value
  };
  if (!payload.trainingId || !payload.date) return showToast('출력할 연수와 날짜를 확인해 주세요.');
  const previewToken = prepareExportPreviewLoading(training, payload);
  try {
    localStorage.setItem(EXPORT_SETTINGS_KEY, JSON.stringify({
      trainingId: payload.trainingId, date: payload.date, columns: String(payload.columns),
      sort: payload.sort, showRate: payload.showRate, outputType: payload.outputType
    }));
    const job = await rpc('start_export', payload);
    upsertExportJob(job);
    setHidden($('exportProgress'), false);
    updateExportPreviewPhase(job);
    await runExportJob(job.jobId, { previewPrepared: true, previewToken });
  } catch (error) {
    setExportPreviewFailure(error.message);
    showToast(error.message, 5200);
  }
}

async function runExportJob(jobId, { previewPrepared = false, previewToken = '' } = {}) {
  const box = $('exportProgress');
  setHidden(box, false);
  let job = state.adminData?.exports.find(item => item.jobId === jobId);
  if (!previewPrepared) {
    previewToken = prepareExportPreviewLoading({
      title: job?.trainingTitle || state.adminData?.trainings.find(item => item.id === job?.trainingId)?.title || '연수'
    }, {
      date: job?.date || '',
      outputType: job?.outputType || 'pdf'
    });
    updateExportPreviewPhase(job);
  }
  try {
    do {
      updateExportPreviewPhase(job);
      const remainingImages = Math.max(0, Number(job?.total || 0) - Number(job?.progress || 0));
      const finalBatchExpected = remainingImages <= 30;
      const pdfPhaseTimer = finalBatchExpected ? setTimeout(() => {
        if (state.pendingPreviewToken === previewToken && $('exportPreviewDialog').open) {
          setExportPreviewPhase('pdf', '실제 서명을 반영한 문서를 A4 PDF로 변환하고 있습니다.');
        }
      }, 700) : null;
      try {
        job = await rpc('continue_export', { jobId });
      } finally {
        clearTimeout(pdfPhaseTimer);
      }
      if (job.busy) {
        await new Promise(resolve => setTimeout(resolve, 800));
        continue;
      }
      const percent = job.total ? Math.round(job.progress / job.total * 100) : job.status === 'preview_ready' ? 100 : 0;
      box.querySelector('progress').value = percent;
      box.querySelector('p').textContent = job.status === 'preview_ready'
        ? '실제 서명이 포함된 A4 미리보기를 만들었습니다.'
        : `서명 이미지를 배치하는 중입니다. ${job.progress}/${job.total}`;
      upsertExportJob(job);
      updateExportPreviewPhase(job);
    } while (job.status === 'processing' || job.status === 'queued');
    if (job.status === 'preview_ready' && state.pendingPreviewToken === previewToken) await openExportPreview(job);
    else if (job.status === 'preview_ready') showToast('미리보기가 준비되었습니다. 출력 내역에서 열 수 있습니다.', 4200);
    else if (job.status === 'failed') throw new Error(job.error || '출력 파일 생성에 실패했습니다.');
  } catch (error) {
    box.querySelector('p').textContent = error.message;
    setExportPreviewFailure(error.message);
    showToast(error.message, 5200);
  }
}

async function downloadExportBlob(jobId, format) {
  let offset = 0;
  let total = null;
  let fileName = `연수_서명등록부.${format}`;
  let mimeType = format === 'pdf' ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const chunks = [];
  do {
    const chunk = await rpc('download_export_chunk', { jobId, format, offset, chunkSize: 1048576 });
    chunks.push(Uint8Array.from(atob(chunk.base64), character => character.charCodeAt(0)));
    offset = chunk.nextOffset;
    total = chunk.totalBytes;
    fileName = chunk.fileName || fileName;
    mimeType = chunk.mimeType || mimeType;
  } while (offset < total);
  return { blob: new Blob(chunks, { type: mimeType }), fileName, mimeType };
}

function saveBlob(blob, fileName) {
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

async function downloadExport(jobId, format) {
  showToast('파일을 내려받는 중입니다…', 5000);
  const file = await downloadExportBlob(jobId, format);
  saveBlob(file.blob, file.fileName);
}

function releaseExportPreview() {
  const preview = state.activePreview;
  if (preview?.blobUrl) URL.revokeObjectURL(preview.blobUrl);
  state.activePreview = null;
  state.pendingPreviewToken = '';
  const frame = $('exportPreviewFrame');
  frame.removeAttribute('srcdoc');
  frame.src = 'about:blank';
  setHidden($('exportPreviewLoading'), false);
}

function closeExportPreview() {
  releaseExportPreview();
  if ($('exportPreviewDialog').open) $('exportPreviewDialog').close();
}

function demoPreviewHtml(job) {
  const names = demoData.staff.map((person, index) => `<tr><td>${index + 1}</td><td>${escapeHtml(person.department)}</td><td>${escapeHtml(person.name)}</td><td>${index < 2 ? '데모 서명' : '미서명'}</td></tr>`).join('');
  return `<!doctype html><html lang="ko"><meta charset="utf-8"><body><h1>데모 연수 서명등록부</h1><p>${escapeHtml(formatKoreanDate(job.date))} · 데모 미리보기</p><table border="1" cellspacing="0" cellpadding="10"><thead><tr><th>번호</th><th>부서</th><th>성명</th><th>서명</th></tr></thead><tbody>${names}</tbody></table><p>실제 운영에서는 등록된 PNG 서명 이미지가 이 칸에 표시됩니다.</p></body></html>`;
}

async function openExportPreview(job) {
  releaseExportPreview();
  const dialog = $('exportPreviewDialog');
  const action = $('confirmExportPreview');
  $('exportPreviewMeta').textContent = `${job.trainingTitle || '연수'} · ${job.date} · ${exportOutputLabel(job.outputType)}`;
  action.textContent = exportPrimaryActionLabel(job.outputType);
  action.disabled = true;
  setHidden($('exportPreviewLoading'), false);
  setExportPreviewPhase('loading', '완성된 미리보기를 불러오고 있습니다.');
  if (!dialog.open) dialog.showModal();
  if (DEMO) {
    $('exportPreviewFrame').srcdoc = demoPreviewHtml(job);
    setHidden($('exportPreviewLoading'), true);
    state.activePreview = { job, blobUrl: '', blob: null, fileName: '데모_서명등록부.pdf' };
    action.disabled = false;
    return;
  }
  try {
    const file = await downloadExportBlob(job.jobId, 'preview');
    const blobUrl = URL.createObjectURL(file.blob);
    state.activePreview = { job, blobUrl, blob: file.blob, fileName: file.fileName };
    const frame = $('exportPreviewFrame');
    frame.addEventListener('load', () => {
      setHidden($('exportPreviewLoading'), true);
      action.disabled = false;
    }, { once: true });
    frame.src = blobUrl;
    setTimeout(() => {
      if (state.activePreview?.blobUrl === blobUrl) {
        setHidden($('exportPreviewLoading'), true);
        action.disabled = false;
      }
    }, 1800);
  } catch (error) {
    setExportPreviewFailure(error.message);
    throw error;
  }
}

async function confirmExportPreview() {
  const preview = state.activePreview;
  if (!preview) return;
  const { job } = preview;
  if (DEMO) return showToast('데모에서는 파일 생성과 인쇄를 실행하지 않습니다.', 4200);
  const button = $('confirmExportPreview');
  try {
    if (job.outputType === 'print') {
      let openedFallback = false;
      try {
        const frameWindow = $('exportPreviewFrame').contentWindow;
        frameWindow.focus();
        frameWindow.print();
      } catch {
        openedFallback = Boolean(window.open(preview.blobUrl, '_blank', 'noopener'));
      }
      rpc('record_print_opened', { jobId: job.jobId }).then(upsertExportJob).catch(() => {});
      showToast(openedFallback ? '새 탭에서 오른쪽 위 인쇄 버튼을 눌러 주세요.' : '인쇄창을 열었습니다.', 5200);
      return;
    }
    button.disabled = true;
    button.textContent = job.outputType === 'xlsx' ? '엑셀 만드는 중…' : 'PDF 준비 중…';
    const completed = await rpc('finalize_export', { jobId: job.jobId });
    upsertExportJob(completed);
    await downloadExport(completed.jobId, job.outputType);
    closeExportPreview();
    showToast(job.outputType === 'xlsx' ? '엑셀 파일을 만들었습니다.' : 'PDF 파일을 만들었습니다.');
  } catch (error) {
    showToast(error.message, 5200);
  } finally {
    button.disabled = false;
    button.textContent = exportPrimaryActionLabel(job.outputType);
  }
}

function requestPurgeConfirmation(expected) {
  return new Promise(resolve => {
    const overlay = document.createElement('div');
    overlay.className = 'purge-confirm-overlay';
    overlay.innerHTML = `
      <div class="purge-confirm-card" role="dialog" aria-modal="true" aria-labelledby="purgeConfirmTitle">
        <form>
          <h3 id="purgeConfirmTitle">서명 원본을 삭제할까요?</h3>
          <p>선택한 출력 파일을 내려받아 보관했는지 확인해 주세요. 삭제한 서명 기록과 이미지는 되돌릴 수 없습니다.</p>
          <label>계속하려면 연수명을 그대로 입력하세요.
            <strong>${escapeHtml(expected)}</strong>
            <input name="confirmation" autocomplete="off" maxlength="100" required>
          </label>
          <div class="button-row">
            <button class="button secondary" type="button" data-action="cancel">취소</button>
            <button class="button danger" type="submit" disabled>원본 삭제</button>
          </div>
        </form>
      </div>`;
    const form = overlay.querySelector('form');
    const input = overlay.querySelector('input');
    const submit = overlay.querySelector('button[type="submit"]');
    const finish = value => {
      overlay.remove();
      resolve(value);
    };
    input.addEventListener('input', () => { submit.disabled = input.value !== expected; });
    form.addEventListener('submit', event => {
      event.preventDefault();
      finish(input.value);
    });
    overlay.querySelector('[data-action="cancel"]').addEventListener('click', () => finish(''));
    overlay.addEventListener('click', event => {
      if (event.target === overlay) finish('');
    });
    $('adminDialog').append(overlay);
    input.focus();
  });
}

async function purgeOriginals(job) {
  const training = state.adminData.trainings.find(item => item.id === job.trainingId);
  const expected = training?.title || job.trainingTitle;
  const confirmation = await requestPurgeConfirmation(expected);
  if (!confirmation) return;
  try {
    const result = await rpc('purge_originals', { jobId: job.jobId, confirmation });
    if (result.job) upsertExportJob(result.job);
    showToast(result.failed ? `${result.deleted}건 삭제, ${result.failed}건 실패. 다시 시도할 수 있습니다.` : `${result.deleted}건의 원본을 삭제했습니다.`, 5200);
  } catch (error) { showToast(error.message, 5200); }
}

async function handleExportJobClick(event) {
  const button = event.target.closest('button[data-action]');
  const row = event.target.closest('[data-job-id]');
  if (!button || !row) return;
  const job = state.adminData.exports.find(item => item.jobId === row.dataset.jobId);
  if (!job) return;
  try {
    if (button.dataset.action === 'download-pdf') await downloadExport(job.jobId, 'pdf');
    if (button.dataset.action === 'download-xlsx') await downloadExport(job.jobId, 'xlsx');
    if (button.dataset.action === 'resume-export') await runExportJob(job.jobId);
    if (button.dataset.action === 'open-preview') await openExportPreview(job);
    if (button.dataset.action === 'purge-originals') await purgeOriginals(job);
  } catch (error) { showToast(error.message, 5200); }
}

function restoreExportSettings(training) {
  $('exportColumns').value = '2';
  $('exportSort').value = 'registration';
  $('exportOutputType').value = 'pdf';
  $('exportShowRate').checked = true;
  $('exportDate').value = training.daily ? todaySeoul() : training.date;
  try {
    const saved = JSON.parse(localStorage.getItem(EXPORT_SETTINGS_KEY) || 'null');
    if (!saved) return;
    if (training.daily && saved.trainingId === training.id && /^\d{4}-\d{2}-\d{2}$/.test(String(saved.date || ''))) $('exportDate').value = saved.date;
    if (['1', '2', '3'].includes(String(saved.columns))) $('exportColumns').value = String(saved.columns);
    if (['registration', 'department', 'name'].includes(saved.sort)) $('exportSort').value = saved.sort;
    if (['pdf', 'xlsx', 'print'].includes(saved.outputType)) $('exportOutputType').value = saved.outputType;
    $('exportShowRate').checked = saved.showRate !== false;
  } catch { /* 손상된 브라우저 설정은 기본값을 사용합니다. */ }
}

function renderAdmin() {
  renderTrainingAdmin();
  fillSettingsForm();
  populateTrainingSelects();
  if (state.adminLoadedAt.staff) renderStaffAdmin();
  if (state.adminLoadedAt.share) renderShareAdmin();
}

function closeAdminAndLogout() {
  closeExportPreview();
  const currentShareToken = state.adminData?.shareToken || shareToken;
  const logoutRequest = state.adminSession ? rpc('logout') : Promise.resolve();
  clearInterval(state.adminSyncTimer);
  stopUnsignedStatusRefresh();
  state.adminSession = '';
  state.adminData = null;
  state.adminLoadedAt = {};
  state.adminSectionPromises = {};
  state.activeExportTrainingId = '';
  state.activeUnsignedTrainingId = '';
  state.unsignedStatusDate = '';
  state.unsignedStatusCache.clear();
  state.unsignedStatusRequest = null;
  state.records = [];
  if ($('adminDialog').open) $('adminDialog').close();
  logoutRequest.catch(() => { /* 이미 만료된 서버 세션은 별도 안내가 필요하지 않습니다. */ });
  if (currentShareToken) {
    shareToken = currentShareToken;
    history.replaceState(null, '', buildShareUrl(baseUrl, shareToken));
    initializePublicApp();
  }
}

async function initializePublicApp() {
  $('schoolDate').textContent = formatKoreanHeaderDate(todaySeoul());
  if (!shareToken) {
    showPanel('invalidPanel');
    return;
  }
  showPanel('loadingPanel');
  try {
    state.publicData = await rpc('get_public_data', { shareToken }, { admin: false });
    if (!state.publicData.privacyReady) throw Object.assign(new Error('관리자가 개인정보 처리 안내를 완료하지 않았습니다.'), { code: 'PRIVACY_NOT_READY' });
    applySettings(state.publicData.settings || {});
    renderTrainings();
    showPanel('trainingPanel');
    if (DEMO) showStatus('데모 화면입니다. 입력 내용은 실제로 저장되지 않습니다.', false);
  } catch (error) {
    const unavailable = error.code === 'INVALID_LINK';
    $('invalidTitle').textContent = unavailable ? '오늘 참여할 수 있는 연수가 없습니다' : '서명 화면을 준비하지 못했습니다';
    $('invalidMessage').textContent = unavailable ? '연수 일정 또는 공유 링크를 다시 확인해 주세요.' : error.message;
    showPanel('invalidPanel');
  }
}

function bindEvents() {
  $('departmentSelect').addEventListener('change', renderStaffForDepartment);
  $('staffSelect').addEventListener('change', () => { $('goToSignature').disabled = !$('staffSelect').value; });
  $('goToSignature').addEventListener('click', goToSignature);
  $('backToTraining').addEventListener('click', () => showPanel('trainingPanel'));
  $('backToPerson').addEventListener('click', () => showPanel('personPanel'));
  $('submitSignature').addEventListener('click', submitSignature);
  $('clearSignature').addEventListener('click', clearSignature);
  $('undoSignature').addEventListener('click', undoSignature);
  $('signAnother').addEventListener('click', () => { state.selectedTraining = null; state.selectedStaff = null; showPanel('trainingPanel'); });
  $('signatureCanvas').addEventListener('pointerdown', startDrawing);
  $('signatureCanvas').addEventListener('pointermove', continueDrawing);
  $('signatureCanvas').addEventListener('pointerup', stopDrawing);
  $('signatureCanvas').addEventListener('pointercancel', stopDrawing);
  new ResizeObserver(resizeCanvas).observe($('signatureCanvas'));
  $('shareButton').addEventListener('click', openShareDialog);
  $('copyShareUrl').addEventListener('click', () => copyText($('shareUrl').value, '공유 링크를 복사했습니다.'));
  ['openPrivacy', 'footerPrivacy'].forEach(id => $(id).addEventListener('click', renderPrivacy));
  $('adminButton').addEventListener('click', openAdminLogin);
  $('adminLoginForm').addEventListener('submit', handleAdminLogin);
  $('adminTabs').addEventListener('click', event => { const button = event.target.closest('[data-admin-tab]'); if (button) switchAdminTab(button.dataset.adminTab); });
  $('closeAdmin').addEventListener('click', closeAdminAndLogout);
  $('adminDialog').addEventListener('cancel', event => {
    event.preventDefault();
    if (!state.adminAuthenticating) closeAdminAndLogout();
  });
  $('newTraining').addEventListener('click', () => openTrainingForm());
  $('cancelTraining').addEventListener('click', () => setHidden($('trainingForm'), true));
  $('trainingForm').addEventListener('submit', saveTraining);
  $('trainingAdminList').addEventListener('click', handleTrainingListClick);
  $('collapseTrainingExport').addEventListener('click', collapseTrainingExport);
  $('collapseTrainingUnsigned').addEventListener('click', collapseTrainingUnsigned);
  $('unsignedStatusDate').addEventListener('change', event => {
    state.unsignedStatusDate = event.currentTarget.value;
    fetchUnsignedStatus({ force: true }).catch(() => {});
  });
  $('unsignedStatusViewToggle').addEventListener('click', event => {
    const button = event.target.closest('[data-unsigned-view]');
    if (!button) return;
    state.unsignedStatusMode = button.dataset.unsignedView;
    const key = unsignedStatusCacheKey(state.activeUnsignedTrainingId, state.unsignedStatusDate);
    renderUnsignedStatusContent(state.unsignedStatusCache.get(key)?.data || null);
  });
  $('refreshUnsignedStatus').addEventListener('click', () => {
    fetchUnsignedStatus({ force: true }).catch(error => showToast(error.message, 4200));
  });
  $('staffNames').addEventListener('compositionstart', () => { state.staffNamesComposing = true; });
  $('staffNames').addEventListener('compositionend', () => { state.staffNamesComposing = false; normalizeStaffNamesField(); });
  $('staffNames').addEventListener('input', event => { if (!state.staffNamesComposing && !event.isComposing) normalizeStaffNamesField(); });
  $('staffAddForm').addEventListener('submit', addStaff);
  $('staffAdminList').addEventListener('click', handleStaffListClick);
  $('renameDepartmentForm').addEventListener('submit', renameDepartment);
  $('downloadRosterTemplate').addEventListener('click', downloadRosterTemplate);
  $('rosterFile').addEventListener('change', importRosterFile);
  $('settingsFaviconFile').addEventListener('change', handleFaviconFile);
  $('resetFavicon').addEventListener('click', () => {
    state.settingsFaviconData = '';
    renderFaviconSetting('기본 아이콘으로 되돌렸습니다. 설정 저장을 눌러 적용하세요.');
  });
  $('settingsForm').addEventListener('submit', saveSettings);
  $('recordTraining').addEventListener('change', syncRecordDateToTraining);
  $('recordFilterForm').addEventListener('submit', loadRecords);
  $('recordList').addEventListener('click', handleRecordClick);
  $('rotateShareToken').addEventListener('click', rotateShareToken);
  $('adminCopyShare').addEventListener('click', () => copyText($('adminShareUrl').value, '공유 링크를 복사했습니다.'));
  $('changePasswordForm').addEventListener('submit', changePassword);
  $('exportForm').addEventListener('submit', startExport);
  $('exportJobList').addEventListener('click', handleExportJobClick);
  $('orphanExportJobList').addEventListener('click', handleExportJobClick);
  $('closeExportPreview').addEventListener('click', closeExportPreview);
  $('confirmExportPreview').addEventListener('click', confirmExportPreview);
  $('exportPreviewDialog').addEventListener('cancel', event => { event.preventDefault(); closeExportPreview(); });
}

bindEvents();
initializePublicApp();
