export const SEOUL_TIME_ZONE = 'Asia/Seoul';

export function parseShareToken(hash = '') {
  const raw = String(hash).replace(/^#/, '');
  if (!raw || raw === 'admin') return '';
  const params = new URLSearchParams(raw);
  const token = params.get('k') || '';
  return /^[A-Za-z0-9_-]{20,64}$/.test(token) ? token : '';
}

export function formatKoreanDate(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value || ''))) return '';
  const [year, month, day] = value.split('-').map(Number);
  return `${year}년 ${month}월 ${day}일`;
}

export function formatKoreanHeaderDate(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value || ''))) return '';
  const [year, month, day] = value.split('-').map(Number);
  const weekdays = ['일', '월', '화', '수', '목', '금', '토'];
  const weekday = weekdays[new Date(Date.UTC(year, month - 1, day)).getUTCDay()];
  return `${month}월 ${day}일 (${weekday})`;
}

export function todaySeoul(now = new Date()) {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: SEOUL_TIME_ZONE,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).formatToParts(now);
  const get = type => parts.find(part => part.type === type)?.value || '';
  return `${get('year')}-${get('month')}-${get('day')}`;
}

export function trainingTimeLabel(training) {
  if (!training) return '';
  const date = training.daily ? '매일' : formatKoreanDate(training.date);
  const time = training.startTime || training.endTime
    ? `${training.startTime || '00:00'} ~ ${training.endTime || '24:00'}`
    : '시간 제한 없음';
  return `${date} · ${time}`;
}

export function groupStaffByDepartment(staff = []) {
  const groups = new Map();
  [...staff]
    .filter(person => person && person.active !== false && person.id && person.name)
    .sort((a, b) => Number(a.sortOrder || 0) - Number(b.sortOrder || 0))
    .forEach(person => {
      const department = String(person.department || '미지정').trim() || '미지정';
      if (!groups.has(department)) groups.set(department, []);
      groups.get(department).push(person);
    });
  return groups;
}

export function normalizeAudienceDepartments(value = []) {
  let source = value;
  if (!Array.isArray(source)) {
    try {
      source = JSON.parse(String(source || '[]'));
    } catch {
      source = [];
    }
  }
  if (!Array.isArray(source)) return [];
  const seen = new Set();
  const departments = [];
  source.forEach(item => {
    const department = String(item || '').trim();
    const key = department.toLocaleLowerCase('ko');
    if (!department || department.length > 50 || seen.has(key)) return;
    seen.add(key);
    departments.push(department);
  });
  return departments.slice(0, 200);
}

export function trainingAudienceMode(training = {}) {
  const mode = String(training.audienceMode || '').trim();
  if (!mode || mode === 'all') return 'all';
  return 'departments';
}

export function isStaffIncludedInTraining(training = {}, person = {}) {
  const storedMode = String(training.audienceMode || '').trim();
  if (!storedMode || storedMode === 'all') return true;
  if (storedMode !== 'departments') return false;
  const departments = normalizeAudienceDepartments(training.audienceDepartments);
  const department = String(person.department || '').trim().toLowerCase();
  return departments.some(item => item.toLowerCase() === department);
}

export function staffForTraining(staff = [], training = {}) {
  return [...staff]
    .filter(person => person && person.active !== false && isStaffIncludedInTraining(training, person))
    .sort((a, b) => Number(a.sortOrder || 0) - Number(b.sortOrder || 0));
}

export function activeDepartments(staff = []) {
  return [...groupStaffByDepartment(staff).keys()];
}

export function normalizeRosterRows(rows = []) {
  if (!Array.isArray(rows) || rows.length === 0) return [];
  const normalized = [];
  const seen = new Set();
  const header = rows[0].map(value => String(value || '').trim().replace(/\s/g, ''));
  const departmentIndex = header.findIndex(value => ['부서', '부서명', '소속'].includes(value));
  const nameIndex = header.findIndex(value => ['성명', '이름', '교직원명'].includes(value));
  const start = departmentIndex >= 0 && nameIndex >= 0 ? 1 : 0;
  const deptColumn = departmentIndex >= 0 ? departmentIndex : 0;
  const personColumn = nameIndex >= 0 ? nameIndex : 1;

  rows.slice(start).forEach(row => {
    const department = String(row?.[deptColumn] ?? '').trim();
    const name = String(row?.[personColumn] ?? '').trim();
    if (!department || !name || department.length > 50 || name.length > 50) return;
    const key = `${department}\u0000${name}`.toLocaleLowerCase('ko');
    if (seen.has(key)) return;
    seen.add(key);
    normalized.push({ department, name });
  });
  return normalized;
}

export function normalizeNameEntryText(value) {
  return String(value || '')
    .replace(/\r\n?/g, '\n')
    .replace(/[\t ,;]+/g, '\n')
    .replace(/\n{2,}/g, '\n');
}

export function splitNames(value) {
  return [...new Set(normalizeNameEntryText(value)
    .split('\n')
    .map(name => name.trim())
    .filter(Boolean))]
    .slice(0, 200);
}

export function isPrivacyReady(settings = {}) {
  return ['schoolName', 'subtitle', 'privacyPurpose', 'privacyItems', 'privacyRetention']
    .every(key => String(settings[key] || '').trim().length > 0);
}

export function isValidAdminPassword(password) {
  const value = String(password || '');
  if (/^\d{4}$/.test(value)) return true;
  return value.length >= 10
    && value.length <= 100
    && /[A-Za-z가-힣]/.test(value)
    && /\d/.test(value);
}

export function validateTraining(training = {}) {
  const errors = [];
  if (!String(training.title || '').trim()) errors.push('연수명을 입력해 주세요.');
  if (!training.daily && !/^\d{4}-\d{2}-\d{2}$/.test(String(training.date || ''))) errors.push('연수 날짜를 입력해 주세요.');
  if (training.startTime && training.endTime && training.startTime >= training.endTime) errors.push('종료 시각은 시작 시각보다 늦어야 합니다.');
  const audienceMode = String(training.audienceMode || '').trim();
  if (audienceMode && audienceMode !== 'all' && audienceMode !== 'departments') errors.push('서명 대상 방식을 다시 선택해 주세요.');
  if (audienceMode === 'departments' && !normalizeAudienceDepartments(training.audienceDepartments).length) {
    errors.push('서명 대상 부서를 하나 이상 선택해 주세요.');
  }
  return errors;
}

export function buildShareUrl(baseUrl, token) {
  const base = String(baseUrl || '').split('#')[0].replace(/\?.*$/, '');
  if (!base || !token) return '';
  return `${base}#k=${encodeURIComponent(token)}`;
}

export function localDuplicateKey(trainingId, staffId, date) {
  return `training-sign:${trainingId}:${staffId}:${date}`;
}

export function sortRecords(records = [], mode = 'registration') {
  return [...records].sort((a, b) => {
    if (mode === 'name') return String(a.name).localeCompare(String(b.name), 'ko') || String(a.department).localeCompare(String(b.department), 'ko');
    if (mode === 'department') return String(a.department).localeCompare(String(b.department), 'ko') || String(a.name).localeCompare(String(b.name), 'ko');
    return Number(a.sortOrder || 0) - Number(b.sortOrder || 0);
  });
}

export function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

export function safeFileName(value) {
  return String(value || '')
    .normalize('NFC')
    .replace(/[\\/:*?"<>|\r\n]+/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 80) || '연수';
}

export function bytesToBase64(bytes) {
  let binary = '';
  const chunk = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunk) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunk));
  }
  return btoa(binary);
}

export function base64ToUint8Array(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
}
