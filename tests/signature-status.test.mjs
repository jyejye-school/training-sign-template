import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve(import.meta.dirname, '..');
const backend = fs.readFileSync(path.join(root, 'apps-script', 'Code.gs'), 'utf8');
const loadHelpers = new Function(`${backend}
return {
  buildTrainingSignatureStatus_: buildTrainingSignatureStatus_,
  getTrainingSignatureStatus_: getTrainingSignatureStatus_,
  startExport_: startExport_,
  buildExportRoster_: buildExportRoster_,
  dispatch_: dispatch_,
  readExportImageBatch_: readExportImageBatch_,
  insertExportImageBatch_: insertExportImageBatch_,
  normalizeDriveDownloadResponse_: normalizeDriveDownloadResponse_,
  continueExport_: continueExport_,
  configureStatus: function(training, staff, signatures) {
    findRow_ = function() { return training; };
    readRows_ = function(definition) {
      if (definition === SHEETS.STAFF) return staff;
      if (definition === SHEETS.SIGNATURES) return signatures;
      return [];
    };
    requireAdminSession_ = function(token) {
      if (token !== 'valid-session') apiError_('SESSION_EXPIRED', '관리자 로그인이 필요합니다.');
      return token;
    };
  },
  configureTerminalExport: function(job) {
    let leaseCalls = 0;
    findRow_ = function() { return job; };
    acquireExportLease_ = function() { leaseCalls += 1; return ''; };
    return function() { return leaseCalls; };
  },
  configureExportRoster: function(staff) {
    readRows_ = function(definition) {
      if (definition === SHEETS.STAFF) return staff;
      return [];
    };
  }
};`);

function createHarness() {
  return loadHelpers();
}

test('미서명 현황은 활성 명단 등록순과 현재 명단 기준 서명률을 사용한다', () => {
  const { buildTrainingSignatureStatus_ } = createHarness();
  const result = buildTrainingSignatureStatus_(
    'training-1',
    '2026-07-20',
    [
      { id: 'staff-2', department: '연구부', name: '박교사', sortOrder: 2 },
      { id: 'staff-1', department: '교무부', name: '김교사', sortOrder: 1 }
    ],
    [
      { staffId: 'staff-2', signTime: '10:12:03', createdAt: '2026-07-20T01:12:03Z' },
      { staffId: 'deleted-staff', signTime: '10:13:00', createdAt: '2026-07-20T01:13:00Z' }
    ]
  );

  assert.equal(result.summary.targetCount, 2);
  assert.equal(result.summary.signedCount, 1);
  assert.equal(result.summary.unsignedCount, 1);
  assert.equal(result.summary.rate, 50);
  assert.equal(result.summary.outsideRosterSignedCount, 1);
  assert.deepEqual(result.people.map(person => person.staffId), ['staff-1', 'staff-2']);
  assert.deepEqual(result.people.map(person => person.status), ['unsigned', 'signed']);
  assert.equal(result.people[1].signTime, '10:12');
});

test('특수한 ID도 일반 객체 속성과 충돌하지 않고 집계한다', () => {
  const { buildTrainingSignatureStatus_ } = createHarness();
  const result = buildTrainingSignatureStatus_(
    'training-2',
    '2026-07-20',
    [
      { id: 'toString', department: '교무부', name: '동명이인1', sortOrder: 1 },
      { id: '__proto__', department: '교무부', name: '동명이인2', sortOrder: 2 }
    ],
    [
      { staffId: 'toString', signTime: '09:00:00', createdAt: '1' },
      { staffId: '__proto__', signTime: '09:01:00', createdAt: '2' }
    ]
  );

  assert.equal(result.summary.signedCount, 2);
  assert.equal(result.summary.outsideRosterSignedCount, 0);
  assert.deepEqual(result.people.map(person => person.status), ['signed', 'signed']);
  assert.ok(result.people.every(person => !('fileId' in person) && !('imageFileId' in person)));
});

test('현황 API는 관리자 세션을 요구하고 고정 날짜와 활성 명단을 서버에서 검증한다', () => {
  const harness = createHarness();
  const training = { id: 'training-0001', date: '2026-07-20', daily: false };
  harness.configureStatus(
    training,
    [
      { id: 'staff-0001', department: '교무부', name: '활성 교사', active: true, sortOrder: 1 },
      { id: 'staff-0002', department: '연구부', name: '비활성 교사', active: false, sortOrder: 2 }
    ],
    [
      { trainingId: training.id, staffId: 'staff-0001', signDate: '2026-07-20', signTime: '09:10:30', createdAt: '1', imageFileId: 'private-1' },
      { trainingId: training.id, staffId: 'staff-0002', signDate: '2026-07-20', signTime: '09:11:30', createdAt: '2', imageFileId: 'private-2' }
    ]
  );

  assert.throws(
    () => harness.dispatch_({ action: 'get_training_signature_status', trainingId: training.id, date: training.date }),
    error => error.apiCode === 'SESSION_EXPIRED'
  );
  assert.throws(
    () => harness.dispatch_({ action: 'get_training_signature_status', sessionToken: 'valid-session', trainingId: training.id, date: '2026-07-19' }),
    error => error.apiCode === 'TRAINING_DATE'
  );

  const result = harness.dispatch_({
    action: 'get_training_signature_status',
    sessionToken: 'valid-session',
    trainingId: training.id,
    date: training.date
  });
  assert.equal(result.summary.targetCount, 1);
  assert.equal(result.summary.signedCount, 1);
  assert.equal(result.summary.outsideRosterSignedCount, 1);
  assert.deepEqual(result.people.map(person => person.name), ['활성 교사']);
  assert.ok(!JSON.stringify(result).includes('private-'));
});

test('미서명 현황은 선택 부서만 대상으로 계산하고 나머지 서명은 대상 외 건수로 남긴다', () => {
  const harness = createHarness();
  const training = {
    id: 'training-0001', date: '2026-07-20', daily: false,
    audienceMode: 'departments', audienceDepartments: '["교무부"]'
  };
  harness.configureStatus(
    training,
    [
      { id: 'teacher', department: '교무부', name: '김교사', active: true, sortOrder: 1 },
      { id: 'office', department: '행정실', name: '박주무관', active: true, sortOrder: 2 },
      { id: 'inactive', department: '교무부', name: '비활성', active: false, sortOrder: 3 }
    ],
    [
      { trainingId: training.id, staffId: 'teacher', signDate: training.date, signTime: '09:10:00', createdAt: '1' },
      { trainingId: training.id, staffId: 'office', signDate: training.date, signTime: '09:11:00', createdAt: '2' },
      { trainingId: training.id, staffId: 'inactive', signDate: training.date, signTime: '09:12:00', createdAt: '3' }
    ]
  );
  const result = harness.getTrainingSignatureStatus_(training.id, training.date);
  assert.deepEqual(result.people.map(person => person.staffId), ['teacher']);
  assert.deepEqual(result.summary, {
    targetCount: 1,
    signedCount: 1,
    unsignedCount: 0,
    rate: 100,
    outsideRosterSignedCount: 2
  });
});

test('출력은 현재 대상만 포함하고 활성 대상 외 서명은 보존하되 출력하지 않는다', () => {
  const harness = createHarness();
  const training = {
    id: 'training-0001', date: '2026-07-20', daily: false,
    audienceMode: 'departments', audienceDepartments: '["교무부"]'
  };
  harness.configureExportRoster([
    { id: 'teacher', department: '교무부', name: '김교사', active: true, sortOrder: 1 },
    { id: 'office', department: '행정실', name: '박주무관', active: true, sortOrder: 2 },
    { id: 'inactive', department: '교무부', name: '퇴직자', active: false, sortOrder: 3 }
  ]);
  const roster = harness.buildExportRoster_(training.id, training.date, 'registration', training, [
    { id: 'sig-teacher', trainingId: training.id, staffId: 'teacher', department: '교무부', name: '김교사', signDate: training.date, signTime: '09:00:00', imageFileId: 'file-1', createdAt: '1' },
    { id: 'sig-office', trainingId: training.id, staffId: 'office', department: '행정실', name: '박주무관', signDate: training.date, signTime: '09:01:00', imageFileId: 'file-2', createdAt: '2' },
    { id: 'sig-inactive', trainingId: training.id, staffId: 'inactive', department: '교무부', name: '퇴직자', signDate: training.date, signTime: '09:02:00', imageFileId: 'file-3', createdAt: '3' },
    { id: 'sig-deleted', trainingId: training.id, staffId: 'deleted', department: '연구부', name: '삭제된 구성원', signDate: training.date, signTime: '09:03:00', imageFileId: 'file-4', createdAt: '4' }
  ]);
  assert.deepEqual(roster.map(person => person.staffId), ['teacher', 'inactive', 'deleted']);
  assert.deepEqual(roster.map(person => person.signatureId), ['sig-teacher', 'sig-inactive', 'sig-deleted']);
  assert.ok(!roster.some(person => person.staffId === 'office'));
});

test('고정 날짜 연수는 다른 날짜의 출력 작업을 서버에서 만들지 않는다', () => {
  const harness = createHarness();
  harness.configureStatus({ id: 'training-0001', title: '고정 연수', date: '2026-07-20', daily: false }, [], []);
  assert.throws(
    () => harness.startExport_({
      trainingId: 'training-0001',
      date: '2026-07-19',
      columns: 2,
      sort: 'registration',
      outputType: 'pdf',
      showRate: true
    }),
    error => error.apiCode === 'TRAINING_DATE'
  );
});

test('기존 _DATA 작업도 서명 이미지 목록을 30개씩 이어서 읽는다', () => {
  const { readExportImageBatch_ } = createHarness();
  const rows = Array.from({ length: 35 }, (_, index) => [index, '부서', `이름${index}`, '', '', `file-${index}`]);
  rows.splice(5, 0, [999, '부서', '미서명', '', '', '']);
  const legacySheet = {
    getLastRow: () => rows.length + 1,
    getRange: () => ({ getValues: () => rows })
  };
  const spreadsheet = {
    getSheetByName: name => name === '_DATA' ? legacySheet : null
  };
  const first = readExportImageBatch_(spreadsheet, 0);
  const second = readExportImageBatch_(spreadsheet, first.processedCount);
  assert.equal(first.total, 35);
  assert.equal(first.batch.length, 30);
  assert.equal(second.batch.length, 5);
});

test('서명 이미지 하나라도 누락되면 배치 전에 출력을 중단해 원본 삭제 가능한 파일을 만들지 않는다', () => {
  const { insertExportImageBatch_ } = createHarness();
  let insertCalls = 0;
  const output = {
    insertImage: function() {
      insertCalls += 1;
      return { setWidth: function() { return this; }, setHeight: function() { return this; }, remove: function() {} };
    }
  };

  assert.throws(
    () => insertExportImageBatch_(output, [
      { layoutIndex: 0, blob: { bytes: [1] } },
      { layoutIndex: 1, blob: null }
    ], { columns: 2 }, 10),
    /원본 삭제를 막기 위해 출력을 중단/
  );
  assert.equal(insertCalls, 0);
});

test('Range 응답은 실제 받은 바이트만큼 이어받고 Range 무시 시 한 번에 끝낸다', () => {
  const { normalizeDriveDownloadResponse_ } = createHarness();
  const partial = normalizeDriveDownloadResponse_(
    206,
    { 'Content-Range': 'bytes 10-11/20' },
    [10, 11],
    10,
    14,
    20
  );
  assert.deepEqual(partial, { bytes: [10, 11], nextOffset: 12 });

  const full = normalizeDriveDownloadResponse_(200, {}, [0, 1, 2, 3, 4, 5], 2, 4, 6);
  assert.deepEqual(full, { bytes: [2, 3, 4, 5], nextOffset: 6 });

  assert.throws(
    () => normalizeDriveDownloadResponse_(206, { 'Content-Range': 'bytes 9-10/20' }, [9, 10], 10, 12, 20),
    error => error.apiCode === 'DOWNLOAD_FAILED'
  );
});

test('완료된 출력은 남은 lease가 있어도 즉시 반환한다', () => {
  const harness = createHarness();
  const leaseCalls = harness.configureTerminalExport({
    jobId: 'export-0001',
    status: 'preview_ready',
    trainingId: 'training-0001',
    date: '2026-07-20',
    outputType: 'pdf'
  });
  const result = harness.continueExport_('export-0001');
  assert.equal(result.status, 'preview_ready');
  assert.equal(leaseCalls(), 0);
});
