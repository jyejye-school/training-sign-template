import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve(import.meta.dirname, '..');
const backend = fs.readFileSync(path.join(root, 'apps-script', 'Code.gs'), 'utf8');
const loadHelpers = new Function('Utilities', 'LockService', `${backend}
return {
  isTrainingPublicOnDate_: isTrainingPublicOnDate_,
  signatureMatchesTrainingDate_: signatureMatchesTrainingDate_,
  parseSignatureSnapshot_: parseSignatureSnapshot_,
  signatureBelongsToExportPurge_: signatureBelongsToExportPurge_,
  exportSignatureTime_: exportSignatureTime_,
  buildTrainingSignatureStatus_: buildTrainingSignatureStatus_,
  getTrainingSignatureStatus_: getTrainingSignatureStatus_,
  listRecords_: listRecords_,
  getPublicData_: getPublicData_,
  readScopedSignatures_: readScopedSignatures_,
  configurePublic: function(settings, staff, trainings, today) {
    requireInitialized_ = function() {};
    requireShareToken_ = function() {};
    readSettings_ = function() { return settings; };
    privacyReady_ = function() { return true; };
    today_ = function() { return today; };
    readRows_ = function(definition) {
      if (definition === SHEETS.STAFF) return staff;
      if (definition === SHEETS.TRAININGS) return trainings;
      return [];
    };
  },
  configureRows: function(training, staff, signatures) {
    findRow_ = function(definition) {
      if (definition === SHEETS.TRAININGS) return training;
      return null;
    };
    readRows_ = function(definition) {
      if (definition === SHEETS.STAFF) return staff;
      if (definition === SHEETS.SIGNATURES) return signatures;
      return [];
    };
  },
  submit: function(training, person, signatures, serverDate, serverTime, requestOverrides) {
    let appended = null;
    let trashed = false;
    let blobName = '';
    requireInitialized_ = function() {};
    requireShareToken_ = function() {};
    id_ = function(value) { return String(value); };
    findRow_ = function(definition) {
      if (definition === SHEETS.TRAININGS) return training;
      if (definition === SHEETS.STAFF) return person;
      return null;
    };
    readRows_ = function(definition) {
      return definition === SHEETS.SIGNATURES ? signatures : [];
    };
    invalidateRows_ = function() {};
    appendObject_ = function(definition, row) {
      appended = row;
      signatures.push(row);
    };
    today_ = function() { return serverDate; };
    formatDate_ = function(date, pattern) {
      if (pattern === 'yyyy-MM-dd') return serverDate;
      if (pattern === 'HH:mm:ss') return serverTime;
      if (pattern === 'HH:mm') return serverTime.slice(0, 5);
      return serverDate;
    };
    getOrCreateTrainingFolder_ = function() {
      return {
        createFile: function(blob) {
          blobName = blob.fileName;
          return {
            getId: function() { return 'private-signature-file'; },
            setTrashed: function(value) { trashed = Boolean(value); }
          };
        }
      };
    };
    const png = new Uint8Array(120);
    png.set([137, 80, 78, 71], 0);
    const request = Object.assign({
      shareToken: 'valid-share-token',
      trainingId: training.id,
      staffId: person.id,
      signatureData: 'data:image/png;base64,' + Buffer.from(png).toString('base64')
    }, requestOverrides || {});
    try {
      return {
        result: submitSignature_(request),
        appended: function() { return appended; },
        trashed: function() { return trashed; },
        blobName: function() { return blobName; }
      };
    } catch (error) {
      error.testState = { appended: appended, trashed: trashed, blobName: blobName };
      throw error;
    }
  },
  validate: function(training, person, today, time) {
    today_ = function() { return today; };
    formatDate_ = function(date, pattern) {
      if (pattern === 'yyyy-MM-dd') return today;
      if (pattern === 'HH:mm') return time;
      if (pattern === 'HH:mm:ss') return time + ':00';
      return today;
    };
    return validateSigningWindow_(training, person);
  }
};`);

function harness() {
  const Utilities = {
    base64Decode: value => Array.from(Buffer.from(value, 'base64')),
    newBlob: (bytes, type, fileName) => ({ bytes, type, fileName }),
    getUuid: () => 'signature-uuid'
  };
  const lock = { waitLock() {}, releaseLock() {} };
  const LockService = { getScriptLock: () => lock };
  return loadHelpers(Utilities, LockService);
}

const activePerson = { id: 'staff-0001', active: true };

test('활성화한 과거 고정 연수는 공개하고 미래·비활성 연수는 공개하지 않는다', () => {
  const { isTrainingPublicOnDate_ } = harness();
  assert.equal(isTrainingPublicOnDate_({ active: true, daily: false, date: '2026-07-13' }, '2026-07-20'), true);
  assert.equal(isTrainingPublicOnDate_({ active: true, daily: false, date: '2026-07-20' }, '2026-07-20'), true);
  assert.equal(isTrainingPublicOnDate_({ active: true, daily: false, date: '2026-07-21' }, '2026-07-20'), false);
  assert.equal(isTrainingPublicOnDate_({ active: false, daily: false, date: '2026-07-13' }, '2026-07-20'), false);
  assert.equal(isTrainingPublicOnDate_({ active: true, daily: true, date: '' }, '2026-07-20'), true);
});

test('공개 데이터 API는 활성 과거 연수와 매일 연수만 실제 응답에 포함한다', () => {
  const api = harness();
  const trainings = [
    { id: 'past-training', title: '지난 연수', active: true, daily: false, date: '2026-07-13', sortOrder: 1 },
    { id: 'future-training', title: '미래 연수', active: true, daily: false, date: '2026-07-21', sortOrder: 2 },
    { id: 'inactive-training', title: '비활성 연수', active: false, daily: false, date: '2026-07-13', sortOrder: 3 },
    { id: 'daily-training', title: '매일 연수', active: true, daily: true, date: '', sortOrder: 4 }
  ];
  api.configurePublic(
    { schoolName: '테스트 학교', privacyPurpose: '연수 확인', privacyItems: '성명', privacyRetention: '출력 뒤 삭제' },
    [{ id: 'staff-0001', active: true, department: '교무부', name: '홍길동', sortOrder: 1 }],
    trainings,
    '2026-07-20'
  );
  const data = api.getPublicData_('valid-share-token');
  assert.deepEqual(data.trainings.map(training => training.id), ['past-training', 'daily-training']);
  assert.equal(data.serverDate, '2026-07-20');
});

test('과거 고정 연수는 당일 운영 시각을 건너뛰고 미래 연수는 직접 제출도 거부한다', () => {
  const api = harness();
  assert.doesNotThrow(() => api.validate(
    { active: true, daily: false, date: '2026-07-13', startTime: '09:00', endTime: '10:00' },
    activePerson,
    '2026-07-20',
    '18:30'
  ));
  assert.throws(
    () => api.validate(
      { active: true, daily: false, date: '2026-07-21', startTime: '', endTime: '' },
      activePerson,
      '2026-07-20',
      '09:30'
    ),
    error => error.apiCode === 'TRAINING_DATE'
  );
  assert.throws(
    () => api.validate(
      { active: false, daily: false, date: '2026-07-13', startTime: '', endTime: '' },
      activePerson,
      '2026-07-20',
      '09:30'
    ),
    error => error.apiCode === 'TRAINING_CLOSED'
  );
});

test('오늘 고정 연수와 매일 연수에는 현재 서울 시각 제한을 계속 적용한다', () => {
  const fixed = { active: true, daily: false, date: '2026-07-20', startTime: '09:00', endTime: '10:00' };
  const daily = { active: true, daily: true, date: '', startTime: '09:00', endTime: '10:00' };
  assert.throws(() => harness().validate(fixed, activePerson, '2026-07-20', '08:59'), error => error.apiCode === 'TOO_EARLY');
  assert.doesNotThrow(() => harness().validate(fixed, activePerson, '2026-07-20', '09:00'));
  assert.doesNotThrow(() => harness().validate(fixed, activePerson, '2026-07-20', '10:00'));
  assert.throws(() => harness().validate(fixed, activePerson, '2026-07-20', '10:01'), error => error.apiCode === 'TOO_LATE');
  assert.throws(() => harness().validate(daily, activePerson, '2026-07-20', '08:59'), error => error.apiCode === 'TOO_EARLY');
  assert.doesNotThrow(() => harness().validate(daily, activePerson, '2026-07-20', '09:30'));
});

test('고정 연수 중복·현황·출력은 실제 제출일이 달라도 같은 연수로 묶는다', () => {
  const api = harness();
  const fixed = { id: 'training-0001', active: true, daily: false, date: '2026-07-13' };
  const daily = { id: 'training-0002', active: true, daily: true, date: '' };
  const signature = {
    trainingId: fixed.id,
    staffId: activePerson.id,
    signDate: '2026-07-20',
    scopeDate: '2026-07-13',
    signTime: '14:25:30',
    createdAt: '2026-07-20T05:25:30.000Z'
  };
  const dailySignature = Object.assign({}, signature, {
    trainingId: daily.id,
    scopeDate: '2026-07-20'
  });

  assert.equal(api.signatureMatchesTrainingDate_(signature, fixed, '2026-07-20'), true);
  assert.equal(api.signatureMatchesTrainingDate_(signature, fixed, '2026-07-13'), true);
  assert.equal(api.signatureMatchesTrainingDate_(dailySignature, daily, '2026-07-20'), true);
  assert.equal(api.signatureMatchesTrainingDate_(dailySignature, daily, '2026-07-21'), false);
  assert.equal(api.exportSignatureTime_(signature, fixed, fixed.date), '07.20 14:25');

  const status = api.buildTrainingSignatureStatus_(
    fixed.id,
    fixed.date,
    [{ id: activePerson.id, active: true, department: '교무부', name: '홍길동', sortOrder: 1 }],
    [signature]
  );
  assert.equal(status.summary.signedCount, 1);
  assert.equal(status.people[0].signDate, '2026-07-20');
  assert.equal(status.people[0].signTime, '14:25');
});

test('과거 고정 연수 제출은 실제 서버 날짜·시각을 저장하고 요청의 날짜 값을 무시한다', () => {
  const api = harness();
  const training = { id: 'training-0001', title: '과거 연수', active: true, daily: false, date: '2026-07-13', startTime: '', endTime: '' };
  const person = { id: 'staff-0001', active: true, department: '교무부', name: '홍길동' };
  const submission = api.submit(training, person, [], '2026-07-20', '14:25:30', {
    signDate: '1999-01-01',
    signTime: '00:00:00',
    createdAt: '1999-01-01T00:00:00.000Z'
  });
  const row = submission.appended();

  assert.equal(submission.result.signDate, '2026-07-20');
  assert.equal(submission.result.signTime, '14:25:30');
  assert.equal(row.signDate, '2026-07-20');
  assert.equal(row.signTime, '14:25:30');
  assert.equal(row.scopeDate, '2026-07-13');
  assert.ok(!String(row.createdAt).startsWith('1999-01-01'));
  assert.match(submission.blobName(), /^2026-07-20_/);
});

test('과거 고정 연수는 실제 제출일이 달라도 재제출을 막고 임시 이미지를 폐기한다', () => {
  const api = harness();
  const training = { id: 'training-0001', title: '과거 연수', active: true, daily: false, date: '2026-07-13', startTime: '', endTime: '' };
  const person = { id: 'staff-0001', active: true, department: '교무부', name: '홍길동' };
  const signatures = [{ trainingId: training.id, staffId: person.id, signDate: '2026-07-13' }];

  assert.throws(
    () => api.submit(training, person, signatures, '2026-07-20', '14:25:30'),
    error => error.apiCode === 'DUPLICATE' && error.testState.trashed === true && error.testState.appended === null
  );
});

test('고정 연수 현황과 기록 조회는 연수일을 선택해 실제 제출일의 서명도 반환한다', () => {
  const api = harness();
  const training = { id: 'training-0001', title: '과거 연수', active: true, daily: false, date: '2026-07-13' };
  const staff = [{ id: 'staff-0001', active: true, department: '교무부', name: '홍길동', sortOrder: 1 }];
  const signatures = [{
    id: 'signature-0001',
    trainingId: training.id,
    staffId: staff[0].id,
    signDate: '2026-07-20',
    scopeDate: '2026-07-13',
    signTime: '14:25:30',
    department: '교무부',
    name: '홍길동',
    createdAt: '2026-07-20T05:25:30.000Z'
  }];
  api.configureRows(training, staff, signatures);

  const status = api.getTrainingSignatureStatus_(training.id, training.date);
  const records = api.listRecords_(training.id, training.date);
  assert.equal(status.summary.signedCount, 1);
  assert.equal(status.people[0].signDate, '2026-07-20');
  assert.equal(records.records.length, 1);
  assert.equal(records.records[0].signDate, '2026-07-20');
});

test('연수 방식을 바꿔도 다른 기준일의 서명은 고정 연수에 섞이지 않는다', () => {
  const api = harness();
  const fixed = { id: 'training-0001', active: true, daily: false, date: '2026-07-13' };
  const july13 = { trainingId: fixed.id, signDate: '2026-07-13', scopeDate: '2026-07-13' };
  const july14 = { trainingId: fixed.id, signDate: '2026-07-14', scopeDate: '2026-07-14' };
  const legacyJuly14 = { trainingId: fixed.id, signDate: '2026-07-14' };

  assert.equal(api.signatureMatchesTrainingDate_(july13, fixed, fixed.date), true);
  assert.equal(api.signatureMatchesTrainingDate_(july14, fixed, fixed.date), false);
  assert.equal(api.signatureMatchesTrainingDate_(legacyJuly14, fixed, fixed.date), false);
});

test('출력 당시 서명 ID 스냅샷은 이후 추가된 서명을 원본 삭제 대상에 넣지 않는다', () => {
  const api = harness();
  const snapshot = api.parseSignatureSnapshot_(JSON.stringify(['signature-1', 'signature-2']));
  const job = { trainingId: 'training-0001', date: '2026-07-13' };
  const included = { id: 'signature-1', trainingId: job.trainingId, signDate: '2026-07-20', scopeDate: job.date };
  const late = { id: 'signature-late', trainingId: job.trainingId, signDate: '2026-07-20', scopeDate: job.date };
  assert.equal(snapshot.has('signature-1'), true);
  assert.equal(snapshot.has('signature-2'), true);
  assert.equal(snapshot.has('signature-late'), false);
  assert.equal(api.signatureBelongsToExportPurge_(included, job, null, snapshot), true);
  assert.equal(api.signatureBelongsToExportPurge_(late, job, null, snapshot), false);
  const changedTraining = { id: job.trainingId, daily: false, date: '2026-07-14' };
  const legacyIncluded = { id: 'legacy-1', trainingId: job.trainingId, signDate: '2026-07-13' };
  const differentDate = { id: 'legacy-2', trainingId: job.trainingId, signDate: '2026-07-14' };
  assert.equal(api.signatureBelongsToExportPurge_(legacyIncluded, job, changedTraining, null), true);
  assert.equal(api.signatureBelongsToExportPurge_(differentDate, job, changedTraining, null), false);
  assert.equal(api.parseSignatureSnapshot_(''), null);
  assert.throws(() => api.parseSignatureSnapshot_('{"bad":true}'), error => error.apiCode === 'EXPORT_DATA');
});
