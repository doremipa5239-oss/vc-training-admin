/****************************************************
 * [7_Operations.gs] 교육 운영 워크스페이스
 * - VupOperationsModel: 명단·체크리스트·문서 필드 정규화 모델
 * - 운영 기록 시트 저장/불러오기, 운영 문서(HWPX/DOCX/TXT) 본문 추출
 ****************************************************/

(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.VupOperationsModel = factory();
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  const MEMBERSHIPS = { member: '회원사', special: '특별회원사', lp: '회원사 LP', general: '비회원·일반' };
  const MAX_PARTICIPANTS = 500;
  function text(value, limit) { return String(value == null ? '' : value).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').trim().slice(0, limit || 200); }
  function pick(value, options, fallback) { return options.indexOf(value) >= 0 ? value : fallback; }
  function money(value, label, nullable) {
    if (nullable && (value == null || value === '')) return null;
    const number = Number(value == null || value === '' ? 0 : value);
    if (!Number.isSafeInteger(number) || number < 0 || number > 1000000000) throw new Error((label || '금액') + '은 0~10억 원의 정수여야 합니다.');
    return number;
  }
  function normalizeState(input) {
    const src = input || {};
    if (src.schemaVersion != null && src.schemaVersion !== 1) throw new Error('지원하지 않는 운영 기록 버전입니다. 최신 관리 화면에서 열어주세요.');
    const fees = {};
    Object.keys(MEMBERSHIPS).forEach(function (key) { fees[key] = money((src.fees || {})[key], MEMBERSHIPS[key] + ' 교육비', true); });
    if (!Array.isArray(src.participants || [])) throw new Error('교육생 명단 형식이 잘못되었습니다.');
    if ((src.participants || []).length > MAX_PARTICIPANTS) throw new Error('교육생은 최대 500명까지 저장할 수 있습니다.');
    const ids = new Set();
    const participants = (src.participants || []).map(function (p, index) {
      if (!p || typeof p !== 'object' || Array.isArray(p)) throw new Error((index + 1) + '행의 교육생 정보를 확인해주세요.');
      const id = text(p.id, 80);
      if (!id || !/^[a-zA-Z0-9_-]+$/.test(id) || ids.has(id)) throw new Error((index + 1) + '행의 교육생 ID가 없거나 중복됩니다.');
      ids.add(id);
      const name = text(p.name, 80);
      if (!name) throw new Error((index + 1) + '행의 이름을 입력해주세요.');
      const age = p.age == null || p.age === '' ? null : Number(p.age);
      if (age !== null && (!Number.isInteger(age) || age < 1 || age > 120)) throw new Error(name + '의 나이를 확인해주세요.');
      const email = text(p.email, 200);
      if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error(name + '의 이메일을 확인해주세요.');
      return {
        id: id, name: name, company: text(p.company, 120), position: text(p.position, 80),
        phone: text(p.phone, 30), email: email, membership: pick(p.membership, Object.keys(MEMBERSHIPS), 'general'),
        status: pick(p.status, ['pending', 'approved', 'cancelled'], 'pending'),
        feeAmount: money(p.feeAmount, name + ' 교육비', true), paidAmount: money(p.paidAmount, name + ' 납부액'),
        paymentMethod: pick(p.paymentMethod, ['', 'card', 'transfer'], ''),
        invoiceStatus: pick(p.invoiceStatus, ['', 'requested', 'issued'], ''),
        attendance: pick(p.attendance, ['unknown', 'present', 'absent'], 'unknown'),
        notes: text(p.notes, 1500), gender: pick(p.gender, ['male', 'female', 'unknown'], 'unknown'),
        age: age, snoring: pick(p.snoring, ['yes', 'no', 'unknown'], 'unknown'),
        roomType: pick(p.roomType, ['twin', 'single', 'none'], 'twin'),
        onsiteIncluded: p.onsiteIncluded !== false && p.onsiteIncluded !== 'false',
        excludedMeals: Array.from(new Set((Array.isArray(p.excludedMeals) ? p.excludedMeals : []).map(function (v) { return text(v, 60); }))).slice(0, 30)
      };
    });
    if (src.tasks != null && !Array.isArray(src.tasks)) throw new Error('체크리스트 형식이 잘못되었습니다.');
    if ((src.tasks || []).length > 150) throw new Error('체크리스트는 150개까지 저장할 수 있습니다.');
    if (Object.keys(src.documentFields || {}).length > 70) throw new Error('문서 필드는 70개까지 저장할 수 있습니다.');
    const taskIds = new Set();
    const tasks = (src.tasks || []).map(function (task) {
      if (!task || !task.id || taskIds.has(String(task.id))) throw new Error('체크리스트 ID가 없거나 중복됩니다.');
      taskIds.add(String(task.id));
      return { id: text(task.id, 80), title: text(task.title || task.label, 200), stage: text(task.stage, 40),
        dueDate: text(task.dueDate, 10), status: pick(task.status, ['todo', 'doing', 'done'], 'todo'),
        owner: text(task.owner, 80), note: text(task.note || task.notes, 500), review: {
          first: pick((task.review || {}).first, ['pending', 'done'], 'pending'),
          second: pick((task.review || {}).second, ['pending', 'done'], 'pending'),
          manager: pick((task.review || {}).manager, ['pending', 'done'], 'pending'),
          director: pick((task.review || {}).director, ['pending', 'done'], 'pending') } };
    });
    const documentFields = {};
    Object.keys(src.documentFields || {}).slice(0, 70).forEach(function (key) {
      if (!/^[a-zA-Z][a-zA-Z0-9]*$/.test(key) || ['__proto__', 'constructor', 'prototype'].indexOf(key) >= 0) return;
      const value = src.documentFields[key];
      if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') documentFields[key] = typeof value === 'string' ? text(value, 6000) : value;
      else if (Array.isArray(value)) { if (value.length > 100) throw new Error('문서의 표는 100행까지 저장할 수 있습니다.'); documentFields[key] = JSON.parse(JSON.stringify(value)); }
      else if (value && typeof value === 'object' && ['planBudget', 'actualBudget', 'fees'].includes(key)) {
        documentFields[key] = {};
        (key === 'fees' ? Object.keys(MEMBERSHIPS) : ['income', 'expense']).forEach(function (field) {
          if (value[field] != null && value[field] !== '') documentFields[key][field] = money(value[field], '문서 금액');
        });
      }
    });
    const plans = src.plans && typeof src.plans === 'object' ? JSON.parse(JSON.stringify(src.plans)) : {};
    const result = { schemaVersion: 1, fees: fees, participants: participants, tasks: tasks, documentFields: documentFields, plans: plans };
    if (JSON.stringify(result).length > 240000) throw new Error('운영 기록이 저장 한도에 도달했습니다. 메모 또는 배정 이력을 줄여주세요.');
    return result;
  }
  function amountDue(person, fees) { return person.status === 'cancelled' ? 0 : person.feeAmount == null ? fees[person.membership] == null ? null : Number(fees[person.membership]) : person.feeAmount; }
  function summarize(state) {
    const result = { total: state.participants.length, approved: 0, pending: 0, cancelled: 0, present: 0, expected: 0, paid: 0, outstanding: 0, refundReview: 0, notes: 0, unpriced: 0 };
    state.participants.forEach(function (p) {
      result[p.status]++;
      if (p.attendance === 'present') result.present++;
      if (p.notes) result.notes++;
      const due = amountDue(p, state.fees);
      if (due === null) result.unpriced++;
      result.expected += due;
      result.paid += p.paidAmount;
      if (due !== null) { result.outstanding += Math.max(0, due - p.paidAmount); result.refundReview += Math.max(0, p.paidAmount - due); }
    });
    result.completedTasks = state.tasks.filter(function (t) { return t.status === 'done'; }).length;
    return result;
  }
  function normalizeHeader(v) { return text(v).toLowerCase().replace(/[\s_()·.-]/g, ''); }
  const HEADERS = {
    id: ['교육생id', 'id'], name: ['이름', '성명', '교육생명', 'name'], company: ['소속', '회사명', '소속기관', '기관명', 'company'],
    position: ['직급', '직위', '직책', 'position'], phone: ['휴대폰번호', '휴대폰', '전화번호', '연락처', 'phone'],
    email: ['이메일', 'email'], membership: ['회원구분', '회원유형', 'membership'], gender: ['성별', 'gender'], age: ['나이', '연령', 'age'],
    snoring: ['코골이', '코골이여부', 'snoring'], roomType: ['객실유형', '객실', '숙박', 'roomtype'], notes: ['특이사항', '메모', 'notes'],
    onsiteIncluded: ['전체배정', '배정대상', 'onsiteincluded'], excludedMeals: ['불참식사id', '불참식사', 'excludedmeals'],
    paidAmount: ['납부액', '입금액', '결제금액', 'paidamount'], feeAmount: ['교육비', '청구액', 'feeamount']
  };
  function importRows(rows, idPrefix) {
    if (!Array.isArray(rows) || rows.length > 2012) throw new Error('명단은 2,000행 이내의 표로 올려주세요.');
    let mapping, start = -1;
    for (let i = 0; i < Math.min(rows.length, 12); i++) {
      const normalized = (rows[i] || []).map(normalizeHeader);
      const candidate = {};
      Object.keys(HEADERS).forEach(function (key) { candidate[key] = normalized.findIndex(function (v) { return HEADERS[key].indexOf(v) >= 0; }); });
      if (candidate.name >= 0) { mapping = candidate; start = i; break; }
    }
    if (start < 0) throw new Error('이름 또는 성명 열 제목을 찾지 못했습니다. 첫 12행 안에 열 제목을 넣어주세요.');
    const result = [];
    rows.slice(start + 1).forEach(function (row, index) {
      if (!(row || []).some(function (v) { return text(v); })) return;
      const person = {};
      Object.keys(mapping).forEach(function (key) { person[key] = mapping[key] >= 0 ? text(row[mapping[key]], 1500) : ''; });
      if (!person.name) throw new Error((start + index + 2) + '행의 이름이 비어 있습니다.');
      person.id = person.id || (idPrefix || 'import') + '_' + (index + 1);
      const memberships = { '회원사': 'member', '특별회원사': 'special', '회원사lp': 'lp', '비회원': 'general', '일반': 'general', '비회원·일반': 'general' };
      person.membership = memberships[person.membership.toLowerCase().replace(/\s/g, '')] || person.membership || 'general';
      if (!Object.prototype.hasOwnProperty.call(MEMBERSHIPS, person.membership)) throw new Error(person.name + '의 회원구분을 확인해주세요.');
      person.gender = ({ '남': 'male', '남성': 'male', '남자': 'male', '여': 'female', '여성': 'female', '여자': 'female' })[person.gender] || person.gender || 'unknown';
      person.snoring = ({ '예': 'yes', '있음': 'yes', 'o': 'yes', '아니오': 'no', '없음': 'no', 'x': 'no' })[person.snoring.toLowerCase()] || person.snoring || 'unknown';
      person.roomType = ({ '1인실': 'single', '2인실': 'twin', '비숙박': 'none', '없음': 'none' })[person.roomType] || person.roomType || 'twin';
      if (person.onsiteIncluded && !['포함', '제외', 'true', 'false'].includes(person.onsiteIncluded)) throw new Error(person.name + '의 전체 배정은 포함 또는 제외로 입력해주세요.');
      person.onsiteIncluded = !['제외', 'false'].includes(person.onsiteIncluded);
      person.excludedMeals = person.excludedMeals ? person.excludedMeals.split(';').map(function (v) { return v.trim(); }).filter(Boolean) : [];
      person.paidAmount = person.paidAmount.replace(/[,원\s]/g, '');
      person.feeAmount = person.feeAmount.replace(/[,원\s]/g, '');
      result.push(person);
    });
    return normalizeState({ participants: result }).participants;
  }
  function csvCell(value) { let v = String(value == null ? '' : value); if (/^[\s\uFEFF]*[=+@-]/.test(v)) v = "'" + v; return '"' + v.replace(/"/g, '""') + '"'; }
  function toCsv(rows) { return '\uFEFF' + rows.map(function (row) { return row.map(csvCell).join(','); }).join('\r\n'); }
  function fieldParticipants(state) { return state.participants.filter(function (p) { return p.onsiteIncluded !== false; }); }
  function summarizeField(state) {
    const people = fieldParticipants(state), rooms = people.filter(function (p) { return p.roomType !== 'none'; });
    const prep = state.tasks.filter(function (t) { return ['preparation', 'onsite'].includes(t.stage); });
    const closing = state.tasks.filter(function (t) { return t.stage === 'closing'; });
    return { people: people.length, excluded: state.participants.length - people.length,
      companies: new Set(people.map(function (p) { return p.company.replace(/\s/g, '').toLowerCase(); }).filter(Boolean)).size,
      lodging: rooms.length, singles: rooms.filter(function (p) { return p.roomType === 'single'; }).length,
      missingProfiles: rooms.filter(function (p) { return p.gender === 'unknown' || p.snoring === 'unknown' || p.age == null || !p.company; }).length,
      mealExceptions: people.filter(function (p) { return p.excludedMeals.length; }).length,
      preparation: prep.length, preparationDone: prep.filter(function (t) { return t.status === 'done'; }).length,
      closing: closing.length, closingDone: closing.filter(function (t) { return t.status === 'done'; }).length };
  }
  function importFieldRows(rows, idPrefix) {
    // The upstream registration system owns approvals, tuition, payments and attendance.
    // Drop those columns before parsing, so unrelated values cannot block the final roster import.
    const ignored = ['membership', 'paidAmount', 'feeAmount'];
    const excludedHeaders = ignored.flatMap(function (key) { return HEADERS[key]; }).map(normalizeHeader);
    if (!Array.isArray(rows) || rows.length > 2012) throw new Error('명단은 2,000행 이내의 표로 올려주세요.');
    const headerIndex = rows.slice(0, 12).findIndex(function (row) { return (row || []).some(function (v) { return HEADERS.name.includes(normalizeHeader(v)); }); });
    const ignoredColumns = headerIndex < 0 ? [] : rows[headerIndex].map(function (v, index) { return excludedHeaders.includes(normalizeHeader(v)) ? index : -1; });
    return importRows(rows.map(function (row, rowIndex) { return (row || []).map(function (v, column) { return rowIndex >= headerIndex && ignoredColumns.includes(column) ? '' : v; }); }), idPrefix);
  }
  return { MEMBERSHIPS: MEMBERSHIPS, normalizeState: normalizeState, summarize: summarize, amountDue: amountDue, importRows: importRows, toCsv: toCsv,
    fieldParticipants: fieldParticipants, summarizeField: summarizeField, importFieldRows: importFieldRows };
});


/** Private course operations. Data in this sheet is never part of a public-site response. */
const OPERATIONS_SHEET_ = '교육운영워크스페이스';
const OPERATIONS_CHUNKS_ = 10;
const OPERATIONS_CHUNK_SIZE_ = 28000;
const OPERATIONS_CHUNK_PREFIX_ = 'VUPJSON1:';
const OPERATIONS_MAX_JSON_ = 240000;
const DOCUMENT_MAX_BYTES_ = 5 * 1024 * 1024;
const DOCUMENT_MAX_EXPANDED_BYTES_ = 16 * 1024 * 1024;
const DOCUMENT_MAX_TEXT_ = 150000;

function getOperationsSheet_() {
  return getOrCreatePortalSheet_(OPERATIONS_SHEET_, ['강의ID', '수정번호', '수정시각', '수정관리자'].concat(
    Array.from({ length: OPERATIONS_CHUNKS_ }, function (_, i) { return '데이터' + (i + 1); })
  ));
}

function readOperations_(siteId) {
  const sheet = getOperationsSheet_();
  const row = findPortalRowByValue_(sheet, 1, siteId);
  if (row < 2) return { row: 0, revision: 0, updatedAt: '', state: VupOperationsModel.normalizeState({}) };
  const values = sheet.getRange(row, 1, 1, 4 + OPERATIONS_CHUNKS_).getValues()[0];
  const revision = Number(values[1]);
  if (String(values[0]) !== String(siteId) || !Number.isSafeInteger(revision) || revision < 0 || revision >= Number.MAX_SAFE_INTEGER) {
    throw new Error('저장된 운영 기록의 수정번호가 올바르지 않습니다. 기존 데이터를 덮어쓰지 말고 관리자에게 문의해주세요.');
  }
  let state;
  try {
    const cells = values.slice(4).map(function (value) { return String(value == null ? '' : value); });
    const prefixed = cells[0].indexOf(OPERATIONS_CHUNK_PREFIX_) === 0;
    const serialized = cells.map(function (cell) {
      if (!prefixed) return cell; // 이전 버전의 접두사 없는 JSON 행도 읽습니다.
      if (cell.indexOf(OPERATIONS_CHUNK_PREFIX_) !== 0 || cell.length > OPERATIONS_CHUNK_SIZE_ + OPERATIONS_CHUNK_PREFIX_.length) throw new Error('조각 형식 오류');
      return cell.slice(OPERATIONS_CHUNK_PREFIX_.length);
    }).join('');
    if (serialized.length > OPERATIONS_MAX_JSON_) throw new Error('저장 한도 초과');
    state = JSON.parse(serialized);
    validateOperationsPayload_(state);
  }
  catch (error) { throw new Error('저장된 운영 기록을 읽을 수 없습니다. 기존 데이터를 덮어쓰지 말고 관리자에게 문의해주세요.'); }
  return { row: row, revision: revision, updatedAt: String(values[2] || ''), state: VupOperationsModel.normalizeState(state) };
}

/** 모델의 JSON 복사 전에 크기·깊이를 검사하여 손상된 입력을 빠르게 거절합니다. */
function validateOperationsPayload_(state) {
  if (!state || typeof state !== 'object' || Array.isArray(state)) throw new Error('운영 기록 형식이 올바르지 않습니다.');
  const stack = [{ value: state, depth: 0 }];
  const seen = new Set();
  let nodes = 0, textLength = 0;
  while (stack.length) {
    const item = stack.pop(), value = item.value;
    if (++nodes > 30000 || item.depth > 30) throw new Error('운영 기록 구조가 너무 복잡합니다. 배정 이력 또는 메모를 줄여주세요.');
    if (typeof value === 'string') {
      textLength += value.length;
      if (textLength > OPERATIONS_MAX_JSON_) throw new Error('운영 기록의 저장 한도를 초과했습니다.');
    } else if (value && typeof value === 'object') {
      if (seen.has(value)) throw new Error('운영 기록에 반복 참조가 있습니다. 다시 불러와주세요.');
      seen.add(value);
      const keys = Object.keys(value);
      if (keys.length > 10000 || keys.length + stack.length > 30000) throw new Error('운영 기록 항목이 너무 많습니다.');
      keys.forEach(function (key) {
        textLength += key.length;
        if (textLength > OPERATIONS_MAX_JSON_) throw new Error('운영 기록의 저장 한도를 초과했습니다.');
        stack.push({ value: value[key], depth: item.depth + 1 });
      });
    } else if (value !== null && typeof value !== 'number' && typeof value !== 'boolean') {
      throw new Error('운영 기록에 지원하지 않는 값이 있습니다.');
    } else if (typeof value === 'number' && !Number.isFinite(value)) {
      throw new Error('운영 기록의 숫자 값을 확인해주세요.');
    }
  }
}

function encodeOperationsChunks_(serialized) {
  const chunks = [];
  let start = 0;
  while (start < serialized.length) {
    let end = Math.min(start + OPERATIONS_CHUNK_SIZE_, serialized.length);
    // UTF-16 서로게이트 쌍을 셀 사이에서 자르지 않아 이모지 등이 손상되지 않게 합니다.
    if (end < serialized.length && /[\uD800-\uDBFF]/.test(serialized.charAt(end - 1)) && /[\uDC00-\uDFFF]/.test(serialized.charAt(end))) end--;
    chunks.push(OPERATIONS_CHUNK_PREFIX_ + serialized.slice(start, end));
    start = end;
  }
  if (chunks.length > OPERATIONS_CHUNKS_) throw new Error('운영 기록의 저장 한도를 초과했습니다.');
  while (chunks.length < OPERATIONS_CHUNKS_) chunks.push(OPERATIONS_CHUNK_PREFIX_);
  return chunks;
}

function portalGetOperationsWorkspace(token, siteId) {
  return withAdminSession_(token, function (admin) {
    const id = normalizePublicSiteId_(siteId);
    assertCourseAccess_(admin, id);
    const record = readOperations_(id);
    return { revision: record.revision, updatedAt: record.updatedAt, state: record.state };
  });
}

function portalSaveOperationsWorkspace(token, payload) {
  return withAdminSession_(token, function (admin) {
    const input = payload || {};
    const id = normalizePublicSiteId_(input.siteId);
    assertCourseAccess_(admin, id);
    if (!Number.isSafeInteger(input.revision) || input.revision < 0 || input.revision >= Number.MAX_SAFE_INTEGER) throw new Error('수정번호가 없습니다. 운영 기록을 다시 불러와주세요.');
    validateOperationsPayload_(input.state);
    const normalized = VupOperationsModel.normalizeState(input.state);
    const serialized = JSON.stringify(normalized);
    if (serialized.length > OPERATIONS_MAX_JSON_) throw new Error('운영 기록의 저장 한도를 초과했습니다.');
    return withCourseMutation_(admin, id, null, function () {
      const current = readOperations_(id);
      if (current.revision !== input.revision) throw new Error('다른 관리자가 먼저 저장했습니다. 현재 내용을 내보낸 뒤 최신 기록을 다시 불러와 병합해주세요.');
      const now = new Date().toISOString();
      // 모든 조각을 텍스트 접두사로 시작해 중간 조각의 '=' 등이 수식으로 평가되지 않게 합니다.
      const chunks = encodeOperationsChunks_(serialized);
      const sheet = getOperationsSheet_();
      sheet.getRange(current.row || sheet.getLastRow() + 1, 1, 1, 4 + OPERATIONS_CHUNKS_).setValues([[id, current.revision + 1, now, admin.id].concat(chunks)]);
      SpreadsheetApp.flush();
      appendAuditLog_('OPERATIONS_SAVED', id, { revision: current.revision + 1, participants: normalized.participants.length, tasks: normalized.tasks.length });
      return { success: true, revision: current.revision + 1, updatedAt: now };
    });
  });
}

/** HWPX and DOCX use a ZIP container. Reject oversized containers before unzip. */
function validateDocumentArchive_(bytes) {
  if (!Array.isArray(bytes) || bytes.length < 22 || bytes.length > DOCUMENT_MAX_BYTES_) throw new Error('문서 크기 또는 압축 형식이 올바르지 않습니다.');
  function u16(i) { return (bytes[i] & 255) | ((bytes[i + 1] & 255) << 8); }
  function u32(i) { return (u16(i) + u16(i + 2) * 65536) >>> 0; }
  function sameBytes(a, b, length) {
    for (let i = 0; i < length; i++) if ((bytes[a + i] & 255) !== (bytes[b + i] & 255)) return false;
    return true;
  }
  function checkExtra(start, length) {
    const end = start + length;
    while (start < end) {
      if (start + 4 > end || u16(start) === 1) throw new Error('ZIP64 또는 손상된 문서는 지원하지 않습니다.');
      start += 4 + u16(start + 2);
      if (start > end) throw new Error('문서 압축 정보를 확인할 수 없습니다.');
    }
  }
  let end = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (u32(i) === 0x06054b50 && i + 22 + u16(i + 20) === bytes.length) { end = i; break; }
  }
  if (end < 0 || u16(end + 4) || u16(end + 6) || u16(end + 8) !== u16(end + 10)) throw new Error('지원하지 않는 문서 압축 형식입니다.');
  const entries = u16(end + 10), size = u32(end + 12), offset = u32(end + 16);
  if (!entries || entries > 500 || size === 0xffffffff || offset === 0xffffffff || offset + size !== end) throw new Error('문서에 파일이 너무 많거나 압축 형식이 올바르지 않습니다.');
  let pos = offset, total = 0;
  const records = [], names = new Set(), ranges = [];
  for (let n = 0; n < entries; n++) {
    if (pos + 46 > end || u32(pos) !== 0x02014b50) throw new Error('손상된 문서입니다.');
    const flags = u16(pos + 8), method = u16(pos + 10), crc = u32(pos + 16);
    const compressed = u32(pos + 20), expanded = u32(pos + 24), local = u32(pos + 42);
    const nameLength = u16(pos + 28), extraLength = u16(pos + 30), next = pos + 46 + nameLength + extraLength + u16(pos + 32);
    total += expanded;
    if ((flags & 0x61) || [0, 8].indexOf(method) < 0 || u16(pos + 34) !== 0 || !nameLength || nameLength > 1024 || next > end) throw new Error('암호화되었거나 지원하지 않는 압축 문서입니다.');
    if (expanded > 8 * 1024 * 1024 || total > DOCUMENT_MAX_EXPANDED_BYTES_ || compressed > DOCUMENT_MAX_BYTES_ || (method === 0 && compressed !== expanded)) throw new Error('문서 압축 해제 크기가 너무 크거나 압축 정보가 일치하지 않습니다.');
    if (local + 30 > offset || u32(local) !== 0x04034b50 || u16(local + 6) !== flags || u16(local + 8) !== method || u16(local + 26) !== nameLength) throw new Error('문서의 압축 헤더가 일치하지 않습니다.');
    const localExtraLength = u16(local + 28), dataStart = local + 30 + nameLength + localExtraLength;
    if (dataStart > offset || !sameBytes(local + 30, pos + 46, nameLength)) throw new Error('문서 파일 이름 또는 압축 범위가 일치하지 않습니다.');
    checkExtra(pos + 46 + nameLength, extraLength);
    checkExtra(local + 30 + nameLength, localExtraLength);
    const name = Utilities.newBlob(bytes.slice(pos + 46, pos + 46 + nameLength)).getDataAsString('UTF-8');
    if (!name || /[\u0000-\u001f\u007f\\]/.test(name) || /^(?:\/|[A-Za-z]:)/.test(name) || name.split('/').indexOf('..') >= 0 || names.has(name)) throw new Error('문서 파일 이름이 올바르지 않거나 중복됩니다.');
    names.add(name);
    let dataEnd = dataStart + compressed;
    if (dataEnd > offset) throw new Error('문서 압축 데이터 범위가 올바르지 않습니다.');
    if (flags & 8) {
      if ((u32(local + 14) && u32(local + 14) !== crc) || (u32(local + 18) && u32(local + 18) !== compressed) || (u32(local + 22) && u32(local + 22) !== expanded)) throw new Error('문서 압축 크기가 일치하지 않습니다.');
      const descriptor = u32(dataEnd) === 0x08074b50 ? dataEnd + 4 : dataEnd;
      if (descriptor + 12 > offset || u32(descriptor) !== crc || u32(descriptor + 4) !== compressed || u32(descriptor + 8) !== expanded) throw new Error('문서 압축 종료 정보가 일치하지 않습니다.');
      dataEnd = descriptor + 12;
    } else if (u32(local + 14) !== crc || u32(local + 18) !== compressed || u32(local + 22) !== expanded) {
      throw new Error('문서 압축 크기가 일치하지 않습니다.');
    }
    ranges.push({ start: local, end: dataEnd });
    records.push({ name: name, expanded: expanded });
    pos = next;
  }
  if (pos !== end) throw new Error('문서 압축 정보를 확인할 수 없습니다.');
  ranges.sort(function (a, b) { return a.start - b.start; });
  for (let i = 1; i < ranges.length; i++) if (ranges[i].start < ranges[i - 1].end) throw new Error('문서 압축 데이터 범위가 겹칩니다.');
  return records;
}

function portalExtractOperatingDocument(token, payload) {
  return withAdminSession_(token, function (admin) {
    const input = payload || {};
    assertCourseAccess_(admin, normalizePublicSiteId_(input.siteId));
    const name = String(input.name || '').trim();
    const data = String(input.base64 || '');
    if (name.length > 200 || !/\.(hwpx|docx|txt)$/i.test(name)) throw new Error('HWPX, DOCX 또는 TXT 문서를 선택해주세요. HWP·PDF는 텍스트를 붙여넣어주세요.');
    if (data.length > Math.ceil(DOCUMENT_MAX_BYTES_ / 3) * 4) throw new Error('문서는 5MB 이내로 올려주세요.');
    if (!data || data.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(data)) throw new Error('문서 데이터가 손상되었습니다. 원본 파일을 다시 선택해주세요.');
    const padding = data.endsWith('==') ? 2 : data.endsWith('=') ? 1 : 0;
    const estimatedBytes = data.length / 4 * 3 - padding;
    if (estimatedBytes > DOCUMENT_MAX_BYTES_) throw new Error('문서는 5MB 이내로 올려주세요.');
    const bytes = Utilities.base64Decode(data);
    if (bytes.length !== estimatedBytes || bytes.length > DOCUMENT_MAX_BYTES_) throw new Error('문서 크기가 올바르지 않습니다.');
    let text;
    if (/\.txt$/i.test(name)) text = Utilities.newBlob(bytes).getDataAsString('UTF-8');
    else {
      const records = validateDocumentArchive_(bytes);
      const expectedNames = new Map(records.map(function (record) { return [record.name, record.expanded]; }));
      const bodyPattern = /\.hwpx$/i.test(name) ? /^Contents\/section\d+\.xml$/ : /^word\/document\.xml$/;
      if (!records.some(function (record) { return bodyPattern.test(record.name); })) throw new Error('문서 본문을 찾지 못했습니다. 텍스트를 직접 붙여넣어주세요.');
      // Utilities.unzip은 스트리밍 크기 제한을 제공하지 않습니다. 헤더 검증 뒤 실제 길이도 대조합니다.
      const files = Utilities.unzip(Utilities.newBlob(bytes, 'application/zip'));
      if (files.length > records.length) throw new Error('문서 압축 정보와 실제 파일이 일치하지 않습니다.');
      const seen = new Set(), bodies = [];
      let expandedTotal = 0;
      files.forEach(function (file) {
        const fileName = file.getName();
        if (!expectedNames.has(fileName) || seen.has(fileName)) throw new Error('문서 압축 정보와 실제 파일이 일치하지 않습니다.');
        seen.add(fileName);
        const size = file.getBytes().length;
        expandedTotal += size;
        if (size !== expectedNames.get(fileName) || size > 8 * 1024 * 1024 || expandedTotal > DOCUMENT_MAX_EXPANDED_BYTES_) throw new Error('문서의 실제 압축 해제 크기가 올바르지 않습니다.');
        if (bodyPattern.test(fileName)) bodies.push(file);
      });
      if (!bodies.length || records.some(function (record) { return bodyPattern.test(record.name) && !seen.has(record.name); })) throw new Error('문서 본문을 찾지 못했습니다.');
      bodies.sort(function (a, b) {
        return Number((a.getName().match(/section(\d+)/) || [])[1] || 0) - Number((b.getName().match(/section(\d+)/) || [])[1] || 0);
      });
      const textParts = [];
      let remaining = DOCUMENT_MAX_TEXT_;
      bodies.forEach(function (file) {
        const part = extractDocumentXmlText_(file.getDataAsString('UTF-8'), remaining);
        remaining -= part.length + (textParts.length ? 1 : 0);
        if (remaining < 0) throw new Error('본문이 너무 깁니다. 필요한 부분만 텍스트로 붙여넣어주세요.');
        textParts.push(part);
      });
      text = textParts.join('\n');
    }
    if (text.length > DOCUMENT_MAX_TEXT_) throw new Error('본문이 너무 깁니다. 필요한 부분만 텍스트로 붙여넣어주세요.');
    return { text: text, sourceName: sanitizeFileName_(name), message: '본문만 추출했습니다. 추출한 내용을 확인한 뒤 운영 정보에 반영해주세요.' };
  });
}

/** XML 엔티티 선언을 거절하고 깊이·노드 수·출력 길이를 제한합니다. */
function extractDocumentXmlText_(source, maxText) {
  if (source.length > 8000000 || /<!\s*(?:DOCTYPE|ENTITY)\b/i.test(source)) throw new Error('지원하지 않는 문서 구조입니다.');
  validateDocumentXmlStructure_(source);
  const root = XmlService.parse(source).getRootElement();
  const stack = [{ element: root, depth: 0, closing: false }], parts = [];
  let nodes = 0, length = 0;
  function append(value) {
    const text = String(value || '');
    length += text.length + (parts.length ? 1 : 0);
    if (length > maxText) throw new Error('본문이 너무 깁니다. 필요한 부분만 텍스트로 붙여넣어주세요.');
    parts.push(text);
  }
  while (stack.length) {
    const item = stack.pop(), element = item.element;
    if (item.closing) { append('\n'); continue; }
    if (++nodes > 50000 || item.depth > 64) throw new Error('문서 구조가 너무 복잡합니다. 필요한 부분만 텍스트로 붙여넣어주세요.');
    const name = element.getName();
    if (name === 'p' || name === 'tr') stack.push({ element: element, closing: true });
    if (name === 't') append(element.getText());
    else {
      const children = element.getChildren();
      if (children.length + nodes > 50000 || children.length + stack.length > 50000) throw new Error('문서에 요소가 너무 많습니다.');
      for (let i = children.length - 1; i >= 0; i--) stack.push({ element: children[i], depth: item.depth + 1, closing: false });
    }
  }
  return parts.join(' ');
}

/** 파서에 전달하기 전에도 깊이와 요소 수를 제한합니다. 인용부호·주석·CDATA 안의 태그는 세지 않습니다. */
function validateDocumentXmlStructure_(source) {
  const tokens = /<(?:!--[\s\S]*?-->|!\[CDATA\[[\s\S]*?\]\]>|\?[\s\S]*?\?>|(?:[^<>"']|"[^"]*"|'[^']*')*>)/g;
  let depth = 0, nodes = 0, match;
  while ((match = tokens.exec(source))) {
    const tag = match[0];
    if (/^<(?:!|\?)/.test(tag)) continue;
    if (/^<\//.test(tag)) {
      if (--depth < 0) throw new Error('문서 XML 구조가 올바르지 않습니다.');
    } else {
      if (++nodes > 50000 || depth > 64) throw new Error('문서 구조가 너무 복잡합니다. 필요한 부분만 텍스트로 붙여넣어주세요.');
      if (!/\/\s*>$/.test(tag)) depth++;
    }
  }
  if (depth !== 0) throw new Error('문서 XML 구조가 올바르지 않습니다.');
}
