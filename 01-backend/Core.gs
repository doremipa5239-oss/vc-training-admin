/****************************************************
 * [2_Core.gs]
 * 진입점·인증 가드 · 사이트 시트 저장소 · 감사 기록 · 캐시/드라이브 · 공통 유틸 · 주소/편집 세션 · 구버전 호환
 ****************************************************/

/****************************************************
 * 웹앱 진입점 · 공통 인증 가드 · 스프레드시트 열기
 ****************************************************/

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  if (!isValidAdminEntryKey_(params.admin)) return renderAdminEntryGuide_();
  const requestedCourseId = String(params.course || params.edit || '').trim();

  const template = HtmlService.createTemplateFromFile('admin');
  template.adminBaseUrl =
    getConfiguredWebAppBaseUrl_('ADMIN_WEB_APP_URL') ||
    String(ScriptApp.getService().getUrl() || '').replace(/[?#].*$/, '');
  template.requestedCourseId = requestedCourseId;
  return renderHtml_(template, 'V-Up 교육 운영 콘솔');
}

/**
 * HTML 파일 분리용 include 함수입니다.
 * admin.html 안에서 아래처럼 사용할 수 있습니다.
 *
 * <?!= include_('admin-style-1'); ?>
 * <?!= include_('admin-style-2'); ?>
 * <?!= include_('admin-style-3'); ?>
 * <?!= include('admin-main'); ?>
 * <?!= include('admin-script-1'); ?>
 * <?!= include('admin-script-2'); ?>
 * <?!= include('admin-script-3'); ?>
 *
 * 교육생용 index 파일은 별도의 공개 Apps Script 프로젝트에 둡니다.
 */
function include_(filename) {
  if (!filename) return '';
  return HtmlService.createHtmlOutputFromFile(String(filename)).getContent();
}

function requireAdmin_() {
  if (!ADMIN_RUNTIME_CONTEXT || !ADMIN_RUNTIME_CONTEXT.id) {
    throw new Error('관리자 인증이 만료되었거나 확인되지 않았습니다. 다시 로그인해주세요.');
  }
  return ADMIN_RUNTIME_CONTEXT;
}

/** 재시도 식별자를 계정·작업에 묶어 다른 관리자의 결과를 재사용할 수 없게 합니다. */
function scopeAdminRequestId_(requestId, operation) {
  const input = String(requestId || '').trim();
  if (!input) return '';
  if (input.length > 120) throw new Error('요청 식별자가 너무 깁니다. 화면을 새로고침해주세요.');
  return hashText_([requireAdmin_().id, String(operation || ''), input].join('|'));
}

function assertExpectedRevision_(expected, actual) {
  if (expected == null) return; // 기존 클라이언트 호환. 새 클라이언트는 항상 버전을 전달합니다.
  if (!Number.isSafeInteger(Number(expected)) || Number(expected) < 0 || Number(expected) !== Number(actual)) {
    throw new Error('다른 관리자가 먼저 수정했습니다. 최신 내용을 다시 불러온 뒤 변경해주세요.');
  }
}

/** 읽기-확인-저장을 같은 락 안에서 실행하여 동시 저장 시 덮어쓰기를 막습니다. */
function withCourseMutation_(currentAdmin, siteId, expectedRevision, callback) {
  return withScriptLock_(function () {
    const id = normalizePublicSiteId_(siteId);
    delete SITE_INFO_RUNTIME_CACHE[id];
    if (ADMIN_REQUEST_CACHE) {
      ADMIN_REQUEST_CACHE.ownerMap = null;
      ADMIN_REQUEST_CACHE.settingsMap = null;
    }
    const site = assertCourseAccess_(currentAdmin, id);
    if (isSiteTrashed_(site)) throw new Error('휴지통의 강의는 복구한 후 수정해주세요.');
    assertExpectedRevision_(expectedRevision, site.revision || 0);
    const result = callback(site);
    SpreadsheetApp.flush();
    return result;
  }, 15000);
}

function renderAdminEntryGuide_() {
  return HtmlService.createHtmlOutput(
    '<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>관리자 주소 확인</title><style>body{margin:0;background:#f7f5fb;color:#221b2f;font-family:Arial,sans-serif}' +
    'main{max-width:560px;margin:12vh auto;padding:32px;background:#fff;border:1px solid #ded7e8;border-radius:18px;box-shadow:0 12px 36px rgba(47,32,66,.1)}' +
    'h1{font-size:24px;margin:0 0 14px}p{font-size:16px;line-height:1.7;margin:0}</style></head><body>' +
    '<main><h1>관리자용 전체 주소를 확인해주세요</h1>' +
    '<p>이 페이지는 긴 보안키가 포함된 관리자용 고정 주소로 접속해야 합니다. 전달받은 관리자 링크를 다시 확인해주세요.</p></main></body></html>'
  ).setTitle('관리자 주소 확인').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function renderHtml_(template, title) {
  return template
    .evaluate()
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function getSpreadsheet_() {
  if (ADMIN_REQUEST_CACHE && ADMIN_REQUEST_CACHE.spreadsheet) return ADMIN_REQUEST_CACHE.spreadsheet;
  const configuredId = String(PropertiesService.getScriptProperties().getProperty('VUP_SPREADSHEET_ID') || SPREADSHEET_ID).trim();
  if (!/^[A-Za-z0-9_-]{20,150}$/.test(configuredId)) throw new Error('관리 데이터 시트 ID 설정을 확인해주세요.');
  const spreadsheet = SpreadsheetApp.openById(configuredId);
  if (ADMIN_REQUEST_CACHE) ADMIN_REQUEST_CACHE.spreadsheet = spreadsheet;
  return spreadsheet;
}

/****************************************************
 * 7. 내부 공통 함수 (2)
 * 교육사이트목록 시트 읽기/쓰기 · 행 변환 · 스크립트 락 · 체크포인트 · 공개 가드
 ****************************************************/

function getOrCreateSitesSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_SITES);

  if (!sheet) {
    sheet = withScriptLock_(function () { return ss.getSheetByName(SHEET_SITES) || ss.insertSheet(SHEET_SITES); });
  }

  ensureSitesHeader_(sheet);
  return sheet;
}

function getSitesSheetForRead_() {
  const sheet = getSpreadsheet_().getSheetByName(SHEET_SITES);
  if (!sheet) throw new Error('교육 사이트 목록 시트를 찾을 수 없습니다.');
  return sheet;
}

function ensureSitesHeader_(sheet) {
  if (sheet.getMaxColumns() < SITE_HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), SITE_HEADERS.length - sheet.getMaxColumns());
  }

  const current = sheet.getRange(1, 1, 1, SITE_HEADERS.length).getValues()[0];
  let shouldRewrite = false;

  for (let i = 0; i < SITE_HEADERS.length; i++) {
    if (String(current[i] || '').trim() !== SITE_HEADERS[i]) {
      shouldRewrite = true;
      break;
    }
  }

  if (shouldRewrite) {
    sheet.getRange(1, 1, 1, SITE_HEADERS.length).setValues([SITE_HEADERS]);
  }

  // 표준 컬럼(33개) 뒤에 남아 있는 옛 헤더 찌꺼기를 지워 컬럼 밀림을 방지합니다.
  const extraColumns = sheet.getMaxColumns() - SITE_HEADERS.length;
  if (extraColumns > 0) {
    const extraHeaders = sheet.getRange(1, SITE_HEADERS.length + 1, 1, extraColumns).getValues()[0];
    if (extraHeaders.some(function (value) { return String(value || '').trim() !== ''; })) {
      sheet.getRange(1, SITE_HEADERS.length + 1, 1, extraColumns).clearContent();
    }
  }

  if (sheet.getFrozenRows() !== 1) sheet.setFrozenRows(1);
}

function findSiteInfo_(sheet, siteId) {
  const rowIndex = findSiteRowIndex_(sheet, siteId);
  if (!rowIndex) return null;

  const row = sheet.getRange(rowIndex, 1, 1, SITE_HEADERS.length).getValues()[0];
  return rowToSiteInfo_(row);
}

function setSiteRowValues_(sheet, rowIndex, updater) {
  const range = sheet.getRange(rowIndex, 1, 1, SITE_HEADERS.length);
  const row = range.getValues()[0];
  const nextRow = typeof updater === 'function' ? (updater(row) || row) : row;
  nextRow[COL.revision - 1] = Math.max(0, Number(row[COL.revision - 1] || 0)) + 1;
  range.setValues([safeSheetRow_(nextRow)]);
  if (nextRow && nextRow[COL.siteId - 1]) {
    SITE_INFO_RUNTIME_CACHE[String(nextRow[COL.siteId - 1])] = rowToSiteInfo_(nextRow);
  }
}

function withScriptLock_(callback, timeoutMs) {
  const lock = LockService.getScriptLock();
  const alreadyOwned = lock.hasLock();
  if (!alreadyOwned) lock.waitLock(Number(timeoutMs || 15000));
  try {
    return callback();
  } finally {
    if (!alreadyOwned) {
      try { lock.releaseLock(); } catch (e) {}
    }
  }
}

function updateSiteRowById_(sheet, siteId, updater) {
  const normalizedSiteId = String(siteId || '').trim();
  if (!normalizedSiteId) throw new Error('교육 사이트 ID가 없습니다.');
  return withScriptLock_(function () {
    const currentRowIndex = findSiteRowIndex_(sheet, normalizedSiteId);
    if (!currentRowIndex) throw new Error('해당 교육 사이트를 찾을 수 없습니다.');
    setSiteRowValues_(sheet, currentRowIndex, updater);
    SpreadsheetApp.flush();
  });
}

function getSiteInfoById_(siteId) {
  const key = String(siteId || '');
  if (key && SITE_INFO_RUNTIME_CACHE[key]) return SITE_INFO_RUNTIME_CACHE[key];
  const sheet = getSitesSheetForRead_();
  return findSiteInfoFast_(sheet, siteId);
}

function findSiteInfoFast_(sheet, siteId) {
  const key = String(siteId || '');
  if (key && SITE_INFO_RUNTIME_CACHE[key]) return SITE_INFO_RUNTIME_CACHE[key];
  const site = findSiteInfo_(sheet, siteId);
  if (key && site) SITE_INFO_RUNTIME_CACHE[key] = site;
  return site;
}

function findSiteRowIndex_(sheet, siteId) {
  if (!siteId) return null;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const values = sheet.getRange(2, COL.siteId, lastRow - 1, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(siteId)) {
      return i + 2;
    }
  }

  return null;
}

function findSiteRowByRequestId_(sheet, requestId) {
  const normalized = String(requestId || '').trim();
  if (!normalized) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, COL.memo, lastRow - 1, 1).getValues();
  const marker = 'requestId:' + normalized;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === marker) return i + 2;
  }
  return null;
}

function rowToSiteInfo_(row) {
  const schedule = parseScheduleJson_(
    row[COL.scheduleJson - 1],
    formatCellValue_(row[COL.day1Date - 1]),
    formatCellValue_(row[COL.day2Date - 1])
  );
  const mealSlots = parseMealSlotsJson_(row[COL.mealSlotsJson - 1]);

  return {
    siteId: String(row[COL.siteId - 1] || ''),
    generation: String(row[COL.generation - 1] || ''),
    title: String(row[COL.title - 1] || ''),
    isNonVup: !/^\d+기$/.test(
      String(row[COL.generation - 1] || '').trim()
    ),
    managerName: String(row[COL.managerName - 1] || ''),
    managerPhone: String(row[COL.managerPhone - 1] || ''),
    courseDays: schedule.courseDays,
    day1Date: schedule.days[0] || '',
    day2Date: schedule.days[1] || '',
    day3Date: schedule.days[2] || '',
    day4Date: schedule.days[3] || '',
    day5Date: schedule.days[4] || '',
    mealSlots: mealSlots,
    linkedSheetId: String(row[COL.linkedSheetId - 1] || ''),
    linkedSheetUrl: String(row[COL.linkedSheetUrl - 1] || ''),
    userUrl: String(row[COL.userUrl - 1] || ''),
    createdAt: formatDate_(row[COL.createdAt - 1]),
    updatedAt: formatDate_(row[COL.updatedAt - 1]),
    status: String(row[COL.status - 1] || ''),
    memo: String(row[COL.memo - 1] || ''),
    materialFolderId: String(row[COL.materialFolderId - 1] || ''),
    materialFolderUrl: String(row[COL.materialFolderUrl - 1] || ''),
    photoFolderId: String(row[COL.photoFolderId - 1] || ''),
    photoFolderUrl: String(row[COL.photoFolderUrl - 1] || ''),
    lecturePdfId: String(row[COL.legacyLecturePdfId - 1] || ''),
    lecturePdfUrl: String(row[COL.legacyLecturePdfUrl - 1] || ''),
    finalCompleted: String(row[COL.finalCompleted - 1] || ''),
    finalCompletedAt: formatDate_(row[COL.finalCompletedAt - 1]),
    dataReviewConfirmedAt: formatDate_(row[COL.dataReviewConfirmedAt - 1]),
    deletedAt: formatDate_(row[COL.deletedAt - 1]),
    purgeAt: formatDate_(row[COL.purgeAt - 1]),
    adminLinkDocumentId: String(row[COL.adminLinkDocumentId - 1] || ''),
    folderNumber: Number(row[COL.folderNumber - 1] || 0),
    workflowCheckpoint: parseWorkflowCheckpoint_(row[COL.workflowCheckpointJson - 1]),
    preTrashStatus: String(row[COL.preTrashStatus - 1] || ''),
    revision: Math.max(0, Number(row[COL.revision - 1] || 0)),
    lastApprovedRevision: Math.max(0, Number(row[COL.lastApprovedRevision - 1] || 0))
  };
}

/** Spreadsheet에 쓰이는 사용자 문자열이 수식으로 실행되지 않도록 보호합니다. */
function safeSheetCell_(value) {
  if (typeof value !== 'string') return value;
  if (!value || value.charAt(0) === "'") return value;
  return /^[\s\u0000-\u001f\u007f\ufeff]*[=+\-@]/.test(value) ? "'" + value : value;
}

function safeSheetRow_(row) {
  return (Array.isArray(row) ? row : []).map(safeSheetCell_);
}

function safeSheetRows_(rows) {
  return (Array.isArray(rows) ? rows : []).map(safeSheetRow_);
}

function parseWorkflowCheckpoint_(rawValue) {
  if (!rawValue) return null;
  try {
    const parsed = JSON.parse(String(rawValue));
    return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : null;
  } catch (e) {
    return null;
  }
}

function saveAdminWorkflowCheckpoint(siteId, checkpoint) {
  requireAdmin_();
  try {
    const normalizedSiteId = normalizePublicSiteId_(siteId);
    const safeCheckpoint = sanitizeWorkflowCheckpoint_(checkpoint);
    const sheet = getOrCreateSitesSheet_(getSpreadsheet_());
    updateSiteRowById_(sheet, normalizedSiteId, function (row) {
      row[COL.workflowCheckpointJson - 1] = JSON.stringify(safeCheckpoint);
      row[COL.updatedAt - 1] = new Date();
      return row;
    });
    clearSiteListCache_();
    return {
      success: true,
      savedAtText: Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss')
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getAdminWorkflowCheckpoint(siteId) {
  requireAdmin_();
  try {
    const normalizedSiteId = normalizePublicSiteId_(siteId);
    const site = getSiteInfoById_(normalizedSiteId);
    if (!site || isSiteTrashed_(site)) throw new Error('운영 단계를 불러올 사이트를 찾을 수 없습니다.');
    return {
      success: true,
      checkpoint: site.workflowCheckpoint || null
    };
  } catch (error) {
    return { success: false, checkpoint: null, message: error.message };
  }
}

function sanitizeWorkflowCheckpoint_(checkpoint) {
  const input = checkpoint && typeof checkpoint === 'object' ? checkpoint : {};
  const allowedSteps = ['basic', 'link', 'meal', 'material', 'photo'];
  const allowedPriorities = ['differentPeople', 'company', 'adjacentSeat', 'roommate', 'sameName'];
  const priorityOrder = Array.isArray(input.priorityOrder)
    ? input.priorityOrder.map(String).filter(function (item, index, list) {
        return allowedPriorities.indexOf(item) !== -1 && list.indexOf(item) === index;
      })
    : [];
  return {
    version: 1,
    activeStep: allowedSteps.indexOf(String(input.activeStep || '')) !== -1 ? String(input.activeStep) : 'basic',
    mealStep: Math.max(1, Math.min(3, Number(input.mealStep || 1))),
    materialStep: Math.max(1, Math.min(2, Number(input.materialStep || 1))),
    photoStep: Math.max(1, Math.min(2, Number(input.photoStep || 1))),
    participantReviewConfirmed: Boolean(input.participantReviewConfirmed),
    priorityOrder: priorityOrder.length === allowedPriorities.length ? priorityOrder : allowedPriorities.slice(),
    priorityConfirmed: Boolean(input.priorityConfirmed),
    tableStrategy: String(input.tableStrategy || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 80),
    revalidation: {
      material: Boolean(input.revalidation && input.revalidation.material),
      photo: Boolean(input.revalidation && input.revalidation.photo)
    },
    updatedAt: Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss")
  };
}

function isSiteTrashed_(site) {
  return Boolean(site && site.deletedAt);
}

function ensurePublicSiteActive_(site) {
  if (!site) throw new Error('해당 교육 사이트 정보를 찾을 수 없습니다.');
  if (isSiteTrashed_(site)) throw new Error('현재 운영이 중단된 교육 사이트입니다. 교육 담당자에게 문의해주세요.');
  if (['비공개', '접근차단', '운영종료'].indexOf(String(site.status || '')) !== -1) {
    throw new Error('관리자가 현재 이 교육 사이트의 사용자 접근을 중단했습니다.');
  }
  return site;
}

function assertPublicModuleAccess_(site, moduleKey) {
  ensurePublicSiteActive_(site);
  const config = getPublicCourseAccessConfig_(site.siteId);
  if (!config.configured) return config;
  if (!config.siteAvailable) {
    throw new Error(config.phase === 'not_open'
      ? '아직 교육 안내 페이지 공개 전입니다.'
      : '교육 안내 페이지의 운영 기간이 종료되었습니다.');
  }
  if (moduleKey && !config.modules[moduleKey]) {
    throw new Error('이 기능은 아직 공개되지 않았습니다. 공개 시각 이후 다시 확인해주세요.');
  }
  return config;
}

function updateSiteLinkedSheet_(siteId, spreadsheetId, spreadsheetUrl) {
  const ss = getSpreadsheet_();
  const sheet = getOrCreateSitesSheet_(ss);
  const rowIndex = findSiteRowIndex_(sheet, siteId);

  if (!rowIndex) throw new Error('해당 교육 사이트를 찾을 수 없습니다.');

  updateSiteRowById_(sheet, siteId, function (row) {
    row[COL.linkedSheetId - 1] = spreadsheetId;
    row[COL.linkedSheetUrl - 1] = spreadsheetUrl;
    row[COL.updatedAt - 1] = new Date();
    return row;
  });
}

/****************************************************
 * 관리자 감사 기록 · 운영 현황 조회
 ****************************************************/

function getOrCreateAuditSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_AUDIT);
  if (!sheet) sheet = ss.insertSheet(SHEET_AUDIT);
  const headers = ['일시', '작업', 'siteId', '관리자계정', '상세JSON'];
  if (sheet.getMaxColumns() < headers.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (headers.some(function (header, index) { return String(current[index] || '') !== header; })) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function appendAuditLog_(action, siteId, detail) {
  try {
    const actor = ADMIN_RUNTIME_CONTEXT && ADMIN_RUNTIME_CONTEXT.id
      ? ADMIN_RUNTIME_CONTEXT.id
      : '공개 API 또는 시스템';
    getOrCreateAuditSheet_().appendRow(safeSheetRow_([
      new Date(),
      String(action || ''),
      String(siteId || ''),
      actor,
      JSON.stringify(detail || {}).slice(0, 5000)
    ]));
  } catch (e) {
    // 감사 기록 실패가 본 작업을 망가뜨리지는 않되, 다음 작업에서 다시 기록됩니다.
  }
}

function getAdminOperationsOverview(siteId, limit) {
  requireAdmin_();
  try {
    const normalizedSiteId = String(siteId || '').trim();
    const safeLimit = Math.max(10, Math.min(100, Number(limit || 40)));
    const sheet = getOrCreateAuditSheet_();
    const lastRow = sheet.getLastRow();
    const startRow = Math.max(2, lastRow - Math.max(safeLimit * 4, safeLimit) + 1);
    const rowCount = lastRow >= startRow ? lastRow - startRow + 1 : 0;
    const values = rowCount ? sheet.getRange(startRow, 1, rowCount, 5).getValues() : [];
    const logs = [];

    for (let index = values.length - 1; index >= 0 && logs.length < safeLimit; index--) {
      const row = values[index];
      const rowSiteId = String(row[2] || '');
      if (normalizedSiteId && rowSiteId !== normalizedSiteId) continue;
      const action = String(row[1] || '');
      const failure = /FAIL|ERROR|BLOCK|DENIED/i.test(action);
      logs.push({
        timestamp: formatDate_(row[0]),
        action: action,
        siteId: rowSiteId,
        operationId: 'AUD-' + String(startRow + index) + '-' + Utilities.formatDate(new Date(row[0] || new Date()), TIMEZONE, 'yyyyMMddHHmmss'),
        failure: failure,
        summary: summarizeAuditDetail_(row[4])
      });
    }

    let editSession = null;
    if (normalizedSiteId) {
      try {
        const rawSession = CacheService.getScriptCache().get(getSiteEditSessionKey_(normalizedSiteId));
        const parsedSession = rawSession ? JSON.parse(rawSession) : null;
        if (parsedSession) {
          editSession = {
            active: true,
            lastSeenAt: parsedSession.lastSeenAt || parsedSession.updatedAt || '',
            concurrent: Boolean(parsedSession.concurrent)
          };
        }
      } catch (e) {}
    }

    const failureCount = logs.filter(function (item) { return item.failure; }).length;
    return {
      success: true,
      siteId: normalizedSiteId,
      diagnosticId: 'VUP-' + Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0, 8),
      checkedAtText: Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
      total: logs.length,
      failureCount: failureCount,
      editSession: editSession,
      logs: logs
    };
  } catch (error) {
    return { success: false, message: error.message, logs: [] };
  }
}

function summarizeAuditDetail_(rawDetail) {
  if (!rawDetail) return '상세 기록 없음';
  try {
    const parsed = typeof rawDetail === 'string' ? JSON.parse(rawDetail) : rawDetail;
    if (!parsed || typeof parsed !== 'object') return String(rawDetail).slice(0, 180);
    const safeParts = [];
    Object.keys(parsed).slice(0, 8).forEach(function (key) {
      if (/phone|email|base64|token|key|url|name/i.test(key)) return;
      const value = parsed[key];
      if (value === null || typeof value === 'object') return;
      safeParts.push(key + ': ' + String(value).slice(0, 90));
    });
    return safeParts.join(' · ') || '개인정보를 제외한 작업 기록';
  } catch (e) {
    return String(rawDetail).replace(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g, '[이메일 숨김]').slice(0, 180);
  }
}

/****************************************************
 * 캐시 헬퍼 · 6. 구글 드라이브 폴더 탐색
 ****************************************************/

function getUserDataCache_(key, forceRefresh) {
  if (forceRefresh) return null;
  try {
    const raw = CacheService.getScriptCache().get(key);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

function putUserDataCache_(key, value) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(value), USER_DATA_CACHE_SECONDS);
  } catch (e) {
    // 응답이 큰 경우 CacheService 제한에 걸릴 수 있으므로 캐시 실패는 무시합니다.
  }
}

function putSiteListCache_(key, value) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(value), SITE_LIST_CACHE_SECONDS);
  } catch (e) {
    // 목록 캐시가 실패해도 다음 조회에서 원본 시트를 읽으면 됩니다.
  }
}

function putDriveBrowserCache_(key, value) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(value), DRIVE_BROWSER_CACHE_SECONDS);
  } catch (e) {
    // 폴더 탐색 캐시는 보조 수단입니다. 실패해도 원본 Drive를 읽으면 됩니다.
  }
}

function clearSiteListCache_() {
  try {
    CacheService.getScriptCache().remove('educationSites:list');
  } catch (e) {
    // 캐시 삭제 실패는 실제 저장 결과에 영향을 주지 않습니다.
  }
}

function clearUserDataCache_(siteId) {
  if (!siteId) return;
  delete SITE_INFO_RUNTIME_CACHE[String(siteId)];
  try {
    CacheService.getScriptCache().removeAll([
      'mealSeatData:' + String(siteId),
      'mealSeatData:v2:' + String(siteId),
      'siteData:' + String(siteId),
      'siteData:v2:' + String(siteId),
      'materialFiles:' + String(siteId),
      'materialFiles:v2:' + String(siteId),
      'photoFiles:' + String(siteId),
      'photoFiles:v2:' + String(siteId),
      'participantLookup:v3:' + String(siteId)
    ]);
  } catch (e) {
    // 캐시 삭제 실패는 실제 저장 결과에 영향을 주지 않습니다.
  }
}

/****************************************************
 * 6. 구글 드라이브 폴더 탐색
 ****************************************************/

function getFolderContents(folderId) {
  requireAdmin_();
  const normalizedFolderId = folderId || 'root';
  const cacheKey = 'driveFolderContents:' + String(normalizedFolderId);
  const cached = getUserDataCache_(cacheKey, false);
  if (cached) return cached;

  const folder = !folderId || folderId === 'root'
    ? DriveApp.getRootFolder()
    : DriveApp.getFolderById(folderId);
  const currentId = folder.getId();
  const rootId = getDriveRootId_();

  const children = [];
  const folders = folder.getFolders();

  while (folders.hasNext()) {
    const f = folders.next();
    children.push({ name: f.getName(), id: f.getId() });
  }

  children.sort((a, b) => a.name.localeCompare(b.name, 'ko'));

  let parentId = null;
  if (currentId !== rootId) {
    const parents = folder.getParents();
    if (parents.hasNext()) parentId = parents.next().getId();
  }

  const result = {
    currentId: currentId,
    currentPath: getFullPath_(folder),
    parentId,
    children
  };
  putDriveBrowserCache_(cacheKey, result);
  return result;
}

/**
 * [신규] 시트 연동 picker용: 폴더 + 구글 시트 파일을 함께 반환합니다.
 * 관리자가 "폴더 → 폴더 → 파일" 경로로 직접 시트를 고를 수 있게 합니다.
 * 엑셀(.xlsx) 원본 파일은 SpreadsheetApp으로 열 수 없으므로,
 * 구글 시트(GOOGLE_SHEETS)로 변환된 파일만 목록에 표시합니다.
 */
function getSheetPickerContents(folderId) {
  requireAdmin_();
  const normalizedFolderId = folderId || 'root';
  const cacheKey = 'sheetPickerContents:' + String(normalizedFolderId);
  const cached = getUserDataCache_(cacheKey, false);
  if (cached) return cached;

  const folder = !folderId || folderId === 'root'
    ? DriveApp.getRootFolder()
    : DriveApp.getFolderById(folderId);
  const currentId = folder.getId();
  const rootId = getDriveRootId_();

  const folders = [];
  const folderIterator = folder.getFolders();
  while (folderIterator.hasNext()) {
    const f = folderIterator.next();
    folders.push({ name: f.getName(), id: f.getId() });
  }
  folders.sort(function (a, b) { return a.name.localeCompare(b.name, 'ko'); });

  const sheets = [];
  const fileIterator = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (fileIterator.hasNext()) {
    const file = fileIterator.next();
    sheets.push({ name: file.getName(), id: file.getId() });
  }
  sheets.sort(function (a, b) { return a.name.localeCompare(b.name, 'ko'); });

  let parentId = null;
  if (currentId !== rootId) {
    const parents = folder.getParents();
    if (parents.hasNext()) parentId = parents.next().getId();
  }

  const result = {
    currentId: currentId,
    currentPath: getFullPath_(folder),
    parentId: parentId,
    folders: folders,
    sheets: sheets
  };
  putDriveBrowserCache_(cacheKey, result);
  return result;
}

function getFullPath_(folder) {
  const parts = [];
  let current = folder;
  const rootId = getDriveRootId_();
  let depth = 0;

  while (current && current.getId() !== rootId && depth < 20) {
    parts.unshift(current.getName());
    const parents = current.getParents();
    if (!parents.hasNext()) break;
    current = parents.next();
    depth++;
  }

  return parts.length ? '내 드라이브 > ' + parts.join(' > ') : '내 드라이브';
}

function getDriveRootId_() {
  if (!DRIVE_ROOT_ID_RUNTIME_CACHE) {
    DRIVE_ROOT_ID_RUNTIME_CACHE = DriveApp.getRootFolder().getId();
  }
  return DRIVE_ROOT_ID_RUNTIME_CACHE;
}

/****************************************************
 * 7. 내부 공통 함수 (3)
 * 일정·날짜·식사 슬롯 정규화, 사이트 ID, 포맷, 파일 형식 검증
 ****************************************************/

/**
 * 관리자 화면에서 선택한 합숙 형태를 실제 운영 일수(1~5일)로 정규화합니다.
 * 저장 형식은 기존 courseDays를 유지하므로 과거 강의 데이터와 호환됩니다.
 */
function normalizeCourseDays_(courseDays) {
  const n = Number(courseDays || 0);
  if (!n || n < 1 || n > 5) {
    throw new Error('합숙 형태를 선택해주세요. (당일치기부터 4박 5일까지)');
  }
  return n;
}

/**
 * data.day1Date ~ data.day5Date 중 courseDays 만큼의 날짜를 추출하고,
 * 각 날짜가 비어있지 않은지 검증합니다. 나머지(course 범위를 넘는 날짜)는 빈 문자열로 채웁니다.
 * 반환값은 항상 길이 5의 배열입니다 (day1~day5 순서).
 */
function extractDayDatesFromData_(data, courseDays) {
  const dates = [];
  const dateKeys = {};
  for (let day = 1; day <= 5; day++) {
    const raw = String((data && data['day' + day + 'Date']) || '').trim();
    if (day <= courseDays) {
      if (!raw) throw new Error(day + '일차 날짜를 입력해주세요.');
      const parsed = parseStrictEducationDate_(raw, day);
      if (dateKeys[parsed.dateKey]) throw new Error(day + '일차 날짜가 앞선 교육 날짜와 겹칩니다. 각 일차의 날짜를 다르게 입력해주세요.');
      dateKeys[parsed.dateKey] = true;
      dates.push(parsed.displayText);
    } else {
      dates.push('');
    }
  }
  return dates;
}

function parseStrictEducationDate_(value, dayNumber) {
  const text = String(value || '').trim();
  const match = text.match(/^(\d{4})\D+(\d{1,2})\D+(\d{1,2})\D*$/);
  if (!match) throw new Error((dayNumber || '') + '일차 날짜를 "2026년 7월 1일" 형식으로 입력해주세요.');
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (year < 2000 || year > 2100 || month < 1 || month > 12 || day < 1 || day > 31 ||
      date.getUTCFullYear() !== year || date.getUTCMonth() + 1 !== month || date.getUTCDate() !== day) {
    throw new Error((dayNumber || '') + '일차 날짜가 실제 달력에 존재하지 않습니다.');
  }
  return {
    dateKey: year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0'),
    displayText: year + '년 ' + month + '월 ' + day + '일'
  };
}

/**
 * 프런트엔드(admin-script.html)가 보내는 mealSlots(식사 슬롯 배열: id/day/type/label/startHour/endHour)를
 * 그대로 신뢰하지 않고, 저장 전에 형태를 검증/정리합니다.
 */
function normalizeMealSlots_(mealSlots, courseDays) {
  if (!Array.isArray(mealSlots)) return [];
  const maximumDay = Number(courseDays || 5);
  const seen = {};

  return mealSlots.map(function (slot, index) {
    if (!slot) throw new Error((index + 1) + '번째 식사 일정이 비어 있습니다.');
    const day = Number(slot.day);
    if (!Number.isInteger(day) || day < 1 || day > maximumDay) {
      throw new Error((index + 1) + '번째 식사 일정의 일차가 교육 기간을 벗어났습니다.');
    }
    if (slot.type !== 'lunch' && slot.type !== 'dinner') {
      throw new Error((index + 1) + '번째 식사 종류가 올바르지 않습니다.');
    }
    const type = slot.type;
    const defaultStart = type === 'dinner' ? 15 : 10;
    const defaultEnd = type === 'dinner' ? 20 : 15;
    const startHour = slot.startHour === '' || slot.startHour == null ? defaultStart : Number(slot.startHour);
    const endHour = slot.endHour === '' || slot.endHour == null ? defaultEnd : Number(slot.endHour);
    if (!Number.isInteger(startHour) || !Number.isInteger(endHour) || startHour < 0 || startHour > 23 || endHour < 1 || endHour > 24 || startHour >= endHour) {
      throw new Error(day + '일차 ' + (type === 'dinner' ? '저녁' : '점심') + ' 공개 시간은 0~24시 범위에서 시작보다 종료가 늦어야 합니다.');
    }
    const duplicateKey = day + '|' + type;
    if (seen[duplicateKey]) {
      throw new Error(day + '일차 ' + (type === 'dinner' ? '저녁' : '점심') + ' 일정이 중복되었습니다.');
    }
    seen[duplicateKey] = true;

    return {
      id: 'day' + day + '_' + type,
      day: day,
      type: type,
      label: String(slot.label || '').trim().slice(0, 80),
      startHour: startHour,
      endHour: endHour
    };
  });
}

function getMealSlotShapeKey_(mealSlots) {
  return (Array.isArray(mealSlots) ? mealSlots : [])
    .map(function (slot) {
      return Number(slot && slot.day || 0) + '|' + String(slot && slot.type || '');
    })
    .filter(function (key) { return key !== '0|'; })
    .sort()
    .join(',');
}

/**
 * '오프라인일정JSON' 컬럼 값을 안전하게 파싱합니다.
 * 값이 없거나(과거 데이터) 파싱에 실패하면, 1~2일차 날짜 컬럼만 있던 구버전 데이터로 간주해
 * courseDays=2 로 폴백합니다.
 */
function parseScheduleJson_(rawJson, legacyDay1Date, legacyDay2Date) {
  if (rawJson) {
    try {
      const parsed = JSON.parse(rawJson);
      if (parsed && Array.isArray(parsed.days)) {
        const days = parsed.days.slice(0, 5);
        while (days.length < 5) days.push('');
        return { courseDays: Number(parsed.courseDays) || days.filter(Boolean).length || 2, days: days };
      }
    } catch (e) {
      // fallthrough to legacy
    }
  }

  const days = [String(legacyDay1Date || ''), String(legacyDay2Date || ''), '', '', ''];
  return { courseDays: 2, days: days };
}

/**
 * '식사슬롯JSON' 컬럼 값을 안전하게 파싱합니다. 과거 데이터는 항상 빈 배열입니다.
 */
function parseMealSlotsJson_(rawJson) {
  if (!rawJson) return [];
  try {
    const parsed = JSON.parse(rawJson);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function normalizeGeneration_(generation) {
  const match = String(generation || '').match(/\d+/);
  if (!match) throw new Error('기수는 숫자를 포함해야 합니다. 예: 62기');
  return match[0];
}

function createSiteId_() {
  // 기존 12자리(48비트)보다 충분히 긴 24자리(96비트)를 사용합니다.
  // 과거 siteId는 읽기 로직을 바꾸지 않으므로 기존 사용자 URL도 그대로 동작합니다.
  return Utilities.getUuid().replace(/-/g, '').slice(0, 24);
}

function extractPhoneLast4_(phone) {
  const digits = String(phone || '').replace(/\D/g, '');
  return digits.length >= 4 ? digits.slice(-4) : '';
}

function formatDate_(date) {
  if (!date) return '';
  if (Object.prototype.toString.call(date) !== '[object Date]') return String(date);
  return Utilities.formatDate(date, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}

function formatCellValue_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, TIMEZONE, 'yyyy년 M월 d일');
  }
  return String(value);
}

function parseKoreanDateParts_(value) {
  const text = String(value || '');
  let match = text.match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  let year;
  let month;
  let day;

  if (match) {
    year = Number(match[1]);
    month = Number(match[2]);
    day = Number(match[3]);
  } else {
    match = text.match(/(\d{1,2})\D+(\d{1,2})/);
    if (!match) return null;
    year = Number(Utilities.formatDate(new Date(), TIMEZONE, 'yyyy'));
    month = Number(match[1]);
    day = Number(match[2]);
  }

  if (!year || !month || !day) return null;

  return {
    dateKey: `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`,
    shortText: `${month}월 ${day}일`
  };
}

/**
 * 지금 이 순간이 어떤 식사 시간대인지 판별합니다.
 * 과거에는 1~2일차, 점심/저녁 4개 슬롯으로 하드코딩되어 있었지만,
 * 이제는 사이트가 실제로 저장한 courseDays(최대 5일)와 mealSlots(관리자가 직접 고른 식사 슬롯,
 * 시작/종료 시각 포함)를 그대로 사용하므로 3일/5일 과정 모두 정확히 동작합니다.
 * mealSlots가 비어있는(과거) 사이트는 항상 null을 반환합니다 - 실시간 표시 기능 없이도
 * 나머지 기능(교재/사진)은 정상 동작해야 하므로 에러를 던지지 않습니다.
 */
function getCurrentMealType_(siteInfo) {
  const runtimeSlots = getMealRuntimeSlots_(siteInfo);
  return getCurrentMealTypeFromRuntimeSlots_(runtimeSlots);
}

function getCurrentMealTypeFromRuntimeSlots_(runtimeSlots) {
  const slots = Array.isArray(runtimeSlots) ? runtimeSlots : [];
  for (let i = 0; i < slots.length; i++) {
    if (slots[i].status === 'current') {
      return {
        dateKey: slots[i].dateKey,
        dateText: slots[i].dateText,
        mealText: slots[i].mealText,
        seatKey: slots[i].seatKey
      };
    }
  }

  return null;
}

function getMealRuntimeSlots_(siteInfo) {
  const mealSlots = Array.isArray(siteInfo.mealSlots) ? siteInfo.mealSlots : [];
  if (!mealSlots.length) return [];

  const now = new Date();
  const todayKey = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd');
  const hour = Number(Utilities.formatDate(now, TIMEZONE, 'H'));
  const minute = Number(Utilities.formatDate(now, TIMEZONE, 'm'));
  const nowMinutes = hour * 60 + minute;

  return mealSlots.map(function (slot, i) {
    const dayNumber = Number(slot.day || 0);
    if (!dayNumber || dayNumber < 1 || dayNumber > 5) return null;

    const dateRaw = siteInfo['day' + dayNumber + 'Date'];
    const parts = parseKoreanDateParts_(dateRaw);
    if (!parts) return null;

    const startHour = Number.isFinite(Number(slot.startHour)) ? Number(slot.startHour) : (slot.type === 'dinner' ? 15 : 10);
    const endHour = Number.isFinite(Number(slot.endHour)) ? Number(slot.endHour) : (slot.type === 'dinner' ? 20 : 15);
    const startMinutes = startHour * 60;
    const endMinutes = endHour * 60;
    const mealText = slot.type === 'dinner' ? '저녁' : '점심';
    let status = 'future';

    if (parts.dateKey < todayKey || (parts.dateKey === todayKey && nowMinutes >= endMinutes)) {
      status = 'past';
    } else if (parts.dateKey === todayKey && nowMinutes >= startMinutes && nowMinutes < endMinutes) {
      status = 'current';
    }

    return {
      dateKey: parts.dateKey,
      dateText: parts.shortText,
      mealText: mealText,
      seatKey: makeMealSeatKey_(dayNumber + '일차 ' + mealText, i),
      startHour: startHour,
      endHour: endHour,
      status: status
    };
  }).filter(Boolean);
}

function sanitizeFileName_(fileName) {
  return String(fileName || '')
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/[\u0000-\u001f\u007f]/g, '')
    .replace(/\s+/g, ' ')
    .trim().slice(0, 180) || '파일';
}

function guessMimeTypeFromName_(fileName, serviceType) {
  const ext = getExtension_(fileName);

  if (ext === 'pdf') return 'application/pdf';
  if (ext === 'ppt') return 'application/vnd.ms-powerpoint';
  if (ext === 'pptx') return 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
  if (ext === 'jpg' || ext === 'jpeg') return 'image/jpeg';
  if (ext === 'png') return 'image/png';
  if (ext === 'gif') return 'image/gif';
  if (ext === 'webp') return 'image/webp';
  if (ext === 'heic') return 'image/heic';
  if (ext === 'heif') return 'image/heif';

  return serviceType === 'material' ? 'application/pdf' : 'application/octet-stream';
}

function validateFileType_(fileName, mimeType, serviceType) {
  const ext = getExtension_(fileName);

  if (serviceType === 'material') {
    if (getAllowedMaterialMimeTypes_().indexOf(mimeType) === -1 || getAllowedMaterialExtensions_().indexOf(ext) === -1) {
      throw new Error('교재 파일은 PDF, PPT 또는 PPTX 형식으로 올려주세요.');
    }
    return;
  }

  const okImageMime = ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'image/heic', 'image/heif'].indexOf(mimeType) !== -1;
  const okImageExt = ['jpg', 'jpeg', 'png', 'gif', 'webp', 'heic', 'heif'].indexOf(ext) !== -1;

  if (!okImageMime || !okImageExt) {
    throw new Error('현장 사진은 jpg, jpeg, png, gif, webp, heic, heif 형식만 올릴 수 있습니다.');
  }
}

function normalizeManagerPhone_(value) {
  const contact = String(value || '').trim();
  if (!contact) throw new Error('담당자 연락처를 입력해주세요.');

  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (emailPattern.test(contact)) return contact;

  // 자동으로 하이픈을 넣거나 숫자 길이를 임의로 바꾸지 않고, 관리자가 쓴 값을 보존합니다.
  const phonePattern = /^\d{2,4}(?:-\d{3,4}){1,2}$/;
  if (phonePattern.test(contact)) return contact;

  throw new Error('전화번호는 반드시 - 을 포함하여 작성하거나 이메일 주소를 입력해주세요.');
}

function validateFileSignature_(fileName, bytes, serviceType) {
  const data = Array.isArray(bytes) ? bytes.slice(0, 16).map(function (value) { return (Number(value) + 256) % 256; }) : [];
  const ext = getExtension_(fileName);
  const startsWith = function (signature) {
    if (data.length < signature.length) return false;
    for (let i = 0; i < signature.length; i++) if (data[i] !== signature[i]) return false;
    return true;
  };
  const asciiAt = function (offset, text) {
    if (data.length < offset + text.length) return false;
    for (let i = 0; i < text.length; i++) if (data[offset + i] !== text.charCodeAt(i)) return false;
    return true;
  };

  if (serviceType === 'material') {
    const validMaterial = ext === 'pdf'
      ? startsWith([0x25, 0x50, 0x44, 0x46, 0x2D])
      : ext === 'ppt'
        ? startsWith([0xD0, 0xCF, 0x11, 0xE0])
        : ext === 'pptx'
          ? startsWith([0x50, 0x4B, 0x03, 0x04])
          : false;
    if (!validMaterial) throw new Error('교재 파일의 확장자와 실제 내용이 일치하지 않습니다. 원본 PDF/PPT/PPTX를 다시 선택해주세요.');
    return;
  }

  let valid = false;
  if (ext === 'jpg' || ext === 'jpeg') valid = startsWith([0xFF, 0xD8, 0xFF]);
  else if (ext === 'png') valid = startsWith([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
  else if (ext === 'gif') valid = asciiAt(0, 'GIF87a') || asciiAt(0, 'GIF89a');
  else if (ext === 'webp') valid = asciiAt(0, 'RIFF') && asciiAt(8, 'WEBP');
  else if (ext === 'heic' || ext === 'heif') {
    const brand = data.length >= 12 ? String.fromCharCode(data[8], data[9], data[10], data[11]).toLowerCase() : '';
    valid = asciiAt(4, 'ftyp') && ['heic', 'heix', 'hevc', 'hevx', 'mif1', 'msf1'].indexOf(brand) !== -1;
  }
  if (!valid) {
    throw new Error('파일 확장자와 실제 이미지 내용이 일치하지 않습니다. 원본 사진을 다시 선택해주세요.');
  }
}

function getAllowedMaterialMimeTypes_() {
  return [
    'application/pdf',
    'application/vnd.ms-powerpoint',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'application/vnd.google-apps.presentation'
  ];
}

function getAllowedMaterialExtensions_() {
  return ['pdf', 'ppt', 'pptx'];
}

function getExtension_(fileName) {
  const parts = String(fileName || '').toLowerCase().split('.');
  return parts.length > 1 ? parts.pop() : '';
}

function formatFileSize_(bytes) {
  const size = Number(bytes || 0);
  if (size < 1024) return size + ' B';
  if (size < 1024 * 1024) return (size / 1024).toFixed(1) + ' KB';
  if (size < 1024 * 1024 * 1024) return (size / 1024 / 1024).toFixed(1) + ' MB';
  return (size / 1024 / 1024 / 1024).toFixed(1) + ' GB';
}

/****************************************************
 * 웹앱 주소·관리자 보안키, 동시 편집 세션,
 *                         관리자 링크 메일/드라이브 문서
 ****************************************************/

function createUserUrl_(siteId) {
  const baseUrl = getConfiguredWebAppBaseUrl_('PUBLIC_WEB_APP_URL');

  if (!baseUrl) {
    throw new Error('스크립트 속성 PUBLIC_WEB_APP_URL에 교육생용 웹앱 배포 주소를 먼저 등록해주세요.');
  }

  return `${baseUrl}?site=${encodeURIComponent(siteId)}`;
}

function createAdminUrl_(siteId) {
  return appendQueryParam_(getAdminEntryUrl_(), 'edit', String(siteId || ''));
}

function getConfiguredWebAppBaseUrl_(propertyName) {
  return String(PropertiesService.getScriptProperties().getProperty(propertyName) || '')
    .trim()
    .replace(/[?#].*$/, '');
}

function getOrCreateAdminEntryKey_() {
  const properties = PropertiesService.getScriptProperties();
  let key = String(properties.getProperty('ADMIN_ENTRY_KEY') || '').trim();
  if (key) return key;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    key = String(properties.getProperty('ADMIN_ENTRY_KEY') || '').trim();
    if (!key) {
      key = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
      properties.setProperty('ADMIN_ENTRY_KEY', key);
    }
    return key;
  } finally {
    lock.releaseLock();
  }
}

function isValidAdminEntryKey_(providedKey) {
  const provided = String(providedKey || '').trim();
  return /^[A-Za-z0-9_-]{32,160}$/.test(provided) && constantTimeEquals_(provided, getOrCreateAdminEntryKey_());
}

function getAdminEntryUrl_() {
  const baseUrl = getConfiguredWebAppBaseUrl_('ADMIN_WEB_APP_URL');
  if (!baseUrl) {
    throw new Error('스크립트 속성 ADMIN_WEB_APP_URL에 관리자용 웹앱 배포 주소를 등록해주세요.');
  }
  const publicBaseUrl = getConfiguredWebAppBaseUrl_('PUBLIC_WEB_APP_URL');
  if (publicBaseUrl && publicBaseUrl === baseUrl) {
    throw new Error('관리자용 배포 주소와 교육생용 배포 주소는 서로 달라야 합니다.');
  }
  return baseUrl + '?admin=' + encodeURIComponent(getOrCreateAdminEntryKey_());
}

function appendQueryParam_(url, name, value) {
  const text = String(url || '');
  const paramValue = String(value == null ? '' : value).trim();
  if (!paramValue) return text;
  return text + (text.indexOf('?') === -1 ? '?' : '&') +
    encodeURIComponent(String(name || '')) + '=' + encodeURIComponent(paramValue);
}

/**
 * 최초 1회: 프로젝트 설정의 스크립트 속성에 ADMIN_WEB_APP_URL과
 * PUBLIC_WEB_APP_URL을 등록한 다음 Apps Script 편집기에서 이 함수를 실행하세요.
 * 실행 로그에 앞으로 사용할 고정 관리자 주소가 표시됩니다.
 */
function initializeWebAppRouting_() {
  // 이 함수는 Apps Script 편집기에서 최초 1회 실행하는 주소 생성 도우미입니다.
  // 관리자 보안 주소는 실행 로그에만 남고 함수 응답에는 포함되지 않습니다.
  // 실제 접근 권한은 관리자 프로젝트 웹 앱 배포의 "나만" 설정이 담당합니다.
  const adminBaseUrl = getConfiguredWebAppBaseUrl_('ADMIN_WEB_APP_URL');
  const publicBaseUrl = getConfiguredWebAppBaseUrl_('PUBLIC_WEB_APP_URL');
  if (!adminBaseUrl || !publicBaseUrl) {
    throw new Error('스크립트 속성 ADMIN_WEB_APP_URL과 PUBLIC_WEB_APP_URL을 모두 등록해주세요.');
  }
  if (adminBaseUrl === publicBaseUrl) {
    throw new Error('관리자용 배포 주소와 교육생용 배포 주소는 서로 달라야 합니다.');
  }
  const result = {
    adminUrl: getAdminEntryUrl_(),
    publicBaseUrl: publicBaseUrl
  };
  console.log('관리자용 고정 주소: ' + result.adminUrl);
  console.log('교육생용 공개 주소: ' + result.publicBaseUrl);
  // 브라우저에서 이 함수를 임의로 호출하더라도 관리자 보안 주소는 응답으로 반환하지 않습니다.
  return { success: true, message: '실행 로그에서 관리자용 고정 주소를 확인해주세요.' };
}

function initializeWebAppRouting() {
  return initializeWebAppRouting_();
}

function getSiteEditSessionKey_(siteId) {
  return 'siteEditSession:' + String(siteId || '');
}

function normalizeEditDeviceToken_(deviceToken) {
  const token = String(deviceToken || '').trim();
  if (!/^[A-Za-z0-9_-]{12,160}$/.test(token)) throw new Error('편집 기기를 확인하지 못했습니다. 페이지를 새로고침해주세요.');
  return token;
}

function claimSiteEditSession(siteId, deviceToken, ignoreAndEnter) {
  requireAdmin_();
  const normalizedSiteId = normalizePublicSiteId_(siteId);
  const token = normalizeEditDeviceToken_(deviceToken);
  return withScriptLock_(function () {
    const site = getSiteInfoById_(normalizedSiteId);
    if (!site || isSiteTrashed_(site)) throw new Error('편집할 교육 사이트를 찾을 수 없습니다.');
    const cache = CacheService.getScriptCache();
    const key = getSiteEditSessionKey_(normalizedSiteId);
    const now = new Date().getTime();
    let state = { editors: [] };
    try { state = JSON.parse(cache.get(key) || '{"editors":[]}'); } catch (e) {}
    state.editors = (state.editors || []).filter(function (editor) {
      return now - Number(editor.heartbeatAt || 0) < 90000;
    });
    const mine = state.editors.filter(function (editor) { return editor.token === token; })[0];
    const others = state.editors.filter(function (editor) { return editor.token !== token; });
    if (mine) {
      mine.heartbeatAt = now;
      if (now - Number(mine.startedAt || now) >= SITE_EDIT_SESSION_MINUTES * 60000) mine.startedAt = now;
    } else if (!others.length || ignoreAndEnter) {
      state.editors.push({ token: token, startedAt: now, heartbeatAt: now, override: Boolean(ignoreAndEnter) });
      if (ignoreAndEnter && others.length) appendAuditLog_('CONCURRENT_EDIT_OVERRIDE', normalizedSiteId, { otherEditorCount: others.length });
    }
    const activeOthers = state.editors.filter(function (editor) { return editor.token !== token; });
    cache.put(key, JSON.stringify(state), 900);
    if (!mine && activeOthers.length && !ignoreAndEnter) {
      const firstStartedAt = Math.min.apply(null, activeOthers.map(function (editor) { return Number(editor.startedAt || now); }));
      const elapsedMinutes = Math.max(0, Math.floor((now - firstStartedAt) / 60000));
      return { success: true, canEnter: false, concurrent: true, waitMinutes: Math.max(1, SITE_EDIT_SESSION_MINUTES - elapsedMinutes) };
    }
    return { success: true, canEnter: true, concurrent: activeOthers.length > 0, waitMinutes: 0 };
  });
}

function heartbeatSiteEditSession(siteId, deviceToken) {
  return claimSiteEditSession(siteId, deviceToken, false);
}

function releaseSiteEditSession(siteId, deviceToken) {
  requireAdmin_();
  const normalizedSiteId = normalizePublicSiteId_(siteId);
  const token = normalizeEditDeviceToken_(deviceToken);
  return withScriptLock_(function () {
    const cache = CacheService.getScriptCache();
    const key = getSiteEditSessionKey_(normalizedSiteId);
    let state = { editors: [] };
    try { state = JSON.parse(cache.get(key) || '{"editors":[]}'); } catch (e) {}
    state.editors = (state.editors || []).filter(function (editor) { return editor.token !== token; });
    if (state.editors.length) cache.put(key, JSON.stringify(state), 900); else cache.remove(key);
    return { success: true };
  });
}

function sendAdminSiteLinks(data) {
  requireAdmin_();
  const result = { success: false, email: null, drive: null };
  try {
    if (!data || !data.siteId) throw new Error('링크를 만들 교육 사이트가 없습니다.');
    const site = getSiteInfoById_(normalizePublicSiteId_(data.siteId));
    if (!site || isSiteTrashed_(site)) throw new Error('링크를 만들 교육 사이트를 찾을 수 없습니다.');
    if (site.finalCompleted !== '완료') throw new Error('모든 단계를 최종 완료한 뒤 관리자 링크를 받을 수 있습니다.');
    const sendEmail = Boolean(data.sendEmail);
    const saveDrive = Boolean(data.saveDrive);
    if (!sendEmail && !saveDrive) throw new Error('이메일로 받기 또는 Google Drive에 저장하기를 하나 이상 선택해주세요.');
    const adminUrl = createAdminUrl_(site.siteId);
    const userUrl = site.userUrl;
    const dateText = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy년 M월 d일');
    const bodyText = buildAdminLinkMailBody_(dateText, site.title, adminUrl, userUrl);

    if (sendEmail) {
      try {
        const email = String(data.email || '').trim();
        if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('받을 이메일 주소를 올바르게 입력해주세요.');
        if (MailApp.getRemainingDailyQuota() < 1) throw new Error('오늘 보낼 수 있는 이메일 수를 모두 사용했습니다. Google Drive 저장을 이용하거나 다음 날 다시 시도해주세요.');
        MailApp.sendEmail({
          to: email,
          subject: '[V-Up 관리자 링크] ' + site.title,
          body: bodyText,
          htmlBody: buildAdminLinkMailHtml_(dateText, site.title, adminUrl, userUrl),
          name: '한국벤처캐피탈연수원'
        });
        result.email = { success: true, message: '이메일 전송 완료' };
        appendAuditLog_('ADMIN_LINK_EMAIL_SENT', site.siteId, { recipient: maskEmail_(email) });
      } catch (emailError) {
        result.email = { success: false, message: emailError.message };
      }
    }

    if (saveDrive) {
      try {
        const folderId = String(data.folderId || 'root');
        const folder = !folderId || folderId === 'root' ? DriveApp.getRootFolder() : DriveApp.getFolderById(folderId);
        if (folderId === site.materialFolderId || folderId === site.photoFolderId) throw new Error('교육생에게 공개되는 교재·사진 폴더에는 관리자 링크를 저장할 수 없습니다. 관리자 전용 폴더를 선택해주세요.');
        const access = folder.getSharingAccess();
        if (access === DriveApp.Access.ANYONE || access === DriveApp.Access.ANYONE_WITH_LINK) throw new Error('외부 공개 폴더에는 관리자 링크를 저장할 수 없습니다. 비공개 관리자 폴더를 선택해주세요.');
        const documentResult = createOrUpdateAdminLinkDocument_(site, folder, adminUrl, userUrl, dateText);
        result.drive = { success: true, message: 'Google Drive 저장 완료', documentUrl: documentResult.url };
        appendAuditLog_('ADMIN_LINK_DRIVE_SAVED', site.siteId, { folderId: folder.getId() });
      } catch (driveError) {
        result.drive = { success: false, message: driveError.message };
      }
    }
    result.success = Boolean((result.email && result.email.success) || (result.drive && result.drive.success));
    result.adminUrl = adminUrl;
    result.userUrl = userUrl;
    return result;
  } catch (error) {
    result.message = error.message;
    return result;
  }
}

function buildAdminLinkMailBody_(dateText, title, adminUrl, userUrl) {
  return [
    '안녕하세요,',
    '한국벤처캐피탈연수원의 "원활한 교육 운영을 위한 교육생 사이트 생성기"를 이용해주시어 감사드립니다.',
    '',
    dateText + ', ' + title + ' 교육 운영을 위해 생성해주신 사이트를 관리하기 용이하시도록',
    '아래와 같이 링크 보내드립니다.',
    '',
    '1. 관리자용 사이트 바로가기',
    adminUrl,
    '2. 교육생 배포용 사이트 바로가기',
    userUrl,
    '',
    '사이트 내 알림 문구, 공지사항 등을 수정하고 싶으시다면 "관리자용 사이트"에,',
    '교육생이 어떤 화면을 보고 있는지 궁금하시다면 "교육생 배포용 사이트"에 접속하여 주시기 바랍니다.',
    '',
    '이용해주심에 감사드리며, 교육생 사이트 생성기 사용과 관련하여 문의사항이 있으신 경우,',
    '언제든지 "오찬희 연구원"에게 문의해주시기 바랍니다.',
    '',
    '따뜻하고 평온한 하루되시기 바랍니다.',
    '감사합니다.'
  ].join('\n');
}

function buildAdminLinkMailHtml_(dateText, title, adminUrl, userUrl) {
  const lines = buildAdminLinkMailBody_(dateText, title, adminUrl, userUrl).split('\n');
  return lines.map(function (line) {
    if (line === adminUrl) return '<a href="' + escapeHtmlServer_(adminUrl) + '" style="display:inline-block;padding:12px 18px;background:#6d28d9;color:#fff;border-radius:8px;text-decoration:none;font-weight:bold">관리자용 사이트 바로가기</a>';
    if (line === userUrl) return '<a href="' + escapeHtmlServer_(userUrl) + '" style="display:inline-block;padding:12px 18px;background:#0f766e;color:#fff;border-radius:8px;text-decoration:none;font-weight:bold">교육생 배포용 사이트 바로가기</a>';
    return line ? escapeHtmlServer_(line) : '&nbsp;';
  }).join('<br>');
}

function createOrUpdateAdminLinkDocument_(site, folder, adminUrl, userUrl, dateText) {
  let doc = null;
  if (site.adminLinkDocumentId) {
    try { doc = DocumentApp.openById(site.adminLinkDocumentId); } catch (e) { doc = null; }
  }
  if (!doc) doc = DocumentApp.create('[V-Up 관리자 링크] ' + site.title + ' - ' + dateText);
  const body = doc.getBody();
  body.clear();
  body.appendParagraph('V-Up 교육생 사이트 관리 링크').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(site.title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('생성·안내일: ' + dateText);
  body.appendParagraph('이 사이트 관리하기').setLinkUrl(adminUrl).setBold(true).setFontSize(16).setForegroundColor('#6D28D9');
  body.appendParagraph(adminUrl);
  body.appendParagraph('교육생 화면 확인하기').setLinkUrl(userUrl).setBold(true).setFontSize(14).setForegroundColor('#0F766E');
  body.appendParagraph(userUrl);
  body.appendParagraph('관리자용 링크는 승인된 관리자에게만 전달해주세요.');
  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  file.moveTo(folder);
  const sheet = getOrCreateSitesSheet_(getSpreadsheet_());
  updateSiteRowById_(sheet, site.siteId, function (row) {
    row[COL.adminLinkDocumentId - 1] = doc.getId();
    row[COL.updatedAt - 1] = new Date();
    return row;
  });
  return { id: doc.getId(), url: doc.getUrl() };
}

function maskEmail_(email) {
  const parts = String(email || '').split('@');
  if (parts.length !== 2) return '***';
  return (parts[0].slice(0, 2) || '*') + '***@' + parts[1];
}

function escapeHtmlServer_(value) {
  return String(value == null ? '' : value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/****************************************************
 * 8. 구버전 호환 함수 (LEGACY)
 * ------------------------------------------------------
 * uploadLecturePdfFile() / getLecturePdf() 는 admin-script.html,
 * admin-main.html 어디에서도 더 이상 호출되지 않는 것을 확인했습니다.
 * (현재는 saveMaterialPdfFiles() / getMaterialPdfFiles() 로 대체되었습니다.)
 *
 * 다만 index.html / index-script.html 은 이번 검토 대상에 포함되지 않아
 * 그쪽에서 이 이름을 그대로 호출하고 있을 가능성을 배제할 수 없으므로,
 * 삭제하지 않고 맨 아래로 옮겨 Step 1~3의 핵심 로직과 분리만 해두었습니다.
 * index 계열 파일에서 미사용이 확인되면 이 섹션 전체를 삭제해도 안전합니다.
 ****************************************************/

function uploadLecturePdfFile(data) {
  requireAdmin_();
  try {
    if (!data || !data.siteId) throw new Error('연결할 사용자 사이트가 없습니다.');
    if (!data.base64) throw new Error('PDF 파일 데이터가 없습니다.');

    return saveMaterialPdfFiles({
      siteId: data.siteId,
      targetFolderId: data.targetFolderId || 'root',
      files: [{
        base64: data.base64,
        mimeType: data.mimeType || 'application/pdf',
        fileName: data.fileName || '교재.pdf'
      }]
    });

  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getLecturePdf(siteId) {
  const res = getMaterialPdfFiles(siteId);
  const first = res.files && res.files.length ? res.files[0] : null;
  return {
    success: res.success,
    pdfUrl: first ? first.viewUrl : '',
    files: res.files || [],
    folderUrl: res.folderUrl || '',
    message: res.message || ''
  };
}
