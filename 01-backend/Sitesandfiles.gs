/****************************************************
 * [3_SitesAndFiles.gs]
 * 교육 사이트 생성·수정·휴지통 · 교재/사진 · 교육생 공개 API · 리소스/업로드
 ****************************************************/

/****************************************************
 * 1. 교육 사이트 생성 / 수정 / 목록 / 점검 / 휴지통
 ****************************************************/

function createEducationSite(data) {
  requireAdmin_();
  const createLock = LockService.getScriptLock();
  const alreadyOwned = createLock.hasLock();
  if (!alreadyOwned) createLock.waitLock(15000);
  try {
    if (!data) throw new Error('입력 데이터가 없습니다.');

    const isNonVup =
      data.isNonVup === true ||
      String(data.isNonVup || '').toLowerCase() === 'true';

    const educationInput = String(data.generation || '').trim();

    if (!educationInput) {
      throw new Error(
        isNonVup
          ? '교육명을 입력해주세요.'
          : '기수를 입력해주세요.'
      );
    }

    /*
    * 일반 숙박교육은 글자 수·연도·기수 조건을 두지 않습니다.
    * V-Up 교육만 "OO기" 형식을 확인합니다.
    */
    if (!isNonVup && !/^\d+기$/.test(educationInput)) {
      throw new Error('기수는 "OO기" 형식으로 입력해주세요.');
    }

    if (!String(data.managerName || '').trim()) {
      throw new Error('담당자 이름을 입력해주세요.');
    }

    const managerPhone = normalizeManagerPhone_(data.managerPhone);

    const courseDays = normalizeCourseDays_(data.courseDays);
    const dayDates = extractDayDatesFromData_(data, courseDays);
    const mealSlots = normalizeMealSlots_(data.mealSlots, courseDays);

    const ss = getSpreadsheet_();
    const sheet = getOrCreateSitesSheet_(ss);
    const requestId = scopeAdminRequestId_(data.requestId, 'course');

    if (requestId) {
      const existingRequestRow = findSiteRowByRequestId_(sheet, requestId);
      if (existingRequestRow) {
        const existingSite = rowToSiteInfo_(
          sheet.getRange(existingRequestRow, 1, 1, SITE_HEADERS.length).getValues()[0]
        );
        const existingOwner = getCourseOwnerId_(existingSite.siteId);
        if (existingOwner && existingOwner !== requireAdmin_().id) {
          throw new Error('다른 관리자의 생성 요청은 재사용할 수 없습니다.');
        }
        return {
          success: true,
          siteId: existingSite.siteId,
          userUrl: existingSite.userUrl,
          title: existingSite.title,
          reused: true,
          message: '이미 완료된 생성 요청의 결과를 다시 불러왔습니다.'
        };
      }
    }

    const rawEducationValue = educationInput;

    let generationText = '';
    let siteTitle = '';

    if (isNonVup) {
      generationText = rawEducationValue;
      siteTitle = rawEducationValue;
    } else {
      const generationNumber = normalizeGeneration_(rawEducationValue);
      generationText = generationNumber + '기';
      siteTitle = `제${generationNumber}기 벤처투자 전문인력 양성과정(V-Up)`;
    }

    const siteId = createSiteId_();
    const userUrl = createUserUrl_(siteId);
    const now = new Date();

    const row = new Array(SITE_HEADERS.length).fill('');
    row[COL.siteId - 1] = siteId;
    row[COL.generation - 1] = generationText;
    row[COL.title - 1] = siteTitle;
    row[COL.managerName - 1] = String(data.managerName || '').trim();
    row[COL.managerPhone - 1] = managerPhone;
    row[COL.day1Date - 1] = dayDates[0] || '';
    row[COL.day2Date - 1] = dayDates[1] || '';
    row[COL.scheduleJson - 1] = JSON.stringify({ courseDays: courseDays, days: dayDates });
    row[COL.mealSlotsJson - 1] = JSON.stringify(mealSlots);
    row[COL.userUrl - 1] = userUrl;
    row[COL.createdAt - 1] = now;
    row[COL.updatedAt - 1] = now;
    row[COL.status - 1] = '사용중';
    if (requestId) row[COL.memo - 1] = 'requestId:' + requestId;

    row[COL.revision - 1] = 1;
    sheet.appendRow(safeSheetRow_(row));
    clearSiteListCache_();
    appendAuditLog_('SITE_CREATED', siteId, { title: siteTitle });

    return {
      success: true,
      siteId,
      userUrl,
      adminUrl: createAdminUrl_(siteId),
      title: siteTitle,
      message: '사용자용 사이트가 생성되었습니다.'
    };

  } catch (error) {
    return { success: false, message: error.message };
  } finally {
    if (!alreadyOwned) { try { createLock.releaseLock(); } catch (e) {} }
  }
}

function updateEducationSiteBasicInfo(data) {
  requireAdmin_();
  try {
    if (!data || !data.siteId) throw new Error('수정할 교육 사이트가 없습니다.');

    const isNonVup =
      data.isNonVup === true ||
      String(data.isNonVup || '').toLowerCase() === 'true';

    const educationInput = String(data.generation || '').trim();

    if (!educationInput) {
      throw new Error(
        isNonVup
          ? '교육명을 입력해주세요.'
          : '기수를 입력해주세요.'
      );
    }

    /*
    * 일반 숙박교육은 글자 수·연도·기수 조건을 두지 않습니다.
    * V-Up 교육만 "OO기" 형식을 확인합니다.
    */
    if (!isNonVup && !/^\d+기$/.test(educationInput)) {
      throw new Error('기수는 "OO기" 형식으로 입력해주세요.');
    }

    if (!String(data.managerName || '').trim()) {
      throw new Error('담당자 이름을 입력해주세요.');
    }

    const managerPhone = normalizeManagerPhone_(data.managerPhone);

    const courseDays = normalizeCourseDays_(data.courseDays);
    const dayDates = extractDayDatesFromData_(data, courseDays);
    const mealSlots = normalizeMealSlots_(data.mealSlots, courseDays);

    const ss = getSpreadsheet_();
    const sheet = getOrCreateSitesSheet_(ss);

    const rawEducationValue = educationInput;

    let generationText = '';
    let siteTitle = '';

    if (isNonVup) {
      generationText = rawEducationValue;
      siteTitle = rawEducationValue;
    } else {
      const generationNumber = normalizeGeneration_(rawEducationValue);
      generationText = generationNumber + '기';
      siteTitle = `제${generationNumber}기 벤처투자 전문인력 양성과정(V-Up)`;
    }

    let resetLinkedData = false;
    updateSiteRowById_(sheet, data.siteId, function (row) {
      assertExpectedRevision_(data.revision, Number(row[COL.revision - 1] || 0));
      const previousSite = rowToSiteInfo_(row);
      resetLinkedData = getMealSlotShapeKey_(previousSite.mealSlots) !== getMealSlotShapeKey_(mealSlots);
      row[COL.generation - 1] = generationText;
      row[COL.title - 1] = siteTitle;
      row[COL.managerName - 1] = String(data.managerName || '').trim();
      row[COL.managerPhone - 1] = managerPhone;
      row[COL.day1Date - 1] = dayDates[0] || '';
      row[COL.day2Date - 1] = dayDates[1] || '';
      row[COL.scheduleJson - 1] = JSON.stringify({ courseDays: courseDays, days: dayDates });
      row[COL.mealSlotsJson - 1] = JSON.stringify(mealSlots);
      if (resetLinkedData) {
        row[COL.linkedSheetId - 1] = '';
        row[COL.linkedSheetUrl - 1] = '';
        row[COL.dataReviewConfirmedAt - 1] = '';   // 교육생 검토 확인 해제
        row[COL.finalCompleted - 1] = '';          // 최종 완료 도장 해제
        row[COL.lastApprovedRevision - 1] = '';    // 승인 이력 초기화
      }
      row[COL.revision - 1] = Number(row[COL.revision - 1] || 0) + 1;
      row[COL.updatedAt - 1] = new Date();
      return row;
    });
    clearUserDataCache_(data.siteId);
    clearSiteListCache_();
    appendAuditLog_('SITE_BASIC_INFO_UPDATED', data.siteId, { mealShapeChanged: resetLinkedData });

    return {
      success: true,
      title: siteTitle,
      resetLinkedData: resetLinkedData,
      preservedFileFolders: true,
      message: resetLinkedData
        ? '기본 정보가 수정되었습니다. 식사 구성 변경으로 식사 배정표만 다시 연결해주세요.'
        : '기본 정보가 수정되었고 기존 식사·교재·사진 연결은 유지되었습니다.'
    };

  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getEducationSites(forceRefresh) {
  requireAdmin_();
  try {
    const cacheKey = 'educationSites:list';
    const cached = getUserDataCache_(cacheKey, !!forceRefresh);
    if (cached) return cached;
    purgeExpiredTrashedSites_();

    const ss = getSpreadsheet_();
    const sheet = getOrCreateSitesSheet_(ss);
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      const emptyResult = { success: true, sites: [] };
      putSiteListCache_(cacheKey, emptyResult);
      return emptyResult;
    }

    const values = sheet.getRange(2, 1, lastRow - 1, SITE_HEADERS.length).getValues();
    const sites = values
      .filter(row => String(row[COL.siteId - 1] || '').trim() !== '')
      .map(row => decorateAdminSiteStatus_(rowToSiteInfo_(row)))
      .filter(function (site) { return !site.deletedAt; })
      .sort(function (a, b) {
        const timeA = new Date(a.createdAt).getTime();
        const timeB = new Date(b.createdAt).getTime();
        const safeA = isNaN(timeA) ? 0 : timeA;
        const safeB = isNaN(timeB) ? 0 : timeB;
        return safeB - safeA;   // 최신 생성 순
      });

    const result = { success: true, sites };
    putSiteListCache_(cacheKey, result);
    return result;

  } catch (error) {
    return { success: false, message: error.message, sites: [] };
  }
}

function decorateAdminSiteStatus_(site) {
  const result = Object.assign({}, site || {});
  if (result.finalCompleted === '완료' &&
      Number(result.lastApprovedRevision || 0) === Number(result.revision || 0)) {
    result.workflowStatusKey = 'complete';
    result.workflowStatusLabel = '운영 준비 완료';
  } else if (!result.userUrl) {
    result.workflowStatusKey = 'attention';
    result.workflowStatusLabel = '기본 정보 확인';
  } else if (!result.linkedSheetUrl) {
    result.workflowStatusKey = 'attention';
    result.workflowStatusLabel = '시트 연결 필요';
  } else if (!result.dataReviewConfirmedAt) {
    result.workflowStatusKey = 'attention';
    result.workflowStatusLabel = '교육생 검토 필요';
  } else if (!result.materialFolderUrl) {
    result.workflowStatusKey = 'progress';
    result.workflowStatusLabel = '교재 준비 중';
  } else if (!result.photoFolderUrl) {
    result.workflowStatusKey = 'progress';
    result.workflowStatusLabel = '사진 준비 중';
  } else {
    result.workflowStatusKey = 'progress';
    result.workflowStatusLabel = '최종 점검 대기';
  }
  return result;
}

function getSiteHealthCheck(siteId) {
  requireAdmin_();
  try {
    return buildSiteHealthCheck_(siteId);
  } catch (error) {
    return { success: false, ready: false, blockerCount: 1, warningCount: 0, items: [], message: error.message };
  }
}

function buildSiteHealthCheck_(siteId) {
  const normalizedSiteId = String(siteId || '').trim();
  if (!normalizedSiteId) throw new Error('점검할 교육 사이트 ID가 없습니다.');
  const sheet = getSitesSheetForRead_();
  const site = findSiteInfoFast_(sheet, normalizedSiteId);
  if (!site) throw new Error('점검할 교육 사이트를 찾을 수 없습니다.');
  if (isSiteTrashed_(site)) throw new Error('휴지통에 있는 사이트는 완료 처리할 수 없습니다.');

  const items = [];
  const add = function (code, label, status, detail) {
    items.push({ code: code, label: label, status: status, detail: detail });
  };

  let participants = [];
  if (!site.linkedSheetId) {
    add('linkedSheet', '교육생·식사 시트 연결', 'blocker', '연결된 Google Sheet가 없습니다.');
  } else {
    try {
      participants = getParticipantsFromLinkedSheet_(site.linkedSheetId);
      add('linkedSheet', '교육생·식사 시트 연결', participants.length ? 'pass' : 'blocker', participants.length ? '교육생 ' + participants.length + '명을 읽었습니다.' : '교육생 데이터를 찾지 못했습니다.');
    } catch (error) {
      add('linkedSheet', '교육생·식사 시트 연결', 'blocker', '연결 시트를 읽지 못했습니다: ' + error.message);
    }
  }

  if (participants.length) {
    add('participantReview', '교육생 정보 관리자 확인', site.dataReviewConfirmedAt ? 'pass' : 'blocker', site.dataReviewConfirmedAt ? '관리자가 이름·소속·직급·전화번호를 확인했습니다.' : '자동 인식한 교육생 정보표를 관리자가 아직 확인하지 않았습니다.');
    const missingPhones = participants.filter(function (person) { return !/^\d{4}$/.test(String(person.phoneLast4 || '')); }).length;
    add('phoneLookup', '본인 확인용 전화번호', missingPhones ? 'blocker' : 'pass', missingPhones ? '전화번호 뒷자리 4개가 없는 교육생이 ' + missingPhones + '명 있습니다.' : '모든 교육생의 전화번호 뒷자리 4개를 확인했습니다.');

    const credentialCounts = {};
    participants.forEach(function (person) {
      const key = String(person.baseNameKey || '') + '|' + String(person.phoneLast4 || '');
      if (person.baseNameKey && person.phoneLast4) credentialCounts[key] = (credentialCounts[key] || 0) + 1;
    });
    const ambiguousCredentials = Object.keys(credentialCounts).filter(function (key) { return credentialCounts[key] > 1; }).length;
    add('duplicateCredential', '동명이인 본인 확인', ambiguousCredentials ? 'blocker' : 'pass', ambiguousCredentials ? '같은 순수 이름과 전화번호 뒷자리가 겹치는 묶음이 ' + ambiguousCredentials + '개 있습니다.' : '이름과 전화번호 뒷자리 조합이 모두 구분됩니다.');

    const inferredNames = participants.filter(function (person) { return person.baseNameInferred; }).length;
    add('inferredName', '수식어가 붙은 이름', inferredNames ? 'warning' : 'pass', inferredNames ? '공통 핵심 이름으로 추정한 교육생이 ' + inferredNames + '명 있습니다. 관리자 확인이 필요합니다.' : '추정 확인이 필요한 이름이 없습니다.');

    const possibleShortCorePairs = getPossibleShortKoreanCoreNamePairs_(participants).length;
    add('possibleShortName', '두 글자 동명이인 후보', possibleShortCorePairs ? 'warning' : 'pass', possibleShortCorePairs ? '앞 수식어를 제외하면 같은 두 글자 이름일 수 있는 조합이 ' + possibleShortCorePairs + '개 있습니다. 명단 표기를 확인해주세요.' : '추가 확인이 필요한 두 글자 이름 조합이 없습니다.');

    const configuredSlots = Array.isArray(site.mealSlots) ? site.mealSlots : [];
    const runtimeSlots = getMealRuntimeSlots_(site);
    if (!configuredSlots.length || runtimeSlots.length !== configuredSlots.length) {
      add('mealSchedule', '식사 일정', 'blocker', '식사 일정 또는 교육 날짜가 올바르게 설정되지 않았습니다.');
    } else {
      add('mealSchedule', '식사 일정', 'pass', '식사 일정 ' + runtimeSlots.length + '개를 확인했습니다.');
      let missingSeats = 0;
      runtimeSlots.forEach(function (slot) {
        participants.forEach(function (person) {
          if (!String(getParticipantMealSeatValue_(person, slot.seatKey) || '').trim()) missingSeats++;
        });
      });
      add('mealSeats', '식사 자리 배정', missingSeats ? 'blocker' : 'pass', missingSeats ? '교육생×식사 기준 미배정 항목이 ' + missingSeats + '개 있습니다.' : '모든 교육생의 식사 자리가 배정되었습니다.');
    }
  }

  const publicConfig = getPublicCourseAccessConfig_(normalizedSiteId);
  const inspectFolder = function (code, label, folderId, mimeTypes, extensions) {
    if (!folderId) {
      add(code, label, 'blocker', '연결된 Google Drive 폴더가 없습니다.');
      return;
    }
    try {
      const folder = DriveApp.getFolderById(folderId);
      const sharingAccess = folder.getSharingAccess();
      const isPubliclyReadable = sharingAccess === DriveApp.Access.ANYONE ||
        sharingAccess === DriveApp.Access.ANYONE_WITH_LINK;
      const moduleKey = code === 'photoFolder' ? 'photos' : 'materials';
      const shouldBePublic = !publicConfig.configured ||
        (publicConfig.siteAvailable && publicConfig.modules[moduleKey]);
      if (shouldBePublic && !isPubliclyReadable) {
        add(code, label, 'blocker', '현재 공개 시점이지만 폴더가 비공개입니다. 자료 공개 정책을 다시 적용해주세요.');
        return;
      }
      if (!shouldBePublic && isPubliclyReadable) {
        add(code + 'Sharing', label + ' 공개 범위', 'warning', '아직 공개 전이거나 운영이 중단된 자료가 링크 공개 상태입니다.');
      }
      const page = listFolderFilesPage_(folderId, mimeTypes, extensions, 0, 1);
      add(code, label, page.files.length ? 'pass' : 'warning', page.files.length
        ? (shouldBePublic ? '현재 공개 정책과 전달 파일을 확인했습니다.' : '비공개 준비 상태의 파일을 확인했습니다.')
        : '폴더가 비어 있습니다.');
    } catch (error) {
      add(code, label, 'blocker', 'Google Drive 폴더를 읽지 못했습니다: ' + error.message);
    }
  };
  inspectFolder('materialFolder', '교재 파일 폴더', site.materialFolderId, getAllowedMaterialMimeTypes_(), getAllowedMaterialExtensions_());
  inspectFolder('photoFolder', '현장 사진 폴더', site.photoFolderId, ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'image/heic', 'image/heif'], ['jpg', 'jpeg', 'png', 'gif', 'webp', 'heic', 'heif']);

  const blockerCount = items.filter(function (item) { return item.status === 'blocker'; }).length;
  const warningCount = items.filter(function (item) { return item.status === 'warning'; }).length;
  return {
    success: true,
    ready: blockerCount === 0,
    blockerCount: blockerCount,
    warningCount: warningCount,
    items: items,
    checkedAtText: Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss')
  };
}

function deleteEducationSite(siteId) {
  requireAdmin_();
  try {
    const ss = getSpreadsheet_();
    const sheet = getOrCreateSitesSheet_(ss);
    const now = new Date();
    const purgeAt = new Date(now.getTime() + SITE_TRASH_DAYS * 24 * 60 * 60 * 1000);
    updateSiteRowById_(sheet, siteId, function (row) {
      row[COL.preTrashStatus - 1] = String(row[COL.status - 1] || '사용중');
      row[COL.deletedAt - 1] = now;
      row[COL.purgeAt - 1] = purgeAt;
      row[COL.status - 1] = '휴지통';
      row[COL.updatedAt - 1] = now;
      return row;
    });
    clearUserDataCache_(siteId);
    clearSiteListCache_();
    applyCourseResourceSharingPolicy_(siteId);
    appendAuditLog_('SITE_MOVED_TO_TRASH', siteId, { purgeAt: formatDate_(purgeAt) });
    return { success: true, recoverableUntil: formatDate_(purgeAt), message: '휴지통으로 이동했습니다. 30일 안에는 같은 주소로 복원할 수 있습니다.' };

  } catch (error) {
    return { success: false, message: error.message };
  }
}

function markFinalCompleted(siteId) {
  requireAdmin_();
  try {
    const health = buildSiteHealthCheck_(siteId);
    if (!health.ready) {
      throw new Error('배포 전 점검에서 수정이 필요한 항목이 ' + health.blockerCount + '개 확인되었습니다. 점검 결과를 먼저 해결해주세요.');
    }
    const ss = getSpreadsheet_();
    const sheet = getOrCreateSitesSheet_(ss);
    const rowIndex = findSiteRowIndex_(sheet, siteId);
    if (!rowIndex) throw new Error('해당 교육 사이트를 찾을 수 없습니다.');

    const now = new Date();
    updateSiteRowById_(sheet, siteId, function (row) {
      row[COL.finalCompleted - 1] = '완료';
      row[COL.finalCompletedAt - 1] = now;
      row[COL.lastApprovedRevision - 1] = Math.max(0, Number(row[COL.revision - 1] || 0)) + 1;
      row[COL.updatedAt - 1] = now;
      return row;
    });
    clearUserDataCache_(siteId);
    clearSiteListCache_();
    appendAuditLog_('SITE_FINAL_COMPLETED', siteId, {});

    return { success: true, adminUrl: createAdminUrl_(siteId), userUrl: findSiteInfo_(sheet, siteId).userUrl };

  } catch (error) {
    return { success: false, message: error.message };
  }
}

/****************************************************
 * 휴지통 · 복원 · 영구 삭제 · 자동 정리 트리거
 ****************************************************/

function getTrashedEducationSites() {
  requireAdmin_();
  try {
    purgeExpiredTrashedSites_();
    const sheet = getOrCreateSitesSheet_(getSpreadsheet_());
    if (sheet.getLastRow() < 2) return { success: true, sites: [] };
    const sites = sheet.getRange(2, 1, sheet.getLastRow() - 1, SITE_HEADERS.length).getValues()
      .map(rowToSiteInfo_)
      .filter(function (site) { return Boolean(site.siteId && site.deletedAt); })
      .sort(function (a, b) { return String(b.deletedAt).localeCompare(String(a.deletedAt)); });
    return { success: true, sites: sites };
  } catch (error) {
    return { success: false, message: error.message, sites: [] };
  }
}

function restoreEducationSite(siteId) {
  requireAdmin_();
  try {
    const sheet = getOrCreateSitesSheet_(getSpreadsheet_());
    updateSiteRowById_(sheet, siteId, function (row) {
      if (!row[COL.deletedAt - 1]) throw new Error('휴지통에 있는 사이트가 아닙니다.');
      row[COL.deletedAt - 1] = '';
      row[COL.purgeAt - 1] = '';
      row[COL.status - 1] = String(row[COL.preTrashStatus - 1] || '사용중');
      row[COL.preTrashStatus - 1] = '';
      row[COL.revision - 1] = Number(row[COL.revision - 1] || 0) + 1;
      row[COL.updatedAt - 1] = new Date();
      return row;
    });
    clearUserDataCache_(siteId);
    clearSiteListCache_();
    applyCourseResourceSharingPolicy_(siteId);
    appendAuditLog_('SITE_RESTORED', siteId, {});
    return { success: true, message: '기존 사용자 주소와 연결 정보를 그대로 복원했습니다.' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function permanentlyDeleteEducationSite_(siteId, actorId) {
  return withScriptLock_(function () {
  const normalizedSiteId = normalizePublicSiteId_(siteId);
  delete SITE_INFO_RUNTIME_CACHE[normalizedSiteId];
  if (String(actorId) !== 'AUTO_PURGE') {
    if (ADMIN_REQUEST_CACHE) ADMIN_REQUEST_CACHE.ownerMap = null;
    assertCourseAccess_(requireAdmin_(), normalizedSiteId);
  }
  const spreadsheet = getSpreadsheet_();
  const sheet = getOrCreateSitesSheet_(spreadsheet);
  const rowIndex = findPortalRowByValue_(sheet, COL.siteId, normalizedSiteId);
  if (!rowIndex) throw new Error('삭제할 강의를 찾을 수 없습니다.');
  const row = sheet.getRange(rowIndex, 1, 1, SITE_HEADERS.length).getValues()[0];
  const site = rowToSiteInfo_(row);
  if (!site.deletedAt) {
    throw new Error('휴지통에 있는 강의만 완전히 삭제할 수 있습니다.');
  }
  const resourceResults = getCourseResourceRecords_(normalizedSiteId).map(function (resource) {
    return trashManagedResource_(resource);
  });
  appendAuditLog_('SITE_PERMANENTLY_REMOVED', normalizedSiteId, {
    actorId: String(actorId || 'AUTO_PURGE'),
    resourceResults: resourceResults
  });
  sheet.deleteRow(rowIndex);
  deletePortalRowsByValue_(SHEET_ADMIN_COURSE_OWNERS, 1, normalizedSiteId);
  deletePortalRowsByValue_(SHEET_ADMIN_COURSE_SETTINGS, 1, normalizedSiteId);
  deletePortalRowsByValue_(SHEET_ADMIN_SURVEYS, 2, normalizedSiteId);
  if (typeof OPERATIONS_SHEET_ !== 'undefined') deletePortalRowsByValue_(OPERATIONS_SHEET_, 1, normalizedSiteId);
  delete SITE_INFO_RUNTIME_CACHE[normalizedSiteId];
  clearUserDataCache_(normalizedSiteId);
  clearSiteListCache_();
  SpreadsheetApp.flush();
  }, 15000);
}

function deletePortalRowsByValue_(sheetName, columnIndex, value) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;
  const values = sheet.getRange(2, columnIndex, sheet.getLastRow() - 1, 1).getValues();
  for (let index = values.length - 1; index >= 0; index--) {
    if (String(values[index][0] || '') === String(value || '')) {
      sheet.deleteRow(index + 2);
    }
  }
}

function purgeExpiredTrashedSites_() {
  const sheet = getOrCreateSitesSheet_(getSpreadsheet_());
  if (sheet.getLastRow() < 2) return;
  const now = Date.now();
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, SITE_HEADERS.length).getValues();
  const expiredSiteIds = [];
  values.forEach(function (row) {
    const deletedAt = row[COL.deletedAt - 1];
    const purgeAt = row[COL.purgeAt - 1];
    if (!deletedAt || !purgeAt) return;
    const purgeTime = purgeAt instanceof Date ? purgeAt.getTime() : new Date(purgeAt).getTime();
    if (isFinite(purgeTime) && purgeTime <= now) {
      expiredSiteIds.push(String(row[COL.siteId - 1] || ''));
    }
  });
  expiredSiteIds.slice(0, 10).forEach(function (siteId) {
    try {
      permanentlyDeleteEducationSite_(siteId, 'AUTO_PURGE');
    } catch (error) {
      appendAuditLog_('SITE_AUTO_PURGE_FAILED', siteId, { message: error.message });
    }
  });
  if (expiredSiteIds.length) {
    SpreadsheetApp.flush();
    clearSiteListCache_();
  }
}

function purgeExpiredTrashedSitesTrigger_() {
  // 끝의 밑줄이 google.script.run 노출을 차단합니다. 트리거 존재 여부는 호출 인증이 아닙니다.
  purgeExpiredTrashedSites_();
}

function setupDailyTrashCleanupTrigger_() {
  const activeEmail = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  const effectiveEmail = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (!activeEmail || activeEmail !== effectiveEmail) throw new Error('프로젝트 소유자가 편집기에서 직접 실행해야 합니다.');
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (['purgeExpiredTrashedSites', 'reconcileCourseResourcePolicies', 'purgeExpiredTrashedSitesTrigger_', 'reconcileCourseResourcePoliciesTrigger_'].indexOf(trigger.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('purgeExpiredTrashedSitesTrigger_')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
  ScriptApp.newTrigger('reconcileCourseResourcePoliciesTrigger_')
    .timeBased()
    .everyMinutes(15)
    .create();
  return '휴지통 자동 정리 및 공개 일정 동기화 트리거를 등록했습니다.';
}

function reconcileCourseResourcePoliciesTrigger_() {
  const sheet = getOrCreateSitesSheet_(getSpreadsheet_());
  if (sheet.getLastRow() < 2) return;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, SITE_HEADERS.length).getValues()
    .map(rowToSiteInfo_)
    .filter(function (site) { return site.siteId && !site.deletedAt; })
    .slice(0, 100)
    .forEach(function (site) { applyCourseResourceSharingPolicy_(site.siteId); });
}

/****************************************************
 * 3. 교재 파일 첨부 · 4. 현장 사진 첨부
 ****************************************************/

function saveMaterialPdfFiles(data) {
  requireAdmin_();
  const result = createServiceFolderAndUploadFiles_({
    siteId: data && data.siteId,
    targetFolderId: data && data.targetFolderId,
    files: data && data.files,
    serviceType: 'material'
  });
  clearUserDataCache_(data && data.siteId);
  return result;
}

function createMaterialPdfFolderOnly(data) {
  requireAdmin_();
  const result = createServiceFolderAndUploadFiles_({
    siteId: data && data.siteId,
    targetFolderId: data && data.targetFolderId,
    files: [],
    serviceType: 'material'
  });
  clearUserDataCache_(data && data.siteId);
  return result;
}

function getMaterialPdfFiles(siteId, forceRefresh) {
  try {
    const normalizedSiteId = normalizePublicSiteId_(siteId);
    const cacheKey = 'materialFiles:v2:' + normalizedSiteId;
    // 공개 API의 forceRefresh 값은 신뢰하지 않습니다. 최대 30초 캐시로 Drive 호출 폭주를 막습니다.
    const cached = getUserDataCache_(cacheKey, false);
    if (cached) {
      const activeSite = getSiteInfoById_(normalizedSiteId);
      ensurePublicSiteActive_(activeSite);
      assertPublicModuleAccess_(activeSite, 'materials');
      return cached;
    }

    const site = getSiteInfoById_(normalizedSiteId);
    ensurePublicSiteActive_(site);
    assertPublicModuleAccess_(site, 'materials');

    const files = listFolderFiles_(site.materialFolderId, getAllowedMaterialMimeTypes_(), getAllowedMaterialExtensions_());
    const result = {
      success: true,
      files,
      checkedAtText: Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm:ss')
    };
    putUserDataCache_(cacheKey, result);
    return result;

  } catch (error) {
    return { success: false, message: error.message, files: [], checkedAtText: '' };
  }
}

/**
 * 관리자 화면 전용 즉시 확인 API입니다.
 * 공개 API의 캐시는 우회시키지 않되 관리자 전용 배포에서는 Drive를
 * 즉시 다시 읽고, 확인된 결과로 공개 캐시까지 안전하게 갱신합니다.
 */
function getAdminMaterialFiles(siteId) {
  requireAdmin_();
  try {
    const normalizedSiteId = normalizePublicSiteId_(siteId);
    const site = getSiteInfoById_(normalizedSiteId);
    ensurePublicSiteActive_(site);
    const files = listFolderFiles_(site.materialFolderId, getAllowedMaterialMimeTypes_(), getAllowedMaterialExtensions_());
    const result = {
      success: true,
      files: files,
      checkedAtText: Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm:ss')
    };
    // 관리자 응답은 공개 캐시에 넣지 않습니다. 공개 전 교재가 노출될 수 있습니다.
    appendAuditLog_('ADMIN_MATERIAL_FILES_REFRESHED', normalizedSiteId, { fileCount: files.length });
    return result;
  } catch (error) {
    appendAuditLog_('ADMIN_MATERIAL_FILES_REFRESH_FAILED', String(siteId || ''), { message: error.message });
    return { success: false, message: error.message, files: [], checkedAtText: '' };
  }
}

/****************************************************
 * 4. 현장 사진 첨부
 ****************************************************/

function savePhotoFiles(data) {
  requireAdmin_();
  const result = createServiceFolderAndUploadFiles_({
    siteId: data && data.siteId,
    targetFolderId: data && data.targetFolderId,
    files: data && data.files,
    serviceType: 'photo'
  });
  clearUserDataCache_(data && data.siteId);
  return result;
}

function createPhotoFolderOnly(data) {
  requireAdmin_();
  const result = createServiceFolderAndUploadFiles_({
    siteId: data && data.siteId,
    targetFolderId: data && data.targetFolderId,
    files: [],
    serviceType: 'photo'
  });
  clearUserDataCache_(data && data.siteId);
  return result;
}

function getPhotoFiles(siteId, forceRefresh) {
  try {
    const normalizedSiteId = normalizePublicSiteId_(siteId);
    const cacheKey = 'photoFiles:v2:' + normalizedSiteId;
    // 공개 API의 forceRefresh 값은 신뢰하지 않습니다. 최대 30초 캐시로 Drive 호출 폭주를 막습니다.
    const cached = getUserDataCache_(cacheKey, false);
    if (cached) {
      const activeSite = getSiteInfoById_(normalizedSiteId);
      ensurePublicSiteActive_(activeSite);
      assertPublicModuleAccess_(activeSite, 'photos');
      return cached;
    }

    const site = getSiteInfoById_(normalizedSiteId);
    ensurePublicSiteActive_(site);
    assertPublicModuleAccess_(site, 'photos');

    const page = listFolderFilesPage_(site.photoFolderId, ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'image/heic', 'image/heif'], ['jpg', 'jpeg', 'png', 'gif', 'webp', 'heic', 'heif'], 0, 48);
    const result = {
      success: true,
      files: page.files,
      nextPageToken: page.nextPageToken,
      hasMore: page.hasMore,
      checkedAtText: Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm:ss')
    };
    putUserDataCache_(cacheKey, result);
    return result;

  } catch (error) {
    return { success: false, message: error.message, files: [], checkedAtText: '' };
  }
}

function getAdminPhotoFiles(siteId) {
  requireAdmin_();
  try {
    const normalizedSiteId = normalizePublicSiteId_(siteId);
    const site = getSiteInfoById_(normalizedSiteId);
    ensurePublicSiteActive_(site);
    const page = listFolderFilesPage_(
      site.photoFolderId,
      ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'image/heic', 'image/heif'],
      ['jpg', 'jpeg', 'png', 'gif', 'webp', 'heic', 'heif'],
      0,
      48
    );
    const result = {
      success: true,
      files: page.files,
      nextPageToken: page.nextPageToken,
      hasMore: page.hasMore,
      checkedAtText: Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm:ss')
    };
    // 공개 전 사진은 관리자 응답을 통해 공개 캐시에 기록하지 않습니다.
    appendAuditLog_('ADMIN_PHOTO_FILES_REFRESHED', normalizedSiteId, { fileCount: page.files.length, hasMore: page.hasMore });
    return result;
  } catch (error) {
    appendAuditLog_('ADMIN_PHOTO_FILES_REFRESH_FAILED', String(siteId || ''), { message: error.message });
    return { success: false, message: error.message, files: [], checkedAtText: '' };
  }
}

function getUserProvidedFiles(siteId) {
  return {
    success: true,
    materials: getMaterialPdfFiles(siteId),
    photos: getPhotoFiles(siteId)
  };
}

/****************************************************
 * 5. 사용자(교육생) 페이지 데이터 · 식사 좌석 본인 조회
 ****************************************************/

function toPublicSiteInfo_(site) {
  if (!site) return null;
  const publicContact = getPublicManagerContact_(site);
  return {
    siteId: site.siteId,
    generation: site.generation,
    title: site.title,
    isNonVup: site.isNonVup,
    managerName: publicContact.name || site.managerName,
    managerPhone: publicContact.landline || publicContact.mobile || '',
    managerLandline: publicContact.landline,
    managerMobile: publicContact.mobile,
    managerEmail: publicContact.email,
    managerOrganization: publicContact.organization,
    managerPosition: publicContact.position,
    courseDays: site.courseDays,
    day1Date: site.day1Date,
    day2Date: site.day2Date,
    day3Date: site.day3Date,
    day4Date: site.day4Date,
    day5Date: site.day5Date,
    mealSlots: site.mealSlots,
    finalCompleted: site.finalCompleted,
    finalCompletedAt: site.finalCompletedAt
  };
}

function getPhotoFilesPage(siteId, pageToken) {
  try {
    const normalizedSiteId = normalizePublicSiteId_(siteId);
    const site = getSiteInfoById_(normalizedSiteId);
    ensurePublicSiteActive_(site);
    assertPublicModuleAccess_(site, 'photos');
    const offset = Math.max(0, Number(pageToken || 0));
    const page = listFolderFilesPage_(site.photoFolderId, ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'image/heic', 'image/heif'], ['jpg', 'jpeg', 'png', 'gif', 'webp', 'heic', 'heif'], offset, 48);
    return { success: true, files: page.files, nextPageToken: page.nextPageToken, hasMore: page.hasMore };
  } catch (error) {
    return { success: false, message: error.message, files: [], nextPageToken: '', hasMore: false };
  }
}

function normalizePublicSiteId_(siteId) {
  const value = String(siteId || '').trim();
  // 기존 UUID/12자리 ID와 새 24자리 ID를 모두 허용하되, 캐시 키·시트 탐색에 과도한 입력이 들어오지 않게 제한합니다.
  if (!/^[A-Za-z0-9_-]{6,128}$/.test(value)) throw new Error('유효한 교육 사이트 주소가 아닙니다.');
  return value;
}

function getSiteData(siteId) {
  try {
    const normalizedSiteId = normalizePublicSiteId_(siteId);
    const cacheKey = 'siteData:v2:' + normalizedSiteId;
    const cached = getUserDataCache_(cacheKey, false);
    if (cached) {
      const activeSite = getSiteInfoById_(normalizedSiteId);
      ensurePublicSiteActive_(activeSite);
      assertPublicModuleAccess_(activeSite, 'progress');
      return cached;
    }

    const sheet = getSitesSheetForRead_();
    const site = findSiteInfoFast_(sheet, normalizedSiteId);

    ensurePublicSiteActive_(site);
    assertPublicModuleAccess_(site, 'progress');
    const mealRuntimeSlots = getMealRuntimeSlots_(site);

    const result = {
      success: true,
      site: toPublicSiteInfo_(site),
      accessConfig: getPublicCourseAccessConfig_(normalizedSiteId),
      participants: [],
      currentMeal: getCurrentMealTypeFromRuntimeSlots_(mealRuntimeSlots),
      mealRuntimeSlots: mealRuntimeSlots,
      checkedAtText: Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm:ss'),
      materialFiles: [],
      photoFiles: []
    };
    putUserDataCache_(cacheKey, result);
    return result;

  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getMealSeatData(siteId, forceRefresh) {
  try {
    const normalizedSiteId = normalizePublicSiteId_(siteId);
    const cacheKey = 'mealSeatData:v2:' + normalizedSiteId;
    // 공개 API의 forceRefresh 값으로 서버 캐시를 우회하지 못하게 합니다.
    const cached = getUserDataCache_(cacheKey, false);
    if (cached) {
      const activeSite = getSiteInfoById_(normalizedSiteId);
      ensurePublicSiteActive_(activeSite);
      assertPublicModuleAccess_(activeSite, 'meal');
      return cached;
    }

    const sheet = getSitesSheetForRead_();
    const site = findSiteInfoFast_(sheet, normalizedSiteId);

    ensurePublicSiteActive_(site);
    assertPublicModuleAccess_(site, 'meal');
    const mealRuntimeSlots = getMealRuntimeSlots_(site);

    const result = {
      success: true,
      site: toPublicSiteInfo_(site),
      accessConfig: getPublicCourseAccessConfig_(normalizedSiteId),
      participants: [],
      hasParticipantData: Boolean(site.linkedSheetId),
      currentMeal: getCurrentMealTypeFromRuntimeSlots_(mealRuntimeSlots),
      mealRuntimeSlots: mealRuntimeSlots,
      checkedAtText: Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm:ss')
    };
    putUserDataCache_(cacheKey, result);
    return result;

  } catch (error) {
    return { success: false, message: error.message, site: null, participants: [], currentMeal: null, mealRuntimeSlots: [], checkedAtText: '' };
  }
}

function getMealLookupRateKey_(siteId) {
  let userKey = '';
  try { userKey = Session.getTemporaryActiveUserKey(); } catch (e) {}
  return 'mealLookupFailures:' + String(siteId || '') + ':' + String(userKey || 'anonymous');
}

function getMealLookupFailureCount_(siteId) {
  try {
    return Number(CacheService.getScriptCache().get(getMealLookupRateKey_(siteId)) || 0);
  } catch (e) {
    return 0;
  }
}

function recordMealLookupFailure_(siteId) {
  try {
    const cache = CacheService.getScriptCache();
    const key = getMealLookupRateKey_(siteId);
    const count = Number(cache.get(key) || 0) + 1;
    cache.put(key, String(count), 600);
  } catch (e) {
    // 조회 제한 캐시가 실패해도 정상 사용자의 조회 결과에는 영향을 주지 않습니다.
  }
}

function sanitizeMealParticipantForUser_(person, seatKeys, includePhoneLast4, publicPersonKey) {
  const mealSeats = {};
  (seatKeys || []).forEach(function (seatKey) {
    mealSeats[seatKey] = getParticipantMealSeatValue_(person, seatKey);
  });
  const result = {
    name: person.name,
    displayName: person.displayName || person.name,
    baseName: person.baseName,
    nameKey: person.nameKey,
    baseNameKey: person.baseNameKey,
    company: person.company,
    position: person.position,
    phoneLast4: includePhoneLast4 ? person.phoneLast4 : '',
    // 내부 personKey에는 전화번호 뒷자리가 포함될 수 있으므로 브라우저에는 노출하지 않습니다.
    personKey: String(publicPersonKey || ''),
    isDuplicateName: person.isDuplicateName,
    duplicateGroupKey: person.duplicateGroupKey,
    mealSeats: mealSeats
  };
  return Object.assign(result, mealSeats);
}

function getParticipantMealSeatValue_(person, seatKey) {
  if (!person || !seatKey) return '';
  if (person[seatKey]) return String(person[seatKey]);
  if (person.mealSeats && person.mealSeats[seatKey]) return String(person.mealSeats[seatKey]);
  return '';
}

function lookupMealSeatForUser(siteId, enteredName, phoneLast4, forceRefresh) {
  try {
    const normalizedSiteId = normalizePublicSiteId_(siteId);
    const normalizedPhoneLast4 = String(phoneLast4 || '').replace(/\D/g, '');
    const normalizedEnteredName = normalizePersonText_(enteredName);
    if (!normalizedEnteredName || normalizedEnteredName.length > 60) {
      throw new Error('이름과 휴대폰 번호 뒷자리 4자리를 올바르게 입력해주세요.');
    }
    const enteredIdentity = analyzePersonName_(normalizedEnteredName);
    if (!enteredIdentity.baseNameKey || !/^\d{4}$/.test(normalizedPhoneLast4)) {
      throw new Error('이름과 휴대폰 번호 뒷자리 4자리를 올바르게 입력해주세요.');
    }
    if (getMealLookupFailureCount_(normalizedSiteId) >= 12) {
      throw new Error('조회 시도가 많아 잠시 보호 중입니다. 10분 뒤 다시 시도하거나 교육 담당자에게 문의해주세요.');
    }

    const sheet = getSitesSheetForRead_();
    const site = findSiteInfoFast_(sheet, normalizedSiteId);
    if (!site || isSiteTrashed_(site)) throw new Error('입력하신 정보와 일치하는 교육생을 찾을 수 없습니다.');
    assertPublicModuleAccess_(site, 'meal');

    const mealRuntimeSlots = getMealRuntimeSlots_(site);
    const currentMeal = getCurrentMealTypeFromRuntimeSlots_(mealRuntimeSlots);
    // 브라우저가 임의로 forceRefresh=true를 보내도 원본 Sheet를 매번 다시 읽지 않습니다.
    const people = getParticipantsForLookup_(normalizedSiteId, site.linkedSheetId, false);
    const matches = people.filter(function (person) {
      return String(person.baseNameKey || '') === String(enteredIdentity.baseNameKey) &&
        String(person.phoneLast4 || '') === normalizedPhoneLast4;
    });

    if (!matches.length) {
      recordMealLookupFailure_(normalizedSiteId);
      throw new Error('입력하신 정보와 일치하는 교육생을 찾을 수 없습니다.');
    }
    if (matches.length > 1) {
      throw new Error('같은 이름과 번호가 2명 이상 확인됩니다. 교육 담당자에게 문의해주세요.');
    }

    const me = matches[0];
    const visibleSlots = mealRuntimeSlots.filter(function (slot) {
      return slot.status === 'current' || slot.status === 'past';
    });
    const seatKeys = visibleSlots.map(function (slot) { return slot.seatKey; });
    const included = {};
    included[me.personKey] = me;

    visibleSlots.forEach(function (slot) {
      const myTable = String(getParticipantMealSeatValue_(me, slot.seatKey) || '').trim();
      if (!myTable) return;
      people.forEach(function (person) {
        if (String(getParticipantMealSeatValue_(person, slot.seatKey) || '').trim() === myTable) {
          included[person.personKey] = person;
        }
      });
    });

    const relatedParticipants = Object.keys(included).map(function (personKey, index) {
      const person = included[personKey];
      return sanitizeMealParticipantForUser_(person, seatKeys, person.personKey === me.personKey, 'visible-' + String(index + 1));
    });

    return {
      success: true,
      site: toPublicSiteInfo_(site),
      participants: relatedParticipants,
      currentMeal: currentMeal,
      mealRuntimeSlots: mealRuntimeSlots,
      checkedAtText: Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm:ss')
    };
  } catch (error) {
    return {
      success: false,
      message: error.message,
      site: null,
      participants: [],
      currentMeal: null,
      mealRuntimeSlots: [],
      checkedAtText: ''
    };
  }
}

/****************************************************
 * 7. 내부 공통 함수 (1)
 * 관리 리소스 등록 · 공유 정책 · 폴더 생성/업로드 · 폴더 파일 목록
 ****************************************************/

function getResourceHeaders_() {
  return ['리소스ID', 'siteId', '종류', '소유구분', '상위ID', '상태', '생성일시', '수정일시', '메타JSON'];
}

function getResourceSheet_() {
  return getOrCreatePortalSheet_(SHEET_ADMIN_RESOURCES, getResourceHeaders_());
}

function registerManagedResource_(siteId, resourceId, kind, parentId, metadata) {
  const normalizedSiteId = String(siteId || '').trim();
  const normalizedResourceId = String(resourceId || '').trim();
  if (!normalizedSiteId || !normalizedResourceId) return;
  const sheet = getResourceSheet_();
  const rowIndex = findPortalRowByValue_(sheet, 1, normalizedResourceId);
  const now = new Date();
  const row = safeSheetRow_([
    normalizedResourceId,
    normalizedSiteId,
    String(kind || 'file'),
    RESOURCE_OWNERSHIP_MANAGED,
    String(parentId || ''),
    'active',
    rowIndex ? sheet.getRange(rowIndex, 7).getValue() || now : now,
    now,
    JSON.stringify(metadata || {}).slice(0, 4000)
  ]);
  if (rowIndex) sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);
}

function getCourseResourceRecords_(siteId) {
  const sheet = getResourceSheet_();
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, getResourceHeaders_().length)
    .getValues()
    .filter(function (row) { return String(row[1] || '') === String(siteId || ''); })
    .map(function (row) {
      return {
        resourceId: String(row[0] || ''),
        siteId: String(row[1] || ''),
        kind: String(row[2] || ''),
        ownership: String(row[3] || RESOURCE_OWNERSHIP_LEGACY),
        parentId: String(row[4] || ''),
        state: String(row[5] || ''),
        rowUpdatedAt: row[7]
      };
    });
}

function trashManagedResource_(resource) {
  const result = {
    resourceId: String(resource && resource.resourceId || ''),
    kind: String(resource && resource.kind || ''),
    ownership: String(resource && resource.ownership || RESOURCE_OWNERSHIP_LEGACY),
    action: 'detached',
    success: true,
    message: ''
  };
  if (result.ownership !== RESOURCE_OWNERSHIP_MANAGED) {
    result.message = '시스템 소유 자료가 아니므로 Drive에서 삭제하지 않았습니다.';
    return result;
  }
  try {
    if (result.kind === 'materialFolder' || result.kind === 'photoFolder') {
      DriveApp.getFolderById(result.resourceId).setTrashed(true);
    } else {
      if (result.kind === 'surveyForm') {
        try { FormApp.openById(result.resourceId).setAcceptingResponses(false); } catch (formError) {}
      }
      DriveApp.getFileById(result.resourceId).setTrashed(true);
    }
    result.action = 'moved_to_drive_trash';
    result.message = 'Google Drive 휴지통으로 이동했습니다.';
  } catch (error) {
    result.success = false;
    result.action = 'failed';
    result.message = error.message;
  }
  return result;
}

function applyManagedFolderSharingPolicy_(siteId, folder, serviceType) {
  if (!folder) return { success: false, message: '폴더가 없습니다.' };
  try {
    const resource = getCourseResourceRecords_(siteId).find(function (item) { return item.resourceId === folder.getId(); });
    if (!resource || resource.ownership !== RESOURCE_OWNERSHIP_MANAGED) return { success: false, protected: true, message: '기존 연결 폴더의 소유 확인이 필요합니다. Google Drive에서 공유 권한을 직접 확인해주세요.' };
    const site = getSiteInfoById_(siteId);
    const config = getPublicCourseAccessConfig_(siteId);
    const moduleKey = serviceType === 'photo' ? 'photos' : 'materials';
    const active = site && !isSiteTrashed_(site) &&
      ['비공개', '접근차단', '운영종료'].indexOf(String(site.status || '')) === -1;
    const shouldPublish = active && (!config.configured || (config.siteAvailable && config.modules[moduleKey]));
    if (shouldPublish) {
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return { success: true, public: true };
    }
    try { folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); }
    catch (privateError) {
      return { success: false, public: false, message: privateError.message };
    }
    return { success: true, public: false };
  } catch (error) {
    appendAuditLog_('RESOURCE_SHARING_POLICY_FAILED', String(siteId || ''), {
      serviceType: serviceType,
      message: error.message
    });
    return { success: false, message: error.message };
  }
}

function summarizeSharingPolicyWarnings_(results) {
  return (results || []).filter(function (item) { return item && item.success === false; }).map(function (item) { return item.message || '일부 자료의 공유 권한을 적용하지 못했습니다.'; });
}

function applyCourseResourceSharingPolicy_(siteId) {
  const site = getSiteInfoById_(siteId);
  if (!site) return [];
  const results = [];
  const config = getPublicCourseAccessConfig_(siteId);
  if (site.materialFolderId) {
    try { results.push(applyManagedFolderSharingPolicy_(siteId, DriveApp.getFolderById(site.materialFolderId), 'material')); }
    catch (error) { results.push({ success: false, type: 'material', message: error.message }); }
  }
  if (site.photoFolderId) {
    try { results.push(applyManagedFolderSharingPolicy_(siteId, DriveApp.getFolderById(site.photoFolderId), 'photo')); }
    catch (error) { results.push({ success: false, type: 'photo', message: error.message }); }
  }
  getCourseResourceRecords_(siteId)
    .filter(function (resource) { return resource.kind === 'surveyForm' && resource.ownership === RESOURCE_OWNERSHIP_MANAGED; })
    .forEach(function (resource) {
      try {
        const accepting = Boolean(config.siteAvailable && config.modules.survey);
        FormApp.openById(resource.resourceId).setAcceptingResponses(accepting);
        results.push({ success: true, type: 'survey', acceptingResponses: accepting });
      } catch (error) {
        results.push({ success: false, type: 'survey', message: error.message });
      }
    });
  return results;
}

function createServiceFolderAndUploadFiles_(options) {
  try {
    if (!options || !options.siteId) throw new Error('연결할 사용자 사이트가 없습니다.');

    const serviceType = options.serviceType;
    if (serviceType !== 'material' && serviceType !== 'photo') {
      throw new Error('올바르지 않은 파일 유형입니다.');
    }

    const ss = getSpreadsheet_();
    const sheet = getOrCreateSitesSheet_(ss);
    const rowIndex = findSiteRowIndex_(sheet, options.siteId);

    if (!rowIndex) throw new Error('해당 교육 사이트를 찾을 수 없습니다.');
    const site = rowToSiteInfo_(sheet.getRange(rowIndex, 1, 1, SITE_HEADERS.length).getValues()[0]);

    const educationTitle = site.isNonVup
      ? String(site.title || site.generation || '교육').trim()
      : `제${normalizeGeneration_(site.generation)}기 벤처투자 전문인력 양성교육(V-Up)`;

    // 이미 이 사이트에 연결된 폴더가 있으면 정확한 ID로 재사용합니다.
    const existingFolderId = serviceType === 'material' ? site.materialFolderId : site.photoFolderId;
    let folder = null;
    let folderCreatedBySystem = false;
    let folderNumber = Number(site.folderNumber || 0) || 1;
    if (existingFolderId) {
      try { folder = DriveApp.getFolderById(existingFolderId); } catch (e) { folder = null; }
    }
    if (!folder) {
      const parentFolder = !options.targetFolderId || options.targetFolderId === 'root'
        ? DriveApp.getRootFolder()
        : DriveApp.getFolderById(options.targetFolderId);
      folderNumber = getOrAssignFolderNumber_(sheet, site, parentFolder, educationTitle);
      const siteKey = String(options.siteId || '').slice(0, 12);
      const folderName = serviceType === 'material'
        ? '[VUP-' + siteKey + '] 교재 [' + sanitizeFileName_(educationTitle) + ' · ' + folderNumber + '번]'
        : '[VUP-' + siteKey + '] 현장사진 [' + sanitizeFileName_(educationTitle) + ' · ' + folderNumber + '번]';
      folder = withScriptLock_(function () {
        const matches = parentFolder.getFoldersByName(folderName);
        const ownedIds = getCourseResourceRecords_(options.siteId).filter(function (item) { return item.ownership === RESOURCE_OWNERSHIP_MANAGED; }).map(function (item) { return item.resourceId; });
        while (matches.hasNext()) { const candidate = matches.next(); if (ownedIds.indexOf(candidate.getId()) !== -1) return candidate; }
        folderCreatedBySystem = true;
        return parentFolder.createFolder(folderName);
      }, 15000);
    }
    if (folderCreatedBySystem) {
      registerManagedResource_(
        options.siteId,
        folder.getId(),
        serviceType === 'material' ? 'materialFolder' : 'photoFolder',
        options.targetFolderId || 'root',
        { folderNumber: folderNumber, title: educationTitle }
      );
    }
    const sharing = applyManagedFolderSharingPolicy_(options.siteId, folder, serviceType);
    if (!sharing.success) throw new Error(sharing.message || '자료 폴더의 공개 권한을 적용하지 못했습니다.');

    const uploadedFiles = uploadBase64FilesToFolder_(folder, options.files || [], serviceType);

    const folderId = folder.getId();
    const folderUrl = folder.getUrl();
    const now = new Date();

    updateSiteRowById_(sheet, options.siteId, function (row) {
      if (serviceType === 'material') {
        row[COL.materialFolderId - 1] = folderId;
        row[COL.materialFolderUrl - 1] = folderUrl;
      } else {
        row[COL.photoFolderId - 1] = folderId;
        row[COL.photoFolderUrl - 1] = folderUrl;
      }
      row[COL.folderNumber - 1] = folderNumber;
      row[COL.updatedAt - 1] = now;
      return row;
    });

    appendAuditLog_(serviceType === 'material' ? 'MATERIAL_FOLDER_CONNECTED' : 'PHOTO_FOLDER_CONNECTED', options.siteId, { folderNumber: folderNumber });

    return {
      success: true,
      folderId,
      folderUrl,
      uploadedCount: uploadedFiles.length,
      uploadedFiles,
      folderNumber: folderNumber,
      duplicateNotice: folderNumber > 1
        ? '이미 동일 교육명으로 생성된 폴더가 존재하여, 이번에는 폴더 맨 끝 번호를 "' + folderNumber + '번"으로 했어요. 이전에 생성한 사이트를 교육에 이용하지 않을 예정이라면, 앞 번호 폴더는 삭제해주세요'
        : '',
      currentFiles: serviceType === 'material'
        ? listFolderFiles_(folderId, getAllowedMaterialMimeTypes_(), getAllowedMaterialExtensions_())
        : listFolderFilesPage_(folderId, ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'image/heic', 'image/heif'], ['jpg', 'jpeg', 'png', 'gif', 'webp', 'heic', 'heif'], 0, 48).files,
      userUrl: site.userUrl,
      linkedSheetUrl: site.linkedSheetUrl,
      message: serviceType === 'material'
        ? '교재 파일이 사용자 사이트에 전달되었습니다.'
        : '현장 사진이 사용자 사이트에 전달되었습니다.'
    };

  } catch (error) {
    return { success: false, message: error.message };
  }
}

function uploadBase64FilesToFolder_(folder, files, serviceType) {
  const prepared = prepareBase64Uploads_(files, serviceType);
  const createdFiles = [];
  try {
    return prepared.map(function (item) {
      const file = folder.createFile(item.blob);
      createdFiles.push(file);
      return fileToPublicInfo_(file);
    });
  } catch (error) {
    let rollbackFailed = false;
    createdFiles.forEach(function (file) { try { file.setTrashed(true); } catch (rollbackError) { rollbackFailed = true; } });
    throw new Error('파일 업로드에 실패했습니다. ' + (rollbackFailed ? '일부 파일을 되돌리지 못했습니다. 연결된 폴더에서 파일을 확인해주세요.' : '이번 요청에서 생성한 파일을 휴지통으로 되돌렸습니다.') + ' ' + error.message);
  }
}

/** 전체 파일 검증을 먼저 완료하므로 마지막 파일이 잘못되어도 앞 파일이 남지 않습니다. */
function prepareBase64Uploads_(files, serviceType) {
  if (!Array.isArray(files)) throw new Error('업로드 파일 목록이 올바르지 않습니다.');
  if (serviceType !== 'material' && serviceType !== 'photo') throw new Error('파일 종류가 올바르지 않습니다.');
  if (files.length > MAX_BASE64_UPLOAD_FILES) throw new Error('한 번에 전달할 수 있는 파일은 최대 10개입니다.');
  let totalBytes = 0;
  const checked = files.map(function (file, index) {
    if (!file || typeof file.base64 !== 'string' || !file.base64) throw new Error((index + 1) + '번째 파일의 내용이 없습니다.');
    if (file.base64.length > Math.ceil(MAX_BASE64_UPLOAD_FILE_BYTES / 3) * 4 + 4096) throw new Error('파일당 최대 8MB까지 직접 전달할 수 있습니다.');
    const base64 = file.base64.replace(/\s/g, '');
    if (!base64 || base64.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(base64)) throw new Error('파일 데이터가 손상되었습니다. 원본 파일을 다시 선택해주세요.');
    const padding = base64.endsWith('==') ? 2 : base64.endsWith('=') ? 1 : 0;
    const byteLength = base64.length / 4 * 3 - padding;
    if (byteLength > MAX_BASE64_UPLOAD_FILE_BYTES) throw new Error('파일당 최대 8MB까지 직접 전달할 수 있습니다.');
    totalBytes += byteLength;
    if (totalBytes > MAX_BASE64_UPLOAD_TOTAL_BYTES) throw new Error('한 번에 직접 전달할 수 있는 총 용량은 25MB입니다.');
    const fileName = sanitizeFileName_(file.fileName || (serviceType === 'material' ? '교재_' + (index + 1) + '.pdf' : '현장사진_' + (index + 1)));
    const mimeType = String(file.mimeType || '').trim().toLowerCase() || guessMimeTypeFromName_(fileName, serviceType);
    validateFileType_(fileName, mimeType, serviceType);
    if (mimeType !== guessMimeTypeFromName_(fileName, serviceType)) throw new Error('파일 확장자와 MIME 유형이 일치하지 않습니다.');
    return { base64: base64, byteLength: byteLength, fileName: fileName, mimeType: mimeType };
  });
  return checked.map(function (item) {
    const bytes = Utilities.base64Decode(item.base64);
    if (bytes.length !== item.byteLength) throw new Error('파일 데이터 길이가 일치하지 않습니다.');
    validateFileSignature_(item.fileName, bytes, serviceType);
    return { blob: Utilities.newBlob(bytes, item.mimeType, item.fileName) };
  });
}

function listFolderFiles_(folderId, allowedMimeTypes, allowedExtensions, strictAccess) {
  if (!folderId) return [];

  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    if (strictAccess) throw new Error('Google Drive 폴더에 접근할 수 없습니다. 폴더가 삭제되었거나 권한이 변경되었는지 확인해주세요.');
    return [];
  }

  const files = [];
  const iterator = folder.getFiles();

  while (iterator.hasNext()) {
    const file = iterator.next();
    const name = file.getName();
    const mimeType = file.getMimeType();
    const extension = getExtension_(name);

    const mimeOk = !allowedMimeTypes || allowedMimeTypes.indexOf(mimeType) !== -1;
    const extOk = !allowedExtensions || allowedExtensions.indexOf(extension) !== -1 || mimeType === 'application/vnd.google-apps.presentation';

    if (mimeOk && extOk) {
      files.push(fileToPublicInfo_(file));
    }
  }

  files.sort((a, b) => {
    const timeCompare = Number(b.updatedAtMs || 0) - Number(a.updatedAtMs || 0);
    if (timeCompare !== 0) return timeCompare;
    return a.name.localeCompare(b.name, 'ko');
  });

  return files;
}

function listFolderFilesPage_(folderId, allowedMimeTypes, allowedExtensions, offset, pageSize) {
  if (!folderId) return { files: [], nextPageToken: '', hasMore: false };
  let folder;
  try { folder = DriveApp.getFolderById(folderId); }
  catch (e) { throw new Error('Google Drive 폴더에 접근할 수 없습니다. 폴더가 삭제되었거나 권한이 변경되었는지 확인해주세요.'); }
  const start = Number(offset == null || offset === '' ? 0 : offset);
  if (!Number.isSafeInteger(start) || start < 0 || start > 10000) throw new Error('페이지 번호가 올바르지 않습니다. 처음부터 다시 불러와주세요.');
  const limit = Math.max(1, Math.min(100, Number(pageSize || 48)));
  const files = [];
  let matchedIndex = 0;
  let hasMore = false;
  const iterator = folder.getFiles();
  let inspected = 0;
  while (iterator.hasNext()) {
    if (++inspected > 20000) throw new Error('폴더에 파일이 너무 많습니다. 교육별로 폴더를 나누어주세요.');
    const file = iterator.next();
    const mimeType = file.getMimeType();
    const extension = getExtension_(file.getName());
    const mimeOk = !allowedMimeTypes || allowedMimeTypes.indexOf(mimeType) !== -1;
    const extOk = !allowedExtensions || allowedExtensions.indexOf(extension) !== -1 || mimeType === 'application/vnd.google-apps.presentation';
    if (!mimeOk || !extOk) continue;
    if (matchedIndex < start) { matchedIndex++; continue; }
    if (files.length >= limit) { hasMore = true; break; }
    files.push(fileToPublicInfo_(file));
    matchedIndex++;
  }
  return { files: files, nextPageToken: hasMore ? String(start + files.length) : '', hasMore: hasMore };
}

function fileToPublicInfo_(file) {
  const fileId = file.getId();
  const name = file.getName();
  const mimeType = file.getMimeType();
  const size = file.getSize();
  const updatedDate = file.getLastUpdated();
  let resourceKey = '';
  try { resourceKey = String(file.getResourceKey() || ''); } catch (e) {}
  const resourceParam = resourceKey ? '&resourcekey=' + encodeURIComponent(resourceKey) : '';
  const isGoogleSlides = mimeType === 'application/vnd.google-apps.presentation';
  const directViewUrl = isGoogleSlides
    ? `https://docs.google.com/presentation/d/${fileId}/preview`
    : `https://drive.google.com/uc?export=view&id=${fileId}${resourceParam}`;
  const directDownloadUrl = isGoogleSlides
    ? `https://docs.google.com/presentation/d/${fileId}/export/pptx`
    : `https://drive.usercontent.google.com/download?id=${fileId}&export=download&confirm=t${resourceParam}`;
  return {
    id: fileId,
    name: name,
    mimeType: mimeType,
    size: size,
    sizeText: formatFileSize_(size),
    viewUrl: directViewUrl,
    downloadUrl: directDownloadUrl,
    thumbnailUrl: `https://drive.google.com/thumbnail?id=${fileId}&sz=w600${resourceParam}`,
    updatedAt: formatDate_(updatedDate),
    updatedAtMs: updatedDate.getTime()
  };
}

function getOrCreateChildFolder_(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(folderName);
}

function getOrAssignFolderNumber_(sheet, site, parentFolder, educationTitle) {
  if (Number(site.folderNumber || 0) > 0) return Number(site.folderNumber);
  let maxNumber = 0;
  if (sheet && typeof sheet.getLastRow === 'function' && sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, SITE_HEADERS.length).getValues().forEach(function (row) {
      const candidate = rowToSiteInfo_(row);
      if (candidate.siteId !== site.siteId && candidate.title === site.title) maxNumber = Math.max(maxNumber, Number(candidate.folderNumber || 0));
    });
  }
  try {
    const folders = parentFolder.getFolders();
    while (folders.hasNext()) {
      const name = folders.next().getName();
      if (name.indexOf('[' + sanitizeFileName_(educationTitle) + ' · ') === -1) continue;
      const match = name.match(/·\s*(\d+)번\]$/);
      if (match) maxNumber = Math.max(maxNumber, Number(match[1]));
    }
  } catch (e) {}
  return maxNumber + 1;
}
