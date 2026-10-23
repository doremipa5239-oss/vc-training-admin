/****************************************************
 * [5_Portal.gs]
 * 관리자 포털 API · 계정/세션 · 프로필·강의 설정·설문·문의
 ****************************************************/

/****************************************************
 * 9. V-Up 관리자 포털 API
 * --------------------------------------------------
 * 브라우저에서는 이 섹션의 portal* 함수만 호출합니다.
 * 기존 관리자 함수는 requireAdmin_()가 실행 컨텍스트를 확인하므로
 * 개발자 도구에서 직접 호출해도 인증 없이는 실행되지 않습니다.
 ****************************************************/

function portalLogin(credentials) {
  const input = credentials || {};
  const adminId = normalizeAdminId_(input.id);
  const password = String(input.password || '');
  if (!adminId || !password) throw new Error('아이디와 비밀번호를 입력해주세요.');
  if (password.length > 256) throw new Error('비밀번호는 256자 이내로 입력해주세요.');
  const rateKey = reserveAdminLoginAttempt_(adminId);
  const registry = synchronizeAdminAccountsFromCode_(true);
  const account = Object.prototype.hasOwnProperty.call(registry.accounts, adminId) ? registry.accounts[adminId] : null;
  if (!account || !account.active || !verifyAdminPassword_(account, password)) {
    Utilities.sleep(250);
    throw new Error('아이디 또는 비밀번호가 올바르지 않습니다.');
  }
  withScriptLock_(function () { PropertiesService.getScriptProperties().deleteProperty(rateKey); });
  const session = createAdminSession_(account);
  const response = buildPortalBootstrap_({ id: account.id, role: account.role, displayName: account.displayName, mustChangePassword: account.mustChangePassword });
  response.token = session.token;
  response.expiresAt = session.expiresAt;
  response.mustReplaceInitialPassword = Boolean(account.mustChangePassword);
  return response;
}

/** 캐시 조기 퇴거/동시 요청에도 실패 횟수가 초기화되지 않도록 원자적으로 횟수를 예약합니다. */
function reserveAdminLoginAttempt_(adminId) {
  return withScriptLock_(function () {
    const properties = PropertiesService.getScriptProperties();
    const seed = getAccountSeed_(adminId);
    const bucket = seed.id ? adminId : 'unknown';
    const key = 'VUP_LOGIN_WINDOW_V2_' + hashText_(bucket).slice(0, 32);
    let window = {};
    try { window = JSON.parse(properties.getProperty(key) || '{}'); } catch (e) {}
    const now = Date.now();
    if (!Number.isFinite(Number(window.expiresAt)) || Number(window.expiresAt) <= now) window = { count: 0, expiresAt: now + ADMIN_LOGIN_LOCK_SECONDS * 1000 };
    if (Number(window.count || 0) >= ADMIN_MAX_LOGIN_FAILURES) throw new Error('로그인 시도가 너무 많습니다. 10분 후 다시 시도해주세요.');
    window.count = Number(window.count || 0) + 1;
    properties.setProperty(key, JSON.stringify(window));
    return key;
  }, 10000);
}

function portalChangePassword(token, payload) {
  return withAdminSession_(token, function (currentAdmin) {
    return withScriptLock_(function () {
      const input = payload || {};
      const currentPassword = String(input.currentPassword || '');
      const newPassword = String(input.newPassword || '');
      if (currentPassword.length > 256 || newPassword.length > 256) throw new Error('비밀번호는 256자 이내로 입력해주세요.');
      if (newPassword.length < 7 || /\s/.test(newPassword) || !/[A-Za-z]/.test(newPassword) || !/\d/.test(newPassword)) {
        throw new Error('새 비밀번호는 7자 이상이며 영문과 숫자를 모두 포함해야 합니다.');
      }
      const registry = synchronizeAdminAccountsFromCode_(true);
      const account = registry.accounts[currentAdmin.id];
      if (!account || !verifyAdminPassword_(account, currentPassword)) {
        throw new Error('현재 비밀번호가 올바르지 않습니다.');
      }
      const salt = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
      account.salt = salt;
      account.passwordHash = hashAdminPasswordV2_(account.id, newPassword, salt);
      account.hashScheme = 'v2';
      account.passwordVersion = Math.max(1, Number(account.passwordVersion || 1)) + 1;
      account.mustChangePassword = false;
      registry.syncedAt = new Date().toISOString();
      const properties = PropertiesService.getScriptProperties();
      properties.setProperty(ADMIN_ACCOUNT_REGISTRY_KEY, JSON.stringify(registry));
      try {
        CacheService.getScriptCache().put(ADMIN_ACCOUNT_REGISTRY_CACHE_KEY, JSON.stringify(registry), 300);
      } catch (cacheError) {}
      portalLogout(token);
      const session = createAdminSession_(account);
      appendAuditLog_('ADMIN_PASSWORD_CHANGED', '', { adminId: currentAdmin.id });
      return {
        success: true,
        token: session.token,
        expiresAt: session.expiresAt,
        message: '비밀번호를 변경했습니다.'
      };
    }, 15000);
  }, { allowPasswordChange: true });
}

function portalRestoreSession(token) {
  return withAdminSession_(token, function (currentAdmin) {
    return buildPortalBootstrap_(currentAdmin);
  }, { allowPasswordChange: true });
}

/** '다음에 변경하기': 이번 로그인에서만 초기 비밀번호 변경 요구를 건너뜁니다. 다음 로그인 때 다시 안내합니다. */
function portalDeferPasswordChange(token) {
  return withAdminSession_(token, function () {
    return withScriptLock_(function () {
      const properties = PropertiesService.getScriptProperties();
      const key = ADMIN_SESSION_PREFIX + hashPortalToken_(String(token || ''));
      let session = null;
      try { session = JSON.parse(properties.getProperty(key) || 'null'); } catch (e) { session = null; }
      if (!session) throw new Error('로그인이 만료되었습니다. 다시 로그인해주세요.');
      session.passwordChangeDeferred = true;
      properties.setProperty(key, JSON.stringify(session));
      return { success: true, message: '다음 로그인 때 다시 안내해드릴게요.' };
    }, 15000);
  }, { allowPasswordChange: true });
}

function portalLogout(token) {
  return withScriptLock_(function () {
    const normalizedToken = String(token || '');
    if (/^[a-f0-9]{64,160}$/i.test(normalizedToken)) {
      const sessionKey = ADMIN_SESSION_PREFIX + hashPortalToken_(normalizedToken);
      PropertiesService.getScriptProperties().deleteProperty(sessionKey);
    }
    return { success: true };
  }, 15000);
}

function portalGetDashboard(token) {
  return withAdminSession_(token, function (currentAdmin) {
    return buildPortalDashboard_(currentAdmin);
  });
}

function portalGetServerTime(token) {
  return withAdminSession_(token, function () {
    return {
      success: true,
      epochMs: Date.now(),
      timezone: TIMEZONE
    };
  });
}

function portalSaveProfile(token, profile) {
  return withAdminSession_(token, function (currentAdmin) {
    const saved = saveAdminProfile_(currentAdmin.id, profile || {});
    const ownerMap = getCourseOwnerMap_();
    Object.keys(ownerMap).forEach(function (siteId) {
      if (String(ownerMap[siteId] || '') === String(currentAdmin.id || '')) {
        clearUserDataCache_(siteId);
      }
    });
    appendAuditLog_('ADMIN_PROFILE_UPDATED', '', {
      email: saved.email,
      organization: saved.organization
    });
    return { success: true, profile: saved, message: '관리자 기본 정보를 저장했습니다.' };
  });
}

function portalListCourses(token, options) {
  return withAdminSession_(token, function (currentAdmin) {
    const input = options || {};
    const courses = getPortalCourses_(currentAdmin);
    const filter = String(input.filter || 'all');
    const scope = String(input.scope || 'operations');
    const search = String(input.search || '').trim().toLowerCase();
    const filtered = courses.filter(function (course) {
      if (scope === 'operations' && ['past', 'blocked'].indexOf(course.phase) !== -1 && filter !== 'blocked') return false;
      if (scope === 'history' && ['past', 'blocked'].indexOf(course.phase) === -1) return false;
      if (filter === 'manage' && ['upcoming', 'active'].indexOf(course.phase) === -1) return false;
      if (filter === 'history' && ['past', 'blocked'].indexOf(course.phase) === -1) return false;
      if (filter === 'blocked' && course.phase !== 'blocked') return false;
      if (!search) return true;
      return [
        course.title,
        course.generation,
        course.managerName,
        course.siteId
      ].join(' ').toLowerCase().indexOf(search) !== -1;
    });
    return { success: true, courses: filtered.map(toPortalCourseSummary_) };
  });
}

function toPortalCourseSummary_(course) {
  return {
    siteId: course.siteId,
    generation: course.generation,
    title: course.title,
    managerName: course.managerName,
    ownerId: course.ownerId,
    phase: course.phase,
    phaseLabel: course.phaseLabel,
    dateSummary: course.dateSummary,
    readinessPercent: course.readinessPercent,
    createdAt: course.createdAt,
    updatedAt: course.updatedAt,
    revision: course.revision
  };
}

function portalListTrash(token) {
  return withAdminSession_(token, function (currentAdmin) {
    const result = getTrashedEducationSites();
    if (!result || !result.success) {
      throw new Error(result && result.message ? result.message : '휴지통을 불러오지 못했습니다.');
    }
    const ownerMap = getCourseOwnerMap_();
    const settingsMap = getCourseSettingsMap_();
    const courses = (result.sites || [])
      .filter(function (site) {
        return currentAdmin.role === 'developer' || ownerMap[site.siteId] === currentAdmin.id;
      })
      .map(function (site) {
        const course = decoratePortalCourse_(
          site,
          settingsMap[site.siteId] || getDefaultCourseSettings_(site.siteId, ownerMap[site.siteId], true),
          ownerMap[site.siteId] || ''
        );
        course.phase = 'trash';
        course.phaseLabel = '휴지통';
        return course;
      });
    return { success: true, courses: courses };
  });
}

function portalMoveCoursesToTrash(token, siteIds) {
  return withAdminSession_(token, function (currentAdmin) {
    const ids = normalizePortalCourseIds_(siteIds);
    updatePortalCourseTrashState_(currentAdmin, ids, true);
    return {
      success: true,
      count: ids.length,
      message: ids.length + '개 강의를 휴지통으로 이동했습니다. 30일 안에는 복구할 수 있습니다.'
    };
  });
}

function portalRestoreCourses(token, siteIds) {
  return withAdminSession_(token, function (currentAdmin) {
    const ids = normalizePortalCourseIds_(siteIds);
    updatePortalCourseTrashState_(currentAdmin, ids, false);
    return {
      success: true,
      count: ids.length,
      message: ids.length + '개 강의를 복구했습니다.'
    };
  });
}

function updatePortalCourseTrashState_(currentAdmin, siteIds, moveToTrash) {
  return withScriptLock_(function () {
    const ids = normalizePortalCourseIds_(siteIds);
    const selected = {};
    ids.forEach(function (siteId) { selected[siteId] = true; });

    if (ADMIN_REQUEST_CACHE) ADMIN_REQUEST_CACHE.ownerMap = null;
    const ownerMap = getCourseOwnerMap_();
    const sheet = getOrCreateSitesSheet_(getSpreadsheet_());
    const rowCount = Math.max(0, sheet.getLastRow() - 1);
    if (!rowCount) throw new Error('처리할 강의를 찾을 수 없습니다.');

    const values = sheet.getRange(2, 1, rowCount, SITE_HEADERS.length).getValues();
    const found = {};
    const updates = [];
    const now = new Date();
    const purgeAt = new Date(now.getTime() + SITE_TRASH_DAYS * 24 * 60 * 60 * 1000);

    values.forEach(function (row, index) {
      const siteId = String(row[COL.siteId - 1] || '');
      if (!selected[siteId]) return;
      if (
        currentAdmin.role !== 'developer' &&
        String(ownerMap[siteId] || '') !== String(currentAdmin.id || '')
      ) {
        throw new Error('관리 권한이 없는 강의가 포함되어 있습니다.');
      }
      if (!moveToTrash && !row[COL.deletedAt - 1]) {
        throw new Error('휴지통에 있는 강의만 복구할 수 있습니다.');
      }
      found[siteId] = true;
      if (moveToTrash) row[COL.preTrashStatus - 1] = String(row[COL.status - 1] || '사용중');
      row[COL.deletedAt - 1] = moveToTrash ? now : '';
      row[COL.purgeAt - 1] = moveToTrash ? purgeAt : '';
      row[COL.status - 1] = moveToTrash
        ? '휴지통'
        : String(row[COL.preTrashStatus - 1] || '사용중');
      if (!moveToTrash) row[COL.preTrashStatus - 1] = '';
      row[COL.updatedAt - 1] = now;
      row[COL.revision - 1] = Math.max(0, Number(row[COL.revision - 1] || 0)) + 1;
      updates.push({ rowIndex: index + 2, values: row });
    });

    const missing = ids.filter(function (siteId) { return !found[siteId]; });
    if (missing.length) throw new Error('처리할 강의를 찾을 수 없습니다: ' + missing.join(', '));

    updates.forEach(function (update) {
      sheet
        .getRange(update.rowIndex, 1, 1, SITE_HEADERS.length)
        .setValues([safeSheetRow_(update.values)]);
    });
    ids.forEach(function (siteId) {
      delete SITE_INFO_RUNTIME_CACHE[siteId];
      clearUserDataCache_(siteId);
      applyCourseResourceSharingPolicy_(siteId);
    });
    clearSiteListCache_();
    appendAuditLog_(
      moveToTrash ? 'PORTAL_COURSES_MOVED_TO_TRASH' : 'PORTAL_COURSES_RESTORED',
      '',
      {
        actorId: currentAdmin.id,
        siteIds: ids,
        purgeAt: moveToTrash ? formatDate_(purgeAt) : ''
      }
    );
  }, 15000);
}

function portalGetTrashImpact(token, siteIds) {
  return withAdminSession_(token, function (currentAdmin) {
    const ids = normalizePortalCourseIds_(siteIds);
    const items = ids.map(function (siteId) {
      assertCourseAccess_(currentAdmin, siteId);
      const site = getSiteInfoById_(siteId);
      const records = getCourseResourceRecords_(siteId);
      const registered = {};
      records.forEach(function (record) { registered[record.resourceId] = true; });
      const legacy = [site.materialFolderId, site.photoFolderId, site.linkedSheetId, site.adminLinkDocumentId]
        .filter(function (id) { return id && !registered[id]; });
      return {
        siteId: siteId,
        title: site.title,
        managedResourceCount: records.filter(function (r) { return r.ownership === RESOURCE_OWNERSHIP_MANAGED; }).length,
        protectedExternalOrLegacyCount: legacy.length + records.filter(function (r) { return r.ownership !== RESOURCE_OWNERSHIP_MANAGED; }).length,
        action: '관리 시트의 강의 기록은 삭제되고, 시스템 소유 자료만 Google Drive 휴지통으로 이동합니다.'
      };
    });
    return {
      success: true,
      items: items,
      confirmationPhrase: '영구 삭제',
      storageNotice: 'Drive 휴지통으로 이동한 자료는 Drive 휴지통을 비운 뒤에야 저장용량이 실제로 확보됩니다.'
    };
  });
}

function portalPermanentlyDeleteCourses(token, siteIds, confirmation) {
  return withAdminSession_(token, function (currentAdmin) {
    if (String(confirmation && confirmation.confirmation || '') !== '영구 삭제') {
      throw new Error('확인 문구 "영구 삭제"를 정확히 입력해주세요.');
    }
    const ids = normalizePortalCourseIds_(siteIds);
    ids.forEach(function (siteId) {
      assertCourseAccess_(currentAdmin, siteId);
    });
    const results = ids.map(function (siteId) {
      try {
        permanentlyDeleteEducationSite_(siteId, currentAdmin.id);
        return { siteId: siteId, success: true };
      } catch (error) {
        return { siteId: siteId, success: false, message: error.message };
      }
    });
    const succeeded = results.filter(function (item) { return item.success; }).length;
    return {
      success: succeeded === ids.length,
      count: succeeded,
      failedCount: ids.length - succeeded,
      results: results,
      message: succeeded + '개 강의 기록을 삭제했습니다. 시스템 소유 자료는 Drive 휴지통으로 이동했으며, 연결 자료는 보호했습니다.'
    };
  });
}

function normalizePortalCourseIds_(siteIds) {
  if (!Array.isArray(siteIds) || !siteIds.length) {
    throw new Error('선택한 강의가 없습니다.');
  }
  const result = [];
  const seen = {};
  siteIds.slice(0, 100).forEach(function (siteId) {
    const normalized = normalizePublicSiteId_(siteId);
    if (!seen[normalized]) {
      seen[normalized] = true;
      result.push(normalized);
    }
  });
  if (!result.length) throw new Error('선택한 강의가 없습니다.');
  return result;
}

function portalGetCourse(token, siteId) {
  return withAdminSession_(token, function (currentAdmin) {
    assertCourseAccess_(currentAdmin, siteId);
    const site = getSiteInfoById_(normalizePublicSiteId_(siteId));
    if (!site) throw new Error('강의를 찾을 수 없습니다.');
    const settings = getCourseSettings_(site.siteId);
    const course = decoratePortalCourse_(site, settings, getCourseOwnerId_(site.siteId));
    return {
      success: true,
      course: course,
      materials: [],
      photos: [],
      surveys: [],
      resourcesLoaded: false,
      health: buildLightweightCourseHealth_(site, settings, null, null)
    };
  });
}

function portalGetCourseResources(token, siteId, type, pageToken) {
  return withAdminSession_(token, function (currentAdmin) {
    const normalizedSiteId = normalizePublicSiteId_(siteId);
    assertCourseAccess_(currentAdmin, normalizedSiteId);
    const site = getSiteInfoById_(normalizedSiteId);
    const resourceType = String(type || '');
    if (resourceType === 'materials') {
      return {
        success: true,
        type: resourceType,
        items: site.materialFolderId
          ? listFolderFiles_(site.materialFolderId, getAllowedMaterialMimeTypes_(), getAllowedMaterialExtensions_(), true)
          : [],
        hasMore: false,
        nextPageToken: ''
      };
    }
    if (resourceType === 'photos') {
      const page = site.photoFolderId
        ? listFolderFilesPage_(site.photoFolderId,
            ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'image/heic', 'image/heif'],
            ['jpg', 'jpeg', 'png', 'gif', 'webp', 'heic', 'heif'],
            Math.max(0, Number(pageToken || 0)), 48)
        : { files: [], hasMore: false, nextPageToken: '' };
      return { success: true, type: resourceType, items: page.files, hasMore: page.hasMore, nextPageToken: page.nextPageToken };
    }
    if (resourceType === 'surveys') {
      return { success: true, type: resourceType, items: getCourseSurveys_(normalizedSiteId), hasMore: false, nextPageToken: '' };
    }
    throw new Error('불러올 자료 종류가 올바르지 않습니다.');
  });
}

function portalCreateCourse(token, payload) {
  return withAdminSession_(token, function (currentAdmin) {
    return withScriptLock_(function () {
      const input = payload || {};
      input.settings = Object.assign({}, input.settings || {});
      const requestedAccessCode = String(input.settings.accessCode || '').replace(/\D/g, '');
      const settingsMap = getCourseSettingsMap_();
      const requestedCodeInUse = Object.keys(settingsMap).some(function (siteId) {
        return String(settingsMap[siteId].accessCode || '') === requestedAccessCode;
      });
      input.settings.accessCode =
        /^\d{4}$/.test(requestedAccessCode) && !requestedCodeInUse
          ? requestedAccessCode
          : createUniqueCourseAccessCode_();
      validateCourseSettingsDraft_('', input.settings || {});
      const profile = getAdminProfile_(currentAdmin.id);
      const managerName = String(input.managerName || profile.name || currentAdmin.displayName || '').trim();
      const managerPhone = String(input.managerPhone || profile.landline || profile.mobile || '').trim();
      const data = Object.assign({}, input, {
        managerName: managerName,
        managerPhone: managerPhone,
        requestId: String(input.requestId || Utilities.getUuid())
      });

      const created = createEducationSite(data);
      if (!created || !created.success) {
        throw new Error(created && created.message ? created.message : '강의를 개설하지 못했습니다.');
      }

      if (created.reused) {
        const owner = getCourseOwnerId_(created.siteId);
        if (owner && owner !== currentAdmin.id) throw new Error('다른 관리자의 생성 요청은 재사용할 수 없습니다.');
        if (owner && getCourseSettingsMap_()[created.siteId]) return { success: true, reused: true, course: decoratePortalCourse_(getSiteInfoById_(created.siteId), getCourseSettings_(created.siteId), owner), message: '이미 개설된 강의를 불러왔습니다.' };
      }
      recordCourseOwner_(created.siteId, currentAdmin.id);
      const settings = saveCourseSettings_(created.siteId, currentAdmin.id, input.settings || {});
      appendAuditLog_('PORTAL_COURSE_CREATED', created.siteId, {
        ownerId: currentAdmin.id,
        accessCodeConfigured: Boolean(settings.accessCode)
      });
      return {
        success: true,
        course: decoratePortalCourse_(
          getSiteInfoById_(created.siteId),
          settings,
          currentAdmin.id
        ),
        message: '새 강의를 개설했습니다. 지금부터 운영 자료를 등록할 수 있습니다.'
      };
    }, 15000);
  });
}

function portalUpdateCourse(token, payload) {
  return withAdminSession_(token, function (currentAdmin) {
    const input = payload || {};
    const siteId = normalizePublicSiteId_(input.siteId);
    return withCourseMutation_(currentAdmin, siteId, input.revision, function () {
      if (input.settings) validateCourseSettingsDraft_(siteId, input.settings);
      const profile = getAdminProfile_(currentAdmin.id);
      const data = Object.assign({}, input, {
        managerName: String(input.managerName || profile.name || currentAdmin.displayName || '').trim(),
        managerPhone: String(input.managerPhone || profile.landline || profile.mobile || '').trim()
      });
      const updated = updateEducationSiteBasicInfo(data);
      if (!updated || !updated.success) {
        throw new Error(updated && updated.message ? updated.message : '강의 정보를 수정하지 못했습니다.');
      }
      if (input.settings) saveCourseSettings_(siteId, getCourseOwnerId_(siteId), input.settings);
      return {
        success: true,
        course: decoratePortalCourse_(
          getSiteInfoById_(siteId),
          getCourseSettings_(siteId),
          getCourseOwnerId_(siteId)
        ),
        message: updated.message
      };
    });
  });
}

function portalSaveCourseSettings(token, siteId, settings) {
  return withAdminSession_(token, function (currentAdmin) {
    const normalizedSiteId = normalizePublicSiteId_(siteId);
    return withCourseMutation_(currentAdmin, normalizedSiteId, settings && settings._revision, function () {
      const currentSettings = getCourseSettings_(normalizedSiteId);
      const safeSettings = Object.assign({}, settings || {}, {
        accessCode: currentSettings.accessCode
      });
      const saved = saveCourseSettings_(
        normalizedSiteId,
        getCourseOwnerId_(normalizedSiteId) || currentAdmin.id,
        safeSettings
      );
      updateSiteRowById_(getOrCreateSitesSheet_(getSpreadsheet_()), normalizedSiteId, function (row) {
        row[COL.updatedAt - 1] = new Date();
        return row;
      });
      const sharingWarnings = summarizeSharingPolicyWarnings_(applyCourseResourceSharingPolicy_(normalizedSiteId));
      clearUserDataCache_(normalizedSiteId);
      appendAuditLog_('COURSE_OPERATIONS_SETTINGS_UPDATED', normalizedSiteId, {
        progressText: saved.progressText,
        hasRememberUrl: Boolean(saved.rememberUrl)
      });
      return {
        success: true,
        warnings: sharingWarnings,
        settings: sanitizeCourseSettingsForAdmin_(saved),
        revision: getSiteInfoById_(normalizedSiteId).revision,
        publicConfig: getPublicCourseAccessConfig_(normalizedSiteId),
        message: '사용자 사이트 운영 설정을 저장했습니다.' + (sharingWarnings.length ? ' 주의: ' + sharingWarnings.join(' ') : '')
      };
    });
  });
}

function portalSetCourseVisibility(token, siteId, visible) {
  return withAdminSession_(token, function (currentAdmin) {
    const normalizedSiteId = normalizePublicSiteId_(siteId);
    return withCourseMutation_(currentAdmin, normalizedSiteId, null, function () {
      const sheet = getOrCreateSitesSheet_(getSpreadsheet_());
      updateSiteRowById_(sheet, normalizedSiteId, function (row) {
        row[COL.status - 1] = visible ? '사용중' : '접근차단';
        row[COL.updatedAt - 1] = new Date();
        return row;
      });
      clearUserDataCache_(normalizedSiteId);
      clearSiteListCache_();
      const sharingWarnings = summarizeSharingPolicyWarnings_(applyCourseResourceSharingPolicy_(normalizedSiteId));
      appendAuditLog_(visible ? 'COURSE_ACCESS_ENABLED' : 'COURSE_ACCESS_BLOCKED', normalizedSiteId, {});
      return {
        success: true,
        warnings: sharingWarnings,
        visible: Boolean(visible),
        message: (visible ? '교육생의 사이트 접근을 다시 허용했습니다.' : '교육생의 사이트 접근을 차단했습니다.') + (sharingWarnings.length ? ' 주의: ' + sharingWarnings.join(' ') : '')
      };
    });
  });
}

function portalUploadCourseFiles(token, payload) {
  return withAdminSession_(token, function (currentAdmin) {
    const input = payload || {};
    const siteId = normalizePublicSiteId_(input.siteId);
    return withCourseMutation_(currentAdmin, siteId, null, function () {
      const type = String(input.type || '');
      if (type !== 'material' && type !== 'photo') throw new Error('파일 종류가 올바르지 않습니다.');
      const files = Array.isArray(input.files) ? input.files : [];
      if (!files.length) throw new Error('업로드할 파일을 선택해주세요.');
      const result = createServiceFolderAndUploadFiles_({
        siteId: siteId,
        targetFolderId: input.targetFolderId || 'root',
        files: files,
        serviceType: type
      });
      if (!result.success) throw new Error(result.message || '파일 업로드에 실패했습니다.');
      clearUserDataCache_(siteId);
      clearSiteListCache_();
      appendAuditLog_(type === 'material' ? 'PORTAL_MATERIAL_UPLOADED' : 'PORTAL_PHOTO_UPLOADED', siteId, {
        count: result.uploadedCount
      });
      return result;
    });
  });
}

function portalDeleteCourseFile(token, payload) {
  return withAdminSession_(token, function (currentAdmin) {
    const input = payload || {};
    const siteId = normalizePublicSiteId_(input.siteId);
    return withCourseMutation_(currentAdmin, siteId, null, function () {
      const type = String(input.type || '');
      if (type !== 'material' && type !== 'photo') throw new Error('파일 종류가 올바르지 않습니다.');
      const site = getSiteInfoById_(siteId);
      const folderId = type === 'material' ? site.materialFolderId : site.photoFolderId;
      if (!folderId) throw new Error('연결된 파일 폴더가 없습니다.');
      const file = DriveApp.getFileById(String(input.fileId || ''));
      let belongsToCourse = false;
      const parents = file.getParents();
      while (parents.hasNext()) {
        if (parents.next().getId() === folderId) {
          belongsToCourse = true;
          break;
        }
      }
      if (!belongsToCourse) throw new Error('이 강의에 속한 파일만 삭제할 수 있습니다.');
      const fileName = file.getName();
      file.setTrashed(true);
      clearUserDataCache_(siteId);
      appendAuditLog_(type === 'material' ? 'PORTAL_MATERIAL_DELETED' : 'PORTAL_PHOTO_DELETED', siteId, {
        fileId: input.fileId,
        fileName: fileName
      });
      return { success: true, message: '파일을 휴지통으로 이동했습니다.' };
    });
  });
}


function portalCreateSurvey(token, payload) {
  return withAdminSession_(token, function (currentAdmin) {
    const input = payload || {};
    const siteId = normalizePublicSiteId_(input.siteId);
    return withCourseMutation_(currentAdmin, siteId, input.revision, function () {
      const site = getSiteInfoById_(siteId);
      const requestId = scopeAdminRequestId_(input.requestId, 'survey:' + siteId);
      const surveySheet = getOrCreatePortalSheet_(SHEET_ADMIN_SURVEYS, getSurveyHeaders_());
      if (requestId) {
        const existingRow = findPortalRowByValue_(surveySheet, 10, requestId);
        if (existingRow) {
          const row = surveySheet.getRange(existingRow, 1, 1, getSurveyHeaders_().length).getValues()[0];
          if (String(row[1]) !== siteId || String(row[2]) !== currentAdmin.id) throw new Error('설문 생성 요청의 관리 권한이 일치하지 않습니다.');
          return {
            success: true,
            reused: true,
            survey: { id: String(row[5] || ''), title: String(row[3] || ''), templateType: String(row[4] || ''), publicUrl: String(row[6] || ''), editUrl: String(row[7] || ''), createdAt: formatDate_(row[8]) },
            message: '이미 완료된 설문 생성 결과를 다시 불러왔습니다.'
          };
        }
      }
      const templateType = ['satisfaction', 'pre', 'custom'].indexOf(String(input.templateType || '')) !== -1
        ? String(input.templateType)
        : 'satisfaction';
      const title = String(input.title || '').trim().slice(0, 150) ||
        site.title + (templateType === 'pre' ? ' 사전 설문' : ' 만족도 조사');
      const form = FormApp.create(title);
      form.setDescription(String(input.description || '').trim().slice(0, 1500) ||
        '더 나은 교육 운영을 위해 의견을 남겨주세요.');
      form.setConfirmationMessage('응답해주셔서 감사합니다.');

      if (templateType === 'satisfaction') {
        form.addScaleItem()
          .setTitle('이번 교육은 전반적으로 만족스러우셨나요?')
          .setBounds(1, 5)
          .setLabels('매우 아쉬움', '매우 만족')
          .setRequired(true);
        form.addScaleItem()
          .setTitle('교육 내용이 실무에 도움이 되었나요?')
          .setBounds(1, 5)
          .setLabels('도움이 적음', '매우 도움됨')
          .setRequired(true);
        form.addParagraphTextItem().setTitle('가장 도움이 된 내용은 무엇인가요?');
        form.addParagraphTextItem().setTitle('개선이 필요한 점이나 추가 의견을 자유롭게 적어주세요.');
      } else if (templateType === 'pre') {
        form.addTextItem().setTitle('성함').setRequired(true);
        form.addTextItem().setTitle('소속 기관').setRequired(true);
        form.addTextItem().setTitle('직급 또는 담당 업무');
        form.addParagraphTextItem().setTitle('이번 교육에서 가장 기대하는 내용은 무엇인가요?');
        form.addMultipleChoiceItem()
          .setTitle('관련 업무 경험')
          .setChoiceValues(['1년 미만', '1~3년', '3~5년', '5년 이상'])
          .setRequired(true);
      } else {
        form.addParagraphTextItem().setTitle('의견을 자유롭게 남겨주세요.').setRequired(true);
      }

      const formFile = DriveApp.getFileById(form.getId());
      const surveyFolder = getOrCreateChildFolder_(DriveApp.getRootFolder(), 'V-Up 교육 설문');
      try { formFile.moveTo(surveyFolder); } catch (e) {}
      registerManagedResource_(siteId, form.getId(), 'surveyForm', surveyFolder.getId(), {
        title: title,
        templateType: templateType
      });
      const now = new Date();
      surveySheet.appendRow(safeSheetRow_([
        Utilities.getUuid(),
        siteId,
        currentAdmin.id,
        title,
        templateType,
        form.getId(),
        form.getPublishedUrl(),
        form.getEditUrl(),
        now,
        requestId
      ]));
      const settings = getCourseSettings_(siteId);
      settings.surveyUrl = form.getPublishedUrl();
      saveCourseSettings_(siteId, getCourseOwnerId_(siteId) || currentAdmin.id, settings);
      applyCourseResourceSharingPolicy_(siteId);
      clearUserDataCache_(siteId);
      appendAuditLog_('PORTAL_SURVEY_CREATED', siteId, {
        templateType: templateType,
        formId: form.getId()
      });
      return {
        success: true,
        survey: {
          id: form.getId(),
          title: title,
          templateType: templateType,
          publicUrl: form.getPublishedUrl(),
          editUrl: form.getEditUrl(),
          createdAt: formatDate_(now)
        },
        message: '설문지를 만들고 사용자 사이트 연결 설정에 반영했습니다.'
      };
    });
  });
}

function portalListInquiries(token) {
  return withAdminSession_(token, function (currentAdmin) {
    return {
      success: true,
      inquiries: getPortalInquiries_(currentAdmin)
    };
  });
}

function portalSubmitInquiry(token, payload) {
  return withAdminSession_(token, function (currentAdmin) {
    const input = payload || {};
    const subject = String(input.subject || '').trim().slice(0, 160);
    const body = String(input.body || '').trim().slice(0, 5000);
    if (!subject) throw new Error('문의 제목을 입력해주세요.');
    if (body.length < 5) throw new Error('문의 내용을 5자 이상 입력해주세요.');
    const profile = getAdminProfile_(currentAdmin.id);
    const now = new Date();
    const inquiryId = 'Q-' + Utilities.formatDate(now, TIMEZONE, 'yyyyMMdd') + '-' +
      Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase();
    getOrCreatePortalSheet_(SHEET_ADMIN_INQUIRIES, getInquiryHeaders_()).appendRow(safeSheetRow_([
      inquiryId,
      currentAdmin.id,
      profile.name || currentAdmin.displayName,
      subject,
      body,
      '접수',
      now,
      '',
      '',
      ''
    ]));
    appendAuditLog_('PORTAL_INQUIRY_SUBMITTED', '', { inquiryId: inquiryId, subject: subject });
    return { success: true, inquiryId: inquiryId, message: '문의가 접수되었습니다.' };
  });
}

function portalReplyInquiry(token, payload) {
  return withAdminSession_(token, function (currentAdmin) {
    requireDeveloperRole_(currentAdmin);
    const input = payload || {};
    const inquiryId = String(input.inquiryId || '').trim();
    const reply = String(input.reply || '').trim().slice(0, 5000);
    if (!inquiryId || !reply) throw new Error('답변 내용을 입력해주세요.');
    const sheet = getOrCreatePortalSheet_(SHEET_ADMIN_INQUIRIES, getInquiryHeaders_());
    const rowIndex = findPortalRowByValue_(sheet, 1, inquiryId);
    if (!rowIndex) throw new Error('문의글을 찾을 수 없습니다.');
    sheet.getRange(rowIndex, 6, 1, 5).setValues([safeSheetRow_([
      '답변완료',
      sheet.getRange(rowIndex, 7).getValue(),
      reply,
      new Date(),
      currentAdmin.id
    ])]);
    appendAuditLog_('PORTAL_INQUIRY_REPLIED', '', { inquiryId: inquiryId });
    return { success: true, message: '문의 답변을 저장했습니다.' };
  });
}

function portalAssignCourseOwner(token, siteId, adminId) {
  return withAdminSession_(token, function (currentAdmin) {
    requireDeveloperRole_(currentAdmin);
    const normalizedSiteId = normalizePublicSiteId_(siteId);
    return withCourseMutation_(currentAdmin, normalizedSiteId, null, function () {
      const targetAdminId = normalizeAdminId_(adminId);
      const registry = synchronizeAdminAccountsFromCode_();
      if (!registry.accounts[targetAdminId] || !registry.accounts[targetAdminId].active) {
        throw new Error('활성 관리자 계정을 찾을 수 없습니다.');
      }
      recordCourseOwner_(normalizedSiteId, targetAdminId);
      const settings = getCourseSettings_(normalizedSiteId);
      saveCourseSettings_(normalizedSiteId, targetAdminId, settings);
      appendAuditLog_('COURSE_OWNER_ASSIGNED', normalizedSiteId, { ownerId: targetAdminId });
      return { success: true, message: '강의 담당 관리자를 변경했습니다.' };
    });
  });
}

/****************************************************
 * 9-A. 계정 · 세션 · 권한
 ****************************************************/

function normalizeAdminId_(value) {
  const id = String(value || '').trim().toLowerCase();
  return /^[a-z0-9._-]{3,40}$/.test(id) ? id : '';
}

function getAccountSeed_(adminId) {
  const id = normalizeAdminId_(adminId);
  for (let i = 0; i < ADMIN_ACCOUNT_SEEDS.length; i++) {
    if (normalizeAdminId_(ADMIN_ACCOUNT_SEEDS[i].id) === id) return ADMIN_ACCOUNT_SEEDS[i];
  }
  return {};
}

function synchronizeAdminAccountsFromCode_(forceFresh) {
  return withScriptLock_(function () {
    const signatureSource = ADMIN_ACCOUNT_SEEDS.map(function (seed) {
      return [
        normalizeAdminId_(seed.id),
        Number(seed.passwordVersion || 1),
        String(seed.role || 'manager'),
        seed.active !== false,
        String(seed.displayName || ''),
        seed.forcePasswordChange !== false
      ].join('|');
    }).join('||');
    const signature = hashText_(signatureSource);
    const scriptCache = CacheService.getScriptCache();
    try {
      const cachedRegistry = JSON.parse(
        scriptCache.get(ADMIN_ACCOUNT_REGISTRY_CACHE_KEY) || '{}'
      );
      if (
        !forceFresh && cachedRegistry &&
        cachedRegistry.signature === signature &&
        cachedRegistry.accounts
      ) {
        return cachedRegistry;
      }
    } catch (cacheError) {}

    const properties = PropertiesService.getScriptProperties();
    let registry = {};
    try {
      registry = JSON.parse(properties.getProperty(ADMIN_ACCOUNT_REGISTRY_KEY) || '{}');
    } catch (e) {
      registry = {};
    }
    if (!registry.accounts || typeof registry.accounts !== 'object') registry.accounts = {};

    if (registry.signature === signature) {
      try {
        scriptCache.put(
          ADMIN_ACCOUNT_REGISTRY_CACHE_KEY,
          JSON.stringify(registry),
          300
        );
      } catch (cacheError) {}
      return registry;
    }

    const nextAccounts = {};
    ADMIN_ACCOUNT_SEEDS.forEach(function (seed) {
      const id = normalizeAdminId_(seed.id);
      if (!id) throw new Error('ADMIN_ACCOUNT_SEEDS에 올바르지 않은 관리자 ID가 있습니다.');
      if (nextAccounts[id]) throw new Error('ADMIN_ACCOUNT_SEEDS에 중복된 관리자 ID가 있습니다: ' + id);
      const previous = registry.accounts[id] || {};
      const version = Math.max(1, Number(seed.passwordVersion || 1));
      const initialPassword = String(seed.initialPassword || '1234');
      let salt = String(previous.salt || '');
      let passwordHash = String(previous.passwordHash || '');
      let hashScheme = String(previous.hashScheme || 'v1');
      let mustChangePassword = previous.mustChangePassword === false ? false : seed.forcePasswordChange !== false;
      const hadPasswordHash = Boolean(passwordHash);
      const seedVersionChanged = previous.seedPasswordVersion != null &&
        Number(previous.seedPasswordVersion) !== version;
      let credentialVersion = Math.max(1, Number(previous.passwordVersion || version));

      if (!passwordHash || seedVersionChanged) {
        const bootstrapPassword = initialPassword || createStrongTemporaryPassword_();
        salt = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
        passwordHash = hashAdminPasswordV2_(id, bootstrapPassword, salt);
        hashScheme = 'v2';
        mustChangePassword = true;
        credentialVersion = Math.max(credentialVersion + (hadPasswordHash ? 1 : 0), version);
        if (!initialPassword) Logger.log('[최초 1회] ' + id + ' 임시 비밀번호: ' + bootstrapPassword);
      }

      nextAccounts[id] = {
        id: id,
        displayName: String(seed.displayName || id).trim().slice(0, 80),
        role: String(seed.role || 'manager') === 'developer' ? 'developer' : 'manager',
        active: seed.active !== false,
        salt: salt,
        passwordHash: passwordHash,
        passwordVersion: credentialVersion,
        seedPasswordVersion: version,
        hashScheme: hashScheme,
        mustChangePassword: mustChangePassword
      };
    });

    registry = {
      version: 1,
      signature: signature,
      syncedAt: new Date().toISOString(),
      accounts: nextAccounts
    };
    properties.setProperty(ADMIN_ACCOUNT_REGISTRY_KEY, JSON.stringify(registry));
    try {
      scriptCache.put(
        ADMIN_ACCOUNT_REGISTRY_CACHE_KEY,
        JSON.stringify(registry),
        300
      );
    } catch (cacheError) {}
    return registry;
  }, 15000);
}

function hashAdminPassword_(adminId, password, salt) {
  return hashText_([String(salt || ''), String(adminId || ''), String(password || '')].join('|'));
}

function hashAdminPasswordV2_(adminId, password, salt) {
  let value = [String(salt || ''), String(adminId || ''), String(password || '')].join('|');
  for (let i = 0; i < ADMIN_PASSWORD_HASH_ITERATIONS; i++) {
    value = sha256Hex_(value + '|' + salt + '|' + i);
  }
  return value;
}

const SHA256_K_ = [
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
  0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
  0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
  0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
  0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2
];

/** Utilities.computeDigest(SHA_256, 문자열, UTF_8)와 같은 결과를 서비스 호출 없이 계산합니다. */
function sha256Hex_(text) {
  const str = String(text || '');
  const bytes = [];
  for (let i = 0; i < str.length; i++) {
    let code = str.charCodeAt(i);
    if (code >= 0xd800 && code <= 0xdbff && i + 1 < str.length) {
      const next = str.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        code = 0x10000 + ((code - 0xd800) << 10) + (next - 0xdc00);
        i++;
      }
    }
    if (code < 0x80) bytes.push(code);
    else if (code < 0x800) bytes.push(0xc0 | (code >> 6), 0x80 | (code & 63));
    else if (code < 0x10000) bytes.push(0xe0 | (code >> 12), 0x80 | ((code >> 6) & 63), 0x80 | (code & 63));
    else bytes.push(0xf0 | (code >> 18), 0x80 | ((code >> 12) & 63), 0x80 | ((code >> 6) & 63), 0x80 | (code & 63));
  }
  const bitLength = bytes.length * 8;
  bytes.push(0x80);
  while (bytes.length % 64 !== 56) bytes.push(0);
  const high = Math.floor(bitLength / 0x100000000);
  bytes.push((high >>> 24) & 255, (high >>> 16) & 255, (high >>> 8) & 255, high & 255);
  bytes.push((bitLength >>> 24) & 255, (bitLength >>> 16) & 255, (bitLength >>> 8) & 255, bitLength & 255);

  const h = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];
  const w = new Array(64);
  for (let offset = 0; offset < bytes.length; offset += 64) {
    for (let t = 0; t < 16; t++) {
      const j = offset + t * 4;
      w[t] = ((bytes[j] << 24) | (bytes[j + 1] << 16) | (bytes[j + 2] << 8) | bytes[j + 3]) | 0;
    }
    for (let t = 16; t < 64; t++) {
      const a = w[t - 15], b = w[t - 2];
      const s0 = ((a >>> 7) | (a << 25)) ^ ((a >>> 18) | (a << 14)) ^ (a >>> 3);
      const s1 = ((b >>> 17) | (b << 15)) ^ ((b >>> 19) | (b << 13)) ^ (b >>> 10);
      w[t] = (w[t - 16] + s0 + w[t - 7] + s1) | 0;
    }
    let a = h[0], b = h[1], c = h[2], d = h[3], e = h[4], f = h[5], g = h[6], k = h[7];
    for (let t = 0; t < 64; t++) {
      const S1 = ((e >>> 6) | (e << 26)) ^ ((e >>> 11) | (e << 21)) ^ ((e >>> 25) | (e << 7));
      const ch = (e & f) ^ (~e & g);
      const temp1 = (k + S1 + ch + SHA256_K_[t] + w[t]) | 0;
      const S0 = ((a >>> 2) | (a << 30)) ^ ((a >>> 13) | (a << 19)) ^ ((a >>> 22) | (a << 10));
      const maj = (a & b) ^ (a & c) ^ (b & c);
      const temp2 = (S0 + maj) | 0;
      k = g; g = f; f = e; e = (d + temp1) | 0;
      d = c; c = b; b = a; a = (temp1 + temp2) | 0;
    }
    h[0] = (h[0] + a) | 0; h[1] = (h[1] + b) | 0; h[2] = (h[2] + c) | 0; h[3] = (h[3] + d) | 0;
    h[4] = (h[4] + e) | 0; h[5] = (h[5] + f) | 0; h[6] = (h[6] + g) | 0; h[7] = (h[7] + k) | 0;
  }
  let hex = '';
  for (let i = 0; i < 8; i++) hex += ('00000000' + (h[i] >>> 0).toString(16)).slice(-8);
  return hex;
}

/** 편집기에서 1회 실행: 새 계산 방식이 기존과 같은 값을 내는지, 얼마나 빠른지 확인합니다. */
function testLoginHashSpeed() {
  const sample = 'speed-test|sample-admin|Sample!Password123';
  let oldValue = sample;
  let started = Date.now();
  for (let i = 0; i < 2000; i++) oldValue = hashText_(oldValue + '|' + i);
  const oldMs = Date.now() - started;
  let newValue = sample;
  started = Date.now();
  for (let i = 0; i < 2000; i++) newValue = sha256Hex_(newValue + '|' + i);
  const newMs = Date.now() - started;
  const scale = ADMIN_PASSWORD_HASH_ITERATIONS / 2000;
  console.log('결과 일치: ' + (oldValue === newValue ? '예 (그대로 써도 됩니다)' : '아니오 (적용하면 안 됩니다)'));
  console.log('기존 방식 예상 로그인 계산 시간: 약 ' + Math.round(oldMs * scale) + 'ms');
  console.log('새 방식 예상 로그인 계산 시간: 약 ' + Math.round(newMs * scale) + 'ms');
}

function verifyAdminPassword_(account, password) {
  const expected = String(account && account.passwordHash || '');
  const actual = String(account && account.hashScheme || 'v1') === 'v2'
    ? hashAdminPasswordV2_(account.id, password, account.salt)
    : hashAdminPassword_(account.id, password, account.salt);
  return constantTimeEquals_(expected, actual);
}

function createStrongTemporaryPassword_() {
  return 'V!' + Utilities.getUuid().replace(/-/g, '').slice(0, 10) +
    Utilities.getUuid().replace(/-/g, '').slice(0, 10) + '9a';
}

function hashText_(value) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value || ''),
    Utilities.Charset.UTF_8
  );
  return bytes.map(function (byte) {
    const normalized = byte < 0 ? byte + 256 : byte;
    return ('0' + normalized.toString(16)).slice(-2);
  }).join('');
}

function hashPortalToken_(token) {
  return hashText_('portal-session|' + String(token || ''));
}

function constantTimeEquals_(left, right) {
  const a = String(left || '');
  const b = String(right || '');
  let diff = a.length ^ b.length;
  const max = Math.max(a.length, b.length);
  for (let i = 0; i < max; i++) {
    diff |= (a.charCodeAt(i % Math.max(a.length, 1)) || 0) ^
      (b.charCodeAt(i % Math.max(b.length, 1)) || 0);
  }
  return diff === 0;
}

function createAdminSession_(account) {
  // 만료 세션은 접근 시 즉시 차단되므로 매 로그인마다 전체 속성을 훑지 않습니다.
  // 낮은 확률의 정리만 수행해 로그인 체감 시간을 일정하게 유지합니다.
  if (Math.random() < 0.05) cleanupExpiredAdminSessions_();
  const token = Utilities.getUuid().replace(/-/g, '') +
    Utilities.getUuid().replace(/-/g, '') +
    Utilities.getUuid().replace(/-/g, '');
  const now = Date.now();
  const expiresAt = now + ADMIN_SESSION_HOURS * 60 * 60 * 1000;
  const session = {
    id: account.id,
    role: account.role,
    displayName: account.displayName,
    passwordVersion: account.passwordVersion,
    issuedAt: now,
    lastSeenAt: now,
    expiresAt: expiresAt
  };
  PropertiesService.getScriptProperties().setProperty(
    ADMIN_SESSION_PREFIX + hashPortalToken_(token),
    JSON.stringify(session)
  );
  return { token: token, expiresAt: new Date(expiresAt).toISOString() };
}

function getAdminSession_(token) {
  return withScriptLock_(function () {
    const rawToken = String(token || '');
    if (!/^[a-f0-9]{64,160}$/i.test(rawToken)) throw new Error('로그인이 필요합니다.');
    const properties = PropertiesService.getScriptProperties();
    const key = ADMIN_SESSION_PREFIX + hashPortalToken_(rawToken);
    const raw = properties.getProperty(key);
    if (!raw) throw new Error('로그인이 만료되었습니다. 다시 로그인해주세요.');
    let session;
    try { session = JSON.parse(raw); }
    catch (e) { session = null; }
    if (!session || !Number.isFinite(Number(session.expiresAt)) || Number(session.expiresAt || 0) <= Date.now()) {
      properties.deleteProperty(key);
      throw new Error('로그인이 만료되었습니다. 다시 로그인해주세요.');
    }
    const now = Date.now();
    const lastSeenAt = Number(session.lastSeenAt || session.issuedAt || 0);
    if (!lastSeenAt || now - lastSeenAt > ADMIN_SESSION_IDLE_MINUTES * 60 * 1000 ||
        now - Number(session.issuedAt || now) > ADMIN_SESSION_ABSOLUTE_HOURS * 60 * 60 * 1000) {
      properties.deleteProperty(key);
      throw new Error('보안을 위해 로그인이 만료되었습니다. 다시 로그인해주세요.');
    }
    const registry = synchronizeAdminAccountsFromCode_(true);
    const account = registry.accounts[session.id];
    if (!account || !account.active) {
      properties.deleteProperty(key);
      throw new Error('사용이 중지된 관리자 계정입니다.');
    }
    if (Number(session.passwordVersion || 0) !== Number(account.passwordVersion || 0)) {
      properties.deleteProperty(key);
      throw new Error('비밀번호가 변경되어 다시 로그인해야 합니다.');
    }
    session.role = account.role;
    session.displayName = account.displayName;
    session.mustChangePassword = Boolean(account.mustChangePassword);
    if (now - lastSeenAt > 5 * 60 * 1000) {
      session.lastSeenAt = now;
      properties.setProperty(key, JSON.stringify(session));
    }
    return session;
  }, 15000);
}

function withAdminSession_(token, callback, options) {
  const session = getAdminSession_(token);
  const previous = ADMIN_RUNTIME_CONTEXT;
  const previousRequestCache = ADMIN_REQUEST_CACHE;
  ADMIN_REQUEST_CACHE = {};
  ADMIN_RUNTIME_CONTEXT = {
    id: session.id,
    role: session.role,
    displayName: session.displayName,
    mustChangePassword: Boolean(session.mustChangePassword)
  };
  try {
    if (ADMIN_RUNTIME_CONTEXT.mustChangePassword && !(options && options.allowPasswordChange)) {
      throw new Error('보안을 위해 초기 비밀번호를 먼저 변경해주세요.');
    }
    return callback(ADMIN_RUNTIME_CONTEXT);
  } finally {
    ADMIN_RUNTIME_CONTEXT = previous;
    ADMIN_REQUEST_CACHE = previousRequestCache;
  }
}

function cleanupExpiredAdminSessions_() {
  const properties = PropertiesService.getScriptProperties();
  const all = properties.getProperties();
  let inspected = 0;
  Object.keys(all).forEach(function (key) {
    if (inspected >= 30 || key.indexOf(ADMIN_SESSION_PREFIX) !== 0) return;
    inspected++;
    try {
      const session = JSON.parse(all[key]);
      if (!session || Number(session.expiresAt || 0) <= Date.now()) properties.deleteProperty(key);
    } catch (e) {
      properties.deleteProperty(key);
    }
  });
}

function requireDeveloperRole_(currentAdmin) {
  if (!currentAdmin || currentAdmin.role !== 'developer') {
    throw new Error('총괄 개발자 계정만 사용할 수 있는 기능입니다.');
  }
}

function buildPortalBootstrap_(currentAdmin) {
  const profile = getAdminProfile_(currentAdmin.id);
  const response = {
    success: true,
    admin: {
      id: currentAdmin.id,
      role: currentAdmin.role,
      displayName: currentAdmin.displayName,
      isDeveloper: currentAdmin.role === 'developer'
    },
    profile: profile,
    mustReplaceInitialPassword: Boolean(currentAdmin.mustChangePassword),
    dashboard: {
      counts: { all: 0, upcoming: 0, active: 0, past: 0, blocked: 0 },
      recentCourses: [],
      upcomingCourse: null,
      attentionCourses: [],
      profileComplete: isProfileComplete_(profile)
    }
  };
  if (currentAdmin.role === 'developer') {
    const registry = synchronizeAdminAccountsFromCode_();
    response.adminAccounts = Object.keys(registry.accounts).map(function (id) {
      const account = registry.accounts[id];
      return {
        id: account.id,
        displayName: account.displayName,
        role: account.role,
        active: account.active
      };
    });
  }
  return response;
}

/****************************************************
 * 9-B. 포털 시트 · 프로필 · 강의 권한
 ****************************************************/

function getOrCreatePortalSheet_(sheetName, headers) {
  const cacheKey = 'portalSheet:' + sheetName;
  if (ADMIN_REQUEST_CACHE && ADMIN_REQUEST_CACHE[cacheKey]) return ADMIN_REQUEST_CACHE[cacheKey];
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = withScriptLock_(function () { return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName); });
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (headers.some(function (header, index) { return String(current[index] || '') !== header; })) {
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#ede9fe')
      .setFontColor('#4c1d95');
  }
  if (sheet.getFrozenRows() !== 1) sheet.setFrozenRows(1);
  if (ADMIN_REQUEST_CACHE) ADMIN_REQUEST_CACHE[cacheKey] = sheet;
  return sheet;
}

function findPortalRowByValue_(sheet, column, value) {
  if (!sheet || sheet.getLastRow() < 2) return null;
  const values = sheet.getRange(2, column, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(value)) return i + 2;
  }
  return null;
}

function getProfileHeaders_() {
  return [
    '관리자ID',
    '이름',
    '유선번호',
    '휴대폰번호',
    '이메일',
    '소속기관',
    '직책',
    '수정일시',
    '유선번호공개',
    '휴대폰번호공개',
    '이메일공개'
  ];
}

function getAdminProfile_(adminId) {
  const profileCacheKey = 'VUP_ADMIN_PROFILE_V2_' + String(adminId || '');
  try {
    const cached = CacheService.getScriptCache().get(profileCacheKey);
    if (cached) return JSON.parse(cached);
  } catch (cacheError) {}

  const sheet = getOrCreatePortalSheet_(SHEET_ADMIN_PROFILES, getProfileHeaders_());
  const rowIndex = findPortalRowByValue_(sheet, 1, adminId);
  const seed = getAccountSeed_(adminId);
  let profile;
  if (!rowIndex) {
    profile = {
      adminId: adminId,
      name: String(seed.displayName || adminId),
      landline: '',
      mobile: '',
      email: '',
      organization: '',
      position: '',
      publicLandline: false,
      publicMobile: false,
      publicEmail: false,
      updatedAt: ''
    };
  } else {
    const row = sheet.getRange(rowIndex, 1, 1, getProfileHeaders_().length).getValues()[0];
    profile = {
      adminId: String(row[0] || ''),
      name: String(row[1] || ''),
      landline: String(row[2] || ''),
      mobile: String(row[3] || ''),
      email: String(row[4] || ''),
      organization: String(row[5] || ''),
      position: String(row[6] || ''),
      publicLandline: row[8] === true || String(row[8] || '').toUpperCase() === 'TRUE',
      publicMobile: row[9] === true || String(row[9] || '').toUpperCase() === 'TRUE',
      publicEmail: row[10] === true || String(row[10] || '').toUpperCase() === 'TRUE',
      updatedAt: formatDate_(row[7])
    };
  }
  try {
    CacheService.getScriptCache().put(
      profileCacheKey,
      JSON.stringify(profile),
      ADMIN_PROFILE_CACHE_SECONDS
    );
  } catch (cacheError) {}
  return profile;
}

function saveAdminProfile_(adminId, profile) {
  const input = profile || {};
  const name = String(input.name || '').trim().slice(0, 80);
  const email = String(input.email || '').trim().toLowerCase().slice(0, 160);
  const organization = String(input.organization || '').trim().slice(0, 120);
  const position = String(input.position || '').trim().slice(0, 80);
  const landline = String(input.landline || '').trim().slice(0, 30);
  const mobile = String(input.mobile || '').trim().slice(0, 30);
  const phonePattern = /^\d{2,4}-\d{3,4}-\d{4}$/;
  if (!name) throw new Error('관리자 이름을 입력해주세요.');
  if (!organization) throw new Error('소속 기관명을 입력해주세요.');
  if (!position) throw new Error('직책을 입력해주세요.');
  if (!landline) throw new Error('유선번호를 입력해주세요.');
  if (!phonePattern.test(landline)) {
    throw new Error('유선번호는 하이픈을 포함해 입력해주세요. 예: 02-3017-0000');
  }
  if (mobile && !phonePattern.test(mobile)) {
    throw new Error('휴대폰번호는 하이픈을 포함해 입력해주세요. 예: 010-1234-5678');
  }
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error('이메일 주소에 @를 포함해 올바르게 입력해주세요.');
  }
  const row = [
    adminId,
    name,
    landline,
    mobile,
    email,
    organization,
    position,
    new Date(),
    Boolean(input.publicLandline),
    Boolean(input.publicMobile),
    Boolean(input.publicEmail)
  ];
  const sheet = getOrCreatePortalSheet_(SHEET_ADMIN_PROFILES, getProfileHeaders_());
  const rowIndex = findPortalRowByValue_(sheet, 1, adminId);
  if (rowIndex) sheet.getRange(rowIndex, 1, 1, row.length).setValues([safeSheetRow_(row)]);
  else sheet.appendRow(safeSheetRow_(row));
  try {
    CacheService.getScriptCache().remove(
      'VUP_ADMIN_PROFILE_V2_' + String(adminId || '')
    );
  } catch (cacheError) {}
  Object.keys(PUBLIC_MANAGER_CONTACT_RUNTIME_CACHE).forEach(function (siteId) {
    delete PUBLIC_MANAGER_CONTACT_RUNTIME_CACHE[siteId];
  });
  return getAdminProfile_(adminId);
}

function getPublicManagerContact_(site) {
  const siteId = String(site && site.siteId || '');
  if (PUBLIC_MANAGER_CONTACT_RUNTIME_CACHE[siteId]) {
    return PUBLIC_MANAGER_CONTACT_RUNTIME_CACHE[siteId];
  }
  const settings = getCourseSettingsMap_()[siteId] || {};
  const ownerId = String(settings.ownerId || getCourseOwnerId_(siteId) || '');
  const profile = ownerId ? getAdminProfile_(ownerId) : {};
  const contact = {
    name: String(profile.name || site.managerName || ''),
    organization: String(profile.organization || ''),
    position: String(profile.position || ''),
    landline: profile.publicLandline ? String(profile.landline || '') : '',
    mobile: profile.publicMobile ? String(profile.mobile || '') : '',
    email: profile.publicEmail ? String(profile.email || '') : ''
  };
  PUBLIC_MANAGER_CONTACT_RUNTIME_CACHE[siteId] = contact;
  return contact;
}

function getCourseOwnerHeaders_() {
  return ['siteId', '관리자ID', '배정일시'];
}

function recordCourseOwner_(siteId, adminId) {
  return withScriptLock_(function () {
    const sheet = getOrCreatePortalSheet_(SHEET_ADMIN_COURSE_OWNERS, getCourseOwnerHeaders_());
    const rowIndex = findPortalRowByValue_(sheet, 1, siteId);
    const row = [siteId, adminId, new Date()];
    if (rowIndex) sheet.getRange(rowIndex, 1, 1, row.length).setValues([safeSheetRow_(row)]);
    else sheet.appendRow(safeSheetRow_(row));
    if (ADMIN_REQUEST_CACHE) ADMIN_REQUEST_CACHE.ownerMap = null;
  }, 15000);
}

function getCourseOwnerMap_() {
  if (ADMIN_REQUEST_CACHE && ADMIN_REQUEST_CACHE.ownerMap) return ADMIN_REQUEST_CACHE.ownerMap;
  const sheet = getOrCreatePortalSheet_(SHEET_ADMIN_COURSE_OWNERS, getCourseOwnerHeaders_());
  const result = {};
  if (sheet.getLastRow() < 2) return result;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function (row) {
    if (row[0]) result[String(row[0])] = String(row[1] || '');
  });
  if (ADMIN_REQUEST_CACHE) ADMIN_REQUEST_CACHE.ownerMap = result;
  return result;
}

function getCourseOwnerId_(siteId) {
  return getCourseOwnerMap_()[String(siteId || '')] || '';
}

function assertCourseAccess_(currentAdmin, siteId) {
  const normalizedSiteId = normalizePublicSiteId_(siteId);
  const site = getSiteInfoById_(normalizedSiteId);
  if (!site) throw new Error('강의를 찾을 수 없습니다.');
  if (currentAdmin.role === 'developer') return site;
  const ownerId = getCourseOwnerId_(normalizedSiteId);
  if (!ownerId || ownerId !== currentAdmin.id) {
    throw new Error('이 강의를 관리할 권한이 없습니다.');
  }
  return site;
}

/****************************************************
 * 9-C. 강의 운영 설정 · 대시보드
 ****************************************************/

function getCourseSettingsHeaders_() {
  return ['siteId', '관리자ID', '운영설정JSON', '수정일시'];
}

function getDefaultCourseSettings_(siteId, ownerId, skipAccessCode) {
  return {
    siteId: String(siteId || ''),
    ownerId: String(ownerId || ''),
    accessCode: skipAccessCode ? '' : createUniqueCourseAccessCode_(),
    rememberUrl: '',
    rememberEnabled: false,
    busInfo: '',
    progressText: '교육 준비 중',
    organizationEditEnabled: true,
    preOpenAt: '',
    preCloseAt: '',
    courseOpenAt: '',
    materialReleaseAt: '',
    mealReleaseAt: '',
    photoReleaseAt: '',
    surveyReleaseAt: '',
    closeAt: '',
    surveyUrl: '',
    updatedAt: ''
  };
}

function getCourseSettingsMap_() {
  if (ADMIN_REQUEST_CACHE && ADMIN_REQUEST_CACHE.settingsMap) return ADMIN_REQUEST_CACHE.settingsMap;
  const sheet = getOrCreatePortalSheet_(SHEET_ADMIN_COURSE_SETTINGS, getCourseSettingsHeaders_());
  const result = {};
  if (sheet.getLastRow() < 2) return result;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues().forEach(function (row) {
    const siteId = String(row[0] || '');
    if (!siteId) return;
    let parsed = {};
    try { parsed = JSON.parse(String(row[2] || '{}')); } catch (e) {}
    result[siteId] = Object.assign(
      getDefaultCourseSettings_(siteId, String(row[1] || ''), true),
      parsed || {},
      { updatedAt: formatDate_(row[3]) }
    );
  });
  if (ADMIN_REQUEST_CACHE) ADMIN_REQUEST_CACHE.settingsMap = result;
  return result;
}

function getCourseSettings_(siteId) {
  const ownerId = getCourseOwnerId_(siteId);
  const map = getCourseSettingsMap_();
  return map[String(siteId || '')] || getDefaultCourseSettings_(siteId, ownerId);
}

function validateCourseSettingsDraft_(siteId, settings) {
  const input = settings || {};
  const accessCode = String(input.accessCode || '').replace(/\D/g, '');
  if (input.accessCode != null && input.accessCode !== '' && !/^\d{4}$/.test(accessCode)) {
    throw new Error('교육 고유 코드는 숫자 4자리로 입력해주세요.');
  }
  ['rememberUrl', 'surveyUrl'].forEach(function (field) {
    const value = String(input[field] || '').trim();
    if (value && !/^https:\/\//i.test(value)) {
      throw new Error(field === 'rememberUrl'
        ? '리멤버 링크는 https:// 주소로 입력해주세요.'
        : '설문 링크는 https:// 주소로 입력해주세요.');
    }
  });
  [
    'preOpenAt',
    'preCloseAt',
    'courseOpenAt',
    'materialReleaseAt',
    'mealReleaseAt',
    'photoReleaseAt',
    'surveyReleaseAt',
    'closeAt'
  ].forEach(function (field) {
    const value = String(input[field] || '').trim();
    if (value && isNaN(new Date(value).getTime())) throw new Error('공개 일정 날짜와 시간을 확인해주세요.');
  });
  validateCourseChronology_(input);
  if (accessCode) {
    const existingSettings = getCourseSettingsMap_();
    const duplicateSiteId = Object.keys(existingSettings).find(function (existingSiteId) {
      return existingSiteId !== String(siteId || '') &&
        String(existingSettings[existingSiteId].accessCode || '') === accessCode;
    });
    if (duplicateSiteId) {
      throw new Error('이미 다른 강의에서 사용 중인 교육 고유 코드입니다. 다른 숫자 4자리를 입력해주세요.');
    }
  }
}

function validateCourseChronology_(settings) {
  const input = settings || {};
  const time = function (field) {
    const value = String(input[field] || '').trim();
    return value ? new Date(value).getTime() : NaN;
  };
  const preOpen = time('preOpenAt');
  const preClose = time('preCloseAt');
  const courseOpen = time('courseOpenAt');
  const close = time('closeAt');
  if ([preOpen, preClose, courseOpen, close].every(isFinite)) {
    if (!(preOpen < preClose && preClose <= courseOpen && courseOpen < close)) {
      throw new Error('공개 일정은 사전 안내 시작 < 사전 안내 종료 ≤ 교육 시작 < 운영 종료 순서여야 합니다.');
    }
    ['materialReleaseAt', 'mealReleaseAt', 'photoReleaseAt', 'surveyReleaseAt'].forEach(function (field) {
      const release = time(field);
      if (isFinite(release) && (release < preOpen || release > close)) {
        throw new Error('교재·식사·사진·설문 공개 시각은 전체 운영 기간 안에 있어야 합니다.');
      }
    });
  }
}

function saveCourseSettings_(siteId, ownerId, settings) {
  return withScriptLock_(function () {
  if (ADMIN_REQUEST_CACHE) ADMIN_REQUEST_CACHE.settingsMap = null;
  const current = getCourseSettings_(siteId);
  const input = Object.assign({}, current, settings || {});
  const urlFields = ['rememberUrl', 'surveyUrl'];
  urlFields.forEach(function (field) {
    const value = String(input[field] || '').trim();
    if (value && !/^https:\/\//i.test(value)) {
      throw new Error(field === 'rememberUrl'
        ? '리멤버 링크는 https:// 주소로 입력해주세요.'
        : '설문 링크는 https:// 주소로 입력해주세요.');
    }
    input[field] = value.slice(0, 1200);
  });
  const accessCode = String(input.accessCode || '').replace(/\D/g, '');
  if (!/^\d{4}$/.test(accessCode)) throw new Error('교육 고유 코드는 숫자 4자리로 입력해주세요.');
  const existingSettings = getCourseSettingsMap_();
  const duplicateSiteId = Object.keys(existingSettings).find(function (existingSiteId) {
    return existingSiteId !== String(siteId) &&
      String(existingSettings[existingSiteId].accessCode || '') === accessCode;
  });
  if (duplicateSiteId) {
    throw new Error('이미 다른 강의에서 사용 중인 교육 고유 코드입니다. 다른 숫자 4자리를 입력해주세요.');
  }
  const dateFields = [
    'preOpenAt',
    'preCloseAt',
    'courseOpenAt',
    'materialReleaseAt',
    'mealReleaseAt',
    'photoReleaseAt',
    'surveyReleaseAt',
    'closeAt'
  ];
  dateFields.forEach(function (field) {
    const value = String(input[field] || '').trim().slice(0, 40);
    if (value && isNaN(new Date(value).getTime())) throw new Error('공개 일정 날짜와 시간을 확인해주세요.');
    input[field] = value;
  });
  validateCourseChronology_(input);
  const safe = {
    siteId: siteId,
    ownerId: ownerId,
    accessCode: accessCode,
    rememberUrl: input.rememberUrl,
    rememberEnabled: Boolean(input.rememberEnabled && input.rememberUrl),
    busInfo: String(input.busInfo || '').trim().slice(0, 1000),
    progressText: String(input.progressText || '교육 준비 중').trim().slice(0, 200),
    organizationEditEnabled: input.organizationEditEnabled !== false,
    preOpenAt: input.preOpenAt,
    preCloseAt: input.preCloseAt,
    courseOpenAt: input.courseOpenAt,
    materialReleaseAt: input.materialReleaseAt,
    mealReleaseAt: input.mealReleaseAt,
    photoReleaseAt: input.photoReleaseAt,
    surveyReleaseAt: input.surveyReleaseAt,
    closeAt: input.closeAt,
    surveyUrl: input.surveyUrl,
    updatedAt: new Date().toISOString()
  };
  const sheet = getOrCreatePortalSheet_(SHEET_ADMIN_COURSE_SETTINGS, getCourseSettingsHeaders_());
  const rowIndex = findPortalRowByValue_(sheet, 1, siteId);
  const row = [siteId, ownerId, JSON.stringify(safe), new Date()];
  if (rowIndex) sheet.getRange(rowIndex, 1, 1, row.length).setValues([safeSheetRow_(row)]);
  else sheet.appendRow(safeSheetRow_(row));
  if (ADMIN_REQUEST_CACHE) ADMIN_REQUEST_CACHE.settingsMap = null;
  return safe;
  }, 15000);
}

function createUniqueCourseAccessCode_() {
  const used = {};
  try {
    const settings = getCourseSettingsMap_();
    Object.keys(settings).forEach(function (siteId) {
      if (settings[siteId].accessCode) used[settings[siteId].accessCode] = true;
    });
  } catch (e) {}
  for (let i = 0; i < 80; i++) {
    const code = String(Math.floor(1000 + Math.random() * 9000));
    if (!used[code]) return code;
  }
  return String(Math.floor(1000 + Math.random() * 9000));
}

function sanitizeCourseSettingsForAdmin_(settings) {
  return Object.assign({}, settings || {});
}

function getPortalCourses_(currentAdmin) {
  const result = getEducationSites(false);
  if (!result || !result.success) throw new Error(result && result.message ? result.message : '강의 목록을 불러오지 못했습니다.');
  const ownerMap = getCourseOwnerMap_();
  const settingsMap = getCourseSettingsMap_();
  return (result.sites || [])
    .filter(function (site) {
      return currentAdmin.role === 'developer' || ownerMap[site.siteId] === currentAdmin.id;
    })
    .map(function (site) {
      return decoratePortalCourse_(
        site,
        settingsMap[site.siteId] || getDefaultCourseSettings_(site.siteId, ownerMap[site.siteId], true),
        ownerMap[site.siteId] || ''
      );
    })
    .sort(function (a, b) {
      return String(b.updatedAt || b.createdAt || '')
        .localeCompare(String(a.updatedAt || a.createdAt || ''));
    });
}

function decoratePortalCourse_(site, settings, ownerId) {
  const phase = getCoursePhase_(site);
  const dates = [];
  for (let day = 1; day <= Number(site.courseDays || 0); day++) {
    if (site['day' + day + 'Date']) dates.push(site['day' + day + 'Date']);
  }
  const readinessItems = [
    Boolean(site.userUrl),
    Boolean(site.linkedSheetId),
    Boolean(site.materialFolderId),
    Boolean(site.photoFolderId),
    Boolean(settings && settings.surveyUrl)
  ];
  return Object.assign({}, site, {
    ownerId: ownerId || '',
    phase: phase,
    phaseLabel: {
      upcoming: '교육 시작 전',
      active: '교육 진행 중',
      past: '종료된 교육',
      blocked: '사용자 사이트 운영 중단'
    }[phase] || '교육 진행 중',
    dates: dates,
    dateSummary: dates.length ? dates[0] + (dates.length > 1 ? ' 외 ' + (dates.length - 1) + '일' : '') : '일정 미등록',
    readinessPercent: Math.round(readinessItems.filter(Boolean).length / readinessItems.length * 100),
    settings: sanitizeCourseSettingsForAdmin_(settings || {})
  });
}

function getCoursePhase_(site) {
  if (['비공개', '접근차단', '운영종료'].indexOf(String(site.status || '')) !== -1) return 'blocked';
  const today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  const dateKeys = [];
  for (let day = 1; day <= Number(site.courseDays || 0); day++) {
    const parts = parseKoreanDateParts_(site['day' + day + 'Date']);
    if (parts) dateKeys.push(parts.dateKey);
  }
  if (!dateKeys.length) return 'active';
  dateKeys.sort();
  if (today < dateKeys[0]) return 'upcoming';
  if (today > dateKeys[dateKeys.length - 1]) return 'past';
  return 'active';
}

function buildPortalDashboard_(currentAdmin) {
  const courses = getPortalCourses_(currentAdmin);
  const counts = { all: courses.length, upcoming: 0, active: 0, past: 0, blocked: 0 };
  courses.forEach(function (course) {
    if (counts[course.phase] != null) counts[course.phase]++;
  });
  const attention = courses.filter(function (course) {
    return ['upcoming', 'active'].indexOf(course.phase) !== -1 && course.readinessPercent < 100;
  }).slice(0, 6);
  const upcomingCourse = courses
    .filter(function (course) { return course.phase === 'upcoming'; })
    .sort(function (a, b) {
      const aParts = parseKoreanDateParts_(a.day1Date);
      const bParts = parseKoreanDateParts_(b.day1Date);
      return String(aParts && aParts.dateKey || '9999-12-31')
        .localeCompare(String(bParts && bParts.dateKey || '9999-12-31'));
    })[0] || null;
  return {
    counts: counts,
    recentCourses: courses.slice(0, 2),
    upcomingCourse: upcomingCourse,
    attentionCourses: attention,
    profileComplete: isProfileComplete_(getAdminProfile_(currentAdmin.id))
  };
}

function isProfileComplete_(profile) {
  return Boolean(
    profile &&
    profile.name &&
    profile.organization &&
    profile.position &&
    profile.landline &&
    profile.email
  );
}

function buildLightweightCourseHealth_(site, settings, materials, photos) {
  const items = [
    { key: 'link', label: '사용자 사이트', ready: Boolean(site.userUrl) },
    { key: 'participants', label: '교육생·식사 데이터', ready: Boolean(site.linkedSheetId) },
    { key: 'materials', label: '교재', ready: materials == null ? Boolean(site.materialFolderId) : Boolean(materials.length) },
    { key: 'photos', label: '현장 사진', ready: photos == null ? Boolean(site.photoFolderId) : Boolean(photos.length) },
    { key: 'survey', label: '설문', ready: Boolean(settings && settings.surveyUrl) }
  ];
  return {
    readyCount: items.filter(function (item) { return item.ready; }).length,
    totalCount: items.length,
    percent: Math.round(items.filter(function (item) { return item.ready; }).length / items.length * 100),
    items: items
  };
}

/****************************************************
 * 9-D. 공개 단계 설정
 ****************************************************/

function getPublicCourseAccessConfig_(siteId) {
  const settings = getCourseSettingsMap_()[String(siteId || '')];
  if (!settings) {
    return {
      configured: false,
      phase: 'legacy',
      siteAvailable: true,
      availableFrom: '',
      preGuidanceEndsAt: '',
      closesAt: '',
      progressText: '',
      busInfo: '',
      organizationEditEnabled: true,
      rememberUrl: '',
      surveyUrl: '',
      modules: {
        bus: false,
        profileEdit: true,
        progress: true,
        meal: true,
        materials: true,
        photos: true,
        remember: false,
        survey: false
      }
    };
  }
  const now = new Date();
  const preOpened = isPortalReleaseReached_(settings.preOpenAt, now, true);
  const courseOpen = isPortalReleaseReached_(settings.courseOpenAt, now, true);
  const closed = settings.closeAt && new Date(settings.closeAt).getTime() < now.getTime();
  return {
    configured: true,
    phase: !preOpened ? 'not_open' : (closed ? 'closed' : (courseOpen ? 'course' : 'before')),
    siteAvailable: Boolean(preOpened && !closed),
    availableFrom: settings.preOpenAt,
    preGuidanceEndsAt: settings.preCloseAt || settings.courseOpenAt,
    closesAt: settings.closeAt,
    progressText: settings.progressText,
    busInfo: '',
    organizationEditEnabled: settings.organizationEditEnabled,
    rememberUrl: preOpened && !closed && settings.rememberEnabled ? settings.rememberUrl : '',
    surveyUrl: preOpened && !closed && isPortalReleaseReached_(settings.surveyReleaseAt, now, courseOpen) ? settings.surveyUrl : '',
    modules: {
      bus: false,
      profileEdit: Boolean(preOpened && !closed && settings.organizationEditEnabled),
      progress: Boolean(preOpened && !closed),
      meal: Boolean(preOpened && !closed && isPortalReleaseReached_(settings.mealReleaseAt, now, courseOpen)),
      materials: Boolean(preOpened && !closed && isPortalReleaseReached_(settings.materialReleaseAt, now, courseOpen)),
      photos: Boolean(preOpened && !closed && isPortalReleaseReached_(settings.photoReleaseAt, now, courseOpen)),
      remember: Boolean(preOpened && !closed && settings.rememberEnabled && settings.rememberUrl),
      survey: Boolean(preOpened && !closed && settings.surveyUrl && isPortalReleaseReached_(settings.surveyReleaseAt, now, courseOpen))
    }
  };
}

function isPortalReleaseReached_(dateText, now, fallback) {
  const text = String(dateText || '').trim();
  if (!text) return Boolean(fallback);
  const time = new Date(text).getTime();
  return !isNaN(time) && time <= now.getTime();
}

/****************************************************
 * 9-E. 식사 자리 자동 배정
 ****************************************************/


function ensureLinkedMealSpreadsheet_(site) {
  if (site.linkedSheetId) {
    try {
      SpreadsheetApp.openById(site.linkedSheetId);
      return site.linkedSheetId;
    } catch (e) {}
  }
  const spreadsheet = SpreadsheetApp.create(
    sanitizeFileName_('V-Up 운영데이터 - ' + String(site.title || site.generation || site.siteId)).slice(0, 180)
  );
  registerManagedResource_(site.siteId, spreadsheet.getId(), 'linkedSheet', '', {
    title: spreadsheet.getName()
  });
  updateSiteLinkedSheet_(site.siteId, spreadsheet.getId(), spreadsheet.getUrl());
  return spreadsheet.getId();
}

/****************************************************
 * 9-F. 설문 · 문의
 ****************************************************/

function getSurveyHeaders_() {
  return ['설문ID', 'siteId', '생성자ID', '제목', '유형', 'FormID', '응답URL', '편집URL', '생성일시', '요청ID'];
}

function getCourseSurveys_(siteId) {
  const sheet = getOrCreatePortalSheet_(SHEET_ADMIN_SURVEYS, getSurveyHeaders_());
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, getSurveyHeaders_().length)
    .getValues()
    .filter(function (row) { return String(row[1] || '') === String(siteId); })
    .map(function (row) {
      return {
        surveyId: String(row[0] || ''),
        siteId: String(row[1] || ''),
        createdBy: String(row[2] || ''),
        title: String(row[3] || ''),
        templateType: String(row[4] || ''),
        formId: String(row[5] || ''),
        publicUrl: String(row[6] || ''),
        editUrl: String(row[7] || ''),
        createdAt: formatDate_(row[8])
      };
    })
    .sort(function (a, b) { return String(b.createdAt).localeCompare(String(a.createdAt)); });
}

function getInquiryHeaders_() {
  return ['문의ID', '작성자ID', '작성자명', '제목', '내용', '상태', '작성일시', '답변', '답변일시', '답변자ID'];
}

function getPortalInquiries_(currentAdmin) {
  const sheet = getOrCreatePortalSheet_(SHEET_ADMIN_INQUIRIES, getInquiryHeaders_());
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, getInquiryHeaders_().length)
    .getValues()
    .filter(function (row) {
      return currentAdmin.role === 'developer' || String(row[1] || '') === currentAdmin.id;
    })
    .map(function (row) {
      return {
        inquiryId: String(row[0] || ''),
        authorId: String(row[1] || ''),
        authorName: String(row[2] || ''),
        subject: String(row[3] || ''),
        body: String(row[4] || ''),
        status: String(row[5] || ''),
        createdAt: formatDate_(row[6]),
        reply: String(row[7] || ''),
        repliedAt: formatDate_(row[8]),
        repliedBy: String(row[9] || '')
      };
    })
    .sort(function (a, b) { return String(b.createdAt).localeCompare(String(a.createdAt)); });
}

/****************************************************
 * 9-G. 최초 설치 도우미
 ****************************************************/

function setupAdminPortal_() {
  const registry = synchronizeAdminAccountsFromCode_();
  getOrCreatePortalSheet_(SHEET_ADMIN_PROFILES, getProfileHeaders_());
  getOrCreatePortalSheet_(SHEET_ADMIN_COURSE_OWNERS, getCourseOwnerHeaders_());
  getOrCreatePortalSheet_(SHEET_ADMIN_COURSE_SETTINGS, getCourseSettingsHeaders_());
  getOrCreatePortalSheet_(SHEET_ADMIN_INQUIRIES, getInquiryHeaders_());
  getOrCreatePortalSheet_(SHEET_ADMIN_SURVEYS, getSurveyHeaders_());
  getOrCreatePortalSheet_(SHEET_ADMIN_RESOURCES, getResourceHeaders_());
  Logger.log('V-Up 관리자 포털 설치 완료: ' + Object.keys(registry.accounts).length + '개 계정');
  Logger.log('관리자 웹앱 주소: ' + (ScriptApp.getService().getUrl() || '웹앱 배포 후 확인'));
}
