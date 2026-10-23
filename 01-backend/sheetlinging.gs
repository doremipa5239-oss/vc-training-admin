/****************************************************
 * [4_SheetLinking.gs]
 * 식사 배치 시트 · 관리자 통합 시트 연동 · 이름 판별 · 좌석/룸메이트 파싱
 ****************************************************/

/****************************************************
 * 2. 식사 자리 배치도 (연동 시트 저장 · 읽기 · 서식)
 ****************************************************/

function saveMealArrangement(data) {
  requireAdmin_();
  try {
    if (!data || !data.siteId) throw new Error('연결할 사용자 사이트가 없습니다.');
    if (!String(data.participantText || '').trim()) throw new Error('참가자 명단 엑셀 파일을 업로드해주세요.');
    if (!data.reviewConfirmed) throw new Error('자동 인식한 교육생의 이름·소속·직급·전화번호를 관리자께서 먼저 확인해주세요.');

    const ss = getSpreadsheet_();
    const sheet = getOrCreateSitesSheet_(ss);
    const site = findSiteInfo_(sheet, data.siteId);

    if (!site) throw new Error('해당 교육 사이트를 찾을 수 없습니다.');

    if (!site.linkedSheetId) {
      throw new Error('처음 연결한 구글 시트를 찾을 수 없습니다. 식사 자리 배치의 첫 단계에서 구글 시트를 다시 연결해주세요.');
    }

    const linkedSheet = upsertMealArrangementSheet_({
      spreadsheetId: site.linkedSheetId,
      participantText: data.participantText || '',
      mealHeaders: data.mealHeaders || [],
      frequentOverlapNames: data.frequentOverlapNames || []
    });

    updateSiteLinkedSheet_(data.siteId, linkedSheet.spreadsheetId, linkedSheet.spreadsheetUrl);
    updateSiteRowById_(sheet, data.siteId, function (row) {
      row[COL.dataReviewConfirmedAt - 1] = new Date();
      row[COL.updatedAt - 1] = new Date();
      return row;
    });
    clearUserDataCache_(data.siteId);
    appendAuditLog_('MEAL_SHEET_UPDATED', data.siteId, {
      participantCount: linkedSheet.participantCount,
      reviewConfirmed: true,
      sheetName: 'AI 생성_교육생 자리 배치'
    });

    return {
      success: true,
      spreadsheetId: linkedSheet.spreadsheetId,
      spreadsheetUrl: linkedSheet.spreadsheetUrl,
      participantCount: linkedSheet.participantCount,
      userUrl: site.userUrl
    };

  } catch (error) {
    return { success: false, message: error.message };
  }
}

function upsertMealArrangementSheet_(options) {
  const spreadsheetId = String(options.spreadsheetId || '').trim();
  if (!spreadsheetId) throw new Error('연동된 구글 시트 ID가 없습니다.');

  const targetSs = SpreadsheetApp.openById(spreadsheetId);
  const targetSheetName = 'AI 생성_교육생 자리 배치';
  let sheet = targetSs.getSheetByName(targetSheetName);
  if (!sheet) {
    sheet = targetSs.insertSheet(targetSheetName);
  } else {
    sheet.clear();
    sheet.clearConditionalFormatRules();
  }

  const mealHeaders = normalizeMealHeaders_(options.mealHeaders);
  const headers = PARTICIPANT_BASE_HEADERS.concat(mealHeaders);
  const frequentOverlapNames = normalizeNameSet_(options.frequentOverlapNames);

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);

  const rows = parseParticipantTextForLinkedSheet_(options.participantText, headers.length);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(safeSheetRows_(rows));
  }

  applyParticipantSheetDesign_(sheet, headers.length, Math.max(rows.length, 1), frequentOverlapNames);
  sheet.autoResizeColumns(1, headers.length);

  targetSs.setActiveSheet(sheet);
  targetSs.moveActiveSheet(targetSs.getNumSheets());
  SpreadsheetApp.flush();

  return {
    spreadsheetId: targetSs.getId(),
    spreadsheetUrl: targetSs.getUrl() + '#gid=' + sheet.getSheetId(),
    participantCount: rows.length
  };
}


function getParticipantsFromLinkedSheet_(spreadsheetId) {
  if (!spreadsheetId) return [];

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheets = ss.getSheets();
  const sheet = ss.getSheetByName('AI 생성_교육생 자리 배치') || sheets[sheets.length - 1];
  const values = getBoundedSheetValues_(sheet, 1200, 60);

  if (values.length < 2) return [];

  const headers = values[0].map(v => String(v || '').trim());
  const dataRows = values.slice(1);

  const participants = dataRows
    .filter(row => String(row[0] || '').trim() !== '')
    .map((row, rowIndex) => {
      const mealSeats = {};
      headers.forEach((header, index) => {
        if (index < PARTICIPANT_BASE_HEADERS.length) return;
        const key = makeMealSeatKey_(header, index - PARTICIPANT_BASE_HEADERS.length);
        mealSeats[key] = String(row[index] || '').trim();
      });

      const name = String(row[0] || '').trim();
      const identity = analyzePersonName_(name);
      const base = {
        name: name,
        displayName: name,
        baseName: identity.baseName,
        nameKey: identity.nameKey,
        baseNameKey: identity.baseNameKey,
        discriminator: identity.discriminator,
        company: String(row[1] || '').trim(),
        position: String(row[2] || '').trim(),
        phoneLast4: extractPhoneLast4_(row[3]),
        personKey: identity.baseNameKey + '|' + (extractPhoneLast4_(row[3]) || 'none') + '|row' + String(rowIndex + 2)
      };

      // index-script.html은 me[currentMeal.seatKey] 형태로 "최상위 속성"에서 좌석번호를 바로 꺼내 씁니다.
      // 과거에는 day1LunchSeat~day2DinnerSeat 4개만 최상위에 박아넣어서 3~5일차 좌석 조회가
      // 조용히 실패했습니다(undefined). mealSeats를 그대로 최상위에 펼쳐서(day1~day5, 점심/저녁 전부)
      // 어떤 일수의 과정이든 currentMeal.seatKey로 바로 접근되도록 합니다.
      // mealSeats 하위객체도 그대로 유지해 기존에 이를 참조하던 코드가 있어도 깨지지 않습니다.
      return Object.assign(base, mealSeats, { mealSeats: mealSeats });
    });

  // 연결 미리보기에서 판별한 "서울홍길동 / 부산홍길동 → 홍길동" 같은
  // 공통 핵심 이름 추정을 실제 교육생 조회 경로에도 동일하게 적용합니다.
  // 원본 표시명은 그대로 유지하고 조회용 baseName/baseNameKey만 보정합니다.
  reconcileParticipantCoreNames_(participants);

  const counts = {};
  participants.forEach(function (person) {
    counts[person.baseNameKey] = (counts[person.baseNameKey] || 0) + 1;
  });
  participants.forEach(function (person) {
    person.isDuplicateName = !!person.baseNameKey && counts[person.baseNameKey] > 1;
    person.duplicateGroupKey = person.isDuplicateName ? person.baseNameKey : '';
  });
  return participants;
}

function parseParticipantTextForLinkedSheet_(text, expectedColumns) {
  const columnCount = Math.max(Number(expectedColumns || 0), PARTICIPANT_HEADERS.length);

  return String(text || '')
    .split(/\r?\n/)
    .map(line => line.trim())
    .filter(Boolean)
    .map((line, index) => {
      let cols = line.split('\t');
      if (cols.length < PARTICIPANT_BASE_HEADERS.length) cols = line.split(',');

      if (cols.length < PARTICIPANT_BASE_HEADERS.length) {
        throw new Error(`${index + 1}번째 줄의 참가자 데이터 형식이 올바르지 않습니다.`);
      }

      const normalized = cols.map(v => String(v || '').trim());
      while (normalized.length < columnCount) normalized.push('');
      return normalized.slice(0, columnCount);
    });
}

function normalizeMealHeaders_(mealHeaders) {
  const cleaned = (Array.isArray(mealHeaders) ? mealHeaders : [])
    .map(v => String(v || '').trim())
    .filter(Boolean)
    .map(v => v.endsWith('자리번호') ? v : `${v} 자리번호`);

  return cleaned.length ? cleaned : DEFAULT_MEAL_HEADERS.slice();
}

function normalizeNameSet_(names) {
  const set = new Set();

  if (Array.isArray(names)) {
    names.forEach(name => {
      const cleaned = String(name || '').trim();
      if (cleaned) set.add(cleaned);
    });
  } else if (typeof names === 'string') {
    names.split(/[,;\n\r]+/).forEach(name => {
      const cleaned = String(name || '').trim();
      if (cleaned) set.add(cleaned);
    });
  }

  return set;
}

function applyParticipantSheetDesign_(sheet, columnCount, dataRowCount, frequentOverlapNames) {
  const totalRows = Math.max(dataRowCount + 1, 2);
  const fullRange = sheet.getRange(1, 1, totalRows, columnCount);

  fullRange
    .setBorder(true, true, true, true, true, true, SAMPLE_EXCEL_STYLE.border, SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange(1, 1, 1, columnCount)
    .setBackground(SAMPLE_EXCEL_STYLE.headerBg)
    .setFontColor(SAMPLE_EXCEL_STYLE.headerFont)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  if (dataRowCount > 0) {
    sheet.getRange(2, 1, dataRowCount, columnCount)
      .setBackground(SAMPLE_EXCEL_STYLE.bodyBg)
      .setFontColor('#1f2937');

    const highlightRanges = [];
    const nameValues = sheet.getRange(2, 1, dataRowCount, 1).getValues();
    nameValues.forEach((row, index) => {
      const name = String(row[0] || '').trim();
      if (frequentOverlapNames && frequentOverlapNames.has(name)) {
        highlightRanges.push('A' + (index + 2));
      }
    });
    if (highlightRanges.length) {
      sheet.getRangeList(highlightRanges)
        .setBackground(FREQUENT_OVERLAP_HIGHLIGHT.bg)
        .setFontColor(FREQUENT_OVERLAP_HIGHLIGHT.font)
        .setFontWeight('bold');
    }
  }

  sheet.setRowHeight(1, 28);
  if (totalRows > 1) sheet.setRowHeights(2, totalRows - 1, 24);

  const widths = [140, 220, 120, 150];
  for (let col = 1; col <= columnCount; col++) {
    sheet.setColumnWidth(col, widths[col - 1] || 145);
  }
}

function makeMealSeatKey_(header, zeroBasedIndex) {
  const text = String(header || '').trim();
  const dayMatch = text.match(/(\d+)일차/);
  const isLunch = text.indexOf('점심') !== -1;
  const isDinner = text.indexOf('저녁') !== -1;

  if (dayMatch && (isLunch || isDinner)) {
    return `day${dayMatch[1]}${isLunch ? 'Lunch' : 'Dinner'}Seat`;
  }

  return `meal${zeroBasedIndex + 1}Seat`;
}

/**
 * 서버에서 샘플 엑셀을 만들 때 사용할 수 있는 룸메이트 샘플 데이터입니다.
 * 현재 관리자 페이지가 브라우저의 XLSX 라이브러리로 샘플을 직접 생성하는 구조라면,
 * 실제 다운로드 버튼 디자인은 admin-script.html의 downloadRoommateSampleExcel()도 같은 스타일로 맞춰야 합니다.
 */
function getRoommateSampleRowsForDownload() {
  requireAdmin_();
  return {
    success: true,
    sheetName: '룸메이트명단',
    fileName: 'V-Up_룸메이트명단_샘플.xlsx',
    headers: ['구분', '룸메이트 1', '룸메이트 2'],
    rows: [
      ['예) 1', '예) 홍길동', '예) 김벤처'],
      ['예) 2', '예) 이투자', '예) 박심사']
    ],
    style: {
      headerFill: 'C6EFCE',
      borderColor: '000000',
      columnWidths: [12, 25, 25],
      rowCount: 200
    }
  };
}

/****************************************************
 * 2-B. 관리자 통합 시트 연동 (openById 직접 읽기)
 * ------------------------------------------------------
 * 프런트엔드(admin-script.html)의 linkGoogleSheet()가 호출합니다.
 * 관리자가 붙여넣은 시트 URL 또는 ID를 받아, 3종 데이터를 파싱해
 * 그대로 브라우저로 돌려줍니다. 브라우저는 이 결과로 기존 배정 알고리즘의
 * 입력 변수(MEAL_PARTICIPANTS / LECTURE_SEAT_GRID / ROOMMATE_PAIRS)를 채웁니다.
 *
 * 반환 형태:
 * {
 *   success: true/false,
 *   message: '...',
 *   spreadsheetId, spreadsheetUrl, spreadsheetName,
 *   participants: [{name, company, position, phone}, ...],
 *   lectureSeatGrid: [[name, name, ...], ...],
 *   roommatePairs: [{a, b}, ...],
 *   detected: { participantSheet, lectureSeatSheet, roommateSheet },
 *   warnings: ['...']   // 치명적이지 않지만 관리자가 확인해야 할 사항
 * }
 ****************************************************/

function readLinkedSheetData(sheetIdOrUrl, excludedSheetNames) {
  requireAdmin_();
  try {
    // [신규] 칩 토글로 사용자가 "이 시트는 쓰지 마세요"라고 지정한 시트 목록
    const excluded = {};
    (excludedSheetNames || []).forEach(function (n) {
      if (n) excluded[String(n)] = true;
    });

    const spreadsheetId = extractSpreadsheetId_(sheetIdOrUrl);
    if (!spreadsheetId) {
      return {
        success: false,
        message: '구글 시트 주소(URL) 또는 ID를 올바르게 입력해주세요.',
        warnings: [],
        diagnostics: {
          stage: 'invalidInput',
          fileOpened: false,
          allSheetNames: [],
          items: {}
        }
      };
    }

    let ss;
    try {
      ss = SpreadsheetApp.openById(spreadsheetId);
    } catch (openError) {
      return {
        success: false,
        message: '시트를 열 수 없습니다. 주소가 정확한지, 이 계정에 시트 열람 권한이 있는지 확인해주세요.',
        warnings: [],
        diagnostics: {
          stage: 'openFailed',
          fileOpened: false,
          allSheetNames: [],
          items: {}
        }
      };
    }

    const warnings = [];
    const detected = { participantSheet: '', lectureSeatSheet: '', roommateSheet: '' };

    // 이 파일에 실제로 어떤 시트들이 있는지 목록화 (진단 리포트용)
    // getSheets()는 Apps Script에서 비교적 비용이 있으므로 한 번만 호출해 재사용합니다.
    const sheets = ss.getSheets();
    const allSheetNames = sheets.map(function (s) { return s.getName(); });

    // 1) 교육생 명단
    const participantResult = readParticipantsFromLinkedSpreadsheet_(ss, excluded, sheets);
    detected.participantSheet = participantResult.sheetName;
    if (participantResult.warning) warnings.push(participantResult.warning);

    // 한 시트를 두 종류로 중복 판정하지 않도록, 이미 채택한 시트는 다음 탐색에서 제외합니다.
    const runtimeExcluded = {};
    Object.keys(excluded).forEach(function (name) { runtimeExcluded[name] = true; });
    if (participantResult.sheetName) runtimeExcluded[participantResult.sheetName] = true;

    // 2) 강의장 좌석 배치도
    const lectureResult = readLectureSeatGridFromLinkedSpreadsheet_(ss, runtimeExcluded, sheets, participantResult.participants);
    detected.lectureSeatSheet = lectureResult.sheetName;
    if (lectureResult.warning) warnings.push(lectureResult.warning);
    if (lectureResult.sheetName) runtimeExcluded[lectureResult.sheetName] = true;

    // 3) 룸메이트 명단
    const roommateResult = readRoommatePairsFromLinkedSpreadsheet_(ss, runtimeExcluded, sheets, participantResult.participants);
    detected.roommateSheet = roommateResult.sheetName;
    if (roommateResult.warning) warnings.push(roommateResult.warning);

    // 진단 리포트: 각 항목이 "어디까지 갔는지"를 상태 코드로 표현
    //  - ok:          정상 인식
    //  - sheetMissing: 해당 시트 자체를 못 찾음 → 이름 변경 or 양식 자동 생성 유도
    //  - emptyOrBad:   시트는 있으나 데이터/헤더를 못 읽음 → 데이터 채우기 유도
    // 시트를 못 찾은 항목은 "혹시 이 시트가 아닐까요?" 추측(guess)도 함께 담습니다.
    const usedSheetNames = {};
    [participantResult.sheetName, lectureResult.sheetName, roommateResult.sheetName].forEach(function (n) {
      if (n) usedSheetNames[n] = true;
    });
    // 사용자가 제외한 시트는 추측 후보에서도 제외
    Object.keys(excluded).forEach(function (n) { usedSheetNames[n] = true; });

    const participantStatus = participantResult.status || (participantResult.participants.length ? 'ok' : (participantResult.sheetName ? 'emptyOrBad' : 'sheetMissing'));
    const lectureStatus = lectureResult.status || (lectureResult.grid.length ? 'ok' : (lectureResult.sheetName ? 'emptyOrBad' : 'sheetMissing'));
    const roommateStatus = roommateResult.status || (roommateResult.pairs.length ? 'ok' : (roommateResult.sheetName ? 'emptyOrBad' : 'sheetMissing'));

    const diagnostics = {
      stage: 'read',
      fileOpened: true,
      allSheetNames: allSheetNames,
      expectedSheetNames: {
        participant: LINKED_SHEET_CONFIG.participantSheetNames[0],
        lectureSeat: LINKED_SHEET_CONFIG.lectureSeatSheetNames[0],
        roommate: LINKED_SHEET_CONFIG.roommateSheetNames[0]
      },
      items: {
        participant: {
          status: participantStatus,
          sheetName: participantResult.sheetName,
          sheetId: participantResult.sheetId || '',
          sheetUrl: participantResult.sheetUrl || '',
          count: participantResult.participants.length,
          message: participantResult.warning || '',
          guess: participantStatus === 'sheetMissing' ? guessSheetFor_(sheets, 'participant', usedSheetNames) : ''
        },
        lectureSeat: {
          status: lectureStatus,
          sheetName: lectureResult.sheetName,
          sheetId: lectureResult.sheetId || '',
          sheetUrl: lectureResult.sheetUrl || '',
          count: countGridNames_(lectureResult.grid),
          message: lectureResult.warning || '',
          guess: lectureStatus === 'sheetMissing' ? guessSheetFor_(sheets, 'lectureSeat', usedSheetNames) : ''
        },
        roommate: {
          status: roommateStatus,
          sheetName: roommateResult.sheetName,
          sheetId: roommateResult.sheetId || '',
          sheetUrl: roommateResult.sheetUrl || '',
          count: roommateResult.pairs.length,
          message: roommateResult.warning || '',
          guess: roommateStatus === 'sheetMissing' ? guessSheetFor_(sheets, 'roommate', usedSheetNames) : ''
        }
      }
    };

    return {
      success: true,
      message: '시트 연동에 성공했습니다.',
      spreadsheetId: spreadsheetId,
      spreadsheetUrl: ss.getUrl(),
      spreadsheetName: ss.getName(),
      participants: participantResult.participants,
      participantLayout: participantResult.layout || null,
      participantReviewRequired: Boolean(participantResult.layout && (participantResult.layout.inferred || (participantResult.layout.inferredFieldNames || []).length)),
      lectureSeatGrid: lectureResult.grid,
      roommatePairs: roommateResult.pairs,
      detected: detected,
      warnings: warnings,
      diagnostics: diagnostics
    };

  } catch (error) {
    return {
      success: false,
      message: error.message,
      warnings: [],
      diagnostics: { stage: 'error', fileOpened: false, allSheetNames: [], items: {} }
    };
  }
}

/**
 * [신규] 표준 양식 시트 자동 생성 (솔루션 C)
 * 관리자의 연동 시트 안에, 올바른 이름·헤더를 갖춘 빈 시트를 만들어줍니다.
 *  - which: 'participant' | 'lectureSeat' | 'roommate'
 * 안전장치: 같은 이름의 시트가 이미 있으면 절대 덮어쓰지 않고 그대로 둡니다.
 *           (데이터 손실 방지 — 양식이 어긋난 경우는 가이드로 안내)
 */
function createStandardSheet(spreadsheetId, which) {
  requireAdmin_();
  try {
    if (!spreadsheetId) throw new Error('연동된 시트를 찾을 수 없습니다.');
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const spec = getStandardSheetSpec_(which);
    const targetName = getUniqueSheetName_(ss, spec.sheetName);
    const sheet = ss.insertSheet(targetName, ss.getNumSheets());

    markGeneratedSheet_(sheet, spec.kind, spec.sheetName);

    if (spec.kind === 'lectureSeat') {
      buildLectureSeatTemplate_(sheet);
    } else if (spec.kind === 'roommate') {
      buildRoommateTemplate_(sheet);
    } else if (spec.kind === 'participant') {
      buildParticipantTemplate_(sheet);
    }

    const sheetUrl = buildSheetUrl_(ss, sheet);
    return {
      success: true,
      message: '선택하신 구글 시트 "' + ss.getName() + '" 내에 "' + targetName + '"라는 이름의 시트를 생성하였습니다. 양식에 맞춰 내용을 채운 후 저장하시고, 해당 페이지로 돌아와주시면, 실시간으로 내용이 반영됩니다. 예시는 지우고 작성해주세요.',
      sheetName: targetName,
      sheetId: sheet.getSheetId(),
      spreadsheetUrl: ss.getUrl(),
      sheetUrl: sheetUrl
    };

  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getStandardSheetSpec_(which) {
  if (which === 'participant') {
    return { kind: 'participant', sheetName: 'AI 생성_교육생 명단' };
  }
  if (which === 'lectureSeat') {
    return { kind: 'lectureSeat', sheetName: 'AI 생성_자리배치도' };
  }
  if (which === 'roommate') {
    return { kind: 'roommate', sheetName: 'AI 생성_룸메이트 명단' };
  }
  throw new Error('올바르지 않은 양식 종류입니다.');
}

function getUniqueSheetName_(ss, baseName) {
  let name = baseName;
  let index = 2;
  while (ss.getSheetByName(name)) {
    name = baseName + '_' + index;
    index++;
  }
  return name;
}

function buildSheetUrl_(ss, sheet) {
  if (!ss || !sheet) return '';
  return ss.getUrl() + '#gid=' + sheet.getSheetId();
}

function markGeneratedSheet_(sheet, kind, originalName) {
  try {
    sheet.addDeveloperMetadata('VUP_SHEET_KIND', kind);
    sheet.addDeveloperMetadata('VUP_GENERATED_SHEET_NAME', originalName || sheet.getName());
  } catch (e) {
    // Developer metadata가 제한되는 환경에서도 양식 생성 자체는 계속 진행합니다.
  }
}

function markLinkedSheetAsKind(spreadsheetId, sheetName, which) {
  requireAdmin_();
  const map = {
    participant: 'participant',
    lectureSeat: 'lectureSeat',
    roommate: 'roommate'
  };
  const kind = map[which];
  if (!kind) {
    return { success: false, message: '어떤 정보로 연결할지 확인하지 못했습니다.' };
  }
  try {
    const id = extractSpreadsheetId_(spreadsheetId);
    const ss = SpreadsheetApp.openById(id);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, message: '"' + sheetName + '" 시트를 찾지 못했습니다.' };
    }
    markGeneratedSheet_(sheet, kind, sheet.getName());
    return {
      success: true,
      message: '"' + sheet.getName() + '" 시트를 ' + getSheetTemplateInfo_(which).sheetName.replace('AI 생성_', '') + ' 항목으로 연결했습니다.',
      spreadsheetId: id,
      sheetName: sheet.getName(),
      sheetId: sheet.getSheetId(),
      sheetUrl: buildSheetUrl_(ss, sheet)
    };
  } catch (e) {
    return { success: false, message: e && e.message ? e.message : String(e) };
  }
}

function getAvailableSheets_(ssOrSheets, excluded) {
  const skip = excluded || {};
  const sheets = Array.isArray(ssOrSheets) ? ssOrSheets : ssOrSheets.getSheets();
  return sheets.filter(function (s) { return !skip[s.getName()]; });
}

function findGeneratedSheetByKind_(ssOrSheets, kind, excluded) {
  const sheets = getAvailableSheets_(ssOrSheets, excluded);
  for (let i = 0; i < sheets.length; i++) {
    const metadata = sheets[i].getDeveloperMetadata ? sheets[i].getDeveloperMetadata() : [];
    for (let j = 0; j < metadata.length; j++) {
      if (metadata[j].getKey && metadata[j].getValue && metadata[j].getKey() === 'VUP_SHEET_KIND' && metadata[j].getValue() === kind) {
        return sheets[i];
      }
    }
  }
  return null;
}

function buildParticipantTemplate_(sheet) {
  sheet.clear();
  sheet.getRange(1, 1, 1, 4).setValues([['이름', '회사명', '직급', '휴대폰 번호']]);
  sheet.getRange(2, 1, 2, 4).setValues([
    ['예) 홍길동', '예) 브이업벤처스', '예) 팀장', '예) 010-1234-5678'],
    ['예) 고길동', '예) 케이브이씨에이', '예) 책임', '예) 010-0000-0000']
  ]);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#C6EFCE').setFontColor('#000000');
  sheet.getRange(2, 1, 2, 4).setFontColor('#94A3B8');
  sheet.getRange(1, 1, 3, 4).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  [120, 180, 120, 160].forEach(function (width, idx) { sheet.setColumnWidth(idx + 1, width); });
  sheet.setFrozenRows(1);
}

function buildLectureSeatTemplate_(sheet) {
  sheet.clear();
  try { sheet.setHiddenGridlines(true); } catch (e) {}

  sheet.getRange('C2:N4').merge().setValue('STAGE').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold').setBackground('#F3F4F6');
  sheet.getRange('A5:B5').merge().setValue('창가').setHorizontalAlignment('center').setBackground('#E5E7EB');
  sheet.getRange('A6:C6').merge().setValue('※ 성명은 가나다순 (가로) 입니다. 예시는 지우고 작성해주세요.').setFontColor('#64748B');

  const numbers = [[]];
  for (let i = 1; i <= 12; i++) numbers[0].push(i);
  sheet.getRange(8, 3, 1, 12).setValues(numbers).setHorizontalAlignment('center').setFontWeight('bold');

  const columns = ['1열', '2열', '3열', '4열', '5열', '6열'];
  for (let i = 0; i < columns.length; i++) {
    const col = 3 + i * 2;
    sheet.getRange(9, col, 1, 2).merge().setValue(columns[i]).setHorizontalAlignment('center').setFontWeight('bold').setBackground('#E0F2FE');
  }

  const rowLabels = [['1줄'], ['2줄'], ['3줄'], ['4줄'], ['5줄']];
  sheet.getRange(10, 1, rowLabels.length, 1).setValues(rowLabels).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(10, 3, 3, 4).setValues([
    ['예) 홍길동', '예) 고길동', '', ''],
    ['예) 김철수', '', '예) 이영희', ''],
    ['', '', '', '']
  ]).setFontColor('#94A3B8');

  sheet.getRange('C8:N9').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(10, 3, 20, 12).setBorder(true, true, true, true, true, true, '#D1D5DB', SpreadsheetApp.BorderStyle.SOLID);
  for (let c = 1; c <= 14; c++) sheet.setColumnWidth(c, c <= 2 ? 56 : 90);
  for (let r = 10; r <= 29; r++) sheet.setRowHeight(r, 32);
}

function buildRoommateTemplate_(sheet) {
  sheet.clear();
  try { sheet.setHiddenGridlines(true); } catch (e) {}

  sheet.getRange('A1:A12').merge().setValue('1인실\n(4박)').setVerticalAlignment('middle').setHorizontalAlignment('center').setFontWeight('bold').setBackground('#E5E7EB');
  sheet.getRange('G1:G12').merge().setValue('트윈\n(예시)').setVerticalAlignment('middle').setHorizontalAlignment('center').setFontWeight('bold').setBackground('#E5E7EB');
  sheet.getRange('H1:H12').merge().setValue('2인실\n(4박)').setVerticalAlignment('middle').setHorizontalAlignment('center').setFontWeight('bold').setBackground('#E5E7EB');

  sheet.getRange(1, 2, 1, 4).setValues([['NO', '객실', '성명', '성별']]).setFontWeight('bold').setBackground('#C6EFCE');
  sheet.getRange(1, 9, 1, 4).setValues([['NO', '객실', '룸메이트', '성별']]).setFontWeight('bold').setBackground('#C6EFCE');
  sheet.getRange(2, 2, 3, 4).setValues([
    ['1', '1001', '예) 홍길동', 'M'],
    ['2', '1002', '예) 고길동', 'M'],
    ['3', '1003', '', '']
  ]).setFontColor('#94A3B8');
  sheet.getRange(2, 9, 3, 4).setValues([
    ['1', '1012', '예) 홍길동 / 고길동', 'M'],
    ['2', '1013', '예) 김철수 / 이영희', 'F'],
    ['3', '1014', '', '']
  ]).setFontColor('#94A3B8');

  sheet.getRange(1, 1, 12, 12).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  [70, 48, 80, 150, 48, 48, 70, 70, 48, 80, 190, 48].forEach(function (width, idx) { sheet.setColumnWidth(idx + 1, width); });
  for (let r = 1; r <= 12; r++) sheet.setRowHeight(r, 30);
}

/**
 * [신규] 식사 자리배치 전용 새 구글 시트(스프레드시트 파일) 생성
 * 관리자가 기존 통합 시트 없이 시작할 때: 교육생 명단 / 자리 배치도 / 객실배치
 * 3개 시트가 표준 양식으로 갖춰진 새 파일을 만들어 바로 연동합니다.
 */
function createDedicatedMealSpreadsheet_(siteId, title) {
  const ss = SpreadsheetApp.create(title);

  // 1) 내가 의도적으로 만들 "표준 시트" 목록
  const standardSheetNames = ['교육생 명단', '식사 자리 배치', '객실 배정'];

  // 2) 생성 직후 존재하는 기본 시트들의 ID를 먼저 기록해 둔다 (이름이 아니라 ID로 식별)
  const defaultSheetIds = ss.getSheets().map(function (s) { return s.getSheetId(); });

  // 3) 표준 시트들을 생성
  standardSheetNames.forEach(function (name) {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  // 4) 처음부터 있던 기본 빈 시트(ID 기준)만 삭제. 단, 전체 시트는 최소 1장 남긴다.
  ss.getSheets().forEach(function (s) {
    const isDefault = defaultSheetIds.indexOf(s.getSheetId()) !== -1;
    const isStandard = standardSheetNames.indexOf(s.getName()) !== -1;
    if (isDefault && !isStandard && ss.getSheets().length > 1) {
      ss.deleteSheet(s);
    }
  });

  return ss;
}

/**
 * [신규] "혹시 이 시트가 아닐까요?" 추측 (솔루션: 관리자 부담 최소화)
 * 시트를 이름으로 못 찾았을 때, 내용/이름 단서로 가장 유력한 시트 하나를 추측합니다.
 *  - participant: 상위 5행 안에 "이름" 헤더 + (회사/소속 또는 전화/연락처) 헤더가 함께 있는 시트
 *  - lectureSeat: 이름에 좌석/자리/강의 단서가 있는 시트 (식사/객실 관련은 제외)
 *  - roommate:   이름에 객실/룸메/숙소/방 단서가 있는 시트
 * 추측이 확실치 않으면 빈 문자열을 반환합니다 (엉뚱한 제안 방지).
 */
function guessSheetFor_(ssOrSheets, which, usedSheetNames) {
  const sheets = getAvailableSheets_(ssOrSheets).filter(function (s) {
    return !usedSheetNames[s.getName()];
  });

  if (which === 'participant') {
    const cand = LINKED_SHEET_CONFIG.participantHeaderCandidates;
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      const lastCol = Math.min(sheet.getLastColumn(), 40);
      const lastRow = Math.min(sheet.getLastRow(), 5);
      if (lastCol < 1 || lastRow < 1) continue;
      const rows = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      for (let r = 0; r < rows.length; r++) {
        const norm = rows[r].map(function (v) { return String(v == null ? '' : v).replace(/\s+/g, '').trim(); });
        const hasName = findHeaderColumn_(norm, cand.name) !== -1;
        const hasSecond =
          findHeaderColumn_(norm, cand.company) !== -1 ||
          findHeaderColumn_(norm, cand.phone) !== -1;
        if (hasName && hasSecond) return sheet.getName();
      }
    }
    return '';
  }

  // 이름 키워드 기반 추측 (좌석/룸메이트)
  const keywordMap = {
    lectureSeat: { include: ['좌석', '자리', '강의장', '강의실'], exclude: ['식사', '객실', '버스', '숙소'] },
    roommate: { include: ['객실', '룸메', '숙소', '방배정', '호실'], exclude: ['식사', '자리', '좌석', '강의'] }
  };
  const rule = keywordMap[which];
  if (!rule) return '';

  for (let i = 0; i < sheets.length; i++) {
    const name = sheets[i].getName();
    const hit = rule.include.some(function (kw) { return name.indexOf(kw) !== -1; });
    const blocked = rule.exclude.some(function (kw) { return name.indexOf(kw) !== -1; });
    if (hit && !blocked) return name;
  }
  return '';
}

/**
 * [신규] 추측 시트를 표준 이름으로 자동 변경 (원클릭 채택)
 * 관리자가 "예, 이 시트가 맞아요"를 누르면, 해당 시트 이름을 표준 이름으로 바꿉니다.
 * 이후 프런트엔드가 자동으로 "다시 불러오기"를 실행해 연동을 완성합니다.
 * 안전장치: 표준 이름의 시트가 이미 있으면 중단(데이터 보호).
 */
function adoptSheetAs(spreadsheetId, sourceSheetName, which) {
  requireAdmin_();
  try {
    if (!spreadsheetId) throw new Error('연동된 시트를 찾을 수 없습니다.');
    const ss = SpreadsheetApp.openById(spreadsheetId);

    let targetName = '';
    if (which === 'participant') targetName = LINKED_SHEET_CONFIG.participantSheetNames[0];
    else if (which === 'lectureSeat') targetName = LINKED_SHEET_CONFIG.lectureSeatSheetNames[0];
    else if (which === 'roommate') targetName = LINKED_SHEET_CONFIG.roommateSheetNames[0];
    else throw new Error('올바르지 않은 양식 종류입니다.');

    const source = ss.getSheetByName(sourceSheetName);
    if (!source) {
      return { success: false, message: '"' + sourceSheetName + '" 시트를 찾을 수 없습니다. "다시 불러오기" 후 다시 시도해주세요.' };
    }
    const existing = ss.getSheetByName(targetName);
    if (existing) {
      return { success: false, message: '"' + targetName + '" 시트가 이미 있어 이름을 바꾸지 않았습니다. 두 시트 중 어느 쪽이 맞는지 확인해주세요.' };
    }

    source.setName(targetName);
    return {
      success: true,
      message: '"' + sourceSheetName + '" 시트의 이름을 "' + targetName + '"(으)로 바꿨습니다.',
      sheetName: targetName
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * URL 또는 ID 문자열에서 스프레드시트 ID만 안전하게 추출합니다.
 * - https://docs.google.com/spreadsheets/d/<ID>/edit#gid=0
 * - https://drive.google.com/open?id=<ID>
 * - 순수 ID 문자열
 */
function extractSpreadsheetId_(input) {
  const raw = String(input || '').trim();
  if (!raw) return '';

  // /spreadsheets/d/<ID>/ 패턴
  const dMatch = raw.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (dMatch && dMatch[1]) return dMatch[1];

  // ?id=<ID> 또는 &id=<ID> 패턴
  const idMatch = raw.match(/[?&]id=([a-zA-Z0-9-_]+)/);
  if (idMatch && idMatch[1]) return idMatch[1];

  // /d/<ID>/ 패턴(일반 드라이브 파일 링크)
  const genericDMatch = raw.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (genericDMatch && genericDMatch[1]) return genericDMatch[1];

  // 순수 ID로 보이는 경우(구글 파일 ID는 보통 25자 이상의 영숫자/하이픈/언더스코어)
  if (/^[a-zA-Z0-9-_]{20,}$/.test(raw)) return raw;

  return '';
}

/**
 * 후보 이름 목록으로 시트를 찾습니다. 정확히 일치하는 시트를 우선하고,
 * 없으면 "포함" 매칭으로 한 번 더 시도합니다.
 */
function findSheetByNameCandidates_(ssOrSheets, candidates, excluded) {
  const sheets = getAvailableSheets_(ssOrSheets, excluded);
  const normalize = function (s) { return String(s || '').replace(/\s+/g, '').toLowerCase(); };

  // 1차: 정확 일치(공백 무시)
  for (let i = 0; i < candidates.length; i++) {
    const target = normalize(candidates[i]);
    for (let j = 0; j < sheets.length; j++) {
      if (normalize(sheets[j].getName()) === target) return sheets[j];
    }
  }
  // 2차: 포함 매칭
  for (let i = 0; i < candidates.length; i++) {
    const target = normalize(candidates[i]);
    if (!target) continue;
    for (let j = 0; j < sheets.length; j++) {
      const name = normalize(sheets[j].getName());
      if (name.indexOf(target) !== -1 || target.indexOf(name) !== -1) return sheets[j];
    }
  }
  return null;
}

/**
 * 교육생 명단 파싱: 헤더 행에서 이름/회사명/직급/휴대폰번호 컬럼을
 * 이름으로 찾아, 데이터 행을 뽑아냅니다.
 */
function readParticipantsFromLinkedSpreadsheet_(ss, excluded, sheets) {
  const sourceSheets = sheets || ss;
  const sheet = findSheetByNameCandidates_(sourceSheets, LINKED_SHEET_CONFIG.participantSheetNames, excluded) ||
    findGeneratedSheetByKind_(sourceSheets, 'participant', excluded) ||
    findBestParticipantSheetByContent_(sourceSheets, excluded);
  if (!sheet) {
    return {
      sheetName: '',
      participants: [],
      warning: '교육생 명단 시트를 찾지 못했습니다. 시트 이름을 "교육생_raw data" 등으로 확인해주세요.'
    };
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return { sheetName: sheet.getName(), sheetId: sheet.getSheetId(), sheetUrl: buildSheetUrl_(ss, sheet), participants: [], warning: '교육생 명단 시트가 비어 있습니다.' };
  }

  const allValues = getBoundedSheetValues_(sheet, 1000, 40);

  // 헤더 위치가 흔들리거나 아예 없더라도 세로로 밀집된 이름 열을 찾아냅니다.
  const layout = detectParticipantLayout_(allValues);
  if (!layout || layout.columns.name === -1) {
    return {
      sheetName: sheet.getName(),
      sheetId: sheet.getSheetId(),
      sheetUrl: buildSheetUrl_(ss, sheet),
      participants: [],
      status: 'headerMissing',
      warning: '교육생 명단 시트("' + sheet.getName() + '")에서 이름 열을 신뢰할 수 있게 판별하지 못했습니다. 이름 열 근처에 "이름/성명/Name" 중 하나를 적어주면 가장 정확합니다.'
    };
  }

  const participants = [];
  for (let r = layout.dataStartRow; r < allValues.length; r++) {
    const row = allValues[r];
    const name = String(row[layout.columns.name] == null ? '' : row[layout.columns.name]).trim();
    if (!name || /^예\)/.test(name)) continue;
    // 일반 문장·날짜·전화번호·합계 행이 이름으로 섞이지 않도록 한 번 더 확인합니다.
    if (!isLikelyPersonNameText_(name)) continue;
    if (isParticipantHeaderText_(name)) continue;

    participants.push({
      name: name,
      company: layout.columns.company !== -1 ? String(row[layout.columns.company] == null ? '' : row[layout.columns.company]).trim() : '',
      position: layout.columns.position !== -1 ? String(row[layout.columns.position] == null ? '' : row[layout.columns.position]).trim() : '',
      phone: layout.columns.phone !== -1 ? String(row[layout.columns.phone] == null ? '' : row[layout.columns.phone]).trim() : '',
      gender: layout.columns.gender !== -1 ? String(row[layout.columns.gender] == null ? '' : row[layout.columns.gender]).trim() : '',
      sourceRow: r + 1
    });
  }

  enrichParticipantIdentityMetadata_(participants);

  const warnings = [];
  if (!participants.length) {
    warnings.push('교육생 명단 시트("' + sheet.getName() + '")에서 이름이 입력된 데이터를 찾지 못했습니다. 데이터가 채워져 있는지 확인해주세요.');
  }
  if (layout.columns.phone === -1) {
    warnings.push('교육생 명단에서 "휴대폰번호" 컬럼을 찾지 못했습니다. 교육생 조회 기능이 제한될 수 있습니다.');
  }
  if (layout.inferred) {
    if (layout.nameColumnConfidence === 'review') {
      warnings.push('"' + sheet.getName() + '" 시트에는 이름 헤더가 없어 이름 열 후보가 둘 이상입니다. 연결 결과 미리보기에서 이름 열을 반드시 확인해주세요.');
    } else {
      warnings.push('"' + sheet.getName() + '" 시트에서 이름 열을 내용 패턴으로 자동 판별했습니다. 연결 결과 미리보기에서 이름·소속·직급 열이 맞는지 확인해주세요.');
    }
  }
  if (layout.inferredFieldNames && layout.inferredFieldNames.length) {
    const fieldLabels = { company: '소속', position: '직급', phone: '전화번호', gender: '성별' };
    warnings.push('표준 헤더가 없는 ' + layout.inferredFieldNames.map(function (key) {
      return fieldLabels[key] || key;
    }).join('·') + ' 열을 내용 패턴으로 추정했습니다. 연결 결과 미리보기에서 확인해주세요.');
  }
  const inferredCoreNames = participants.filter(function (person) { return person.baseNameInferred; });
  if (inferredCoreNames.length) {
    const examples = inferredCoreNames.slice(0, 4).map(function (person) {
      return '"' + person.displayName + '" → "' + person.baseName + '"';
    });
    warnings.push('한글 수식어가 붙은 이름을 공통 핵심 이름으로 추정했습니다(' + examples.join(', ') + '). 동명이인 확인 화면에서 반드시 확인해주세요.');
  }
  const possibleShortCorePairs = getPossibleShortKoreanCoreNamePairs_(participants);
  if (possibleShortCorePairs.length) {
    const examples = possibleShortCorePairs.slice(0, 3).map(function (pair) {
      return '"' + pair.left + '" / "' + pair.right + '" → "' + pair.coreName + '"';
    });
    warnings.push('두 글자 이름 앞에 임의 수식어가 붙었을 가능성이 있습니다(' + examples.join(', ') + '). 실제 동명이인이면 명단을 "' + possibleShortCorePairs[0].coreName + ' A/B"처럼 고쳐 구분해주세요.');
  }

  return {
    sheetName: sheet.getName(),
    sheetId: sheet.getSheetId(),
    sheetUrl: buildSheetUrl_(ss, sheet),
    participants: participants,
    layout: layout,
    warning: warnings.join(' ')
  };
}

/****************************************************
 * 이름·헤더 판별, 교육생 명단 레이아웃 자동 인식,
 *                        동명이인 처리, 명단 기준 이름 매칭
 ****************************************************/

/**
 * 정규화된 헤더 배열에서 후보 목록과 정확히 일치하는 첫 컬럼 인덱스를 찾습니다.
 */
function findHeaderColumn_(normalizedHeaderRow, candidates) {
  for (let c = 0; c < normalizedHeaderRow.length; c++) {
    const cell = normalizeHeaderToken_(normalizedHeaderRow[c]);
    if (!cell) continue;
    for (let i = 0; i < candidates.length; i++) {
      if (cell === normalizeHeaderToken_(candidates[i])) return c;
    }
  }
  return -1;
}

function normalizeHeaderToken_(value) {
  return String(value == null ? '' : value)
    .replace(/\s+/g, '')
    .replace(/[._\-/:()[\]{}]/g, '')
    .trim()
    .toLowerCase();
}

function isParticipantHeaderText_(value) {
  const token = normalizeHeaderToken_(value);
  const cand = LINKED_SHEET_CONFIG.participantHeaderCandidates;
  return Object.keys(cand).some(function (key) {
    return cand[key].some(function (header) { return normalizeHeaderToken_(header) === token; });
  });
}

function normalizePersonText_(value) {
  let text = String(value == null ? '' : value).trim();
  try { text = text.normalize('NFKC'); } catch (ignore) {}
  return text.replace(/[\u200B-\u200D\uFEFF]/g, '').replace(/\s+/g, ' ').trim();
}

/**
 * 표시명과 조회용 순수 이름을 분리합니다.
 * 홍길동 A / 홍길동1 / 작은 홍길동 / 큰홍길동처럼 관리자가 붙인 표식을
 * 제거하되, 원본 표시명은 절대 덮어쓰지 않습니다.
 */
function analyzePersonName_(value) {
  const displayName = normalizePersonText_(value).replace(/^예\)\s*/, '');
  let working = displayName;
  let discriminator = '';

  const prefixRules = [
    { re: /^(작은|막내|어린)[\s._\-:]*/i, code: 'small' },
    { re: /^(큰|맏|나이많은)[\s._\-:]*/i, code: 'big' },
    { re: /^(첫째|첫번째)[\s._\-:]*/i, code: '1' },
    { re: /^(둘째|두번째)[\s._\-:]*/i, code: '2' },
    { re: /^(셋째|세번째)[\s._\-:]*/i, code: '3' }
  ];
  for (let i = 0; i < prefixRules.length; i++) {
    if (prefixRules[i].re.test(working)) {
      discriminator = prefixRules[i].code;
      working = working.replace(prefixRules[i].re, '');
      break;
    }
  }

  const bracketSuffix = working.match(/[\s._\-]*[([{]\s*([^()[\]{}]{1,20})\s*[)\]}]\s*$/i);
  if (bracketSuffix) {
    discriminator = normalizeDiscriminator_(bracketSuffix[1]);
    working = working.slice(0, bracketSuffix.index).trim();
  } else {
    const separatedSuffix = working.match(/^(.*?)[\s._\-]+([A-Za-z]|\d{1,2})\s*$/i);
    const attachedKoreanSuffix = working.match(/^([가-힣]{2,7})([A-Za-z]|\d{1,2})$/i);
    if (separatedSuffix && separatedSuffix[1].trim()) {
      discriminator = normalizeDiscriminator_(separatedSuffix[2]);
      working = separatedSuffix[1].trim();
    } else if (attachedKoreanSuffix) {
      discriminator = normalizeDiscriminator_(attachedKoreanSuffix[2]);
      working = attachedKoreanSuffix[1];
    }
  }

  // "키큰 홍길동", "오전반 남궁민수"처럼 띄어 쓴 임의 수식어는
  // 마지막 토큰이 신뢰도 높은 한글 이름일 때만 보수적으로 분리합니다.
  const separatedKoreanParts = working.split(/[\s._\-:]+/).filter(Boolean);
  if (separatedKoreanParts.length > 1) {
    const lastPart = separatedKoreanParts[separatedKoreanParts.length - 1];
    if (/^[가-힣]{2,5}$/.test(lastPart) && getPersonNameConfidence_(lastPart) >= 0.8) {
      if (!discriminator) discriminator = separatedKoreanParts.slice(0, -1).join('');
      working = lastPart;
    }
  }

  working = working
    .replace(/^(성명|이름|교육생|참가자)\s*[:：-]\s*/i, '')
    .replace(/\s*(님|선생님)\s*$/i, '')
    .replace(/^[\s._\-:]+|[\s._\-:]+$/g, '')
    .trim();

  const koreanOnly = working.replace(/[^가-힣]/g, '');
  let baseName = working;
  if (/^[가-힣]{2,7}$/.test(koreanOnly) && koreanOnly.length >= 2) {
    baseName = koreanOnly;
  } else {
    baseName = working.replace(/\s+/g, ' ').trim();
  }

  return {
    displayName: displayName,
    baseName: baseName,
    nameKey: normalizePersonNameKey_(displayName),
    baseNameKey: normalizePersonNameKey_(baseName),
    discriminator: discriminator
  };
}

function normalizeDiscriminator_(value) {
  const text = String(value || '').trim().toUpperCase();
  const wordMap = { '첫째': '1', '둘째': '2', '셋째': '3' };
  if (wordMap[text]) return wordMap[text];
  if (/^[A-Z]$/.test(text)) return String(text.charCodeAt(0) - 64);
  return text.replace(/\D/g, '') || text;
}

function normalizePersonNameKey_(value) {
  return normalizePersonText_(value)
    .replace(/\s+/g, '')
    .replace(/[._\-:'"`~!@#$%^&*+=,/?\\|()[\]{}<>]/g, '')
    .toLowerCase();
}

function isLikelyPersonNameText_(value) {
  const text = normalizePersonText_(value).replace(/^예\)\s*/, '');
  if (!text || text.length > 60 || isParticipantHeaderText_(text)) return false;
  if (/https?:\/\/|www\.|@/.test(text)) return false;
  if (/^\d+(?:[.\-/]\d+)*$/.test(text)) return false;
  if (String(text).replace(/\D/g, '').length >= 7) return false;
  if (/^(남|여|남자|여자|male|female|m|f)$/i.test(text)) return false;
  if (/^(stage|window|door|entrance|exit|teacher|staff|total|sum|무대|창가|창문|교탁|통로|출입구|좌석|배치도|강의장|객실|룸메이트|구분|합계|총계|총원|인원|비고|담당자)$/i.test(text)) return false;
  if (/(주식회사|유한회사|대학교|협회|재단|센터|본부|부서)$/.test(text)) return false;

  const identity = analyzePersonName_(text);
  if (/^[가-힣]{2,7}$/.test(identity.baseName)) return true;
  if (/^[A-Za-z][A-Za-z .'-]{1,49}$/.test(identity.baseName)) {
    const words = identity.baseName.split(/[\s.-]+/).filter(Boolean);
    return words.length >= 1 && words.length <= 6;
  }
  return false;
}

/**
 * 헤더가 없는 표에서 회사명·부서명과 사람 이름을 구분하기 위한 보조 점수입니다.
 * 이 값만으로 이름을 확정하지 않고, 한 열 안의 고유값 비율·연속성·반복률과 함께 사용합니다.
 */
function getPersonNameConfidence_(value) {
  const text = normalizePersonText_(value).replace(/^예\)\s*/, '');
  if (!isLikelyPersonNameText_(text)) return 0;

  const identity = analyzePersonName_(text);
  const baseName = identity.baseName;
  if (/^[가-힣]{2,5}$/.test(baseName)) {
    const commonSingleSurnames = '김이박최정강조윤장임한오서신권황안송전홍유고문양손배백허남심노하곽성차주우구민진지엄채원천방공현함변염여추도소석선설마길연위표명기반왕금옥육인맹제모탁국어은편용';
    const commonDoubleSurnames = ['남궁', '황보', '제갈', '선우', '독고', '사공', '서문', '동방'];
    const startsWithKnownSurname = commonSingleSurnames.indexOf(baseName.charAt(0)) !== -1 ||
      commonDoubleSurnames.some(function (surname) { return baseName.indexOf(surname) === 0; });
    const organizationLike = /(주식회사|유한회사|전자|자동차|금융|증권|캐피탈|벤처스|파트너스|협회|재단|대학교|대학|연수원|연구소|센터|본부|부서|사업부|그룹|공사|공단|은행|보험|투자|기획|전략)$/i.test(baseName);
    const positionLike = /^(대표|대표이사|이사|상무|전무|부장|차장|과장|대리|사원|팀장|본부장|센터장|실장|매니저|파트너|심사역|연구원|위원)$/i.test(baseName);
    if (positionLike) return 0.05;
    if (organizationLike) return startsWithKnownSurname ? 0.30 : 0.12;
    if (startsWithKnownSurname) return 1;
    return baseName.length <= 4 ? 0.62 : 0.45;
  }

  if (/^[A-Za-z][A-Za-z .'-]{1,49}$/.test(baseName)) {
    const words = baseName.split(/[\s.-]+/).filter(Boolean);
    if (/^(front|screen|back|rear|vip|staff|stage|window|door|entrance|exit)$/i.test(baseName)) return 0.05;
    if (words.length >= 2) return 0.88;
    return /^[A-Z]{2,}$/.test(baseName) ? 0.20 : 0.45;
  }
  return 0;
}

function looksLikePhoneValue_(value) {
  const text = String(value == null ? '' : value).trim();
  const digits = text.replace(/\D/g, '');
  return digits.length >= 8 && digits.length <= 15;
}

function isLikelyPositionValue_(value) {
  const text = normalizePersonText_(value).replace(/\s+/g, '');
  if (!text || text.length > 24) return false;
  return /^(대표|대표이사|CEO|CFO|COO|이사|상무|전무|부장|차장|과장|대리|사원|팀장|본부장|센터장|실장|국장|원장|매니저|파트너|심사역|수석|책임|선임|주임|연구원|위원|교수|변호사|회계사|Manager|Director|Partner|Associate|Analyst|VP)$/i.test(text);
}

function getOrganizationValueConfidence_(value) {
  const text = normalizePersonText_(value);
  if (!text || text.length > 80 || /@|https?:\/\//i.test(text) || looksLikePhoneValue_(text)) return 0;
  if (/^\d+(?:[.\-/]\d+)*$/.test(text) || isLikelyPositionValue_(text)) return 0;
  if (/(주식회사|유한회사|전자|자동차|금융|증권|캐피탈|벤처스|파트너스|협회|재단|대학교|대학|연수원|연구소|센터|본부|공사|공단|은행|보험|투자|그룹|법인)$/i.test(text)) return 1;
  if (/(회사|기관|소속|기업|학교|병원|재단|협회)/i.test(text)) return 0.85;
  if (/^[가-힣A-Za-z0-9&().·\-\s]{2,40}$/.test(text)) return 0.30;
  return 0;
}

function detectParticipantLayout_(values) {
  if (!values || !values.length) return null;
  const cand = LINKED_SHEET_CONFIG.participantHeaderCandidates;
  const scanRows = Math.min(values.length, 30);
  let best = null;

  for (let r = 0; r < scanRows; r++) {
    const row = values[r] || [];
    const nameCol = findHeaderColumn_(row, cand.name);
    if (nameCol === -1) continue;
    const columns = {
      name: nameCol,
      company: findHeaderColumn_(row, cand.company),
      position: findHeaderColumn_(row, cand.position),
      phone: findHeaderColumn_(row, cand.phone),
      gender: findHeaderColumn_(row, cand.gender)
    };
    let fieldCount = 0;
    Object.keys(columns).forEach(function (key) { if (columns[key] !== -1) fieldCount++; });
    let namesBelow = 0;
    for (let rr = r + 1; rr < Math.min(values.length, r + 41); rr++) {
      if (isLikelyPersonNameText_((values[rr] || [])[nameCol])) namesBelow++;
    }
    const score = 30 + fieldCount * 8 + namesBelow * 2 - r * 0.2;
    if (!best || score > best.score) {
      best = { headerRowIndex: r, dataStartRow: r + 1, columns: columns, inferred: false, score: score };
    }
  }

  if (best) {
    // 이름 헤더만 표준이고 나머지 헤더가 "기관 정보", "담당", "연락 가능한 번호"처럼
    // 임의 문구여도 실제 데이터 패턴으로 소속·직급·전화·성별 열을 보완합니다.
    const inferredFields = inferParticipantFieldColumns_(values, best.columns.name, best.dataStartRow);
    best.inferredFieldNames = [];
    ['company', 'position', 'phone', 'gender'].forEach(function (key) {
      if (best.columns[key] !== -1 || inferredFields[key] === -1) return;
      best.columns[key] = inferredFields[key];
      best.inferredFieldNames.push(key);
    });
    return best;
  }

  const maxCols = Math.min(40, Math.max.apply(null, values.slice(0, scanRows).map(function (row) { return (row || []).length; }).concat([0])));
  let bestColumn = null;
  const columnCandidates = [];
  for (let c = 0; c < maxCols; c++) {
    const hits = [];
    const normalizedHitCounts = {};
    let confidenceTotal = 0;
    let nonEmpty = 0;
    let longestRun = 0;
    let currentRun = 0;
    for (let r = 0; r < values.length; r++) {
      const cell = (values[r] || [])[c];
      if (String(cell == null ? '' : cell).trim()) nonEmpty++;
      if (isLikelyPersonNameText_(cell)) {
        hits.push(r);
        const nameKey = analyzePersonName_(cell).baseNameKey;
        if (nameKey) normalizedHitCounts[nameKey] = (normalizedHitCounts[nameKey] || 0) + 1;
        confidenceTotal += getPersonNameConfidence_(cell);
        currentRun++;
        if (currentRun > longestRun) longestRun = currentRun;
      } else if (String(cell == null ? '' : cell).trim()) {
        currentRun = 0;
      }
    }
    if (hits.length < 2) continue;
    const density = hits.length / Math.max(1, nonEmpty);
    const nameKeys = Object.keys(normalizedHitCounts);
    const uniqueRatio = nameKeys.length / Math.max(1, hits.length);
    const dominantRatio = nameKeys.reduce(function (maxValue, key) {
      return Math.max(maxValue, normalizedHitCounts[key] / Math.max(1, hits.length));
    }, 0);
    const averageConfidence = confidenceTotal / Math.max(1, hits.length);
    const score = hits.length * 4 + density * 10 + longestRun * 2 +
      uniqueRatio * 12 + averageConfidence * 18 - dominantRatio * 10;
    const candidate = {
      col: c,
      hits: hits,
      score: score,
      uniqueRatio: uniqueRatio,
      dominantRatio: dominantRatio,
      averageConfidence: averageConfidence
    };
    columnCandidates.push(candidate);
    if (!bestColumn || score > bestColumn.score) {
      bestColumn = candidate;
    }
  }
  if (!bestColumn) return null;

  columnCandidates.sort(function (a, b) { return b.score - a.score; });
  const secondBest = columnCandidates.length > 1 ? columnCandidates[1] : null;
  const confidenceGap = secondBest ? bestColumn.score - secondBest.score : bestColumn.score;

  const dataStartRow = bestColumn.hits[0];
  const inferredColumns = inferParticipantFieldColumns_(values, bestColumn.col, dataStartRow);
  inferredColumns.name = bestColumn.col;
  return {
    headerRowIndex: -1,
    dataStartRow: dataStartRow,
    columns: inferredColumns,
    inferred: true,
    score: bestColumn.score,
    nameColumnConfidence: confidenceGap >= 8 && bestColumn.averageConfidence >= 0.55 ? 'high' : 'review',
    nameColumnCandidates: columnCandidates.slice(0, 3).map(function (candidate) {
      return {
        columnIndex: candidate.col,
        score: Math.round(candidate.score * 10) / 10,
        sampleRows: candidate.hits.slice(0, 5)
      };
    })
  };
}

function inferParticipantFieldColumns_(values, nameCol, dataStartRow) {
  const cand = LINKED_SHEET_CONFIG.participantHeaderCandidates;
  const columns = { name: nameCol, company: -1, position: -1, phone: -1, gender: -1 };
  const headerScanEnd = Math.min(values.length, Math.max(10, dataStartRow + 1));
  for (let r = 0; r < headerScanEnd; r++) {
    const row = values[r] || [];
    ['company', 'position', 'phone', 'gender'].forEach(function (key) {
      if (columns[key] === -1) columns[key] = findHeaderColumn_(row, cand[key]);
    });
  }

  const maxCols = Math.min(40, Math.max.apply(null, values.slice(0, Math.min(values.length, 50)).map(function (row) { return (row || []).length; }).concat([0])));
  const endRow = Math.min(values.length, dataStartRow + 80);
  if (columns.phone === -1) {
    let bestPhone = { col: -1, count: 0 };
    for (let c = 0; c < maxCols; c++) {
      if (c === nameCol) continue;
      let count = 0;
      for (let r = dataStartRow; r < endRow; r++) if (looksLikePhoneValue_((values[r] || [])[c])) count++;
      if (count > bestPhone.count) bestPhone = { col: c, count: count };
    }
    if (bestPhone.count >= 2) columns.phone = bestPhone.col;
  }

  if (columns.gender === -1) {
    let bestGender = { col: -1, count: 0 };
    for (let c = 0; c < maxCols; c++) {
      if (c === nameCol || c === columns.phone) continue;
      let count = 0;
      for (let r = dataStartRow; r < endRow; r++) {
        if (/^(남|여|남자|여자|m|f|male|female)$/i.test(String((values[r] || [])[c] || '').trim())) count++;
      }
      if (count > bestGender.count) bestGender = { col: c, count: count };
    }
    if (bestGender.count >= 2) columns.gender = bestGender.col;
  }

  if (columns.position === -1) {
    let bestPosition = { col: -1, count: 0 };
    for (let c = 0; c < maxCols; c++) {
      if (c === nameCol || c === columns.phone || c === columns.gender) continue;
      let count = 0;
      for (let r = dataStartRow; r < endRow; r++) {
        if (isLikelyPositionValue_((values[r] || [])[c])) count++;
      }
      if (count > bestPosition.count) bestPosition = { col: c, count: count };
    }
    const scannedRows = Math.max(1, endRow - dataStartRow);
    if (bestPosition.count >= 2 && bestPosition.count / scannedRows >= 0.25) columns.position = bestPosition.col;
  }

  if (columns.company === -1) {
    let bestCompany = null;
    for (let c = 0; c < maxCols; c++) {
      if (c === nameCol || c === columns.phone || c === columns.gender || c === columns.position) continue;
      let nonEmpty = 0;
      let confidenceTotal = 0;
      const distinct = {};
      for (let r = dataStartRow; r < endRow; r++) {
        const raw = (values[r] || [])[c];
        const text = normalizePersonText_(raw);
        if (!text || /^\d+(?:[.\-/]\d+)*$/.test(text) || /@/.test(text)) continue;
        nonEmpty++;
        confidenceTotal += getOrganizationValueConfidence_(text);
        distinct[normalizeHeaderToken_(text)] = true;
      }
      if (nonEmpty < 2) continue;
      const averageConfidence = confidenceTotal / nonEmpty;
      const distanceBonus = Math.max(0, 3 - Math.abs(c - nameCol));
      const repeatedBonus = nonEmpty > Object.keys(distinct).length ? 1.5 : 0;
      const score = averageConfidence * 10 + distanceBonus + repeatedBonus;
      const candidate = { col: c, score: score, averageConfidence: averageConfidence, nonEmpty: nonEmpty };
      if (!bestCompany || candidate.score > bestCompany.score) bestCompany = candidate;
    }
    if (bestCompany && bestCompany.averageConfidence >= 0.45) columns.company = bestCompany.col;
  }
  return columns;
}

function findBestParticipantSheetByContent_(ssOrSheets, excluded) {
  const sheets = Array.isArray(ssOrSheets) ? ssOrSheets : ssOrSheets.getSheets();
  let best = null;
  sheets.forEach(function (sheet) {
    if (excluded && excluded[sheet.getName()]) return;
    const values = getBoundedSheetValues_(sheet, 180, 40);
    const layout = detectParticipantLayout_(values);
    if (!layout) return;
    const compactName = sheet.getName().replace(/\s+/g, '').toLowerCase();
    let score = layout.score || 0;
    if (/(교육생|참가자|명단|raw|roster|student|participant)/i.test(compactName)) score += 25;
    if (/(좌석|자리|배치도|객실|룸메|숙소|seat|room)/i.test(compactName)) score -= 35;
    if (!best || score > best.score) best = { sheet: sheet, score: score };
  });
  return best && best.score >= 12 ? best.sheet : null;
}

function enrichParticipantIdentityMetadata_(participants) {
  const counts = {};
  (participants || []).forEach(function (person, index) {
    const identity = analyzePersonName_(person.name);
    person.displayName = identity.displayName;
    person.baseName = identity.baseName;
    person.nameKey = identity.nameKey;
    person.baseNameKey = identity.baseNameKey;
    person.discriminator = identity.discriminator;
    person.phoneLast4 = extractPhoneLast4_(person.phone);
    person.personKey = identity.baseNameKey + '|' + (person.phoneLast4 || 'row' + String(person.sourceRow || index + 1));
  });
  reconcileParticipantCoreNames_(participants || []);
  (participants || []).forEach(function (person) {
    counts[person.baseNameKey] = (counts[person.baseNameKey] || 0) + 1;
  });
  (participants || []).forEach(function (person) {
    person.isDuplicateName = !!person.baseNameKey && counts[person.baseNameKey] > 1;
    person.duplicateGroupKey = person.isDuplicateName ? person.baseNameKey : '';
  });
  return participants;
}

function getParticipantsForLookup_(siteId, spreadsheetId, forceRefresh) {
  const cacheKey = 'participantLookup:v3:' + String(siteId || '');
  if (!forceRefresh) {
    const cached = getUserDataCache_(cacheKey, false);
    if (Array.isArray(cached)) return cached;
  }

  const participants = getParticipantsFromLinkedSheet_(spreadsheetId);
  try {
    CacheService.getScriptCache().put(
      cacheKey,
      JSON.stringify(participants),
      PARTICIPANT_LOOKUP_CACHE_SECONDS
    );
  } catch (e) {
    // 대규모 명단이 CacheService 단일 값 한도를 넘으면 캐시 없이 원본 시트를 사용합니다.
  }
  return participants;
}

function getLongestLikelySharedKoreanNameSuffix_(leftValue, rightValue) {
  const left = String(leftValue || '').replace(/[^가-힣]/g, '');
  const right = String(rightValue || '').replace(/[^가-힣]/g, '');
  if (!left || !right || left === right) return '';
  const maxLength = Math.min(5, left.length, right.length);
  for (let length = maxLength; length >= 3; length--) {
    const leftSuffix = left.slice(-length);
    if (leftSuffix !== right.slice(-length)) continue;
    if (getPersonNameConfidence_(leftSuffix) >= 0.8) return leftSuffix;
  }
  if (left.length >= 2 && right.length >= 2 && left.slice(-2) === right.slice(-2)) {
    const suffix = left.slice(-2);
    const compoundSurnames = { '남궁': true, '황보': true, '제갈': true, '선우': true, '독고': true, '사공': true, '서문': true, '동방': true };
    const leftPrefix = left.slice(0, -2);
    const rightPrefix = right.slice(0, -2);
    const hasModifierLength = (left.length >= 4 && right.length >= 4) || left === suffix || right === suffix;
    if (hasModifierLength && !(compoundSurnames[leftPrefix] && compoundSurnames[rightPrefix]) && getPersonNameConfidence_(suffix) >= 0.8) return suffix;
  }
  return '';
}

function getPossibleShortKoreanCoreNamePairs_(participants) {
  const people = participants || [];
  const compoundSurnames = { '남궁': true, '황보': true, '제갈': true, '선우': true, '독고': true, '사공': true, '서문': true, '동방': true };
  const result = [];
  const seen = {};
  for (let i = 0; i < people.length; i++) {
    for (let j = i + 1; j < people.length; j++) {
      const left = String(people[i].baseName || '').replace(/[^가-힣]/g, '');
      const right = String(people[j].baseName || '').replace(/[^가-힣]/g, '');
      if (left.length < 4 || right.length < 4 || left === right) continue;
      const coreName = left.slice(-2);
      if (coreName !== right.slice(-2) || getPersonNameConfidence_(coreName) < 0.8) continue;
      const leftPrefix = left.slice(0, -2);
      const rightPrefix = right.slice(0, -2);
      // 남궁민수/황보민수처럼 알려진 복성 두 개가 우연히 같은 이름을 쓰는 경우는 자동 질문에서 제외합니다.
      if (compoundSurnames[leftPrefix] && compoundSurnames[rightPrefix]) continue;
      const key = [String(people[i].displayName || people[i].name), String(people[j].displayName || people[j].name)].sort().join('||');
      if (seen[key]) continue;
      seen[key] = true;
      result.push({
        left: String(people[i].displayName || people[i].name || ''),
        right: String(people[j].displayName || people[j].name || ''),
        coreName: coreName
      });
    }
  }
  return result;
}

function reconcileParticipantCoreNames_(participants) {
  const proposals = {};
  (participants || []).forEach(function (person, index) { proposals[index] = {}; });
  for (let i = 0; i < (participants || []).length; i++) {
    for (let j = i + 1; j < participants.length; j++) {
      const left = String(participants[i].baseName || '');
      const right = String(participants[j].baseName || '');
      const shared = getLongestLikelySharedKoreanNameSuffix_(left, right);
      if (!shared) continue;
      if (left !== shared) proposals[i][shared] = true;
      if (right !== shared) proposals[j][shared] = true;
    }
  }

  (participants || []).forEach(function (person, index) {
    const candidates = Object.keys(proposals[index] || {}).sort(function (a, b) { return b.length - a.length; });
    if (candidates.length !== 1) return;
    const inferredBaseName = candidates[0];
    if (!inferredBaseName || inferredBaseName === person.baseName) return;
    person.baseName = inferredBaseName;
    person.baseNameKey = normalizePersonNameKey_(inferredBaseName);
    person.baseNameInferred = true;
    person.baseNameInferenceReason = 'sharedKoreanSuffix';
    person.personKey = person.baseNameKey + '|' + (person.phoneLast4 || 'row' + String(person.sourceRow || index + 1));
  });
  return participants;
}

function buildParticipantNameReference_(participants) {
  const reference = { exact: {}, base: {}, people: participants || [] };
  (participants || []).forEach(function (person) {
    const identity = analyzePersonName_(person.name);
    const nameKey = person.nameKey || identity.nameKey;
    const baseKey = person.baseNameKey || identity.baseNameKey;
    if (!reference.exact[nameKey]) reference.exact[nameKey] = [];
    reference.exact[nameKey].push(person);
    if (!reference.base[baseKey]) reference.base[baseKey] = [];
    reference.base[baseKey].push(person);
  });
  return reference;
}

function resolveNameAgainstParticipantReference_(value, reference) {
  const identity = analyzePersonName_(value);
  if (!identity.nameKey) return { status: 'notFound', matches: [], identity: identity };
  const exact = (reference && reference.exact && reference.exact[identity.nameKey]) || [];
  if (exact.length === 1) return { status: 'exact', person: exact[0], matches: exact, identity: identity };
  if (exact.length > 1) return { status: 'ambiguous', matches: exact, identity: identity };

  const base = (reference && reference.base && reference.base[identity.baseNameKey]) || [];
  if (base.length === 1) return { status: 'base', person: base[0], matches: base, identity: identity };
  if (base.length > 1 && identity.discriminator) {
    const discriminatorMatches = base.filter(function (person) {
      return String(person.discriminator || analyzePersonName_(person.name).discriminator || '') === String(identity.discriminator);
    });
    if (discriminatorMatches.length === 1) {
      return { status: 'discriminator', person: discriminatorMatches[0], matches: discriminatorMatches, identity: identity };
    }
  }
  if (base.length > 1) return { status: 'ambiguous', matches: base, identity: identity };
  return { status: 'notFound', matches: [], identity: identity };
}

function getBoundedSheetValues_(sheet, maxRows, maxCols) {
  const rowCount = Math.min(Math.max(sheet.getLastRow(), 1), maxRows || 1000);
  const colCount = Math.min(Math.max(sheet.getLastColumn(), 1), maxCols || 40);
  return sheet.getRange(1, 1, rowCount, colCount).getValues();
}

/****************************************************
 * 강의장 좌석 배치도 · 룸메이트 명단 파싱
 ****************************************************/

function findBestSeatSheetByContent_(ssOrSheets, excluded, participants) {
  const sheets = Array.isArray(ssOrSheets) ? ssOrSheets : ssOrSheets.getSheets();
  const reference = buildParticipantNameReference_(participants);
  let best = null;
  sheets.forEach(function (sheet) {
    if (excluded && excluded[sheet.getName()]) return;
    const values = getBoundedSheetValues_(sheet, 160, 40);
    let matched = 0;
    let likely = 0;
    const rows = {};
    const cols = {};
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < (values[r] || []).length; c++) {
        const value = values[r][c];
        const resolved = resolveNameAgainstParticipantReference_(value, reference);
        if (resolved.status !== 'notFound') {
          matched++;
          rows[r] = true;
          cols[c] = true;
        } else if (isLikelyPersonNameText_(value)) {
          likely++;
        }
      }
    }
    const rowSpread = Object.keys(rows).length;
    const colSpread = Object.keys(cols).length;
    let score = matched * 5 + likely * 0.5 + rowSpread + colSpread * 4;
    const name = sheet.getName().replace(/\s+/g, '');
    if (/(좌석|자리|배치|강의|seat)/i.test(name)) score += 25;
    if (/(교육생|참가자|raw|명단|객실|룸메|숙소)/i.test(name)) score -= 25;
    if (colSpread < 2) score -= 20;
    if (!best || score > best.score) best = { sheet: sheet, score: score, matched: matched };
  });
  return best && best.score >= 12 && (best.matched >= 2 || !(participants || []).length) ? best.sheet : null;
}

/**
 * 강의장 좌석 배치도 파싱: anchor(C10)를 기준으로, 값이 있는 범위 전체를
 * 좌석 그리드로 읽습니다. 라벨(무대/열/줄 등)은 걸러냅니다.
 * anchor 위치에서 아무 값도 없으면, 시트 전체에서 이름 밀집 영역을 자동 탐색합니다.
 */
function readLectureSeatGridFromLinkedSpreadsheet_(ss, excluded, sheets, participants) {
  const sourceSheets = sheets || ss;
  const sheet = findSheetByNameCandidates_(sourceSheets, LINKED_SHEET_CONFIG.lectureSeatSheetNames, excluded) ||
    findGeneratedSheetByKind_(sourceSheets, 'lectureSeat', excluded) ||
    findBestSeatSheetByContent_(sourceSheets, excluded, participants);
  if (!sheet) {
    return {
      sheetName: '',
      grid: [],
      warning: '강의장 좌석 배치도 시트를 찾지 못했습니다. 시트 이름을 "자리 배치도" 등으로 확인해주세요.'
    };
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return { sheetName: sheet.getName(), sheetId: sheet.getSheetId(), sheetUrl: buildSheetUrl_(ss, sheet), grid: [], warning: '강의장 좌석 배치도 시트가 비어 있습니다.' };
  }

  const allValues = getBoundedSheetValues_(sheet, 160, 40);
  const skip = LINKED_SHEET_CONFIG.lectureSeatSkipKeywords;
  const nameReference = buildParticipantNameReference_(participants);

  const isSeatName = function (value) {
    const text = String(value == null ? '' : value).trim();
    if (!text || /^예\)/.test(text)) return false;
    for (let i = 0; i < skip.length; i++) {
      if (text === skip[i]) return false;
      // "1열", "2줄", "3열" 등 라벨 패턴 제거
    }
    if (/^\d+\s*(열|줄|번)?$/.test(text)) return false;
    if (/^[A-Za-z]$/.test(text)) return false; // 단일 알파벳(성별표시 F/M 등)
    const resolved = resolveNameAgainstParticipantReference_(text, nameReference);
    if (resolved.status !== 'notFound') return true;
    return isLikelyPersonNameText_(text);
  };

  // anchor 기준 스캔
  const anchorRow0 = LINKED_SHEET_CONFIG.lectureSeatAnchorRow - 1;
  const anchorCol0 = LINKED_SHEET_CONFIG.lectureSeatAnchorCol - 1;

  let grid = extractSeatGrid_(allValues, anchorRow0, anchorCol0, isSeatName);

  // 고정 anchor보다 자동 탐색 영역에서 더 많은 이름을 찾으면 자동 탐색 결과를 채택합니다.
  // 따라서 좌석표가 C10보다 위/왼쪽에서 시작해도 놓치지 않습니다.
  const auto = findSeatBlockAutomatically_(allValues, isSeatName, nameReference);
  if (auto && countGridNames_(auto) > countGridNames_(grid)) grid = auto;

  const nameCount = countGridNames_(grid);
  let warning = '';
  if (!nameCount) {
    warning = '강의장 좌석 배치도 시트("' + sheet.getName() + '")에서 좌석에 적힌 이름을 찾지 못했습니다. C10 부근부터 교육생 이름이 입력되어 있는지 확인해주세요.';
  }

  return { sheetName: sheet.getName(), sheetId: sheet.getSheetId(), sheetUrl: buildSheetUrl_(ss, sheet), grid: grid, warning: warning };
}

/**
 * 주어진 2차원 배열에서 (startRow0, startCol0)을 좌상단 기준으로,
 * 값이 있는 직사각형 범위를 추출합니다. 뒤쪽의 빈 행/열은 잘라냅니다.
 */
function extractSeatGrid_(allValues, startRow0, startCol0, isSeatName) {
  if (startRow0 < 0) startRow0 = 0;
  if (startCol0 < 0) startCol0 = 0;

  const rawGrid = [];
  for (let r = startRow0; r < allValues.length; r++) {
    const rowVals = [];
    for (let c = startCol0; c < allValues[r].length; c++) {
      const v = allValues[r][c];
      rowVals.push(isSeatName(v) ? String(v).trim() : '');
    }
    rawGrid.push(rowVals);
  }

  return trimGrid_(rawGrid);
}

/**
 * 뒤쪽의 완전히 빈 행과 열을 잘라냅니다.
 */
function trimGrid_(grid) {
  // 빈 행 제거(끝쪽) + 값이 하나도 없는 행은 통째로 제거하지 않고 유지(좌석 공백 표현 가능)
  let lastNonEmptyRow = -1;
  let lastNonEmptyCol = -1;
  for (let r = 0; r < grid.length; r++) {
    for (let c = 0; c < grid[r].length; c++) {
      if (grid[r][c]) {
        if (r > lastNonEmptyRow) lastNonEmptyRow = r;
        if (c > lastNonEmptyCol) lastNonEmptyCol = c;
      }
    }
  }
  if (lastNonEmptyRow === -1) return [];

  const trimmed = [];
  for (let r = 0; r <= lastNonEmptyRow; r++) {
    const row = [];
    for (let c = 0; c <= lastNonEmptyCol; c++) {
      row.push((grid[r] && grid[r][c]) || '');
    }
    trimmed.push(row);
  }
  // 앞쪽/중간의 완전 빈 행은 남겨두되, 전부 빈 행이 맨 위에 연속되면 제거
  while (trimmed.length && trimmed[0].every(function (v) { return !v; })) {
    trimmed.shift();
  }
  return trimmed;
}

function countGridNames_(grid) {
  let n = 0;
  for (let r = 0; r < grid.length; r++) {
    for (let c = 0; c < grid[r].length; c++) {
      if (grid[r][c]) n++;
    }
  }
  return n;
}

/**
 * anchor 위치가 어긋난 시트를 위해, 이름으로 판정되는 셀이 가장 많이 밀집한
 * 좌상단 지점을 찾아 그 지점부터 그리드를 추출합니다.
 */
function findSeatBlockAutomatically_(allValues, isSeatName, nameReference) {
  const cells = [];
  for (let r = 0; r < allValues.length; r++) {
    for (let c = 0; c < allValues[r].length; c++) {
      if (!isSeatName(allValues[r][c])) continue;
      const resolved = resolveNameAgainstParticipantReference_(allValues[r][c], nameReference);
      cells.push({ r: r, c: c, strong: resolved.status !== 'notFound' });
    }
  }
  if (!cells.length) return null;

  // 서로 가까운 이름 셀을 하나의 배치 영역으로 묶고, 기준 명단과 일치하는 이름이
  // 가장 많은 영역을 선택합니다. 시트 제목이나 담당자 이름이 표 밖에 있어도 섞이지 않습니다.
  const visited = {};
  let bestCluster = null;
  for (let i = 0; i < cells.length; i++) {
    if (visited[i]) continue;
    const queue = [i];
    const cluster = [];
    visited[i] = true;
    while (queue.length) {
      const index = queue.shift();
      const current = cells[index];
      cluster.push(current);
      for (let j = 0; j < cells.length; j++) {
        if (visited[j]) continue;
        const next = cells[j];
        if (Math.abs(current.r - next.r) <= 4 && Math.abs(current.c - next.c) <= 5) {
          visited[j] = true;
          queue.push(j);
        }
      }
    }
    const strongCount = cluster.filter(function (cell) { return cell.strong; }).length;
    const score = strongCount * 20 + cluster.length * 2;
    if (!bestCluster || score > bestCluster.score) bestCluster = { cells: cluster, score: score };
  }
  if (!bestCluster || !bestCluster.cells.length) return null;

  let minRow = allValues.length;
  let maxRow = 0;
  let minCol = 999;
  let maxCol = 0;
  bestCluster.cells.forEach(function (cell) {
    minRow = Math.min(minRow, cell.r);
    maxRow = Math.max(maxRow, cell.r);
    minCol = Math.min(minCol, cell.c);
    maxCol = Math.max(maxCol, cell.c);
  });

  const grid = [];
  for (let r = minRow; r <= maxRow; r++) {
    const row = [];
    for (let c = minCol; c <= maxCol; c++) {
      const value = (allValues[r] || [])[c];
      row.push(isSeatName(value) ? String(value).trim() : '');
    }
    grid.push(row);
  }
  return trimGrid_(grid);
}

/**
 * 룸메이트 명단 파싱: "이름/이름" 또는 "이름 / 이름" 패턴 셀을 스캔합니다.
 * 우선 지정된 열(K열)을 살펴보고, 충분히 못 찾으면 시트 전체를 스캔합니다.
 */
function readRoommatePairsFromLinkedSpreadsheet_(ss, excluded, sheets, participants) {
  const sourceSheets = sheets || ss;
  const sheet = findSheetByNameCandidates_(sourceSheets, LINKED_SHEET_CONFIG.roommateSheetNames, excluded) ||
    findGeneratedSheetByKind_(sourceSheets, 'roommate', excluded) ||
    findBestRoommateSheetByContent_(sourceSheets, excluded, participants);
  if (!sheet) {
    // 룸메이트는 선택 정보이므로, 시트가 없어도 치명적이지 않습니다.
    return {
      sheetName: '',
      pairs: [],
      warning: '룸메이트(객실배치) 시트를 찾지 못했습니다. 2인실 정보가 없다면 무시해도 됩니다.'
    };
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return { sheetName: sheet.getName(), sheetId: sheet.getSheetId(), sheetUrl: buildSheetUrl_(ss, sheet), pairs: [], warning: '' };
  }

  const allValues = getBoundedSheetValues_(sheet, 500, 40);
  const nameReference = buildParticipantNameReference_(participants);

  const parsePairCell = function (value) {
    return splitRoommatePairText_(value, nameReference);
  };

  const collectFromColumn = function (col0) {
    const found = [];
    for (let r = 0; r < allValues.length; r++) {
      if (col0 >= allValues[r].length) continue;
      const pair = parsePairCell(allValues[r][col0]);
      if (pair) found.push(pair);
    }
    return found;
  };

  // 룸메이트 명단은 보통 한 열 안에 여러 건이 모여 있습니다. 두 건 이상 모인 열은
  // 복수로 인정하되, 다른 열에 한 건만 외따로 끼어 있으면 자동 채택하지 않습니다.
  const columnGroups = [];
  let isolatedPairCount = 0;
  const columnCount = Math.max.apply(null, allValues.map(function (row) { return (row || []).length; }).concat([0]));
  for (let c = 0; c < columnCount; c++) {
    const found = collectFromColumn(c);
    if (found.length >= 2) columnGroups.push({ col: c, pairs: found });
    else if (found.length === 1) isolatedPairCount++;
  }
  let pairs = [];
  columnGroups.forEach(function (group) { pairs = pairs.concat(group.pairs); });

  // "룸메이트1 | 룸메이트2"처럼 두 칸으로 나뉜 표도 연속된 행 묶음 단위로 인식합니다.
  const adjacentRowPairs = [];
  for (let r = 0; r < allValues.length; r++) {
    const rowCandidates = [];
    for (let c = 0; c < (allValues[r] || []).length; c++) {
      const text = String(allValues[r][c] == null ? '' : allValues[r][c]).trim();
      if (!text || isParticipantHeaderText_(text) || /^(구분|방|호실|객실|룸메이트\s*1|룸메이트\s*2)$/i.test(text)) continue;
      const resolved = resolveNameAgainstParticipantReference_(text, nameReference);
      if (resolved.status !== 'notFound' || isLikelyPersonNameText_(text)) rowCandidates.push({ value: text, col: c });
    }
    if (rowCandidates.length === 2 && Math.abs(rowCandidates[0].col - rowCandidates[1].col) <= 3) {
      adjacentRowPairs.push({ row: r, a: rowCandidates[0].value, b: rowCandidates[1].value });
    }
  }
  const adjacentClusters = [];
  adjacentRowPairs.forEach(function (pair) {
    const last = adjacentClusters[adjacentClusters.length - 1];
    if (!last || pair.row - last[last.length - 1].row > 2) adjacentClusters.push([pair]);
    else last.push(pair);
  });
  adjacentClusters.filter(function (cluster) { return cluster.length >= 2; }).forEach(function (cluster) {
    cluster.forEach(function (pair) { pairs.push({ a: pair.a, b: pair.b }); });
  });

  let reviewWarning = '';
  if (!pairs.length && (isolatedPairCount || adjacentRowPairs.length)) {
    reviewWarning = '룸메이트로 보이는 값이 1건만 떨어져 있어 자동 반영하지 않았습니다. 실제 룸메이트 명단이라면 관리자 확인 후 같은 열에 모아주세요.';
  } else if (isolatedPairCount) {
    reviewWarning = '룸메이트 명단이 모여 있는 열은 반영했고, 다른 열에 한 건만 떨어진 값은 오인 가능성이 있어 제외했습니다.';
  }

  // 중복 제거
  const seen = {};
  const uniquePairs = [];
  pairs.forEach(function (p) {
    const key = p.a < p.b ? p.a + '||' + p.b : p.b + '||' + p.a;
    if (!seen[key]) {
      seen[key] = true;
      uniquePairs.push(p);
    }
  });

  return { sheetName: sheet.getName(), sheetId: sheet.getSheetId(), sheetUrl: buildSheetUrl_(ss, sheet), pairs: uniquePairs, warning: reviewWarning };
}

function splitRoommatePairText_(value, nameReference) {
  const text = normalizePersonText_(value).replace(/^예\)\s*/, '');
  if (!text) return null;

  // '/' 외에도 &, 쉼표, '+', '|', 세미콜론, '·', '와/과/랑/이랑/및', and/with를 허용합니다.
  // 영어 이름의 "Last, First" 오판을 줄이기 위해 쉼표는 양쪽이 기준 명단 이름이거나
  // 이름처럼 보일 때에만 최종 채택합니다.
  const separatorPatterns = [
    /\s*[\/／&＆+]\s*/,
    /\s*[|｜;；]\s*/,
    /\s+(?:and|with)\s+/i,
    /\s*(?:이랑|와|과|랑|및)\s+/,
    /\s*[·ㆍ]\s*/,
    /\s*[,，]\s*/
  ];
  for (let i = 0; i < separatorPatterns.length; i++) {
    const parts = text.split(separatorPatterns[i]).map(function (part) { return part.trim(); }).filter(Boolean);
    if (parts.length !== 2) continue;
    if (/^룸메이트(?:\s*[12])?$/i.test(parts[0]) || /^룸메이트(?:\s*[12])?$/i.test(parts[1])) continue;
    const first = resolveNameAgainstParticipantReference_(parts[0], nameReference);
    const second = resolveNameAgainstParticipantReference_(parts[1], nameReference);
    if (
      i === separatorPatterns.length - 1 &&
      first.status === 'notFound' && second.status === 'notFound' &&
      /^[A-Za-z'-]+$/.test(parts[0]) && /^[A-Za-z'-]+$/.test(parts[1])
    ) {
      // "Doe, John" 같은 영문 성명 한 개를 두 사람으로 쪼개지 않습니다.
      continue;
    }
    const firstOk = first.status !== 'notFound' || isLikelyPersonNameText_(parts[0]);
    const secondOk = second.status !== 'notFound' || isLikelyPersonNameText_(parts[1]);
    if (firstOk && secondOk) return { a: parts[0], b: parts[1] };
  }
  return null;
}

function findBestRoommateSheetByContent_(ssOrSheets, excluded, participants) {
  const sheets = Array.isArray(ssOrSheets) ? ssOrSheets : ssOrSheets.getSheets();
  const reference = buildParticipantNameReference_(participants);
  let best = null;
  sheets.forEach(function (sheet) {
    if (excluded && excluded[sheet.getName()]) return;
    const values = getBoundedSheetValues_(sheet, 220, 40);
    let pairCount = 0;
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < (values[r] || []).length; c++) {
        if (splitRoommatePairText_(values[r][c], reference)) pairCount++;
      }
    }
    let score = pairCount * 12;
    const name = sheet.getName().replace(/\s+/g, '');
    if (/(객실|룸메|숙소|방배정|room)/i.test(name)) score += 25;
    if (/(좌석|자리|강의|교육생|raw)/i.test(name)) score -= 25;
    if (!best || score > best.score) best = { sheet: sheet, score: score, pairCount: pairCount };
  });
  return best && best.pairCount >= 1 && best.score >= 10 ? best.sheet : null;
}
