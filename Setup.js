/**
 * Inventory Management — Setup.gs
 * 설정 / 설치 / 유지보수 함수 + 공용 상수
 *
 * 시트 구성:
 *   - Transaction : 입력 폼 UI (검색 영역 + 트랜잭션 폼 + 헬퍼 열 W/X/Y/Z)
 *   - Ledger      : 거래 원장 (append-only, 재고의 단일 진실 공급원)
 *   - Inventory   : Ledger 연산 결과(정규화 재고 목록) — 스크립트가 기록
 *   - Dashboard   : Inventory 를 피벗한 아이템 × 위치 매트릭스
 *   - Settings    : 마스터 데이터 (Category/Item/Serial여부, Location, Tour, USER)
 *
 * 이름 있는 범위(Named Ranges):
 *   - LIST_CATEGORY = Settings!$A$2:$A$1000
 *   - LIST_ITEM     = Settings!$B$2:$B$1000
 *   - LIST_BUCKET   = Settings!$E$2:$F$1000   (물리 위치 + 투어 패키지를 하나의 "버킷"으로 취급)
 *
 * 참고: 아래 공용 상수는 전역(global)이라 Transaction.gs 등 다른 파일에서도 사용됩니다.
 */

// Ledger 열 인덱스 (0-based) — 전역 공용
const LEDGER_COL = {
  TIMESTAMP: 0, TYPE: 1, CATEGORY: 2, ITEM_ID: 3, ITEM_NAME: 4, SERIAL: 5,
  FROM: 6, TO: 7, QTY: 8, WORKER: 9, NOTE: 10
};

const EXTERNAL_VENDOR = 'EXTERNAL (VENDOR)';
const EXTERNAL_SCRAP = 'EXTERNAL (SCRAP)';

// 단위(Unit) 드롭다운 허용 목록 — Settings_Item E열 검증에 사용
const UNIT_OPTIONS = ['EA', 'box', 'roll', 'set', 'pack', 'pair', 'm'];

/**
 * 필수 시트를 가져오되, 없으면 실제 시트 목록과 함께 알림을 띄우고 null 을 반환한다.
 * (호출부에서 null 이면 즉시 return 하여 크래시 대신 명확한 안내로 중단)
 */
function getRequiredSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    const names = ss.getSheets().map(function (s) { return s.getName(); }).join(', ');
    SpreadsheetApp.getUi().alert(
      '❌ 에러: "' + name + '" 시트를 찾을 수 없습니다.\n' +
      '시트 이름(대소문자·공백 포함)을 확인해 주세요.\n\n' +
      '현재 시트 목록: ' + names
    );
  }
  return sheet;
}

/**
 * 스프레드시트를 열 때 상단에 '📦 Inventory' 메뉴를 추가한다.
 * (설치/유지보수 함수를 편집기 없이 시트에서 바로 실행할 수 있게 해줌)
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 Inventory')
    .addItem('트랜잭션 제출 (Submit)', 'submitTransaction')
    .addSeparator()
    .addItem('🚀 전체 초기화 (Initialize)', 'initializeSystem')
    .addItem('입력폼 재설정 (Setup Form)', 'setupInputSheet')
    .addItem('재고 새로고침 (Rebuild Inventory)', 'rebuildInv')
    .addItem('대시보드 재구성 (Rebuild Dashboard)', 'rebuildDashboard')
    .addItem('단위(Unit) 도입/갱신 (Setup Unit)', 'setupUnitColumn')
    .addItem('설치형 편집트리거 제거 (Use Fast onEdit)', 'removeInstallableEditTriggers')
    .addItem('시리얼 텍스트 변환 (Migrate Serials)', 'migrateSerialsToText')
    .addItem('옛 원본 탭 아카이브 (Archive Origin Tabs)', 'archiveOriginTabs')
    .addItem('투어 위치 정리 (Cleanup Tour Locations)', 'cleanupTourLocations')
    .addItem('위치 비우기·이동 (Relocate Stock)', 'relocateAllStock')
    .addItem('원장 서식 정리 (Normalize Ledger)', 'normalizeLedgerFormat')
    .addToUi();
}

/**
 * [빈 스프레드시트용] 시스템 전체를 처음부터 세운다.
 * 필요한 시트 5개(Transaction/Ledger/Inventory/Dashboard/Settings), 헤더, 마스터 시딩,
 * 이름범위(LIST_*), Inventory/Dashboard 수식, 입력폼을 순서대로 생성한다.
 */
/**
 * 🌟 [수정] 신설 시트(Settings_Location, Settings_User)를 포함하여 시스템 전체 초기화
 */
function initializeSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  renameSheetIfNeeded_(ss, 'Input_Transaction', 'Transaction');
  renameSheetIfNeeded_(ss, 'INV', 'Inventory');
  renameSheetIfNeeded_(ss, 'Settings', 'Settings_Item'); // 🌟 기존 Settings 시트명 자동 변경

  const ledger = getOrCreateSheet_(ss, 'Ledger');
  setupLedgerSheet_(ledger);

  const settingsItem = getOrCreateSheet_(ss, 'Settings_Item'); // 🌟 참조 시트명 변경
  const settingsLoc = getOrCreateSheet_(ss, 'Settings_Location');
  const settingsUser = getOrCreateSheet_(ss, 'Settings_User');
  setupSettingsSheet_(settingsItem, settingsLoc, settingsUser);

  defineNamedRanges_(ss, settingsItem, settingsLoc, settingsUser);

  const inv = getOrCreateSheet_(ss, 'Inventory');
  rebuildInv_(ss);

  const dashboard = getOrCreateSheet_(ss, 'Dashboard');
  setupDashboardSheet_(dashboard);

  const scrap = getOrCreateSheet_(ss, 'Scrap');
  setupScrapSheet_(scrap);

  setupInputSheet();
  removeDefaultSheetIfEmpty_(ss);

  SpreadsheetApp.flush();
  ui.alert('✅ 네이밍 통일 완료!\n모든 마스터 시트가 Settings_ 프리픽스로 통일되었으며 수식이 정상 연동되었습니다.');
}

  /** Inventory_Origin 헤더 설정 및 시트 잠금(보호) */
  function setupOriginSheet_(sheet) {
  // Item ID가 추가된 6열 구조
  const headers = ['Category', 'Item ID', 'Item Name', 'Location', 'Serial Number', 'Quantity'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.getRange('E:E').setNumberFormat('@'); // 시리얼 열이 D -> E열로 밀려남
  sheet.setFrozenRows(1);
  
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (protections.length === 0) {
    sheet.protect().setDescription('⚠️ 기초 재고 구역 (수정 금지)');
  }
}

/** 구버전 이름의 시트가 있고 새 이름 시트가 없으면 이름을 바꾼다 (데이터 보존) */
function renameSheetIfNeeded_(ss, oldName, newName) {
  const oldSheet = ss.getSheetByName(oldName);
  if (oldSheet && !ss.getSheetByName(newName)) {
    oldSheet.setName(newName);
  }
}

/** 시트를 가져오되 없으면 새로 만든다 */
function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/** 원장(Ledger) 헤더 및 형식 설정 (기존 데이터는 건드리지 않음) */
function setupLedgerSheet_(sheet) {
  const headers = ['Timestamp', 'Type', 'Category', 'Item ID', 'Item Name', 'Serial Number',
                   'From', 'To', 'Quantity', 'Worker', 'Note'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.getRange('A:A').setNumberFormat('yyyy-MM-dd HH:mm:ss'); // 🌟 A열 전체를 초 단위 시간 포맷으로 고정
  sheet.getRange('F:F').setNumberFormat('@'); 
  sheet.setFrozenRows(1);
}

/**
 * 🌟 [수정] 3개로 분리된 마스터 시트의 헤더 및 데이터 시딩 통합 처리
 */
/**
 * 마스터 시트 세팅 (하드코딩 데이터 시딩 완벽 제거, 오직 헤더만 관리)
 */
function setupSettingsSheet_(sheet, locSheet, userSheet) {
  // ① Settings (물품 마스터) 헤더만 고정
  sheet.getRange('A1:E1').setValues([['Item ID', 'Category', 'Item Name', 'Manage Serial', 'Unit']]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // ② Settings_Location (장소 마스터) 헤더만 고정
  locSheet.getRange('A1:B1').setValues([['Location', 'Tour Package']]).setFontWeight('bold');
  locSheet.setFrozenRows(1);

  // ③ Settings_User (사용자 마스터) 헤더만 고정
  userSheet.getRange('A1').setValue('USER').setFontWeight('bold');
  userSheet.setFrozenRows(1);

  // 🌟 데이터 시딩 로직 전면 삭제 완료
  // 이제부터 시스템은 스크립트 내부가 아닌, 시트에 직접 입력된 최신 데이터만을 100% 신뢰하고 참조합니다.
}

/**
 * 🌟 [수정] 이름 범위가 참조하는 시트를 각각 분리된 독립 시트로 수정
 */
function defineNamedRanges_(ss, settings, settingsLoc, settingsUser) {
  ss.setNamedRange('LIST_CATEGORY', settings.getRange('B2:B1000'));
  ss.setNamedRange('LIST_ITEM', settings.getRange('C2:C1000'));
  ss.setNamedRange('LIST_BUCKET', settingsLoc.getRange('A2:B1000')); // 🌟 Settings_Location 참조
}

/**
 * 상단 메뉴에서 수동으로 '재고 새로고침'을 실행할 때 연결되는 함수
 */
function rebuildInv() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const ledger = ss.getSheetByName('Ledger');
  const settings = ss.getSheetByName('Settings_Item'); // ✅ Settings_Item 으로 수정 완료
  
  if (!ledger || !settings) {
    ui.alert('❌ 에러: Ledger 시트 또는 Settings_Item 시트를 찾을 수 없습니다.');
    return;
  }
  
  // 실제 재고 연산 함수 실행
  const rowCount = rebuildInv_(ss);
  
  ui.alert('✅ 완료: ' + rowCount + '개의 품목 재고가 성공적으로 새로고침(동기화) 되었습니다.');
}

/**
 * Inventory — Ledger를 연산해 "정규화된 재고 목록"으로 정리하는 데이터 계층.
 * 한 행 = (Category, Item, Location, Serial, Quantity), 재고 수량이 0보다 큰 항목만.
 *   - 시리얼 관리 품목: 시리얼별로 한 행 (수량 보통 1)
 *   - 비시리얼 품목: (품목, 위치)별로 한 행, Serial 칸은 공백
 * 반환값: 기록한 재고 행 수.
 */
function rebuildInv_(ss) {
  const ledger = ss.getSheetByName('Ledger');
  const settings = ss.getSheetByName('Settings_Item'); // 🌟 시트명 변경
  const inv = getOrCreateSheet_(ss, 'Inventory');
  if (!ledger || !settings) return 0;

  const catOf = {};
  const serialOf = {};
  const nameOf = {};
  const unitOf = {};
  const sLast = settings.getLastRow();

  if (sLast >= 2) {
    const sv = settings.getRange(2, 1, sLast - 1, 5).getValues(); // A:E (E=Unit)
    for (let i = 0; i < sv.length; i++) {
      const id = sv[i][0];
      const name = sv[i][2];
      if (id) {
        catOf[id] = sv[i][1];
        nameOf[id] = name;
        serialOf[id] = String(sv[i][3]).toUpperCase() === 'YES';
        unitOf[id] = sv[i][4]; // E: Unit
      }
    }
  }

  const map = {};

  const rows = getLedgerRows_(ledger);
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const id = r[LEDGER_COL.ITEM_ID]; 
    const qty = Number(r[LEDGER_COL.QTY]) || 0;
    const isSer = !!serialOf[id];
    const serial = isSer ? normSerial_(r[LEDGER_COL.SERIAL]) : '';
    const to = r[LEDGER_COL.TO];
    const from = r[LEDGER_COL.FROM];
    if (!isExternal_(to)) addInvQty_(map, catOf, id, to, serial, qty);
    if (!isExternal_(from)) addInvQty_(map, catOf, id, from, serial, -qty);
  }

  const out = [];
  const keys = Object.keys(map);
  for (let i = 0; i < keys.length; i++) {
    const e = map[keys[i]];
    if (e.qty > 0) out.push([e.cat, e.item, nameOf[e.item] || 'UNKNOWN', e.loc, e.serial, e.qty, unitOf[e.item] || '']);
  }
  
  out.sort(function (a, b) {
    return String(a[0]).localeCompare(String(b[0])) || 
           String(a[2]).localeCompare(String(b[2])) || 
           String(a[3]).localeCompare(String(b[3])) || 
           String(a[4]).localeCompare(String(b[4]));   
  });
  
  inv.clearContents();
  inv.getRange(1, 1, inv.getMaxRows(), inv.getMaxColumns()).setFontWeight('normal');
  
  const headers = ['Category', 'Item ID', 'Item Name', 'Location', 'Serial Number', 'Quantity', 'Unit'];
  inv.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  inv.getRange('E:E').setNumberFormat('@');
  
  if (out.length > 0) {
    inv.getRange(2, 1, out.length, headers.length).setValues(out);
  }
  inv.setFrozenRows(1);

  const latestVal = ledger.getRange('A2').getValue();
  if (latestVal instanceof Date) {
    const timeStr = Utilities.formatDate(latestVal, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
    inv.getRange('G1').setValue('🕒 최종 업데이트: ' + timeStr).setFontWeight('bold').setFontColor('#4a86e8');
  } else {
    inv.getRange('G1').setValue('🕒 최종 업데이트: 기록 없음').setFontWeight('bold').setFontColor('#4a86e8');
  }

  return out.length;
}

/** rebuildInv_ 내부: (품목|위치|시리얼) 키에 수량 누적 */
function addInvQty_(map, catOf, item, loc, serial, delta) {
  const key = item + '' + loc + '' + serial;
  if (!map[key]) {
    map[key] = { cat: catOf[item] || '', item: item, loc: loc, serial: serial, qty: 0 };
  }
  map[key].qty += delta;
}

/**
 * Dashboard — Inventory(정규화 재고 목록)를 기반으로 아이템 × 위치 매트릭스를 만든다.
 *
 * 위치 열 순서는 Settings_Location(LIST_BUCKET)을 읽어 스크립트가 결정한다:
 *   KR OFFICE → GMP WH(합계) → GMP (통로) → 18/19 세부구역 → US OFFICE
 *   → (기타 물리 위치) → 투어 패키지(G…)
 * 'GMP WH' 는 실제 위치가 아니라 통로 + 모든 세부구역 재고를 합산한 요약 열이다.
 *
 * 값은 라이브 수식(Inventory 참조)으로 채워지므로 Inventory가 갱신되면 자동 반영된다.
 * Settings_Location 을 편집한 뒤에는 메뉴 '대시보드 재구성' 으로 열 구성을 갱신한다.
 */
function setupDashboardSheet_(sheet) {
  const ss = sheet.getParent();
  sheet.clearContents();

  sheet.getRange('A1').setValue('📊 [Inventory Dashboard]').setFontWeight('bold');
  sheet.getRange('A2').setValue('Category').setFontWeight('bold');
  sheet.getRange('B2').setValue('Item Name').setFontWeight('bold');
  sheet.getRange('C2').setValue('Unit').setFontWeight('bold');
  sheet.getRange('D2').setValue('Total').setFontWeight('bold');

  sheet.getRange('E1').setFormula(
    '=IF(Ledger!A2="Timestamp", "🕒 기록 없음", "🕒 최종 업데이트: " & TEXT(Ledger!A2, "yyyy-mm-dd HH:mm:ss"))'
  ).setFontWeight('bold').setFontColor('#4a86e8');

  // A3: Category → Item Name 이중 오름차순 정렬 (2열 배열이 A·B로 스필)
  sheet.getRange('A3').setFormula(
    '=IFERROR(SORT(FILTER({LIST_CATEGORY, LIST_ITEM}, LIST_ITEM<>""), 1, TRUE, 2, TRUE), "")'
  );

  // C3: Unit — 품목명으로 Settings_Item(E열)에서 단위 조회
  sheet.getRange('C3').setFormula(
    '=IFERROR(BYROW(CHOOSECOLS(SORT(FILTER({LIST_CATEGORY, LIST_ITEM}, LIST_ITEM<>""), 1, TRUE, 2, TRUE), 2), LAMBDA(i, IFERROR(VLOOKUP(i, Settings_Item!$C:$E, 3, FALSE), ""))), "")'
  );

  // D3: Total — A3와 동일한 정렬식(SORT/FILTER)으로 품목 순서를 맞춰 총합
  sheet.getRange('D3').setFormula(
    '=IFERROR(BYROW(CHOOSECOLS(SORT(FILTER({LIST_CATEGORY, LIST_ITEM}, LIST_ITEM<>""), 1, TRUE, 2, TRUE), 2), LAMBDA(i, SUMIFS(Inventory!$F:$F, Inventory!$C:$C, i))), "")'
  );

  // E열 이후: 위치 매트릭스 (스크립트가 정한 순서 + GMP WH 합계 열)
  const columns = buildDashboardColumns_(ss);
  if (columns.length > 0) {
    const MATRIX_START = 5; // E열부터 위치 매트릭스 시작 (A~D = Category/Item/Unit/Total)
    const headers = columns.map(function (c) { return c.label; });
    sheet.getRange(2, MATRIX_START, 1, headers.length).setValues([headers]).setFontWeight('bold');

    // 헤더(2행) 배경색으로 GMP WH ⊃ 하위 구역 계층을 시각화
    const bgOf = { 'gmp-sum': '#a4c2f4', 'gmp-member': '#d9e7fb' }; // 합계=연한 파랑 / 하위=더 연한 파랑
    const backgrounds = columns.map(function (c) { return bgOf[c.kind] || null; });
    sheet.getRange(2, MATRIX_START, 1, columns.length).setBackgrounds([backgrounds]);

    const formulas = columns.map(function (c) {
      const terms = c.locs.map(function (loc) {
        return 'SUMIFS(Inventory!$F:$F, Inventory!$C:$C, i, Inventory!$D:$D, "' +
               String(loc).replace(/"/g, '""') + '")';
      }).join(' + ');
      return '=IFERROR(BYROW(CHOOSECOLS(SORT(FILTER({LIST_CATEGORY, LIST_ITEM}, LIST_ITEM<>""), 1, TRUE, 2, TRUE), 2), LAMBDA(i, ' + terms + ')), "")';
    });
    sheet.getRange(3, MATRIX_START, 1, formulas.length).setFormulas([formulas]);

    // GMP WH 하위 구역(연속된 gmp-member 열)을 열 그룹으로 묶어 접기/펼치기 제공
    clearColumnGroups_(sheet);
    let startIdx = -1, count = 0;
    columns.forEach(function (c, i) {
      if (c.kind === 'gmp-member') { if (startIdx === -1) startIdx = i; count++; }
    });
    if (count > 0) {
      const startCol = MATRIX_START + startIdx;
      sheet.getRange(1, startCol, 1, count).shiftColumnGroupDepth(1);
      sheet.getColumnGroup(startCol, 1).collapse(); // 기본 접힘 (GMP WH 합계만 보이도록)
    }
  }

  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(4);
}

/** Dashboard의 기존 열 그룹을 모두 제거한다 (재구성 시 그룹 중첩 방지). */
function clearColumnGroups_(sheet) {
  const maxCols = sheet.getMaxColumns();
  let changed = true, guard = 0;
  while (changed && guard < 100) {
    changed = false; guard++;
    for (let c = 1; c <= maxCols; c++) {
      const d = sheet.getColumnGroupDepth(c);
      if (d > 0) {
        try { sheet.getColumnGroup(c, d).remove(); changed = true; break; } catch (e) {}
      }
    }
  }
}


/**
 * Settings_Location(LIST_BUCKET)을 읽어 Dashboard 위치 열 구성을 순서대로 만든다.
 * 반환: [{ label, locs: [실제위치...] }] — locs 가 여러 개면 합계(요약) 열.
 */
function buildDashboardColumns_(ss) {
  const loc = ss.getSheetByName('Settings_Location');
  if (!loc) return [];

  const flat = function (rng) {
    return rng.getValues()
      .map(function (r) { return String(r[0]).trim(); })
      .filter(function (v) { return v !== ''; });
  };
  const physical = flat(loc.getRange('A2:A1000'));
  const tour = flat(loc.getRange('B2:B1000'));

  const isZone = function (s) { return /^\d/.test(s); };          // 18-1-01, 19-1-10 …
  const AISLE = 'GMP (통로)';
  const aisle = physical.filter(function (s) { return s === AISLE; });
  const zones = physical.filter(isZone);
  const gmpMembers = aisle.concat(zones);                          // 통로 → 세부구역
  const placed = {};
  ['KR OFFICE', 'US OFFICE', AISLE].forEach(function (s) { placed[s] = true; });
  zones.forEach(function (s) { placed[s] = true; });

  const columns = [];
  if (physical.indexOf('KR OFFICE') !== -1) columns.push({ label: 'KR OFFICE', locs: ['KR OFFICE'] });
  if (gmpMembers.length > 0) columns.push({ label: 'GMP WH', locs: gmpMembers, kind: 'gmp-sum' });   // 합계 열
  gmpMembers.forEach(function (l) { columns.push({ label: l, locs: [l], kind: 'gmp-member' }); });   // GMP WH 하위 구역
  if (physical.indexOf('US OFFICE') !== -1) columns.push({ label: 'US OFFICE', locs: ['US OFFICE'] });
  physical.forEach(function (l) { if (!placed[l]) columns.push({ label: l, locs: [l] }); }); // 기타 물리 위치
  tour.forEach(function (l) { columns.push({ label: l, locs: [l] }); });            // 투어 패키지

  return columns;
}


/**
 * [메뉴] Dashboard 위치 열 구성을 다시 만든다.
 * Settings_Location 을 편집(위치 추가/삭제/순서 변경)한 뒤 실행하면 반영된다.
 */
function rebuildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const dash = ss.getSheetByName('Dashboard');
  if (!dash) {
    ui.alert('❌ 에러: Dashboard 시트를 찾을 수 없습니다.');
    return;
  }
  setupDashboardSheet_(dash);
  SpreadsheetApp.flush();
  ui.alert('✅ 완료: 대시보드 위치 열 구성을 새로 만들었습니다.');
}


/** 새 스프레드시트에 기본 생성되는 빈 'Sheet1' 을 (비어있고 다른 시트가 있으면) 삭제 */
function removeDefaultSheetIfEmpty_(ss) {
  const sh = ss.getSheetByName('Sheet1') || ss.getSheetByName('시트1');
  if (sh && ss.getSheets().length > 1 && sh.getLastRow() === 0 && sh.getLastColumn() === 0) {
    ss.deleteSheet(sh);
  }
}

/**
 * 1. Transaction 시트의 UI 구조 및 수식을 초기화하는 함수
 * (기초 재고 및 성능 최적화를 위해 Ledger 대신 Inventory 시트를 참조하도록 수식 개선)
 */
/**
 * 🌟 [수정] 폼 생성 시 Settings_Item 참조 및 VLOOKUP 수식 업데이트
 */
function setupInputSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = getRequiredSheet_(ss, 'Settings_Item'); // 🌟 시트명 변경
  const userSheet = getRequiredSheet_(ss, 'Settings_User'); 
  if (!settingsSheet || !userSheet) return;

  let inputSheet = ss.getSheetByName('Transaction');
  if (!inputSheet) inputSheet = ss.insertSheet('Transaction');
  
  inputSheet.getRange(1, 1, inputSheet.getMaxRows(), inputSheet.getMaxColumns()).clearDataValidations();
  inputSheet.clear();

  inputSheet.getRange('B2').setValue('🔍 Item Search').setFontWeight('bold');
  inputSheet.getRange('B3').setValue('Category');
  inputSheet.getRange('B4').setValue('Item');
  inputSheet.getRange('B6').setValue('Location').setFontWeight('bold');
  inputSheet.getRange('C6').setValue('Current Qty').setFontWeight('bold');

  inputSheet.getRange('E2').setValue('📦 Transaction Form').setFontWeight('bold');
  inputSheet.getRange('E3').setValue('Type');
  inputSheet.getRange('E4').setValue('Item');
  inputSheet.getRange('E5').setValue('SERIAL NO.');
  inputSheet.getRange('E6').setValue('FROM');
  inputSheet.getRange('E7').setValue('TO');
  inputSheet.getRange('E8').setValue('Quantity');
  inputSheet.getRange('E9').setValue('USER');
  inputSheet.getRange('E10').setValue('Note');

  inputSheet.getRange('B7').setFormula(
    '=IFERROR(LET(item, $C$4, buckets, TOCOL(LIST_BUCKET, 1, TRUE), ' +
    'balances, BYROW(buckets, LAMBDA(b, SUMIFS(Inventory!$F:$F, Inventory!$C:$C, item, Inventory!$D:$D, b))), ' +
    'FILTER({buckets, balances}, balances <> 0)), "Item not selected or out of stock.")'
  );
  inputSheet.getRange('F4').setFormula('=$C$4');
  inputSheet.getRange('W1').setFormula(
    '=IFERROR(UNIQUE(FILTER(TOCOL(LIST_BUCKET, 1, TRUE), ' +
    'BYROW(TOCOL(LIST_BUCKET, 1, TRUE), LAMBDA(b, SUMIFS(Inventory!$F:$F, Inventory!$C:$C, $C$4, Inventory!$D:$D, b))) > 0)), "")'
  );
  inputSheet.getRange('X1').setFormula(
    '=IFERROR(FILTER(Inventory!E:E, Inventory!C:C=$C$4, Inventory!D:D=$F$6, Inventory!F:F>0, Inventory!E:E<>"N/A", Inventory!E:E<>""), "")'
  );
  inputSheet.getRange('Y1').setFormula('=IFERROR(TOCOL(LIST_BUCKET, 1, TRUE), "")');
  inputSheet.getRange('Z1').setFormula(
    '=IFERROR(IF(C3="", FILTER(LIST_ITEM, LIST_ITEM<>""), FILTER(LIST_ITEM, LIST_CATEGORY=C3, LIST_ITEM<>"")), "")'
  );
  
  // 🌟 [핵심 변경] 수식 내부의 시트명 Settings! -> Settings_Item! 으로 교체
  inputSheet.getRange('AA1').setFormula('=IFERROR(VLOOKUP($C$4, Settings_Item!$C:$D, 2, FALSE), "NO")');

  inputSheet.getRange('F3').setValue('ADD');  
  inputSheet.getRange('F5').setNumberFormat('@').setValue('N/A'); 
  inputSheet.getRange('F8').setValue(1);      

  const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(['ADD', 'MOVE', 'REMOVE'], true).setAllowInvalid(false).build();
  const categoryRule = SpreadsheetApp.newDataValidation().requireValueInRange(settingsSheet.getRange('B2:B'), true).setAllowInvalid(false).build();
  const searchItemRule = SpreadsheetApp.newDataValidation().requireValueInRange(inputSheet.getRange('Z1:Z'), true).setAllowInvalid(false).build();
  const toLocRule = SpreadsheetApp.newDataValidation().requireValueInRange(inputSheet.getRange('Y1:Y'), true).setAllowInvalid(false).build();
  const userRule = SpreadsheetApp.newDataValidation().requireValueInRange(userSheet.getRange('A2:A'), true).setAllowInvalid(true).build();

  inputSheet.getRange('F3').setDataValidation(typeRule);       
  inputSheet.getRange('C3').setDataValidation(categoryRule);   
  inputSheet.getRange('C4').setDataValidation(searchItemRule); 
  inputSheet.getRange('F7').setDataValidation(toLocRule);      
  inputSheet.getRange('F9').setDataValidation(userRule);       

  inputSheet.getRange('C3:C4').setBackground('#fff2cc');  
  inputSheet.getRange('F3:F4').setBackground('#e2efda');  
  inputSheet.getRange('F5:F10').setBackground('#e2efda');
  SpreadsheetApp.flush();  
}

/**
 * [1회 실행] updateDynamicUI 를 '수정 시(onEdit)' 트리거로 자동 등록한다.
 * - Apps Script 편집기에서 이 함수를 한 번만 실행하면 트리거가 만들어진다.
 * - 이미 같은 트리거가 있으면 중복 생성하지 않는다(중복 실행 방지).
 */
function createTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 기존에 updateDynamicUI 로 연결된 onEdit 트리거가 있는지 확인
  const existing = ScriptApp.getProjectTriggers();
  for (let i = 0; i < existing.length; i++) {
    const t = existing[i];
    if (t.getHandlerFunction() === 'updateDynamicUI' &&
        t.getEventType() === ScriptApp.EventType.ON_EDIT) {
      ui.alert('ℹ️ 이미 updateDynamicUI(수정 시) 트리거가 등록되어 있습니다. 추가 작업이 필요 없습니다.');
      return;
    }
  }

  // 없으면 새로 등록
  ScriptApp.newTrigger('updateDynamicUI')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ui.alert('✅ 완료: updateDynamicUI(수정 시) 트리거를 등록했습니다.');
}

/**
 * [1회 실행] 원장(Ledger)에 숫자로 저장된 기존 시리얼 번호를 텍스트로 일괄 변환한다.
 * - E열(Serial Number)을 텍스트 형식으로 고정한 뒤, 숫자로 들어간 값을 문자열로 다시 기록한다.
 * - 'N/A' 및 이미 텍스트인 값은 그대로 둔다.
 * - 정수 시리얼의 소수점 표기(예: 123456789.0) 문제도 함께 정리된다.
 */
function migrateSerialsToText() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const ledgerSheet = getRequiredSheet_(ss, 'Ledger');
  if (!ledgerSheet) return;
  const last = ledgerSheet.getLastRow();
  if (last < 2) {
    ui.alert('ℹ️ 원장에 데이터가 없습니다.');
    return;
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const serialRange = ledgerSheet.getRange(2, LEDGER_COL.SERIAL + 1, last - 1, 1); // E열
    serialRange.setNumberFormat('@'); // 열 형식을 텍스트로 고정

    const values = serialRange.getValues();
    let changed = 0;
    for (let i = 0; i < values.length; i++) {
      const v = values[i][0];
      if (typeof v === 'number') {
        // 정수는 소수점 없이, 그 외에는 일반 문자열로 변환
        const asText = Number.isInteger(v) ? String(v) : String(v);
        values[i][0] = asText;
        changed++;
      }
    }

    if (changed > 0) {
      serialRange.setValues(values);
      SpreadsheetApp.flush();
    }
    ui.alert('✅ 완료: 숫자로 저장된 시리얼 ' + changed + '건을 텍스트로 변환했습니다.');

  } catch (error) {
    ui.alert('❌ 오류: ' + error.toString());
  } finally {
    lock.releaseLock();
  }
}

/**
 * Scrap — Ledger 시트에서 폐기(EXTERNAL (SCRAP)) 처리된 내역만 실시간으로 필터링해서 보여주는 시트
 */
function setupScrapSheet_(sheet) {
  sheet.clearContents();
  
  sheet.getRange('A1').setValue('🗑️ [Scrap / Disposal Log]').setFontWeight('bold');
  
  // Ledger 시트의 H열(To)이 'EXTERNAL (SCRAP)'인 행 전체를 쿼리해서 가져오는 수식
  // A열(Timestamp)은 내림차순(최신순)으로 정렬합니다.
  sheet.getRange('A2').setFormula(
    '=IFERROR(QUERY(Ledger!A:K, "SELECT * WHERE H = \'EXTERNAL (SCRAP)\' ORDER BY A DESC", 1), "No scrap records found.")'
  );

  sheet.setFrozenRows(2); // 헤더(2행) 고정
  
  // A열(Timestamp)의 시간 포맷이 숫자로 깨져 보이지 않도록 포맷 고정
  sheet.getRange('A:A').setNumberFormat('yyyy-MM-dd HH:mm:ss');
  // F열(Serial Number) 텍스트 포맷 고정
  sheet.getRange('F:F').setNumberFormat('@');
}

/**
 * [메뉴/1회성] 마이그레이션 이전 원본 스냅샷 탭(Inventory_Origin *)을 아카이브한다.
 * - 이름 끝에 '_archive' 를 붙이고 시트를 숨긴다. (데이터는 이미 Ledger 로 이관됨)
 * - 이미 '_archive' 가 붙은 탭은 건너뛴다(중복 실행 안전).
 */
function archiveOriginTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const done = [];
  ss.getSheets().forEach(function (sh) {
    const name = sh.getName();
    if (name.indexOf('Inventory_Origin') === 0 && name.indexOf('_archive') === -1) {
      const newName = name + '_archive';
      sh.setName(newName);
      sh.hideSheet();
      done.push(name + ' → ' + newName + ' (숨김)');
    }
  });
  if (done.length === 0) {
    ui.alert('ℹ️ 아카이브할 Inventory_Origin 탭이 없습니다 (이미 처리되었을 수 있습니다).');
  } else {
    ui.alert('✅ 아카이브 완료:\n\n' + done.join('\n') +
             '\n\n숨긴 탭은 아무 탭이나 우클릭 → "숨겨진 시트" 에서 다시 볼 수 있습니다.');
  }
}

/**
 * [메뉴] 단위(Unit) 도입/갱신 — 한 번 실행으로 아래를 처리한다.
 *   1) Settings_Item E1 헤더 'Unit' 보장
 *   2) E2:E1000 에 단위 드롭다운(UNIT_OPTIONS) 적용
 *   3) Item Name 이 있는데 Unit 이 빈 행을 'EA' 로 일괄 채움
 *   4) Inventory(Unit 열) + Dashboard(Unit 열) 갱신
 * roll/box 등 예외 품목만 이후 E열에서 직접 바꾸면 된다. (중복 실행 안전)
 */
function setupUnitColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const s = getRequiredSheet_(ss, 'Settings_Item');
  if (!s) return;

  // 1) 헤더
  s.getRange('E1').setValue('Unit').setFontWeight('bold');

  // 2) 드롭다운 검증
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(UNIT_OPTIONS, true).setAllowInvalid(false).build();
  s.getRange('E2:E1000').setDataValidation(rule);

  // 3) 빈 Unit 칸을 'EA' 로 일괄 채움 (Item Name 이 있는 행만)
  const last = s.getLastRow();
  let filled = 0;
  if (last >= 2) {
    const names = s.getRange(2, 3, last - 1, 1).getValues(); // C: Item Name
    const units = s.getRange(2, 5, last - 1, 1).getValues();  // E: Unit
    for (let i = 0; i < units.length; i++) {
      if (String(names[i][0]).trim() !== '' && String(units[i][0]).trim() === '') {
        units[i][0] = 'EA';
        filled++;
      }
    }
    if (filled > 0) s.getRange(2, 5, units.length, 1).setValues(units);
  }

  // 4) Inventory / Dashboard 갱신
  rebuildInv_(ss);
  const dash = ss.getSheetByName('Dashboard');
  if (dash) setupDashboardSheet_(dash);

  SpreadsheetApp.flush();
  ui.alert('✅ Unit 도입/갱신 완료\n' +
           '- 헤더 및 드롭다운 설정\n' +
           '- 빈 칸 ' + filled + '개를 EA 로 채움\n' +
           '- Inventory / Dashboard 갱신\n\n' +
           'roll/box 등 예외 품목만 Settings_Item E열에서 바꿔 주세요.');
}

/**
 * [메뉴] 투어 위치(Settings_Location B열) 중 현재고가 0인 항목만 안전하게 삭제한다.
 * - 재고가 남은 투어 위치는 그대로 두고 경고로 알려준다(재고 orphan 방지).
 * - 삭제 후 남은 항목을 위로 압축하고 Dashboard 를 갱신한다.
 * - 재고가 남은 위치는 이동/폐기로 비운 뒤 다시 실행하면 삭제된다. (중복 실행 안전)
 */
function cleanupTourLocations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const loc = getRequiredSheet_(ss, 'Settings_Location');
  if (!loc) return;

  // 최신 재고 반영 후, 위치별 현재고 합계 집계
  rebuildInv_(ss);
  const inv = ss.getSheetByName('Inventory');
  const stockByLoc = {};
  if (inv) {
    const invLast = inv.getLastRow();
    if (invLast >= 2) {
      const rows = inv.getRange(2, 4, invLast - 1, 3).getValues(); // D:Location, E:Serial, F:Quantity
      rows.forEach(function (r) {
        const l = String(r[0]).trim();
        if (l) stockByLoc[l] = (stockByLoc[l] || 0) + (Number(r[2]) || 0);
      });
    }
  }

  const last = loc.getLastRow();
  if (last < 2) { ui.alert('ℹ️ Settings_Location 에 투어 패키지가 없습니다.'); return; }

  const bVals = loc.getRange(2, 2, last - 1, 1).getValues(); // B2:B (Tour Package)
  const keep = [];
  const kept = [];
  let removed = 0;
  for (let i = 0; i < bVals.length; i++) {
    const name = String(bVals[i][0]).trim();
    if (name === '') continue;
    if ((stockByLoc[name] || 0) > 0) { keep.push([bVals[i][0]]); kept.push(name + ' (재고 ' + stockByLoc[name] + ')'); }
    else removed++;
  }

  // B열을 비우고, 유지할 항목만 위로 압축해 다시 기록
  loc.getRange(2, 2, bVals.length, 1).clearContent();
  if (keep.length > 0) loc.getRange(2, 2, keep.length, 1).setValues(keep);

  const dash = ss.getSheetByName('Dashboard');
  if (dash) setupDashboardSheet_(dash);
  SpreadsheetApp.flush();

  let msg = '✅ 투어 위치 정리 완료\n- 재고 0인 투어 위치 ' + removed + '개 삭제';
  if (kept.length > 0) {
    msg += '\n\n⚠️ 재고가 남아 유지된 위치 (' + kept.length + '개):\n- ' + kept.join('\n- ') +
           '\n\n해당 위치의 재고를 이동/폐기해 0으로 만든 뒤 다시 실행하면 삭제됩니다.';
  }
  ui.alert(msg);
}

/**
 * [메뉴] 위치 비우기(이동) — 특정 위치의 현재고 전체를 다른 위치로 MOVE 처리한다.
 * - FROM/TO 를 입력받아, Inventory 기준 FROM 위치의 모든 재고를 정상 MOVE 원장으로 기록.
 * - 시리얼 품목은 시리얼별(수량 1)로 각각 이동. 실행 전 확인창을 띄운다.
 * - 이동 후 Inventory / Dashboard 를 갱신한다. (예: G3 KR → KR OFFICE)
 */
function relocateAllStock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const fromResp = ui.prompt('위치 비우기(이동)', 'FROM(출발) 위치명을 정확히 입력하세요:', ui.ButtonSet.OK_CANCEL);
  if (fromResp.getSelectedButton() !== ui.Button.OK) return;
  const from = String(fromResp.getResponseText()).trim();
  const toResp = ui.prompt('위치 비우기(이동)', 'TO(도착) 위치명을 정확히 입력하세요:', ui.ButtonSet.OK_CANCEL);
  if (toResp.getSelectedButton() !== ui.Button.OK) return;
  const to = String(toResp.getResponseText()).trim();

  if (!from || !to) { ui.alert('❌ FROM/TO 위치를 모두 입력해야 합니다.'); return; }
  if (from === to) { ui.alert('❌ FROM 과 TO 가 동일합니다.'); return; }

  rebuildInv_(ss); // 최신 재고 반영
  const inv = ss.getSheetByName('Inventory');
  const ledger = getRequiredSheet_(ss, 'Ledger');
  if (!inv || !ledger) return;

  const invLast = inv.getLastRow();
  const rows = invLast > 1 ? inv.getRange(2, 1, invLast - 1, 6).getValues() : []; // A~F
  const moves = rows.filter(function (r) { return String(r[3]).trim() === from && (Number(r[5]) || 0) > 0; });
  if (moves.length === 0) { ui.alert('ℹ️ "' + from + '" 위치에 이동할 재고가 없습니다.'); return; }

  const summary = moves.map(function (m) {
    const ser = String(m[4]).trim();
    return '• ' + m[2] + (ser && ser !== 'N/A' ? ' [' + ser + ']' : '') + ' × ' + (Number(m[5]) || 0);
  }).join('\n');
  const confirm = ui.alert('이동 확인',
    from + ' → ' + to + ' 로 아래 재고를 이동합니다 (' + moves.length + '건):\n\n' + summary,
    ui.ButtonSet.OK_CANCEL);
  if (confirm !== ui.Button.OK) return;

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ts = new Date();
    const newRows = moves.map(function (m) {
      const ser = String(m[4]).trim();
      return [ts, 'MOVE', m[0], m[1], m[2], (ser && ser !== '') ? ser : 'N/A',
              from, to, Number(m[5]) || 0, 'SYSTEM', '위치 정리 이동 (' + from + ' → ' + to + ')'];
    });
    ledger.insertRowsAfter(1, newRows.length);
    ledger.getRange(2, 1, newRows.length, 11).clearFormat(); // 헤더 서식 상속 제거
    ledger.getRange(2, 1, newRows.length, 11).setValues(newRows);
    ledger.getRange(2, 1, newRows.length, 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');
    ledger.getRange(2, LEDGER_COL.SERIAL + 1, newRows.length, 1).setNumberFormat('@');

    rebuildInv_(ss);
    const dash = ss.getSheetByName('Dashboard');
    if (dash) setupDashboardSheet_(dash);
    SpreadsheetApp.flush();
    ui.alert('✅ 완료: ' + newRows.length + '건을 ' + from + ' → ' + to + ' 로 이동했습니다.');
  } catch (e) {
    ui.alert('❌ 오류: ' + e.toString());
  } finally {
    lock.releaseLock();
  }
}

/**
 * [메뉴] 원장(Ledger) 서식 정리 — 기존 데이터 행의 상속된 헤더 서식(굵게·회색배경)을
 * 일반 서식으로 되돌린다. 타임스탬프/시리얼의 숫자서식은 건드리지 않는다.
 */
function normalizeLedgerFormat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const ledger = getRequiredSheet_(ss, 'Ledger');
  if (!ledger) return;
  const last = ledger.getLastRow();
  if (last < 2) { ui.alert('ℹ️ 원장에 데이터가 없습니다.'); return; }

  ledger.getRange(2, 1, last - 1, 11)
    .setFontWeight('normal')
    .setFontColor(null)
    .setBackground(null); // 회색배경 제거(no-fill) — 숫자서식은 유지

  SpreadsheetApp.flush();
  ui.alert('✅ 원장 서식 정리 완료: ' + (last - 1) + '개 행을 일반 서식으로 되돌렸습니다.');
}

/**
 * [메뉴] 설치형(installable) updateDynamicUI 편집 트리거를 제거한다.
 * 폼 반응은 이제 더 빠른 단순(simple) onEdit 트리거가 처리하므로, 설치형은 중복이라 제거한다.
 */
function removeInstallableEditTriggers() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  for (let i = 0; i < triggers.length; i++) {
    const t = triggers[i];
    if (t.getHandlerFunction() === 'updateDynamicUI' && t.getEventType() === ScriptApp.EventType.ON_EDIT) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  }
  ui.alert('✅ 설치형 편집 트리거 ' + removed + '개를 제거했습니다.\n' +
           '이제 더 빠른 단순(onEdit) 트리거가 폼 반응을 처리합니다.');
}
