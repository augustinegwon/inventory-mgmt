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
  TIMESTAMP: 0, TYPE: 1, CATEGORY: 2, ITEM: 3, SERIAL: 4,
  FROM: 5, TO: 6, QTY: 7, WORKER: 8, NOTE: 9
};

const EXTERNAL_VENDOR = 'EXTERNAL (VENDOR)';
const EXTERNAL_SCRAP = 'EXTERNAL (SCRAP)';

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
    .addItem('편집 트리거 등록 (Install Trigger)', 'createTriggers')
    .addItem('시리얼 텍스트 변환 (Migrate Serials)', 'migrateSerialsToText')
    .addToUi();
}

/**
 * [빈 스프레드시트용] 시스템 전체를 처음부터 세운다.
 * 필요한 시트 5개(Transaction/Ledger/Inventory/Dashboard/Settings), 헤더, 마스터 시딩,
 * 이름범위(LIST_*), Inventory/Dashboard 수식, 입력폼을 순서대로 생성한다.
 */
function initializeSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 0) 구버전 시트 이름 정비 (이미 초기화했던 경우 대비)
  renameSheetIfNeeded_(ss, 'Input_Transaction', 'Transaction');
  renameSheetIfNeeded_(ss, 'INV', 'Inventory');

  // 1) 원장
  const ledger = getOrCreateSheet_(ss, 'Ledger');
  setupLedgerSheet_(ledger);

  // 2) 마스터(Settings) — 헤더 + (비어있으면) 마스터 데이터 시딩
  const settings = getOrCreateSheet_(ss, 'Settings');
  setupSettingsSheet_(settings);

  // 3) 이름범위 — 수식이 참조하므로 시트/데이터 준비 후 먼저 정의
  defineNamedRanges_(ss, settings);

  // 4) Inventory — Ledger를 연산해 "정규화된 재고 목록"으로 정리하는 데이터 계층
  getOrCreateSheet_(ss, 'Inventory');
  rebuildInv_(ss);

  // 5) 대시보드 — Inventory를 피벗(SUMIFS)해서 보여주는 표시 계층
  const dashboard = getOrCreateSheet_(ss, 'Dashboard');
  setupDashboardSheet_(dashboard);

  // 6) 입력폼 (이름범위/Settings 참조)
  setupInputSheet();

  // 7) 기본으로 생기는 빈 'Sheet1' 정리
  removeDefaultSheetIfEmpty_(ss);

  SpreadsheetApp.flush();
  ui.alert(
    '✅ 초기화 완료!\n\n' +
    'Transaction / Ledger / Inventory / Dashboard / Settings 시트와 이름범위(LIST_*)를 생성했습니다.\n\n' +
    '다음 단계: 메뉴 → 📦 Inventory → "편집 트리거 등록"을 한 번 실행해 주세요.'
  );
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
  const headers = ['Timestamp', 'Type', 'Category', 'Item', 'Serial Number',
                   'From', 'To', 'Quantity', 'Worker', 'Note'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.getRange('E:E').setNumberFormat('@'); // 시리얼 텍스트 고정
  sheet.setFrozenRows(1);
}

/** 마스터(Settings) 헤더 + (A2가 비어있을 때만) 마스터 데이터 시딩 */
function setupSettingsSheet_(sheet) {
  sheet.getRange('A1').setValue('Category').setFontWeight('bold');
  sheet.getRange('B1').setValue('Item').setFontWeight('bold');
  sheet.getRange('C1').setValue('Manage Serial').setFontWeight('bold');
  sheet.getRange('E1').setValue('Location').setFontWeight('bold');
  sheet.getRange('F1').setValue('Tour Package').setFontWeight('bold');
  sheet.getRange('H1').setValue('USER').setFontWeight('bold');
  sheet.setFrozenRows(1);

  // 이미 데이터가 있으면 덮어쓰지 않음
  if (sheet.getRange('A2').getValue() !== '') return;

  // [Category, Item, Manage Serial]
  const items = [
    ['SERVER', 'SWITCH (NETGEAR MS108UP)', 'YES'],
    ['SERVER', 'SWITCH (TP-LINK)', 'YES'],
    ['SERVER', 'OPTICAL CABLE', 'NO'],
    ['SERVER', 'LAN CABLE', 'NO'],
    ['SERVER', 'HDMI CABLE', 'NO'],
    ['SERVER', 'HDMI SPLITTER', 'YES'],
    ['SERVER', 'TRIPOD', 'NO'],
    ['SERVER', 'AUDIO PROCESSOR (SB1815)', 'YES'],
    ['SERVER', 'ANDROID PHONE (GALAXY A25)', 'YES'],
    ['SERVER', 'IPAD MINI', 'YES'],
    ['SERVER', 'SOUND LEVEL METER', 'YES'],
    ['SERVER', 'LIGHT METER (LX1330B)', 'YES'],
    ['SERVER', 'USB-C MICROPHONE', 'NO'],
    ['FLOOR', 'LONG CHARGING CABLE (NEW)', 'NO'],
    ['FLOOR', 'LONG CHARGING CABLE (OLD)', 'NO'],
    ['FLOOR', 'OFFICIAL CHARGING CABLE', 'NO'],
    ['FLOOR', 'HARNESS', 'NO'],
    ['FLOOR', 'ROUND CHARGER (A TYPE)', 'NO'],
    ['FLOOR', 'ROUND CHARGER (C TYPE)', 'NO'],
    ['FLOOR', 'SQUARE CHARGER', 'NO'],
    ['FLOOR', 'DAISY-CHAIN EXTENSION CORD', 'NO'],
    ['FLOOR', '3M EXTENSION CORD (4-PORT)', 'NO'],
    ['FLOOR', '1M EXTENSION CORD (10-PORT)', 'NO'],
    ['FLOOR', '15M EXTENSION CORD (1-PORT)', 'NO'],
    ['FLOOR', 'VR HYGIENE MASK', 'NO'],
    ['CONTENT UPDATE', 'USB C-C CABLE', 'NO'],
    ['CONTENT UPDATE', 'USB A-C CABLE', 'NO'],
    ['CONTENT UPDATE', 'USB-C 4-PORT HUB', 'NO'],
    ['CONTENT UPDATE', 'USB C-LAN ADAPTER', 'NO'],
    ['CONTENT UPDATE', 'USB DRIVE', 'YES'],
    ['STORAGE', 'HEADSET CASE OLD', 'NO'],
    ['STORAGE', 'HEADSET CASE NEW', 'NO'],
    ['STORAGE', 'MISC CASE OLD', 'NO'],
    ['STORAGE', 'MISC CASE NEW', 'NO'],
    ['STORAGE', '27 GAL TOTE (= 100L STORAGE BOX)', 'NO'],
    ['STORAGE', 'BIKE LOCK', 'NO'],
    ['PERIPHERAL', 'FACIAL SPACER', 'NO'],
    ['PERIPHERAL', 'HEADSET STRAP', 'NO'],
    ['PERIPHERAL', 'LENS PROTECTOR L', 'NO'],
    ['PERIPHERAL', 'LENS PROTECTOR R', 'NO'],
    ['PERIPHERAL', 'QUEST 3 CONTROLLER L&R', 'NO']
  ];
  sheet.getRange(2, 1, items.length, 3).setValues(items);

  const locations = ['KR OFFICE', 'GMP WH A', 'GMP WH B', 'GMP WH C', 'GMP WH D',
                     'GMP WH E', 'GMP WH F', 'GMP WH G', 'GMP WH H', 'GMP WH I'];
  sheet.getRange(2, 5, locations.length, 1).setValues(locations.map(function (v) { return [v]; }));

  const tours = ['G1 US1', 'G1 US2', 'G2 MX', 'G3 KR', 'G4 JP1', 'G4 JP2', 'G4 JP3',
                 'G4 JP4', 'G4 JP5', 'G4 JP6', 'G5 TW1', 'G5 TW2', 'G6 CN', 'G7 HK/MO',
                 'G8 SG', 'G9 EU1', 'G9 EU2', 'G10 TH', 'G11 PH', 'G12 EX'];
  sheet.getRange(2, 6, tours.length, 1).setValues(tours.map(function (v) { return [v]; }));

  const users = ['AUGGIE@AMAZEVR.COM', 'ALEX@AMAZEVR.COM', 'JAE@AMAZEVR.COM'];
  sheet.getRange(2, 8, users.length, 1).setValues(users.map(function (v) { return [v]; }));
}

/** 이름범위 정의(있으면 갱신). 수식이 참조하므로 필수 */
function defineNamedRanges_(ss, settings) {
  ss.setNamedRange('LIST_CATEGORY', settings.getRange('A2:A1000'));
  ss.setNamedRange('LIST_ITEM', settings.getRange('B2:B1000'));
  ss.setNamedRange('LIST_BUCKET', settings.getRange('E2:F1000')); // 물리 위치 + 투어
}

/**
 * [메뉴] Inventory 시트를 현재 Ledger 기준으로 다시 계산한다.
 */
function rebuildInv() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  if (!ss.getSheetByName('Ledger') || !ss.getSheetByName('Settings')) {
    ui.alert('❌ 에러: Ledger 또는 Settings 시트를 찾을 수 없습니다. 먼저 "전체 초기화"를 실행해 주세요.');
    return;
  }
  const n = rebuildInv_(ss);
  SpreadsheetApp.flush();
  ui.alert('✅ Inventory 새로고침 완료: 재고 항목 ' + n + '건을 정리했습니다.');
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
  const settings = ss.getSheetByName('Settings');
  const inv = getOrCreateSheet_(ss, 'Inventory');
  if (!ledger || !settings) return 0;

  // Settings 에서 품목 → Category / 시리얼관리여부 맵 구성
  const catOf = {};
  const serialOf = {};
  const sLast = settings.getLastRow();
  if (sLast >= 2) {
    const sv = settings.getRange(2, 1, sLast - 1, 3).getValues(); // A:C
    for (let i = 0; i < sv.length; i++) {
      const item = sv[i][1];
      if (item !== '' && item != null) {
        catOf[item] = sv[i][0];
        serialOf[item] = String(sv[i][2]).toUpperCase() === 'YES';
      }
    }
  }

  // Ledger 를 훑어 (품목|위치|시리얼) 별 순재고 집계
  const rows = getLedgerRows_(ledger);
  const map = {};
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const item = r[LEDGER_COL.ITEM];
    const qty = Number(r[LEDGER_COL.QTY]) || 0;
    const isSer = !!serialOf[item];
    const serial = isSer ? normSerial_(r[LEDGER_COL.SERIAL]) : '';
    const to = r[LEDGER_COL.TO];
    const from = r[LEDGER_COL.FROM];
    if (!isExternal_(to)) addInvQty_(map, catOf, item, to, serial, qty);
    if (!isExternal_(from)) addInvQty_(map, catOf, item, from, serial, -qty);
  }

  // 재고 > 0 인 항목만 출력, Category→Item→Location→Serial 순 정렬
  const out = [];
  const keys = Object.keys(map);
  for (let i = 0; i < keys.length; i++) {
    const e = map[keys[i]];
    if (e.qty > 0) out.push([e.cat, e.item, e.loc, e.serial, e.qty]);
  }
  out.sort(function (a, b) {
    return String(a[0]).localeCompare(String(b[0])) ||
           String(a[1]).localeCompare(String(b[1])) ||
           String(a[2]).localeCompare(String(b[2])) ||
           String(a[3]).localeCompare(String(b[3]));
  });

  // Inventory 시트에 기록 (헤더 + 데이터)
  inv.clearContents();
  const headers = ['Category', 'Item', 'Location', 'Serial Number', 'Quantity'];
  inv.getRange(1, 1, 1, 5).setValues([headers]).setFontWeight('bold');
  inv.getRange('D:D').setNumberFormat('@'); // 시리얼 텍스트 고정
  if (out.length > 0) {
    inv.getRange(2, 1, out.length, 5).setValues(out);
  }
  inv.setFrozenRows(1);
  return out.length;
}

/** rebuildInv_ 내부: (품목|위치|시리얼) 키에 수량 누적 */
function addInvQty_(map, catOf, item, loc, serial, delta) {
  const key = item + ' ' + loc + ' ' + serial;
  if (!map[key]) {
    map[key] = { cat: catOf[item] || '', item: item, loc: loc, serial: serial, qty: 0 };
  }
  map[key].qty += delta;
}

/**
 * Dashboard — Inventory(정규화 재고 목록)를 SUMIFS 로 피벗해 "아이템 × 위치" 매트릭스로 보여주는 표시 계층.
 * 계산 원천은 Ledger 가 아니라 Inventory. (Ledger ▶ Inventory ▶ Dashboard 파이프라인)
 */
function setupDashboardSheet_(sheet) {
  sheet.getRange('A1').setValue('📊 [Inventory Dashboard]').setFontWeight('bold');
  sheet.getRange('A2').setValue('Item').setFontWeight('bold');
  sheet.getRange('B2').setValue('Total').setFontWeight('bold');

  // A3: 아이템 목록(세로 spill)
  sheet.getRange('A3').setFormula('=FILTER(LIST_ITEM, LIST_ITEM<>"")');

  // B3: 아이템별 총 수량 = Inventory 수량 합계
  sheet.getRange('B3').setFormula(
    '=IFERROR(LET(items, FILTER(LIST_ITEM, LIST_ITEM<>""), ' +
    'BYROW(items, LAMBDA(i, SUMIFS(Inventory!$E:$E, Inventory!$B:$B, i)))), "")'
  );

  // C2: 버킷 헤더(가로 spill)
  sheet.getRange('C2').setFormula('=IFERROR(TRANSPOSE(TOCOL(LIST_BUCKET, 1, TRUE)), "")');

  // C3: 아이템 × 버킷 매트릭스 = Inventory 를 (Item, Location) 으로 피벗
  sheet.getRange('C3').setFormula(
    '=IFERROR(LET(items, FILTER(LIST_ITEM, LIST_ITEM<>""), buckets, TOCOL(LIST_BUCKET, 1, TRUE), ' +
    'MAKEARRAY(ROWS(items), ROWS(buckets), LAMBDA(r, c, ' +
    'SUMIFS(Inventory!$E:$E, Inventory!$B:$B, INDEX(items, r), Inventory!$C:$C, INDEX(buckets, c))))), "")'
  );

  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(2); // A(Item) + B(Total) 고정
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
 * (현재 배포된 시트 상태 — 검색 영역 + LET/LAMBDA 헬퍼 수식 — 를 재현하도록 갱신)
 */
function setupInputSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 필수 시트 존재 여부를 먼저 확인 (없으면 아무것도 건드리지 않고 중단)
  const settingsSheet = getRequiredSheet_(ss, 'Settings');
  if (!settingsSheet) return;

  let inputSheet = ss.getSheetByName('Transaction');
  if (!inputSheet) {
    inputSheet = ss.insertSheet('Transaction');
  }
  // 기존의 유효성 검사 및 내용 전체 초기화 (충돌 방지)
  // ※ clearDataValidations() 는 Range 메서드이므로 시트 전체 범위에 적용해야 함
  inputSheet.getRange(1, 1, inputSheet.getMaxRows(), inputSheet.getMaxColumns())
    .clearDataValidations();
  inputSheet.clear();

  // 기본 가이드 및 텍스트 UI 배치 — 검색 영역
  inputSheet.getRange('B2').setValue('🔍 Item Search').setFontWeight('bold');
  inputSheet.getRange('B3').setValue('Category');
  inputSheet.getRange('B4').setValue('Item');
  inputSheet.getRange('B6').setValue('Location').setFontWeight('bold');
  inputSheet.getRange('C6').setValue('Current Qty').setFontWeight('bold');

  // 트랜잭션 폼
  inputSheet.getRange('E2').setValue('📦 Transaction Form').setFontWeight('bold');
  inputSheet.getRange('E3').setValue('Type');
  inputSheet.getRange('E4').setValue('Item');
  inputSheet.getRange('E5').setValue('SERIAL NO.');
  inputSheet.getRange('E6').setValue('FROM');
  inputSheet.getRange('E7').setValue('TO');
  inputSheet.getRange('E8').setValue('Quantity');
  inputSheet.getRange('E9').setValue('USER');
  inputSheet.getRange('E10').setValue('Note');

  // 선택 아이템의 위치별 현재고 요약 (B7 앵커, {버킷, 잔고} 2열 spill)
  inputSheet.getRange('B7').setFormula(
    '=IFERROR(LET(item, $C$4, buckets, TOCOL(LIST_BUCKET, 1, TRUE), ' +
    'balances, BYROW(buckets, LAMBDA(b, SUMIFS(Ledger!$H:$H, Ledger!$D:$D, item, Ledger!$G:$G, b) - SUMIFS(Ledger!$H:$H, Ledger!$D:$D, item, Ledger!$F:$F, b))), ' +
    'FILTER({buckets, balances}, balances <> 0)), "Item not selected or out of stock.")'
  );

  // 트랜잭션 폼의 Item 은 검색 Item(C4) 을 미러링
  inputSheet.getRange('F4').setFormula('=$C$4');

  // 헬퍼 열(Helper Columns) — 현재 배포본 수식
  // W1: MOVE/REMOVE 시 해당 아이템의 재고가 있는 '버킷' 목록 (재고>0)
  inputSheet.getRange('W1').setFormula(
    '=IFERROR(UNIQUE(FILTER(TOCOL(LIST_BUCKET, 1, TRUE), ' +
    'BYROW(TOCOL(LIST_BUCKET, 1, TRUE), LAMBDA(b, SUMIFS(Ledger!$H:$H, Ledger!$D:$D, $C$4, Ledger!$G:$G, b) - SUMIFS(Ledger!$H:$H, Ledger!$D:$D, $C$4, Ledger!$F:$F, b))) > 0)), "")'
  );
  // X1: MOVE/REMOVE 시 FROM(F6) 위치에 재고가 남아있는 '시리얼 번호' 목록
  inputSheet.getRange('X1').setFormula(
    '=IFERROR(LET(i, $C$4, loc, $F$6, ' +
    'sers, UNIQUE(FILTER(Ledger!E:E, (Ledger!D:D=i)*(Ledger!E:E<>"")*(Ledger!E:E<>"N/A"))), ' +
    'bals, BYROW(sers, LAMBDA(s, SUMIFS(Ledger!H:H, Ledger!D:D, i, Ledger!E:E, s, Ledger!G:G, loc) - SUMIFS(Ledger!H:H, Ledger!D:D, i, Ledger!E:E, s, Ledger!F:F, loc))), ' +
    'FILTER(sers, bals>0)), "")'
  );
  // Y1: TO 드롭다운을 위한 전체 버킷 목록
  inputSheet.getRange('Y1').setFormula('=IFERROR(TOCOL(LIST_BUCKET, 1, TRUE), "")');
  // Z1: 검색 Item 드롭다운 (Category 로 필터)
  inputSheet.getRange('Z1').setFormula(
    '=IFERROR(IF(C3="", FILTER(LIST_ITEM, LIST_ITEM<>""), FILTER(LIST_ITEM, LIST_CATEGORY=C3, LIST_ITEM<>"")), "")'
  );

  // 🌟 유효성 검사 전에 통과 가능한 기본값 먼저 입력
  inputSheet.getRange('F3').setValue('ADD'); // Type 기본값
  inputSheet.getRange('F5').setNumberFormat('@').setValue('N/A'); // Serial No 기본값(텍스트 형식 고정)
  inputSheet.getRange('F8').setValue(1);     // Quantity 기본값

  // 시리얼이 숫자로 저장되어 텍스트 시리얼과 매칭이 깨지는 문제 방지 → 원장 시리얼 열 텍스트 고정
  const ledgerSheet = ss.getSheetByName('Ledger');
  if (ledgerSheet) {
    ledgerSheet.getRange('E:E').setNumberFormat('@');
  }

  // 드롭다운 규칙 정의 (데이터 유효성 검사)
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['ADD', 'MOVE', 'REMOVE'], true)
    .setAllowInvalid(false)
    .build();

  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(settingsSheet.getRange('A2:A'), true)
    .setAllowInvalid(false)
    .build();

  const searchItemRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(inputSheet.getRange('Z1:Z'), true)
    .setAllowInvalid(false)
    .build();

  const toLocRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(inputSheet.getRange('Y1:Y'), true)
    .setAllowInvalid(false)
    .build();

  // USER 드롭다운 (Settings!H 마스터). 신규 사용자도 허용하도록 경고만(allowInvalid=true)
  const userRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(settingsSheet.getRange('H2:H'), true)
    .setAllowInvalid(true)
    .build();

  // 유효성 검사 규칙 적용
  inputSheet.getRange('F3').setDataValidation(typeRule);      // Type 셀
  inputSheet.getRange('C3').setDataValidation(categoryRule);  // 검색 Category
  inputSheet.getRange('C4').setDataValidation(searchItemRule);// 검색 Item
  inputSheet.getRange('F7').setDataValidation(toLocRule);     // TO 셀
  inputSheet.getRange('F9').setDataValidation(userRule);      // USER 셀

  // UI 디자인 정리 (배경색 설정)
  inputSheet.getRange('C3:C4').setBackground('#fff2cc'); // 검색 창
  inputSheet.getRange('F3:F4').setBackground('#e2efda'); // 트랜잭션 창
  inputSheet.getRange('F5:F10').setBackground('#e2efda');
  SpreadsheetApp.flush(); // 시트에 즉시 반영
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
