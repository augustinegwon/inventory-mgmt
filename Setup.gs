/**
 * Inventory Management — Setup.gs
 * 설정 / 설치 / 유지보수 함수 + 공용 상수
 *
 * 시트 구성:
 *   - Input_Transaction : 입력 폼 UI (검색 영역 + 트랜잭션 폼 + 헬퍼 열 W/X/Y/Z)
 *   - Ledger            : 거래 원장 (append-only, 재고의 단일 진실 공급원)
 *   - Dashboard         : 아이템 × 위치 재고 매트릭스 (수식으로 자동 계산)
 *   - Settings          : 마스터 데이터 (Category/Item/Serial여부, Location, Tour, USER)
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
    .addItem('입력폼 초기 설정 (Setup)', 'setupInputSheet')
    .addItem('편집 트리거 등록 (Install Trigger)', 'createTriggers')
    .addItem('시리얼 텍스트 변환 (Migrate Serials)', 'migrateSerialsToText')
    .addToUi();
}

/**
 * 1. Input_Transaction 시트의 UI 구조 및 수식을 초기화하는 함수
 * (현재 배포된 시트 상태 — 검색 영역 + LET/LAMBDA 헬퍼 수식 — 를 재현하도록 갱신)
 */
function setupInputSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 필수 시트 존재 여부를 먼저 확인 (없으면 아무것도 건드리지 않고 중단)
  const settingsSheet = getRequiredSheet_(ss, 'Settings');
  if (!settingsSheet) return;

  let inputSheet = ss.getSheetByName('Input_Transaction');
  if (!inputSheet) {
    inputSheet = ss.insertSheet('Input_Transaction');
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
