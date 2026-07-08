/**
 * Inventory Management — Google Apps Script
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
 */

// Ledger 열 인덱스 (0-based)
const LEDGER_COL = {
  TIMESTAMP: 0, TYPE: 1, CATEGORY: 2, ITEM: 3, SERIAL: 4,
  FROM: 5, TO: 6, QTY: 7, WORKER: 8, NOTE: 9
};

const EXTERNAL_VENDOR = 'EXTERNAL (VENDOR)';
const EXTERNAL_SCRAP = 'EXTERNAL (SCRAP)';

/**
 * 1. Input_Transaction 시트의 UI 구조 및 수식을 초기화하는 함수
 * (현재 배포된 시트 상태 — 검색 영역 + LET/LAMBDA 헬퍼 수식 — 를 재현하도록 갱신)
 */
function setupInputSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let inputSheet = ss.getSheetByName('Input_Transaction');
  if (!inputSheet) {
    inputSheet = ss.insertSheet('Input_Transaction');
  }
  // 기존의 유효성 검사 및 내용 전체 초기화 (충돌 방지)
  inputSheet.clearDataValidations();
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
    .requireValueInRange(ss.getSheetByName('Settings').getRange('A2:A'), true)
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
    .requireValueInRange(ss.getSheetByName('Settings').getRange('H2:H'), true)
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
 * 2. 사용자가 Type 이나 Item 을 바꿨을 때 시리얼 번호 및 수량 칸을 제어하는 동적 UI 함수
 */
function updateDynamicUI(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== 'Input_Transaction') return;

  const cellA1 = range.getA1Notation();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Input_Transaction');

  // 시리얼(F5) 직접 입력값은 공백 제거 + 대문자로 정규화 (텍스트 유지, 매칭 정합성 보장)
  if (cellA1 === 'F5') {
    const rawValue = range.getValue();
    if (typeof rawValue === 'string') {
      const norm = rawValue.trim().toUpperCase();
      if (norm !== rawValue) range.setValue(norm);
    }
  }

  const type = inputSheet.getRange('F3').getValue();
  const itemName = inputSheet.getRange('F4').getValue();
  const serialCell = inputSheet.getRange('F5');
  const fromLocCell = inputSheet.getRange('F6');  // FROM 셀
  const qtyCell = inputSheet.getRange('F8');

  // 아이템의 시리얼 관리 여부 확인 (Settings 시트 참조)
  const isSerial = isSerialManaged_(ss, itemName);

  // ★ FROM 드롭다운 동적 제어: MOVE/REMOVE 일 때는 재고가 있는 위치만 필터링 (W열 참조)
  const locRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(inputSheet.getRange('W1:W'))
    .setAllowInvalid(false).build();

  if (type === 'MOVE' || type === 'REMOVE') {
    fromLocCell.setDataValidation(locRule);
  } else {
    fromLocCell.clearDataValidations(); // ADD 시에는 제약 해제
  }

  // 시리얼 관리 여부 및 거래 조건에 따른 UI 제어 로직
  if (isSerial === 'YES') {
    qtyCell.setValue(1);
    qtyCell.setBackground('#e1e1e1').setFontColor('#7f8c8d'); // 수량 칸 비활성화 (시리얼은 무조건 1개)
    serialCell.setNumberFormat('@'); // 시리얼 텍스트 형식 고정

    if (type === 'ADD') {
      serialCell.clearDataValidations(); // 새 시리얼 입력 가능하게 해제
      if (serialCell.getValue() === 'N/A') {
        serialCell.setValue('');
      }
      serialCell.setBackground('#ffffff').setFontColor('#000000');
    } else if (type === 'MOVE' || type === 'REMOVE') {
      // 헬퍼 열 X열(재고가 있는 시리얼 목록)을 참조하여 드롭다운 생성
      const serialRule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(inputSheet.getRange('X1:X'))
        .setAllowInvalid(false).build();
      serialCell.setDataValidation(serialRule);
      serialCell.setBackground('#fff2cc').setFontColor('#000000');
    }
  } else {
    // 시리얼 관리를 안 하는 품목일 경우
    serialCell.clearDataValidations();
    serialCell.setValue('N/A');
    serialCell.setBackground('#e1e1e1').setFontColor('#7f8c8d'); // 시리얼 칸 비활성화

    qtyCell.setBackground('#ffffff').setFontColor('#000000'); // 수량 입력창 활성화
    if (qtyCell.getValue() == 1 && cellA1 === 'F4') {
      qtyCell.setValue('');
    }
  }
}

/**
 * 3. [제출] 버튼을 눌렀을 때 입력 데이터를 원장(Ledger)에 기록하는 함수
 */
function submitTransaction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Input_Transaction');
  const ledgerSheet = ss.getSheetByName('Ledger');
  const settingsSheet = ss.getSheetByName('Settings');
  const ui = SpreadsheetApp.getUi();

  // 입력 필드 값 가져오기
  const type = inputSheet.getRange('F3').getValue();
  const item = inputSheet.getRange('F4').getValue();
  const rawSerial = inputSheet.getRange('F5').getValue();
  const fromLoc = inputSheet.getRange('F6').getValue();
  const toLoc = inputSheet.getRange('F7').getValue();
  const rawQty = inputSheet.getRange('F8').getValue();
  const worker = inputSheet.getRange('F9').getValue();
  const note = inputSheet.getRange('F10').getValue();

  // 필수값 누락 검증 (ADD 검증 강화)
  if (!type || !item || !worker) {
    ui.alert('❌ 에러: Type, Item, USER 필드는 필수 입력 항목입니다.');
    return;
  }

  // M3. 수량은 양의 정수여야 함
  const qty = Number(rawQty);
  if (!(qty > 0) || Math.floor(qty) !== qty) {
    ui.alert('❌ 에러: Quantity 는 0보다 큰 정수여야 합니다. (입력값: ' + rawQty + ')');
    return;
  }

  if ((type === 'MOVE' || type === 'REMOVE') && !fromLoc) {
    ui.alert('❌ 에러: MOVE 또는 REMOVE 처리 시 FROM(출발지) 위치를 지정해야 합니다.');
    return;
  }

  if ((type === 'ADD' || type === 'MOVE') && !toLoc) {
    ui.alert('❌ 에러: ADD 또는 MOVE 처리 시 TO(목적지) 위치를 지정해야 합니다.');
    return;
  }

  // M2. MOVE 에서 출발지 = 목적지 는 의미 없는 기록
  if (type === 'MOVE' && fromLoc === toLoc) {
    ui.alert('❌ 에러: FROM 과 TO 가 동일합니다. 서로 다른 위치를 선택해 주세요.');
    return;
  }

  // 품목의 Category / 시리얼 관리 여부 조회
  const settingsData = settingsSheet.getRange('A2:C' + settingsSheet.getLastRow()).getValues();
  let category = '';
  let isSerial = 'NO';
  let itemExists = false;
  for (let i = 0; i < settingsData.length; i++) {
    if (settingsData[i][1] === item) {
      category = settingsData[i][0];
      isSerial = (String(settingsData[i][2]).toUpperCase() === 'YES') ? 'YES' : 'NO';
      itemExists = true;
      break;
    }
  }

  if (!itemExists) {
    ui.alert('❌ 에러: Settings 마스터 정보에 등록되지 않은 품목명입니다.');
    return;
  }

  // 시리얼 정규화 (H3): 항상 공백 제거 + 대문자 텍스트. 비시리얼 품목은 N/A 고정.
  let serial = 'N/A';
  if (isSerial === 'YES') {
    serial = (rawSerial === '' || rawSerial == null) ? '' : String(rawSerial).trim().toUpperCase();
    if (!serial || serial === 'N/A') {
      ui.alert('❌ 에러: 시리얼 관리 품목은 SERIAL NO. 를 반드시 입력해야 합니다.');
      return;
    }
    // 시리얼 품목은 수량 1 고정
    if (qty !== 1) {
      ui.alert('❌ 에러: 시리얼 관리 품목의 수량은 1 이어야 합니다.');
      return;
    }
  }

  // 락(Lock) 획득을 통한 동시성 제어 및 데이터 정합성 보장
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 최대 10초 대기

    // 락 안에서 최신 원장을 읽어 재고 검증 (H1/H2)
    const rows = getLedgerRows_(ledgerSheet);

    if (isSerial === 'YES') {
      if (type === 'ADD') {
        // H2. 이미 시스템 내 재고로 존재하는 시리얼 중복 ADD 차단
        if (serialInStock_(rows, item, serial) > 0) {
          ui.alert('❌ 에러: 시리얼 "' + serial + '" 은(는) 이미 재고에 존재합니다. 중복 ADD 할 수 없습니다.');
          return;
        }
      } else { // MOVE / REMOVE
        // H2. 해당 시리얼이 FROM 위치에 실제로 있는지 확인
        if (bucketBalance_(rows, item, fromLoc, serial) <= 0) {
          ui.alert('❌ 에러: 시리얼 "' + serial + '" 은(는) "' + fromLoc + '" 위치에 재고가 없습니다.');
          return;
        }
      }
    } else if (type === 'MOVE' || type === 'REMOVE') {
      // H1. 비시리얼 품목: FROM 위치의 현재고보다 많은 수량 출고 차단
      const available = bucketBalance_(rows, item, fromLoc, null);
      if (qty > available) {
        ui.alert('❌ 에러: 재고 부족. "' + item + '" @ "' + fromLoc +
                 '" 현재고 ' + available + ' 개인데 ' + qty + ' 개를 출고할 수 없습니다.');
        return;
      }
    }

    // Ledger에 기록할 최종 행 데이터 배열 생성
    const timestamp = new Date();
    const finalFrom = (type === 'ADD') ? EXTERNAL_VENDOR : fromLoc;
    const finalTo = (type === 'REMOVE') ? EXTERNAL_SCRAP : toLoc;

    // 원장 데이터 추가 실행 (시리얼은 텍스트로 강제하기 위해 setValues 사용)
    const newRow = ledgerSheet.getLastRow() + 1;
    ledgerSheet.getRange(newRow, 1, 1, 10).setValues([[
      timestamp, type, category, item, serial,
      finalFrom, finalTo, qty, worker, note
    ]]);
    ledgerSheet.getRange(newRow, LEDGER_COL.SERIAL + 1).setNumberFormat('@'); // 시리얼 텍스트 고정

    // 제출 완료 후 입력 폼 초기화 (Type, User 정보는 편의상 유지)
    // F4 는 '=$C$4' 수식이므로 건드리지 않고 검색 Item(C4)만 비운다
    inputSheet.getRange('C4').setValue(''); // 검색 Item 초기화 → F4 자동으로 빈값
    inputSheet.getRange('F5').setValue('N/A'); // Serial 초기화
    inputSheet.getRange('F6').setValue(''); // FROM 초기화
    inputSheet.getRange('F7').setValue(''); // TO 초기화
    inputSheet.getRange('F8').setValue(1);  // 수량 초기화
    inputSheet.getRange('F10').setValue(''); // Note 초기화

    SpreadsheetApp.flush();
    ui.alert('✅ 성공: 트랜잭션이 원장(Ledger)에 정상적으로 기록되었습니다.');

  } catch (error) {
    ui.alert('❌ 오류 발생: 다른 사용자가 작업 중이거나 시스템 지연이 발생했습니다. 다시 시도해 주세요.\n' + error.toString());
  } finally {
    lock.releaseLock(); // 반드시 락 해제
  }
}

/* ────────────────────────────────────────────────────────────
 * 재고 계산 헬퍼
 * ──────────────────────────────────────────────────────────── */

/** 원장 데이터 행 배열(헤더 제외)을 반환 */
function getLedgerRows_(ledgerSheet) {
  const last = ledgerSheet.getLastRow();
  if (last < 2) return [];
  return ledgerSheet.getRange(2, 1, last - 1, 10).getValues();
}

/** EXTERNAL (VENDOR/SCRAP) 등 시스템 외부 버킷 여부 */
function isExternal_(bucket) {
  return String(bucket).indexOf('EXTERNAL') === 0;
}

/** 시리얼 값 정규화 (공백 제거 + 대문자 텍스트) */
function normSerial_(v) {
  return String(v).trim().toUpperCase();
}

/**
 * 특정 아이템의 특정 버킷 현재고.
 * serial 이 주어지면 해당 시리얼만, null 이면 아이템 전체(수량 합산).
 */
function bucketBalance_(rows, item, bucket, serial) {
  const s = serial == null ? null : normSerial_(serial);
  let bal = 0;
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (r[LEDGER_COL.ITEM] !== item) continue;
    if (s !== null && normSerial_(r[LEDGER_COL.SERIAL]) !== s) continue;
    const q = Number(r[LEDGER_COL.QTY]) || 0;
    if (r[LEDGER_COL.TO] === bucket) bal += q;
    if (r[LEDGER_COL.FROM] === bucket) bal -= q;
  }
  return bal;
}

/**
 * 특정 시리얼이 시스템 내(실제 버킷)에 남아있는 총 재고.
 * >0 이면 어딘가에 재고로 존재. (EXTERNAL 버킷은 제외)
 */
function serialInStock_(rows, item, serial) {
  const s = normSerial_(serial);
  let bal = 0;
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (r[LEDGER_COL.ITEM] !== item) continue;
    if (normSerial_(r[LEDGER_COL.SERIAL]) !== s) continue;
    const q = Number(r[LEDGER_COL.QTY]) || 0;
    if (!isExternal_(r[LEDGER_COL.TO])) bal += q;
    if (!isExternal_(r[LEDGER_COL.FROM])) bal -= q;
  }
  return bal;
}

/** Settings 에서 아이템의 시리얼 관리 여부('YES'/'NO') 조회 */
function isSerialManaged_(ss, itemName) {
  const settingsSheet = ss.getSheetByName('Settings');
  const last = settingsSheet.getLastRow();
  if (last < 2) return 'NO';
  const mapping = settingsSheet.getRange(2, 2, last - 1, 2).getValues(); // B:C
  for (let i = 0; i < mapping.length; i++) {
    if (mapping[i][0] === itemName) {
      return String(mapping[i][1]).toUpperCase() === 'YES' ? 'YES' : 'NO';
    }
  }
  return 'NO';
}
