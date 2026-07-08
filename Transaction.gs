/**
 * Inventory Management — Transaction.gs
 * 거래 처리(동적 UI + 제출) + 재고 계산 헬퍼
 *
 * 공용 상수(LEDGER_COL, EXTERNAL_VENDOR, EXTERNAL_SCRAP)는 Setup.gs 에 정의되어 있으며
 * Apps Script 전역 범위를 공유하므로 이 파일에서 그대로 사용한다.
 */

/**
 * 사용자가 Type 이나 Item 을 바꿨을 때 시리얼 번호 및 수량 칸을 제어하는 동적 UI 함수.
 * (Transaction 시트의 '수정 시(onEdit)' 트리거로 연결되어야 동작 — Setup.gs 의 createTriggers 참고)
 */
function updateDynamicUI(e) {
  // 이벤트 없이(편집기에서 직접) 호출되면 e.range 접근에서 크래시하므로 방어
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== 'Transaction') return;

  const cellA1 = range.getA1Notation();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Transaction');

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
 * [제출] 버튼을 눌렀을 때 입력 데이터를 원장(Ledger)에 기록하는 함수.
 */
function submitTransaction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 필수 시트 확인 (없으면 명확한 안내 후 중단)
  const inputSheet = getRequiredSheet_(ss, 'Transaction');
  if (!inputSheet) return;
  const ledgerSheet = getRequiredSheet_(ss, 'Ledger');
  if (!ledgerSheet) return;
  const settingsSheet = getRequiredSheet_(ss, 'Settings');
  if (!settingsSheet) return;

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

    // 원장 기록 후 Inventory(정규화 재고 목록) 재계산 → Dashboard 는 Inventory 를 참조하므로 자동 갱신
    rebuildInv_(ss);

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
  if (!settingsSheet) return 'NO'; // Settings 시트가 없으면 조용히 기본값
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
