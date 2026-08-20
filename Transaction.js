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
/**
 * 🌟 [수정] Settings_Item 시트 단독 자동 채번
 */
function updateDynamicUI(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  // 🌟 시트명 Settings_Item 으로 변경
  if (sheetName === 'Settings_Item') {
    if ((range.getColumn() === 2 || range.getColumn() === 3) && range.getRow() > 1) {
      const idCell = sheet.getRange(range.getRow(), 1); 
      if (idCell.getValue() === '') {
        const lastRow = sheet.getLastRow();
        const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
        let maxNum = 1000;
        for (let i = 0; i < ids.length; i++) {
          const match = String(ids[i]).match(/ITM-(\d+)/);
          if (match) {
            const num = parseInt(match[1], 10);
            if (num > maxNum) maxNum = num;
          }
        }
        idCell.setValue('ITM-' + (maxNum + 1));
      }
    }
    return;
  }

  if (sheetName !== 'Transaction') return;

  const col = range.getColumn();
  const row = range.getRow();
  if (!((col === 3 && row === 4) || (col === 6 && (row === 3 || row === 5)))) return;

  const inputSheet = sheet;

  if (range.getA1Notation() === 'F5') {
    const rawValue = e.value; 
    if (typeof rawValue === 'string') {
      const norm = rawValue.trim().toUpperCase();
      if (norm !== rawValue) range.setValue(norm);
    }
  }

  const formValues = inputSheet.getRange('F3:F8').getValues();
  const type = formValues[0][0];      
  const qtyVal = formValues[5][0];    
  const isSerial = inputSheet.getRange('AA1').getValue(); 

  const serialCell = inputSheet.getRange('F5');
  const fromLocCell = inputSheet.getRange('F6');
  const toLocCell = inputSheet.getRange('F7'); 
  const qtyCell = inputSheet.getRange('F8');

  if (type === 'MOVE' || type === 'REMOVE') {
    const locRule = SpreadsheetApp.newDataValidation().requireValueInRange(inputSheet.getRange('W1:W')).setAllowInvalid(false).build();
    fromLocCell.setDataValidation(locRule);
    fromLocCell.setBackground('#e2efda').setFontColor('#000000');
    if (fromLocCell.getValue() === 'N/A') fromLocCell.setValue('');
  } else { 
    fromLocCell.clearDataValidations().setValue('N/A').setBackground('#e1e1e1').setFontColor('#7f8c8d');
  }

  if (type === 'ADD' || type === 'MOVE') {
    const toLocRule = SpreadsheetApp.newDataValidation().requireValueInRange(inputSheet.getRange('Y1:Y')).setAllowInvalid(false).build();
    toLocCell.setDataValidation(toLocRule);
    toLocCell.setBackground('#e2efda').setFontColor('#000000');
    if (toLocCell.getValue() === 'N/A') toLocCell.setValue('');
  } else { 
    toLocCell.clearDataValidations().setValue('N/A').setBackground('#e1e1e1').setFontColor('#7f8c8d');
  }

  if (isSerial === 'YES') {
    qtyCell.setValue(1).setBackground('#e1e1e1').setFontColor('#7f8c8d');
    serialCell.setNumberFormat('@');
    if (type === 'ADD') {
      serialCell.clearDataValidations();
      if (serialCell.getValue() === 'N/A') serialCell.setValue('');
      serialCell.setBackground('#e2efda').setFontColor('#000000');
    } else if (type === 'MOVE' || type === 'REMOVE') {
      const serialRule = SpreadsheetApp.newDataValidation().requireValueInRange(inputSheet.getRange('X1:X')).setAllowInvalid(false).build();
      serialCell.setDataValidation(serialRule);
      serialCell.setBackground('#fff2cc').setFontColor('#000000');
    }
  } else {
    serialCell.clearDataValidations().setValue('N/A').setBackground('#e1e1e1').setFontColor('#7f8c8d');
    qtyCell.setBackground('#e2efda').setFontColor('#000000');
    if (qtyVal == 1 && col === 3 && row === 4) qtyCell.setValue('');
  }
}

/**
 * [제출] 버튼을 눌렀을 때 입력 데이터를 원장(Ledger)에 기록하는 함수.
 */
/**
 * 🌟 [수정] 트랜잭션 제출 시 Settings_Item 참조
 */
function submitTransaction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const inputSheet = getRequiredSheet_(ss, 'Transaction');
  if (!inputSheet) return;
  const ledgerSheet = getRequiredSheet_(ss, 'Ledger');
  if (!ledgerSheet) return;
  const settingsSheet = getRequiredSheet_(ss, 'Settings_Item'); // 🌟 시트명 변경
  if (!settingsSheet) return;

  const formValues = inputSheet.getRange('F3:F10').getValues();
  const type = formValues[0][0];      
  const itemName = formValues[1][0];  
  const rawSerial = formValues[2][0]; 
  const fromLoc = formValues[3][0];   
  const toLoc = formValues[4][0];     
  const rawQty = formValues[5][0];    
  const worker = formValues[6][0];    
  const note = formValues[7][0];      

  if (!type || !itemName || !worker) {
    ui.alert('❌ 에러: Type, Item, USER 필드는 필수 입력 항목입니다.');
    return;
  }

  const qty = Number(rawQty);
  if (!(qty > 0) || Math.floor(qty) !== qty) {
    ui.alert('❌ 에러: Quantity 는 0보다 큰 정수여야 합니다.');
    return;
  }

  if ((type === 'MOVE' || type === 'REMOVE') && !fromLoc) {
    ui.alert('❌ 에러: MOVE/REMOVE 시 FROM 위치를 지정해야 합니다.');
    return;
  }

  if ((type === 'ADD' || type === 'MOVE') && !toLoc) {
    ui.alert('❌ 에러: ADD/MOVE 시 TO 위치를 지정해야 합니다.');
    return;
  }

  if (type === 'MOVE' && fromLoc === toLoc) {
    ui.alert('❌ 에러: FROM과 TO가 동일합니다.');
    return;
  }

  const settingsData = settingsSheet.getRange('A2:D' + settingsSheet.getLastRow()).getValues();
  let itemId = '';
  let category = '';
  let isSerial = 'NO';
  let itemExists = false;
  
  for (let i = 0; i < settingsData.length; i++) {
    if (settingsData[i][2] === itemName) { 
      itemId = settingsData[i][0];       
      category = settingsData[i][1];     
      isSerial = (String(settingsData[i][3]).toUpperCase() === 'YES') ? 'YES' : 'NO';
      itemExists = true;
      break;
    }
  }

  if (!itemExists || !itemId) {
    ui.alert('❌ 에러: 마스터 정보에 등록되지 않은 품목이거나 ID가 없습니다.');
    return;
  }

  let serial = 'N/A';
  if (isSerial === 'YES') {
    serial = (rawSerial === '' || rawSerial == null) ? '' : String(rawSerial).trim().toUpperCase();
    if (!serial || serial === 'N/A') {
      ui.alert('❌ 에러: 시리얼 관리 품목은 SERIAL NO. 를 반드시 입력해야 합니다.');
      return;
    }
    if (qty !== 1) {
      ui.alert('❌ 에러: 시리얼 관리 품목의 수량은 1 이어야 합니다.');
      return;
    }
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const invSheet = ss.getSheetByName('Inventory');
    const lastInvRow = invSheet ? invSheet.getLastRow() : 0;
    const invData = lastInvRow > 1 ? invSheet.getRange(2, 1, lastInvRow - 1, 6).getValues() : [];

    if (isSerial === 'YES') {
      if (type === 'ADD') {
        if (getActualSerialInStock_(invData, itemId, serial) > 0) {
          ui.alert('❌ 에러: 시리얼 "' + serial + '" 은(는) 이미 시스템 재고에 존재합니다.');
          return;
        }
      } else {
        if (getActualInventoryBalance_(invData, itemId, fromLoc, serial) <= 0) {
          ui.alert('❌ 에러: 시리얼 "' + serial + '" 은(는) "' + fromLoc + '" 위치에 재고가 없습니다.');
          return;
        }
      }
    } else if (type === 'MOVE' || type === 'REMOVE') {
      const available = getActualInventoryBalance_(invData, itemId, fromLoc, null);
      if (qty > available) {
        ui.alert('❌ 에러: 재고 부족. 현재고 ' + available + ' 개인데 ' + qty + ' 개를 출고할 수 없습니다.');
        return;
      }
    }

    const timestamp = new Date();
    const finalFrom = (type === 'ADD') ? EXTERNAL_VENDOR : fromLoc;
    const finalTo = (type === 'REMOVE') ? EXTERNAL_SCRAP : toLoc;

    ledgerSheet.insertRowAfter(1);
    const targetRange = ledgerSheet.getRange(2, 1, 1, 11);
    targetRange.clearFormat(); // 헤더(1행) 서식(굵게·회색배경) 상속 제거
    targetRange.setValues([[
      timestamp, type, category, itemId, itemName, serial,
      finalFrom, finalTo, qty, worker, note
    ]]);
    ledgerSheet.getRange(2, 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');
    ledgerSheet.getRange(2, LEDGER_COL.SERIAL + 1).setNumberFormat('@');

    inputSheet.getRangeList(['C4', 'F6', 'F7', 'F10']).clearContent();
    inputSheet.getRangeList(['F5']).setValue('N/A');
    inputSheet.getRangeList(['F8']).setValue(1);

    rebuildInv_(ss);

    ss.toast('📦 원장(Ledger)에 정상 기록되었습니다.', '제출 완료', 3);

  } catch (error) {
    ui.alert('❌ 오류 발생: ' + error.toString());
  } finally {
    lock.releaseLock();
  }
}

/* ────────────────────────────────────────────────────────────
 * 재고 계산 헬퍼 (초고속 인메모리 배열 기반 최적화 버전)
 * ──────────────────────────────────────────────────────────── */

/** * 🌟 정규화 완료된 Inventory 시트를 기준으로 특정 위치의 현재고 수량을 빠르게 조회 
 */
function getActualInventoryBalance_(invData, itemId, loc, serial) {
  let total = 0;
  const targetSerial = serial ? String(serial).trim().toUpperCase() : null;
  
  for (let i = 0; i < invData.length; i++) {
    if (invData[i][1] === itemId && invData[i][3] === loc) { 
      if (targetSerial !== null) {
        if (String(invData[i][4]).trim().toUpperCase() === targetSerial) { 
          return Number(invData[i][5]) || 0; 
        }
      } else {
        total += Number(invData[i][5]) || 0;
      }
    }
  }
  return total;
}

/** * 🌟 특정 시리얼 번호가 시스템 내에 존재하는지 판별 
 */
function getActualSerialInStock_(invData, itemId, serial) {
  const targetSerial = String(serial).trim().toUpperCase();
  
  for (let i = 0; i < invData.length; i++) {
    if (invData[i][1] === itemId && String(invData[i][4]).trim().toUpperCase() === targetSerial) {
      return Number(invData[i][5]) || 0;
    }
  }
  return 0;
}

/* ────────────────────────────────────────────────────────────
 * 필수 유틸리티 헬퍼 함수 복구 영역
 * ──────────────────────────────────────────────────────────── */

/** 원장 데이터 행 배열(헤더 제외)을 반환 */
function getLedgerRows_(ledgerSheet) {
  const last = ledgerSheet.getLastRow();
  if (last < 2) return [];
  return ledgerSheet.getRange(2, 1, last - 1, 11).getValues(); // 타임스탬프가 추가되어 11열
}

/** EXTERNAL (VENDOR/SCRAP) 등 시스템 외부 위치 여부 판별 */
function isExternal_(bucket) {
  return String(bucket).indexOf('EXTERNAL') === 0;
}

/** 시리얼 값 정규화 (공백 제거 + 대문자 텍스트 변환) */
function normSerial_(v) {
  return String(v).trim().toUpperCase();
}