/**
 * [1회성 실행] Inventory_Origin 시트의 기초 재고를 Ledger의 ADD 트랜잭션으로 일괄 이관합니다.
 */
function migrateOriginToLedger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const origin = ss.getSheetByName('Inventory_Origin');
  const ledger = ss.getSheetByName('Ledger');
  const ui = SpreadsheetApp.getUi();

  if (!origin || !ledger) {
    ui.alert('❌ 에러: Inventory_Origin 또는 Ledger 시트를 찾을 수 없습니다.');
    return;
  }

  const oLast = origin.getLastRow();
  if (oLast < 2) {
    ui.alert('ℹ️ Inventory_Origin 시트에 이관할 데이터가 없습니다.');
    return;
  }

  // Origin 데이터 읽기
  const originData = origin.getRange(2, 1, oLast - 1, 6).getValues();
  const newLedgerRows = [];
  const timestamp = new Date();

  for (let i = 0; i < originData.length; i++) {
    const cat = originData[i][0];
    const id = originData[i][1];
    const name = originData[i][2];
    const loc = originData[i][3];
    const serial = originData[i][4];
    const qty = Number(originData[i][5]);

    if (id && loc && qty > 0) {
      // [Timestamp, Type, Category, Item ID, Item Name, Serial, From, To, Qty, Worker, Note]
      newLedgerRows.push([
        timestamp,
        'ADD',
        cat,
        id,
        name,
        serial ? String(serial) : 'N/A',
        'EXTERNAL (VENDOR)',             // From (외부 기초 재고)
        loc,                             // To (현재 위치)
        qty,
        'SYSTEM',                        // Worker (시스템 자동 이관)
        'Migrated from Inventory_Origin' // Note (이관 기록 메모)
      ]);
    }
  }

  if (newLedgerRows.length === 0) {
    ui.alert('ℹ️ 이관할 유효한 재고 데이터가 없습니다.');
    return;
  }

  // 앞서 요청하신 대로 Ledger의 1행(헤더) 바로 아래에 일괄 삽입
  ledger.insertRowsAfter(1, newLedgerRows.length);
  ledger.getRange(2, 1, newLedgerRows.length, 11).setValues(newLedgerRows);
  ledger.getRange(2, 6, newLedgerRows.length, 1).setNumberFormat('@'); // F열(시리얼) 텍스트 포맷 유지

  // 이관 완료 후 재고 새로고침
  rebuildInv_(ss);

  SpreadsheetApp.flush();
  ui.alert('✅ 마이그레이션 완료!\n총 ' + newLedgerRows.length + '건의 기초 재고가 Ledger에 ADD 기록으로 이관되었습니다.\n이제 Inventory_Origin 시트를 삭제하셔도 됩니다.');
}