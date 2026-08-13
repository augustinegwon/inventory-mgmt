/**
 * [1회성 마이그레이션 스크립트] 
 * Settings에 ITM-1001부터 ID를 발급하고, Ledger와 Inventory_Origin의 기존 기록에 매핑합니다.
 */
function migrateToIdSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName('Settings');
  const ledger = ss.getSheetByName('Ledger');
  const origin = ss.getSheetByName('Inventory_Origin');

  // 1. Settings에 ID 발급
  const sLast = settings.getLastRow();
  const idMap = {};
  let nextIdNum = 1001;

  if (sLast >= 2) {
    const names = settings.getRange(2, 3, sLast - 1, 1).getValues(); // C열(Item Name)
    const ids = [];
    for (let i = 0; i < names.length; i++) {
      const name = names[i][0];
      if (name) {
        const newId = 'ITM-' + nextIdNum++;
        ids.push([newId]);
        idMap[name] = newId; // 이름과 ID를 짝지어 기억
      } else {
        ids.push(['']);
      }
    }
    settings.getRange(2, 1, ids.length, 1).setValues(ids); // A열에 ID 쓰기
  }

  // 2. Ledger에 ID 매핑
  if (ledger) {
    const lLast = ledger.getLastRow();
    if (lLast >= 2) {
      const lNames = ledger.getRange(2, 5, lLast - 1, 1).getValues(); // E열(Item Name)
      const lIds = [];
      for (let i = 0; i < lNames.length; i++) {
        const name = lNames[i][0];
        lIds.push([idMap[name] || '']); // 이름으로 발급된 ID 찾기
      }
      ledger.getRange(2, 4, lIds.length, 1).setValues(lIds); // D열에 ID 쓰기
    }
  }

  // 3. Inventory_Origin에 ID 매핑
  if (origin) {
    const oLast = origin.getLastRow();
    if (oLast >= 2) {
      const oNames = origin.getRange(2, 3, oLast - 1, 1).getValues(); // C열(Item Name)
      const oIds = [];
      for (let i = 0; i < oNames.length; i++) {
        const name = oNames[i][0];
        oIds.push([idMap[name] || '']);
      }
      origin.getRange(2, 2, oIds.length, 1).setValues(oIds); // B열에 ID 쓰기
    }
  }

  SpreadsheetApp.getUi().alert('✅ 마이그레이션 완료! 모든 물품에 ID가 부여되고 기존 기록과 연결되었습니다.');
}