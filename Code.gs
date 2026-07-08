/**
 * Inventory Management — Google Apps Script (baseline)
 *
 * 이 파일은 현재 스프레드시트에 붙어 있는 App Script를 그대로 가져온 "기준(baseline)" 버전입니다.
 * 앞으로의 개선은 이 파일을 기준으로 diff를 관리합니다.
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

/**
 * 1. Input_Transaction 시트의 UI 구조 및 수식을 초기화하는 함수
 * (유효성 검사 충돌 에러를 방지하기 위해 데이터 입력 후 규칙을 적용하도록 순서 교정)
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
  // 기본 가이드 및 텍스트 UI 배치
  inputSheet.getRange('B2').setValue('🔍 Item Search').setFontWeight('bold');
  inputSheet.getRange('B3').setValue('Category');
  inputSheet.getRange('B4').setValue('Item');
  inputSheet.getRange('B6').setValue('Location');
  inputSheet.getRange('B7').setValue('Current Qty').setFontWeight('bold');
  inputSheet.getRange('E2').setValue('📦 Transaction Form').setFontWeight('bold');
  inputSheet.getRange('E3').setValue('Type');
  inputSheet.getRange('E4').setValue('Item');
  inputSheet.getRange('E5').setValue('SERIAL NO.');
  inputSheet.getRange('E6').setValue('FROM');
  inputSheet.getRange('E7').setValue('TO');
  inputSheet.getRange('E8').setValue('Quantity');
  inputSheet.getRange('E9').setValue('USER');
  inputSheet.getRange('E10').setValue('Note');
  // 헬퍼 열(Helper Columns) 수식 배치 (W1, X1, Y1)
  // W1: MOVE/REMOVE 시 해당 아이템의 재고가 있는 '위치' 목록 필터링
  inputSheet.getRange('W1').setFormula(`=IFERROR(FILTER(Settings!$E$2:$E, SUMIFS(Ledger!$H$2:$H, Ledger!$D$2:$D, F4, Ledger!$G$2:$G, Settings!$E$2:$E) - SUMIFS(Ledger!$H$2:$H, Ledger!$D$2:$D, F4, Ledger!$F$2:$F, Settings!$E$2:$E) > 0), "")`);
  // X1: MOVE/REMOVE 시 해당 아이템의 재고가 남아있는 '시리얼 번호' 목록 필터링
  inputSheet.getRange('X1').setFormula(`=IFERROR(FILTER(Settings!$C$2:$C, Settings!$B$2:$B=F4, SUMIFS(Ledger!$H$2:$H, Ledger!$D$2:$D, F4, Ledger!$E$2:$E, Settings!$C$2:$C, Ledger!$G$2:$G, F6) - SUMIFS(Ledger!$H$2:$H, Ledger!$D$2:$D, F4, Ledger!$E$2:$E, Settings!$C$2:$C, Ledger!$F$2:$F, F6) > 0), "")`);
  // Y1: TO 드롭다운을 위한 전체 위치 정보 불러오기
  inputSheet.getRange('Y1').setFormula(`=IFERROR(FILTER(Settings!$E$2:$E, Settings!$E$2:$E<>""), "")`);
  // 🌟 [에러 해결 핵심] 유효성 검사 규칙을 적용하기 전에 통과 가능한 대문자 '기본값'을 먼저 셀에 입력합니다.
  inputSheet.getRange('F3').setValue('ADD'); // Type 기본값
  inputSheet.getRange('F5').setValue('N/A'); // Serial No 기본값
  inputSheet.getRange('F8').setValue(1);     // Quantity 기본값
  // 드롭다운 규칙 정의 (데이터 유효성 검사)
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['ADD', 'MOVE', 'REMOVE'], true)
    .setAllowInvalid(false)
    .build();

  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName('Settings').getRange('A2:A'), true)
    .setAllowInvalid(false)
    .build();

  const toLocRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(inputSheet.getRange('Y1:Y'), true)
    .setAllowInvalid(false)
    .build();

  // 유효성 검사 규칙 적용
  inputSheet.getRange('F3').setDataValidation(typeRule);      // Type 셀
  inputSheet.getRange('C3').setDataValidation(categoryRule);  // Item Search Category 셀
  inputSheet.getRange('F7').setDataValidation(toLocRule);     // TO 셀
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

  // 사용자가 입력한 값 대문자로 강제 변환 (정합성 보장)
  if (cellA1 === 'F3' || cellA1 === 'F4' || cellA1 === 'F6') {
    const rawValue = range.getValue();
    if (typeof rawValue === 'string') {
      range.setValue(rawValue.toUpperCase());
    }
  }

  const type = inputSheet.getRange('F3').getValue();
  const itemName = inputSheet.getRange('F4').getValue();
  const serialCell = inputSheet.getRange('F5');
  const fromLocCell = inputSheet.getRange('F6');  // FROM 셀
  const qtyCell = inputSheet.getRange('F8');

  // 아이템의 시리얼 관리 여부 확인 (Settings 시트 참조)
  const settingsSheet = ss.getSheetByName('Settings');
  const mapping = settingsSheet.getRange('B2:C' + settingsSheet.getLastRow()).getValues();
  let isSerial = 'NO';
  for (let i = 0; i < mapping.length; i++) {
    if (mapping[i][0] === itemName) {
      isSerial = mapping[i][1];
      break;
    }
  }

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

  // 입력 필드 값 가져오기
  const type = inputSheet.getRange('F3').getValue();
  const item = inputSheet.getRange('F4').getValue();
  const serial = inputSheet.getRange('F5').getValue();
  const fromLoc = inputSheet.getRange('F6').getValue();
  const toLoc = inputSheet.getRange('F7').getValue();
  const qty = inputSheet.getRange('F8').getValue();
  const worker = inputSheet.getRange('F9').getValue();
  const note = inputSheet.getRange('F10').getValue();

  // 필수값 누락 검증 (ADD 검증 강화)
  if (!type || !item || !qty || !worker) {
    SpreadsheetApp.getUi().alert('❌ 에러: Type, Item, Quantity, USER 필드는 필수 입력 항목입니다.');
    return;
  }

  if ((type === 'MOVE' || type === 'REMOVE') && !fromLoc) {
    SpreadsheetApp.getUi().alert('❌ 에러: MOVE 또는 REMOVE 처리 시 FROM(출발지) 위치를 지정해야 합니다.');
    return;
  }

  if ((type === 'ADD' || type === 'MOVE') && !toLoc) {
    SpreadsheetApp.getUi().alert('❌ 에러: ADD 또는 MOVE 처리 시 TO(목적지) 위치를 지정해야 합니다.');
    return;
  }

  // 품목의 Category 정보 Settings 시트에서 조회
  const settingsData = settingsSheet.getRange('A2:B' + settingsSheet.getLastRow()).getValues();
  let category = '';
  let itemExists = false;
  for (let i = 0; i < settingsData.length; i++) {
    if (settingsData[i][1] === item) {
      category = settingsData[i][0];
      itemExists = true;
      break;
    }
  }

  if (!itemExists) {
    SpreadsheetApp.getUi().alert('❌ 에러: Settings 마스터 정보에 등록되지 않은 품목명입니다.');
    return;
  }

  // 락(Lock) 획득을 통한 동시성 제어 및 데이터 정합성 보장
  const lock = LockService.getScriptLock();
  try {
    // 최대 10초 대기
    lock.waitLock(10000);

    // Ledger에 기록할 최종 행 데이터 배열 생성
    const timestamp = new Date();
    const finalFrom = (type === 'ADD') ? 'EXTERNAL (VENDOR)' : fromLoc;
    const finalTo = (type === 'REMOVE') ? 'EXTERNAL (SCRAP)' : toLoc;
    const finalSerial = (serial === '') ? 'N/A' : serial;

    // 원장 데이터 추가 실행
    ledgerSheet.appendRow([
      timestamp,
      type,
      category,
      item,
      finalSerial,
      finalFrom,
      finalTo,
      qty,
      worker,
      note
    ]);

    // 제출 완료 후 입력 폼 초기화 (Type, User 정보는 편의상 유지)
    inputSheet.getRange('F4').setValue(''); // Item 초기화
    inputSheet.getRange('F5').setValue('N/A'); // Serial 초기화
    inputSheet.getRange('F6').setValue(''); // FROM 초기화
    inputSheet.getRange('F7').setValue(''); // TO 초기화
    inputSheet.getRange('F8').setValue(1);  // 수량 초기화
    inputSheet.getRange('F10').setValue(''); // Note 초기화

    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert('✅ 성공: 트랜잭션이 원장(Ledger)에 정상적으로 기록되었습니다.');

  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ 오류 발생: 다른 사용자가 작업 중이거나 시스템 지연이 발생했습니다. 다시 시도해 주세요.\n' + error.toString());
  } finally {
    lock.releaseLock(); // 반드시 락 해제
  }
}
