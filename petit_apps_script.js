// 쁘띠치과 서류발급 요청 관리 - Google Apps Script v2
// ★ 기존 코드 전체 삭제 후 이 코드로 교체하세요

const SHEET_NAME = '요청목록';
const HEADERS = ['ID','환자명','신청일','마감일','서류종류','수량','총금액','특이사항','원장메모','승인상태','등록일시'];

function doGet(e) {
  const p = e.parameter;
  const action = p.action;
  const callback = p.callback; // JSONP 콜백
  
  let result;
  
  if (action === 'getAll') {
    result = getAll();
  } else if (action === 'add') {
    result = addReqFromGet(p);
  } else if (action === 'updateStatus') {
    result = updateStatus(p.id, p.status, p.memo || '');
  } else if (action === 'delete') {
    result = deleteReq(p.id);
  } else {
    result = { ok: false, msg: '알 수 없는 요청' };
  }
  
  const json = JSON.stringify(result);
  
  // JSONP 방식
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// GET으로 신규 추가
function addReqFromGet(p) {
  const sheet = getSheet();
  const id = new Date().getTime().toString();
  
  sheet.appendRow([
    id,
    p.patient || '',
    p.date || '',
    p.due || '',
    p.docs || '',
    p.qty || 1,
    p.total || 0,
    p.memo || '',
    p.docMemo || '',
    p.status || 'wait',
    new Date().toLocaleString('ko-KR')
  ]);
  
  return { ok: true, id: id };
}

// 전체 조회
function getAll() {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { ok: true, data: [] };
  
  const data = rows.slice(1).map(r => ({
    id: r[0].toString(),
    patient: r[1],
    date: r[2],
    due: r[3],
    docs: r[4],
    qty: r[5],
    total: r[6],
    memo: r[7],
    docMemo: r[8],
    status: r[9],
    createdAt: r[10]
  }));
  
  return { ok: true, data: data.reverse() };
}

// 상태 변경
function updateStatus(id, status, memo) {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === id.toString()) {
      sheet.getRange(i + 1, 10).setValue(status);
      if (memo) sheet.getRange(i + 1, 9).setValue(memo);
      return { ok: true };
    }
  }
  return { ok: false, msg: '항목을 찾을 수 없음' };
}

// 삭제
function deleteReq(id) {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === id.toString()) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, msg: '항목을 찾을 수 없음' };
}

// 시트 가져오기
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#007A54')
      .setFontColor('#FFFFFF');
  }
  return sheet;
}
