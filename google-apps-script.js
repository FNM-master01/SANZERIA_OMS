// ============================================================
// 宅配会員管理システム — Google Apps Script
// このコードをGoogle Apps Scriptに貼り付けてください
// ============================================================

const SHEET_NAME = '会員データ';
const HEADERS = ['会員番号', '氏名', '電話番号', '郵便番号', '住所', '区分', 'メモ', '登録日'];

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 130);
    sheet.setColumnWidth(4, 90);
    sheet.setColumnWidth(5, 250);
    sheet.setColumnWidth(6, 70);
    sheet.setColumnWidth(7, 250);
    sheet.setColumnWidth(8, 100);
  }
  return sheet;
}

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const result = ContentService.createTextOutput();
  result.setMimeType(ContentService.MimeType.JSON);

  try {
    const params = e.parameter;
    const action = params.action;

    if (action === 'getAll') {
      result.setContent(JSON.stringify(getAllMembers()));
    } else if (action === 'save') {
      const members = JSON.parse(params.data);
      result.setContent(JSON.stringify(saveAllMembers(members)));
    } else if (action === 'add') {
      const member = JSON.parse(params.data);
      result.setContent(JSON.stringify(addMember(member)));
    } else if (action === 'update') {
      const member = JSON.parse(params.data);
      result.setContent(JSON.stringify(updateMember(member)));
    } else if (action === 'delete') {
      result.setContent(JSON.stringify(deleteMember(params.id)));
    } else {
      result.setContent(JSON.stringify({ error: '不明なアクション: ' + action }));
    }
  } catch (err) {
    result.setContent(JSON.stringify({ error: err.message }));
  }

  return result;
}

function getAllMembers() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, members: [] };

  const members = data.slice(1).map(row => ({
    id:   String(row[0]),
    name: String(row[1]),
    tel:  String(row[2]),
    zip:  String(row[3]),
    addr: String(row[4]),
    type: String(row[5]),
    memo: String(row[6]),
    date: String(row[7])
  })).filter(m => m.id && m.id !== '');

  return { success: true, members };
}

function saveAllMembers(members) {
  const sheet = getSheet();
  // ヘッダー行以外を削除
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

  if (members.length > 0) {
    const rows = members.map(m => [m.id, m.name, m.tel, m.zip||'', m.addr||'', m.type||'通常', m.memo||'', m.date||'']);
    sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
  }
  return { success: true, count: members.length };
}

function addMember(member) {
  const sheet = getSheet();
  sheet.appendRow([member.id, member.name, member.tel, member.zip||'', member.addr||'', member.type||'通常', member.memo||'', member.date||'']);
  return { success: true, id: member.id };
}

function updateMember(member) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === member.id) {
      sheet.getRange(i + 1, 1, 1, HEADERS.length).setValues([[
        member.id, member.name, member.tel, member.zip||'', member.addr||'', member.type||'通常', member.memo||'', member.date||''
      ]]);
      return { success: true };
    }
  }
  return { error: '会員が見つかりません: ' + member.id };
}

function deleteMember(id) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: '会員が見つかりません: ' + id };
}
