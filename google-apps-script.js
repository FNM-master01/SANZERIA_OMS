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
  // C列（電話番号）・D列（郵便番号）を常にテキスト形式に設定（0落ち防止）
  sheet.getRange('C:C').setNumberFormat('@STRING@');
  sheet.getRange('D:D').setNumberFormat('@STRING@');
  return sheet;
}

// 日付を常に yyyy/MM/dd 形式に統一
function formatDate(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy/MM/dd');
  }
  return String(value || '');
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
    if (!e || !e.parameter) {
      result.setContent(JSON.stringify({ error: 'リクエストパラメータがありません' }));
      return result;
    }

    const params = e.parameter;
    const action = params.action;

    if (action === 'getAll') {
      result.setContent(JSON.stringify(getAllMembers()));
    } else if (action === 'save') {
      if (!params.data) {
        result.setContent(JSON.stringify({ error: 'dataが指定されていません' }));
        return result;
      }
      const members = JSON.parse(params.data);
      if (!Array.isArray(members)) {
        result.setContent(JSON.stringify({ error: 'membersは配列である必要があります' }));
        return result;
      }
      result.setContent(JSON.stringify(saveAllMembers(members)));
    } else if (action === 'add') {
      if (!params.data) {
        result.setContent(JSON.stringify({ error: 'dataが指定されていません' }));
        return result;
      }
      const member = JSON.parse(params.data);
      result.setContent(JSON.stringify(addMember(member)));
    } else if (action === 'update') {
      if (!params.data) {
        result.setContent(JSON.stringify({ error: 'dataが指定されていません' }));
        return result;
      }
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
    date: formatDate(row[7])  // 日付フォーマット統一
  })).filter(m => m.id && m.id !== '');

  return { success: true, members };
}

function saveAllMembers(members) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

  if (members.length > 0) {
    // 電話番号・郵便番号列をテキスト形式に設定してから書き込む
    sheet.getRange(2, 3, members.length, 1).setNumberFormat('@STRING@');
    sheet.getRange(2, 4, members.length, 1).setNumberFormat('@STRING@');
    const rows = members.map(m => [
      String(m.id),
      m.name,
      String(m.tel||''),
      String(m.zip||''),
      m.addr||'',
      m.type||'通常',
      m.memo||'',
      m.date||''
    ]);
    sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
  }
  return { success: true, count: members.length };
}

function addMember(member) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();

  // 会員番号の重複チェック
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(member.id)) {
      return { error: '同じ会員番号が既に存在します: ' + member.id };
    }
  }

  const lastRow = sheet.getLastRow() + 1;
  // テキスト形式を先に設定してから書き込む（0落ち防止）
  sheet.getRange(lastRow, 3).setNumberFormat('@STRING@');
  sheet.getRange(lastRow, 4).setNumberFormat('@STRING@');
  sheet.getRange(lastRow, 1, 1, HEADERS.length).setValues([[
    String(member.id),
    member.name,
    String(member.tel||''),
    String(member.zip||''),
    member.addr||'',
    member.type||'通常',
    member.memo||'',
    member.date||''
  ]]);
  return { success: true, id: member.id };
}

function updateMember(member) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(member.id)) {  // 型を統一
      // テキスト形式を先に設定してから書き込む（0落ち防止）
      sheet.getRange(i + 1, 3).setNumberFormat('@STRING@');
      sheet.getRange(i + 1, 4).setNumberFormat('@STRING@');
      sheet.getRange(i + 1, 1, 1, HEADERS.length).setValues([[
        String(member.id),
        member.name,
        String(member.tel||''),
        String(member.zip||''),
        member.addr||'',
        member.type||'通常',
        member.memo||'',
        member.date||''
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
    if (String(data[i][0]) === String(id)) {  // 型を統一
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: '会員が見つかりません: ' + id };
}
