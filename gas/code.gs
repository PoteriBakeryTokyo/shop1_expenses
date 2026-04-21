// ===== 設定 =====
const SHEET_NAME = ''; // シート名（空の場合はアクティブシートを使用）

// スプレッドシートの列インデックス（1始まり）
const COL_DATE     = 1; // A: 日付
const COL_NAME     = 2; // B: 名前
const COL_CATEGORY = 3; // C: 内容
const COL_PAYMENT  = 4; // D: 支払方法
const COL_AMOUNT   = 5; // E: 税込金額
const COL_BALANCE  = 6; // F: 残高
const COL_RECEIPT  = 7; // G: レシートURL

// レシート保存先フォルダ名（Google Drive）
const RECEIPT_FOLDER_NAME = 'レシート（三茶）';
// ================

function doPost(e) {
  try {
    const payload = JSON.parse(e.parameter.payload);
    if (payload.action === 'submitExpense') {
      return submitExpense(payload);
    }
    return jsonResponse({ status: 'error', message: '不明なアクション' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function submitExpense(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getActiveSheet();
  if (!sheet) {
    return jsonResponse({ status: 'error', message: 'シートが見つかりません' });
  }

  const lastRow = sheet.getLastRow();
  const amount = Number(data.amount) || 0;

  // 前行の残高を取得して新しい残高を計算
  let prevBalance = 0;
  if (lastRow >= 2) {
    const prevBalanceCell = sheet.getRange(lastRow, COL_BALANCE).getValue();
    prevBalance = Number(prevBalanceCell) || 0;
  }
  const newBalance = prevBalance - amount;

  // レシート画像をGoogle Driveに保存
  let receiptUrl = '';
  if (data.receiptBase64) {
    try {
      receiptUrl = saveReceiptToDrive(data.receiptBase64, data.receiptMime, data.date);
    } catch (err) {
      // レシート保存に失敗してもデータは記録する
      console.error('レシート保存エラー:', err.message);
    }
  }

  // スプレッドシートに行を追加（残高も含む）
  sheet.appendRow([
    data.date,
    data.name,
    data.category,
    data.paymentMethod,
    amount,
    newBalance,
    receiptUrl,
  ]);

  return jsonResponse({ status: 'ok', balance: newBalance });
}

function saveReceiptToDrive(base64, mime, date) {
  const folders = DriveApp.getFoldersByName(RECEIPT_FOLDER_NAME);
  const folder = folders.hasNext()
    ? folders.next()
    : DriveApp.createFolder(RECEIPT_FOLDER_NAME);

  const fileName = 'receipt_' + date + '_' + Date.now() + '.jpg';
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mime || 'image/jpeg', fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
