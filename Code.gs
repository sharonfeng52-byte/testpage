// 報名表 → Google Sheet 接收程式
const SHEET_ID = '18Ij0xASYVxtpQoYvvFWMcnSvzOimjKIPh5Vn0j-vBgE';
const TAB_NAME = '報名表';
const HEADERS = ['報名時間', '家長姓名', '學生姓名', '就讀年級', '學生年齡', '聯絡電話', '電郵', '備註', '來源頁面'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName(TAB_NAME) || ss.insertSheet(TAB_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    const clean = v => {
      v = String(v == null ? '' : v).slice(0, 500);
      return /^[=+\-@]/.test(v) ? "'" + v : v; // 防止試算表公式注入
    };
    sh.appendRow([
      Utilities.formatDate(new Date(), 'Asia/Hong_Kong', 'yyyy-MM-dd HH:mm:ss'),
      clean(d.parent), clean(d.student), clean(d.grade), clean(d.age),
      clean(d.phone), clean(d.email), clean(d.note), clean(d.page)
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('報名系統運作中');
}
