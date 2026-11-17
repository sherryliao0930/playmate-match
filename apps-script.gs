/**
 * 畢旅玩伴配對 — 收件程式（Google Apps Script）
 *
 * 用法：
 *  1. 建一個新的 Google 試算表，上方選「擴充功能 → Apps Script」
 *  2. 把這整個檔案的內容貼進去，覆蓋原本的 myFunction
 *  3. 改下面 TEACHER_KEY 成你自己的密碼
 *  4. 右上角「部署 → 新增部署作業 → 類型選『網頁應用程式』」
 *       執行身分：我
 *       誰可以存取：任何人
 *  5. 複製得到的網址（.../exec 結尾），貼到 index.html 最上面的 CONFIG.endpoint
 */

// ★ 改成你自己的密碼，並把同一組密碼填進 index.html 的 CONFIG.teacherKey
const TEACHER_KEY = 'change-me-1234';

const SHEET_NAME = '回覆';
const HEADER = ['時間', '座號', '姓名', '膽量', '喜歡類型', '九族玩伴', '劍湖山玩伴',
                '避免同組', '原因', '備註'];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d || !d.seat) return json({ ok: false, error: 'no seat' });
    getSheet().appendRow([
      new Date(),
      Number(d.seat),
      String(d.name || ''),
      Number(d.courage) || 3,
      (d.cats || []).join('|'),
      (d.nine || []).join('|'),
      (d.jan || []).join('|'),
      (d.avoid || []).join('|'),      // 「避免同組」：學生私下寫給老師的，請勿外流
      (d.avoidWhy || []).join('|'),   // 對應的原因，順序與上一欄相同
      String(d.note || '')
    ]);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action !== 'list') return json({ ok: true, msg: 'ready' });
  if (p.key !== TEACHER_KEY) return json({ ok: false, error: 'bad key' });

  const values = getSheet().getDataRange().getValues().slice(1);
  const bySeat = {};                       // 同一人重複填，以最後一次為準
  values.forEach(function (r) {
    const seat = Number(r[1]);
    if (!seat) return;
    bySeat[seat] = {
      seat: seat,
      name: String(r[2] || ''),
      courage: Number(r[3]) || 3,
      cats: split(r[4]),
      nine: split(r[5]).map(Number),
      jan: split(r[6]).map(Number),
      avoid: split(r[7]).map(Number),
      avoidWhy: String(r[8] == null ? '' : r[8]).split('|'),
      note: String(r[9] || '')
    };
  });
  return json({ ok: true, rows: Object.keys(bySeat).map(function (k) { return bySeat[k]; }) });
}

function split(v) {
  return String(v == null ? '' : v).split('|').filter(function (x) { return x !== ''; });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
  }
  return sh;
}
