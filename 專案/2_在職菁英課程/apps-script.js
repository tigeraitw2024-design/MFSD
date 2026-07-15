/**
 * Google Apps Script · 在職菁英課程報名後端
 *
 * 綁定 Sheet: https://docs.google.com/spreadsheets/d/1EcRHVdfrx720kvCbYSN2FjyVq_TgYTojDONs2baR9I4/
 *
 * 分頁:
 *   梯次目錄 …………………… 課程場次資料;前端讀 CSV(要「發布到網路」)
 *   在職菁英報名表單 …………… 報名資料;doPost 自動寫入 + 寄確認信
 *
 * ════════ 部署步驟(建 Sheet 之後做這幾步) ════════
 *   1. 打開新 Sheet → 擴充功能 → Apps Script → 貼這整份到 Code.gs(覆蓋預設程式碼)
 *   2. 存檔(Ctrl+S)
 *   3. 上方函式下拉選 setupCourseSheet → Run
 *      → 會跳授權視窗,一路同意(需要 Gmail 寄信權限)
 *   4.(建議)函式選 testCourseEmail → Run → 你會收到範例信,確認版型
 *   5. 部署 → 新增部署 → 網頁應用程式
 *      執行身份:我
 *      存取權限:任何人
 *      按「部署」→ 拿到 Web App URL(這就是 SHEET_WEBHOOK)
 *   6. 把那串 URL 貼給 Robin,他會填進 index.html 的 CONFIG.SHEET_WEBHOOK
 *
 *   ⚠️ 未來改這份程式後要重部署:「管理部署 → 鉛筆編輯 → 版本『新版本』→ 部署」
 *      千萬不要按「新增部署」— 會產生新 URL,前端就抓不到。
 */

const SHEET_ID = '1EcRHVdfrx720kvCbYSN2FjyVq_TgYTojDONs2baR9I4';
const SHEET_COURSE = '在職菁英報名表單';

const COURSE_HEADERS = [
  '時間戳',
  // 學員
  '中文姓名', '職稱', 'E-Mail', '性別', '用餐選擇',
  // 公司
  '統編', '公司名稱', '產業別',
  // 輔導
  '輔導的單位', '是否為受輔導廠商',
  // 梯次
  '報名梯次',
  // 聯絡窗口
  '聯絡人姓名', '聯絡人手機', '聯絡人E-Mail',
  // 其他
  '報名留言', '同意狀態',
  // 系統
  '寄信狀態'
];
const COURSE_MAIL_STATUS_COL = 18;   // 寄信狀態固定在最後(第 18 欄)

// ════════ 入口 ════════
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.type === 'course_signup') return logCourseSignup(data);
    return jsonResp({ ok: false, error: '未知的 type: ' + data.type });
  } catch (err) {
    return jsonResp({ ok: false, error: err.toString() });
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'counts') return courseCounts();
  return ContentService.createTextOutput('在職菁英課程報名 API 運作中 ✓');
}

// 回傳各梯次目前報名人數
function courseCounts() {
  const counts = {};
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_COURSE);
    if (sheet && sheet.getLastRow() > 1) {
      // 報名梯次 = 第 12 欄
      const values = sheet.getRange(2, 12, sheet.getLastRow() - 1, 1).getValues();
      values.forEach(r => {
        const key = String(r[0] || '').trim();
        if (key) counts[key] = (counts[key] || 0) + 1;
      });
    }
  } catch (err) { /* 出錯回空 */ }
  return ContentService.createTextOutput(JSON.stringify(counts))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════ 寫入:課程報名(寫資料 + 寄確認信 + 記錄寄信狀態) ════════
function logCourseSignup(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_COURSE);
  if (!sheet) sheet = ss.insertSheet(SHEET_COURSE, ss.getSheets().length);
  ensureHeaders(sheet, COURSE_HEADERS);
  sheet.appendRow([
    data.timestamp || new Date().toISOString(),
    // 學員
    data.name || '',
    data.jobTitle || '',
    data.email || '',
    data.gender || '',
    data.meal || '',
    // 公司
    data.taxId || '',
    data.companyName || '',
    data.industry || '',
    // 輔導
    data.advisor || '',
    data.isAdvised || '',
    // 梯次
    data.cohort || '',
    // 聯絡窗口
    data.contactName || '',
    data.contactPhone || '',
    data.contactEmail || '',
    // 其他
    data.note || '',
    data.consent || '',
    // 寄信狀態(下面寄完再填回)
    ''
  ]);
  const row = sheet.getLastRow();

  let mailStatus;
  try {
    sendCourseConfirmEmail(data);
    mailStatus = '✅ 已寄出 ' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
  } catch (err) {
    mailStatus = '❌ 寄信失敗:' + err.message;
  }
  sheet.getRange(row, COURSE_MAIL_STATUS_COL).setValue(mailStatus);

  return jsonResp({ ok: true });
}

// ════════ 寄送:報名確認信 ════════
function sendCourseConfirmEmail(data) {
  const to = String(data.email || '').trim();
  if (!to) throw new Error('無 Email');

  const parts = String(data.cohort || '').split('｜');
  const city     = (parts[0] || '—').trim();
  const datetime = (parts[1] || '—').trim();
  const place    = (parts[2] || '—').trim();

  const subject = '【報名確認】AI 賦能 · 生產力躍升與製造業應用趨勢';
  const body =
    (data.name || '') + ' ' + (data.jobTitle || '') + ' 您好,\n\n' +
    '感謝您報名「AI 賦能 · 生產力躍升與製造業應用趨勢」在職菁英課程,\n' +
    '30 小時實戰,把 Low-Code、RAG、AI Agent 帶回工廠。\n' +
    '我們已收到您的報名資料。\n\n' +
    '▌您的報名資訊\n' +
    '　公司名稱:' + (data.companyName || '') + '\n' +
    '　公司統編:' + (data.taxId || '') + '\n' +
    '　產業別  :' + (data.industry || '') + '\n' +
    '　報名學員:' + (data.name || '') + ' / ' + (data.jobTitle || '') + '\n' +
    '　用餐選擇:' + (data.meal || '') + '\n' +
    '　輔導單位:' + (data.advisor || '') + '\n' +
    '　是否受輔導:' + (data.isAdvised || '') + '\n' +
    '　報名梯次:\n' +
    '　　．開課縣市:' + city + '\n' +
    '　　．上課時間:' + datetime + '\n' +
    '　　．上課地點:' + place + '\n' +
    '　公司聯絡窗口:' + (data.contactName || '') + ' / ' + (data.contactPhone || '') + '\n\n' +
    '▌課程資訊\n' +
    '　．30 小時實體課程(4 天,週末上課)\n' +
    '　．結訓帶 3 個具體成果回公司直接用:\n' +
    '　　1. 用 AI 產出公司自己的行銷內容(貼文、圖片、短影音)\n' +
    '　　2. 做出企業自己的 AI 工具(Low-Code + AI Agent)\n' +
    '　　3. 打造部門級的知識庫(RAG + MCP 跨系統串接)\n\n' +
    '▌行前提醒\n' +
    '　．請攜帶個人筆電,以便現場實作\n' +
    '　．課程當天請提早 10 分鐘報到\n' +
    '　．如需改期或取消,請於開課 3 日前來信告知\n\n' +
    '如有任何問題,歡迎隨時與我們聯繫,期待課堂上見!\n\n' +
    '──────────────────\n' +
    '本課程聯絡窗口\n' +
    '．虎智科技  紀先生   02-66058192\n' +
    '．許雅婷   07-2625889 分機 117\n';

  MailApp.sendEmail({ to: to, subject: subject, body: body, name: '虎智科技 TigerAI' });
}

// ════════ 測試工具 ════════

// 部署後手動跑一次,把「在職菁英報名表單」分頁建好並補齊欄位
function setupCourseSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_COURSE);
  if (!sheet) sheet = ss.insertSheet(SHEET_COURSE, ss.getSheets().length);
  ensureHeaders(sheet, COURSE_HEADERS);
  Logger.log('分頁「' + SHEET_COURSE + '」就緒(' + COURSE_HEADERS.length + ' 欄)');
}

// 寄一封範例確認信到你自己信箱,確認版型 OK
function testCourseEmail() {
  sendCourseConfirmEmail({
    name: '王大明', jobTitle: '生產部經理',
    email: Session.getActiveUser().getEmail(),
    gender: '男', meal: '葷',
    companyName: '測試股份有限公司', taxId: '12345678',
    industry: '製造業',
    advisor: '工研院感測中心', isAdvised: '是',
    contactName: '李小華', contactPhone: '0912345678', contactEmail: 'contact@example.com',
    cohort: '台中｜8/17-18 + 8/24-25｜勤益科大',
    note: '素食一份'
  });
  Logger.log('已寄測試信至 ' + Session.getActiveUser().getEmail());
}

// ════════ 工具函式 ════════

function ensureHeaders(sheet, headers) {
  const width = headers.length;
  const lastCol = sheet.getLastColumn();
  let needsUpdate = false;
  if (lastCol < width) {
    needsUpdate = true;
  } else {
    const current = sheet.getRange(1, 1, 1, width).getValues()[0];
    needsUpdate = headers.some((h, i) => current[i] !== h);
  }
  if (needsUpdate) {
    sheet.getRange(1, 1, 1, width).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#F1EBD7');
  }
}

function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
