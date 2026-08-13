/**
 * Google Apps Script · 企業落地課程報名後端
 *
 * ⚠️ 本專案獨立於 2_在職菁英課程,需另建 Sheet + 部署新 Web App
 *
 * 綁 Sheet:自建新 Sheet 後把 SHEET_ID 換上
 *
 * 分頁:
 *   梯次目錄 …………… 課程場次資料(可選,前端讀 CSV,或直接用 index.html 的 COHORTS_FALLBACK)
 *   企業落地報名表單 … 報名資料;doPost 自動寫入 + 寄確認信
 *
 * ════════ 部署步驟(SOP 跟在職菁英課程一樣) ════════
 *   1. 新建空白 Google Sheet,複製 URL 裡的 SPREADSHEET_ID 貼到下面 SHEET_ID
 *   2. Sheet → 擴充功能 → Apps Script → 貼這整份到 Code.gs
 *   3. Ctrl+S 存檔
 *   4. 函式下拉選 setupCourseSheet → Run(第一次會跳授權,同意 Gmail 寄信權限)
 *   5.(建議)選 testCourseEmail → Run → 收到範例信確認版型
 *   6. 部署 → 新增部署 → 網頁應用程式
 *      執行身份:我  ·  存取權限:任何人
 *      按「部署」→ 拿到 Web App URL
 *   7. 把那串 URL 填到 course-enterprise/index.html 的 CONFIG.SHEET_WEBHOOK
 *
 *   未來改這份程式後要重部署:「管理部署 → 鉛筆編輯 → 版本『新版本』→ 部署」
 *   ⚠️ 不要按「新增部署」— 會產生新 URL,前端就抓不到。
 */

const SHEET_ID = 'YOUR_NEW_ENTERPRISE_SHEET_ID_HERE';  // ⚠️ 建立 Sheet 後填上
const SHEET_COURSE = '企業落地報名表單';

const COURSE_HEADERS = [
  '時間戳',
  '中文姓名', '職稱', 'E-Mail', '性別', '用餐選擇',
  '統編', '公司名稱', '產業別',
  '輔導的單位', '是否為受輔導廠商',
  '報名梯次',
  '聯絡人姓名', '聯絡人手機', '聯絡人E-Mail',
  '報名留言', '同意狀態',
  '寄信狀態'
];
const COURSE_MAIL_STATUS_COL = 18;

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
  return ContentService.createTextOutput('企業落地課程報名 API 運作中 ✓');
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

// ════════ 寫入:課程報名(寫資料 + 寄確認信) ════════
function logCourseSignup(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_COURSE);
  if (!sheet) sheet = ss.insertSheet(SHEET_COURSE, ss.getSheets().length);
  ensureHeaders(sheet, COURSE_HEADERS);
  sheet.appendRow([
    data.timestamp || new Date().toISOString(),
    data.name || '', data.jobTitle || '', data.email || '', data.gender || '', data.meal || '',
    data.taxId || '', data.companyName || '', data.industry || '',
    data.advisor || '', data.isAdvised || '',
    data.cohort || '',
    data.contactName || '', data.contactPhone || '', data.contactEmail || '',
    data.note || '', data.consent || '',
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

  const subject = '【報名確認】從 AI 導論到企業落地 · 製造業 AI 轉型一次學會';
  const body =
    (data.name || '') + ' ' + (data.jobTitle || '') + ' 您好,\n\n' +
    '感謝您報名「從 AI 導論到企業落地」課程,\n' +
    '30 小時 4 大單元,把 AI 導論、生成式設計、AI Agent、RAG 帶回公司。\n' +
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
    '　．30 小時 4 大單元(4 天實體)\n' +
    '　．單元 01 · AI 導論(林京賢)\n' +
    '　．單元 02 · AI 設計(林京賢)\n' +
    '　．單元 03 · AI Agent(謝侑霖)\n' +
    '　．單元 04 · AI 落地 · Workshop(謝侑霖)\n' +
    '　．結訓帶 3 個具體成果回公司直接用:\n' +
    '　　1. AI 加速的工作與設計產出\n' +
    '　　2. 企業自己的 AI Agent 工具\n' +
    '　　3. 部門級 RAG 知識庫(LLM Wiki)\n\n' +
    '▌行前提醒\n' +
    '　．請攜帶個人筆電,以便現場實作\n' +
    '　．課程當天請提早 10 分鐘報到\n' +
    '　．開課前 3 天將發送行前通知信\n' +
    '　．如需改期或取消,請於開課 3 日前來信告知\n\n' +
    '如有任何問題,歡迎隨時與我們聯繫,期待課堂上見!\n\n' +
    '──────────────────\n' +
    '本課程聯絡窗口\n' +
    '．虎智科技 紀先生 02-66058192\n' +
    '．工研院 許雅婷 07-2625889 分機 117\n';

  MailApp.sendEmail({ to: to, subject: subject, body: body, name: '虎智科技 TigerAI' });
}

// ════════ 測試工具 ════════

function setupCourseSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_COURSE);
  if (!sheet) sheet = ss.insertSheet(SHEET_COURSE, ss.getSheets().length);
  ensureHeaders(sheet, COURSE_HEADERS);
  Logger.log('分頁「' + SHEET_COURSE + '」就緒(' + COURSE_HEADERS.length + ' 欄)');
}

function testCourseEmail() {
  sendCourseConfirmEmail({
    name: '王大明', jobTitle: '生產部經理',
    email: Session.getActiveUser().getEmail(),
    gender: '男', meal: '葷',
    companyName: '測試股份有限公司', taxId: '12345678',
    industry: '機械設備業',
    advisor: '感測中心', isAdvised: '是',
    contactName: '李小華', contactPhone: '0912345678', contactEmail: 'contact@example.com',
    cohort: '新竹｜9月30日 · 10月1日 · 10月6日 · 10月7日｜新竹縣工業會',
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
    sheet.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#E0F2FE');
  }
}

function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
