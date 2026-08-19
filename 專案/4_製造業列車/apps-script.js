/**
 * Google Apps Script · 製造業 AI 巡迴列車 報名後端
 *
 * ⚠️ 本專案獨立於 3_企業落地,需另建 Sheet + 部署新 Web App
 *
 * ════════ 部署步驟(SOP 跟前面課程一樣) ════════
 *   1. 新建 Google Sheet,複製 URL 裡的 SPREADSHEET_ID 貼到下面 SHEET_ID
 *   2. Sheet → 擴充功能 → Apps Script → 貼這整份到 Code.gs
 *   3. Ctrl+S 存檔
 *   4. 函式下拉選 setupCourseSheet → Run(第一次會跳授權,同意 Gmail 寄信權限)
 *   5.(建議)選 testCourseEmail → Run → 收到範例信確認版型
 *   6. 部署 → 新增部署 → 網頁應用程式
 *      執行身份:我  ·  存取權限:任何人
 *      按「部署」→ 拿到 Web App URL
 *   7. 把那串 URL 填到 course-tour/index.html 的 CONFIG.SHEET_WEBHOOK
 *
 *   ⚠️ 未來改這份程式後要重部署:「管理部署 → 鉛筆編輯 → 版本『新版本』→ 部署」
 *      千萬不要按「新增部署」— 會產生新 URL,前端就抓不到。
 */

const SHEET_ID = 'YOUR_NEW_TOUR_SHEET_ID_HERE';   // ⚠️ 建立 Sheet 後填上
const SHEET_COURSE = '製造業列車報名表單';

const COURSE_HEADERS = [
  '時間戳',
  '中文姓名', '職稱', 'E-Mail', '手機', '性別', '用餐選擇',
  '公司名稱', '統編', '產業別',
  '報名場次',
  '報名留言', '同意狀態',
  '寄信狀態'
];
const COURSE_MAIL_STATUS_COL = 14;

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
  return ContentService.createTextOutput('製造業 AI 巡迴列車 報名 API 運作中 ✓');
}

// 回傳各場次目前報名人數
function courseCounts() {
  const counts = {};
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_COURSE);
    if (sheet && sheet.getLastRow() > 1) {
      // 報名場次 = 第 11 欄
      const values = sheet.getRange(2, 11, sheet.getLastRow() - 1, 1).getValues();
      values.forEach(r => {
        const key = String(r[0] || '').trim();
        if (key) counts[key] = (counts[key] || 0) + 1;
      });
    }
  } catch (err) { /* 出錯回空 */ }
  return ContentService.createTextOutput(JSON.stringify(counts))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════ 寫入:報名 + 寄確認信 ════════
function logCourseSignup(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_COURSE);
  if (!sheet) sheet = ss.insertSheet(SHEET_COURSE, ss.getSheets().length);
  ensureHeaders(sheet, COURSE_HEADERS);
  sheet.appendRow([
    data.timestamp || new Date().toISOString(),
    data.name || '', data.jobTitle || '', data.email || '', data.phone || '', data.gender || '', data.meal || '',
    data.companyName || '', data.taxId || '', data.industry || '',
    data.cohort || '',
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

  const subject = '【報名確認】製造業 AI 巡迴列車 · 用 AI 蓋工廠 × 學資安裝門鎖';
  const body =
    (data.name || '') + ' ' + (data.jobTitle || '') + ' 您好,\n\n' +
    '感謝您報名「製造業 AI 巡迴列車」6 小時實戰課程,\n' +
    'Morris + Victor 雙講師 · n8n × Antigravity + AI Agent 資安一次補齊。\n' +
    '我們已收到您的報名資料。\n\n' +
    '▌您的報名資訊\n' +
    '　公司名稱:' + (data.companyName || '') + '\n' +
    (data.taxId ? '　公司統編:' + data.taxId + '\n' : '') +
    '　產業別  :' + (data.industry || '') + '\n' +
    '　報名學員:' + (data.name || '') + ' / ' + (data.jobTitle || '') + '\n' +
    '　手機    :' + (data.phone || '') + '\n' +
    '　用餐選擇:' + (data.meal || '') + '\n' +
    '　報名場次:\n' +
    '　　．開課縣市:' + city + '\n' +
    '　　．上課時間:' + datetime + '\n' +
    '　　．上課地點:' + place + '\n\n' +
    '▌課程資訊\n' +
    '　．6 小時實體實戰(一日完訓)\n' +
    '　．上半場 · Morris:AI 落地地圖、Antigravity SOP → Skill、n8n 全流程、企業 AI 治理\n' +
    '　．下半場 · Victor:LLM 資安死角、Agent 攻擊面、OpenClaw Gateway 實作、治理實務\n' +
    '　．完訓帶 3 個具體成果:\n' +
    '　　1. 能跑的 AI Skill(SOP 版)\n' +
    '　　2. 能跑的 n8n 自動化流程\n' +
    '　　3. 能守的 OpenClaw 安全 Gateway\n\n' +
    '▌行前提醒\n' +
    '　．請攜帶個人筆電(MacOS / Windows 皆可)\n' +
    '　．課程當天請提早 10 分鐘報到\n' +
    '　．開課前 3 天將發送行前通知信\n' +
    '　．如需改期或取消,請於開課 3 日前來信告知\n\n' +
    '如有任何問題,歡迎隨時與我們聯繫,期待課堂上見!\n\n' +
    '──────────────────\n' +
    '本課程聯絡窗口\n' +
    '．虎智科技  紀先生  02-66058192\n';

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
    phone: '0912345678', gender: '男', meal: '葷',
    companyName: '測試股份有限公司', taxId: '12345678',
    industry: '機械設備業',
    cohort: '台北｜3月15日(六) 09:00-17:00｜虎智科技教室',
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
