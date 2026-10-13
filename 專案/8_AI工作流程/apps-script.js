/**
 * Google Apps Script · AI 工作流程(嘉義班) 報名後端
 *
 * ⚠️ 本專案獨立於其他課程,需另建 Sheet + 部署新 Web App
 *
 * ════════ 部署步驟(SOP 跟前面課程一樣) ════════
 *   1. 新建 Google Sheet,複製 URL 裡的 SPREADSHEET_ID 貼到下面 SHEET_ID
 *   2. Sheet → 擴充功能 → Apps Script → 貼這整份到 Code.gs
 *   3. Ctrl+S 存檔
 *   4. 函式下拉選 setupAll → Run(第一次會跳授權,同意 Gmail 寄信權限)
 *      會自動建:報名表單分頁 + 梯次目錄分頁(+ 3 列範例資料)
 *   5.(建議)選 testCourseEmail → Run → 收到範例信確認版型
 *   6. 部署 → 新增部署 → 網頁應用程式
 *      執行身份:我  ·  存取權限:任何人
 *      按「部署」→ 拿到 Web App URL
 *   7. 把那串 URL 填到 course-ai-workflow/index.html 的 CONFIG.SHEET_WEBHOOK
 *   8. 檔案 → 共用 → 發布到網路 → 選「梯次目錄」+ 逗號分隔(.csv) → 發布
 *      把 CSV URL 填到 index.html 的 CONFIG.COHORTS_CSV_URL
 *
 *   ⚠️ 未來改這份程式後要重部署:「管理部署 → 鉛筆編輯 → 版本『新版本』→ 部署」
 *      千萬不要按「新增部署」— 會產生新 URL,前端就抓不到。
 */

const SHEET_ID = 'PASTE_SHEET_ID_HERE';      // ← 新建 Sheet 後把 URL 裡 /d/ 跟 /edit 中間那串貼這
const SHEET_COURSE = 'AI工作流程報名表單';
const SHEET_COHORT = '梯次目錄';

const COURSE_HEADERS = [
  '時間戳',
  '中文姓名', '職稱', 'E-Mail', '手機', '性別', '用餐選擇',
  '公司名稱', '統編', '邀請單位',
  '報名場次',
  '報名留言', '同意狀態',
  '寄信狀態',
  '單場講師',   // col 15 · 空=雙講師 · 有值=單場(Sheet 填什麼帶什麼)
  '單場主題'    // col 16 · 單場的課程主題(Sheet 帶)
];
const COURSE_MAIL_STATUS_COL = 14;

// 梯次目錄:單場講師 + 單場主題兩欄搭配使用
//   單場講師欄空白 → 雙講師場;有填 → 單場專場,顯示什麼寫什麼
//   單場主題欄 → 這場的課程內容/主題敘述(前端提示條與信件都會直接引用)
const COHORT_HEADERS = ['開課縣市', '課程日期', '課程時間', '課程地點', '單場講師', '單場主題'];
// 範例場次(setupAll 第一次跑會塞這 1 列 · 之後你直接改 Sheet)
const COHORT_SEED = [
  ['嘉義', '2026/11/9 ~ 11/17(四日)', '每日 09:00-17:30 · 共 30h', '嘉義市西區國華街 245 號 5 樓', '', '']
];

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
  return ContentService.createTextOutput('AI 工作流程 · 嘉義班 報名 API 運作中 ✓');
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
    data.companyName || '', data.taxId || '', data.referral || '',
    data.cohort || '',
    data.note || '', data.consent || '',
    '',                            // col 14 · 寄信狀態(下方 setValue 填)
    data.soloInstructor || '',     // col 15 · 單場講師
    data.soloTopic || ''           // col 16 · 單場主題
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
// 兩種模式:雙講師(現行寫死內容) · 通用單場(講師名/主題全由 Sheet 帶,套到任何講師)
function sendCourseConfirmEmail(data) {
  const to = String(data.email || '').trim();
  if (!to) throw new Error('無 Email');

  const parts = String(data.cohort || '').split('｜');
  const city     = (parts[0] || '—').trim();
  const datetime = (parts[1] || '—').trim();
  const place    = (parts[2] || '—').trim();
  const solo     = String(data.soloInstructor || '').trim();   // 有值 = 單場
  const topic    = String(data.soloTopic || '').trim() || '單場專題';

  let subject, headline, courseBlock;
  if (solo) {
    // ── 通用單場版 · 講師/主題從 Sheet 拉,不 hardcode ──
    subject = '【報名確認】AI 工作流程 · ' + solo + ' 單場 · ' + topic;
    headline =
      '感謝您報名「AI 工作流程 · ' + solo + ' 單場」,\n' +
      '本場為 ' + solo + ' 單講師專場 · 主題:' + topic + '。\n' +
      '我們已收到您的報名資料。\n\n';
    courseBlock =
      '▌課程資訊\n' +
      '　．本場為單講師專場(非四日 30h 完整版)\n' +
      '　．講師:' + solo + '\n' +
      '　．主題:' + topic + '\n\n';
  } else {
    // ── 雙講師版(AI 工作流程 嘉義班)──
    subject = '【報名確認】AI 工作流程 · 嘉義班 · 從 AI 工具到智慧工作流程';
    headline =
      '感謝您報名「AI 工作流程 · 嘉義班」四日 30 小時實體課程,\n' +
      '張惟翔(AI 增幅者)+ 謝侑霖(AI 自動症)雙講師 · 把 AI 真正裝進製造現場與部門工作。\n' +
      '我們已收到您的報名資料。\n\n';
    courseBlock =
      '▌課程資訊\n' +
      '　．四日 30 小時實體實戰(2026/11/9、11/10、11/16、11/17)\n' +
      '　．每日 09:00-17:30 · 單日 7.5h\n' +
      '　．雙講師:張惟翔(AI 增幅者)· 謝侑霖(AI 自動症)\n' +
      '　．5 大主題:\n' +
      '　　1. 生成式 AI 理論、通用模型認知與提示詞工程\n' +
      '　　2. 企業 AI Agent 建置與調校\n' +
      '　　3. AI 簡報與營運行銷流程自動化\n' +
      '　　4. 企業流程自動化設計與建構\n' +
      '　　5. 技術文件 AI 化與導入落地規劃\n' +
      '　．完訓後能看懂 AI、拆得解流程、建得起應用、推得動落地\n\n';
  }

  const body =
    (data.name || '') + ' ' + (data.jobTitle || '') + ' 您好,\n\n' +
    headline +
    '▌您的報名資訊\n' +
    '　公司名稱:' + (data.companyName || '') + '\n' +
    (data.taxId ? '　公司統編:' + data.taxId + '\n' : '') +
    '　邀請單位:' + (data.referral || '') + '\n' +
    '　報名學員:' + (data.name || '') + ' / ' + (data.jobTitle || '') + '\n' +
    '　手機    :' + (data.phone || '') + '\n' +
    '　用餐選擇:' + (data.meal || '') + '\n' +
    '　報名場次:\n' +
    '　　．開課縣市:' + city + '\n' +
    '　　．上課時間:' + datetime + '\n' +
    '　　．上課地點:' + place + '\n' +
    (solo ? '　　．講師模式:' + solo + ' · ' + topic + '\n' : '') +
    '\n' +
    courseBlock +
    '▌行前提醒\n' +
    '　．請攜帶個人筆電(MacOS / Windows 皆可)\n' +
    '　．課程當天請提早 10 分鐘報到\n' +
    '　．開課前 3 天將發送行前通知信\n' +
    '　．交通:嘉義火車站步行約 12 分鐘(約 800 公尺)\n' +
    '　．停車:大樓 3、4 樓設有停車場(費用依現場公告)\n' +
    '　．如需改期或取消,請於開課 3 日前來信告知\n\n' +
    '如有任何問題,歡迎隨時與我們聯繫,期待課堂上見!\n\n' +
    '──────────────────\n' +
    '本課程聯絡窗口\n' +
    '．虎智科技  紀先生  02-66058192\n' +
    '．郭哿埼  KE CHI  0972629215\n';

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

// ⭐ 一鍵建全部:報名表單 + 梯次目錄 + 塞範例資料
function setupAll() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // 1. 報名表單分頁
  let sheet = ss.getSheetByName(SHEET_COURSE);
  if (!sheet) sheet = ss.insertSheet(SHEET_COURSE, ss.getSheets().length);
  ensureHeaders(sheet, COURSE_HEADERS);
  Logger.log('✅ 分頁「' + SHEET_COURSE + '」就緒(' + COURSE_HEADERS.length + ' 欄)');

  // 2. 梯次目錄分頁
  let cohortSheet = ss.getSheetByName(SHEET_COHORT);
  if (!cohortSheet) cohortSheet = ss.insertSheet(SHEET_COHORT, 0);   // 插到最前面
  ensureHeaders(cohortSheet, COHORT_HEADERS);
  if (cohortSheet.getLastRow() < 2) {
    cohortSheet.getRange(2, 1, COHORT_SEED.length, COHORT_HEADERS.length).setValues(COHORT_SEED);
    Logger.log('✅ 分頁「' + SHEET_COHORT + '」就緒 + 塞 ' + COHORT_SEED.length + ' 列範例場次');
  } else {
    Logger.log('✅ 分頁「' + SHEET_COHORT + '」就緒(已有資料,未覆蓋)');
  }

  Logger.log('全部建置完成 · 現在到「檔案 → 共用 → 發布到網路」把「梯次目錄」發布成 CSV');
}

function testCourseEmail() {
  sendCourseConfirmEmail({
    name: '王大明', jobTitle: '生產部經理',
    email: Session.getActiveUser().getEmail(),
    phone: '0912345678', gender: '男', meal: '葷',
    companyName: '測試電子股份有限公司', taxId: '12345678',
    referral: '虎智科技',
    cohort: '嘉義｜2026/11/9 ~ 11/17 每日 09:00-17:30｜嘉義市西區國華街 245 號 5 樓',
    note: '素食一份'
  });
  Logger.log('已寄雙講師版測試信至 ' + Session.getActiveUser().getEmail());
}

// 通用單場版測試 · 把 soloInstructor / soloTopic 換成任何值都能跑
function testCourseEmailSolo() {
  sendCourseConfirmEmail({
    name: '王大明', jobTitle: '生產部經理',
    email: Session.getActiveUser().getEmail(),
    phone: '0912345678', gender: '男', meal: '葷',
    companyName: '測試電子股份有限公司', taxId: '12345678',
    referral: '虎智科技',
    cohort: '嘉義｜2026/11/9 09:00-17:30｜嘉義市西區國華街 245 號 5 樓',
    soloInstructor: '張惟翔 · AI 增幅者',
    soloTopic: '提示詞工程 + Agent 建置單日工作坊',
    note: ''
  });
  Logger.log('已寄通用單場版測試信至 ' + Session.getActiveUser().getEmail());
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
    sheet.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#DCE8DC');
  }
}

function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
