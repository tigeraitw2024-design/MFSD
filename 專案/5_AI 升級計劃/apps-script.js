/**
 * Google Apps Script · AI 升級計劃(在職菁英課程)報名後端
 * 課程:從自動化到智動化：手把手帶您做出智慧製造AI 升級計畫
 *
 * 這份程式「綁定」你開啟 Apps Script 時所在的那份 Google Sheet,
 * 不用再填 Sheet ID。
 *
 * 分頁:
 *   第 1 個分頁(gid=0)…… 梯次目錄:前端讀「發布到網路」的 CSV
 *       A 欄=開課縣市  B 欄=課程日期  C 欄=課程時間  D 欄=課程地點(第 1 列標題,資料從第 2 列起)
 *   AI升級計劃報名表單 ……… 報名資料:doPost 自動建立、寫入 + 寄確認信
 *
 * ════════ 部署步驟 ════════
 *   1. 打開新 Sheet → 擴充功能 → Apps Script → 把 Code.gs 的預設內容全部刪掉,貼上這整份
 *   2. 存檔(Ctrl+S)
 *   3. 上方函式下拉選 setupCourseSheet → 執行
 *      → 會跳授權視窗,一路同意(需要 Sheet 與 Gmail 寄信權限)
 *      → 會自動建好「AI升級計劃報名表單」分頁與欄位
 *   4.(建議)函式選 testCourseEmail → 執行 → 你會收到範例信,確認版型
 *   5. 部署 → 新增部署 → 類型選「網頁應用程式」
 *      執行身份:我
 *      存取權限:任何人
 *      按「部署」→ 複製 Web App URL(這就是 SHEET_WEBHOOK)
 *   6. 把那串 URL 貼給 Claude,填進 index.html 的 CONFIG.SHEET_WEBHOOK
 *
 *   ⚠️ 未來改這份程式後要重部署:「管理部署 → 鉛筆編輯 → 版本選『新版本』→ 部署」
 *      千萬不要按「新增部署」,那會產生新 URL,前端就抓不到。
 */

const SHEET_COURSE = 'AI升級計劃報名表單';

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

// 綁定這份 Sheet(從 Sheet 的「擴充功能 → Apps Script」開出來的專案才有效)
function getSS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('找不到綁定的 Sheet,請從 Sheet 的「擴充功能 → Apps Script」建立這份程式');
  return ss;
}

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
  return ContentService.createTextOutput('AI 升級計劃報名 API 運作中 ✓');
}

// 回傳各梯次目前報名人數(只回數字,不含個資)
function courseCounts() {
  const counts = {};
  try {
    const sheet = getSS().getSheetByName(SHEET_COURSE);
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
  const ss = getSS();
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

  const subject = '【報名確認】從自動化到智動化：手把手帶您做出智慧製造AI 升級計畫';
  const body =
    (data.name || '') + ' ' + (data.jobTitle || '') + ' 您好,\n\n' +
    '感謝您報名「從自動化到智動化：手把手帶您做出智慧製造AI 升級計畫」在職菁英課程,\n' +
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
    '．工研院  許雅婷  07-2625889 分機 117\n';

  MailApp.sendEmail({ to: to, subject: subject, body: body, name: '虎智科技 TigerAI' });
}

// ════════ 測試工具 ════════

// 貼完程式先手動跑一次:建好「AI升級計劃報名表單」分頁並補齊欄位、順便完成授權
function setupCourseSheet() {
  const ss = getSS();
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
    industry: '機械設備業',
    advisor: '感測中心', isAdvised: '是',
    contactName: '李小華', contactPhone: '0912345678', contactEmail: 'contact@example.com',
    cohort: '台中｜2026/10/17-18 + 10/24-25 09:00-17:00｜地點另行公告',
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
    sheet.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#FDECDF');
  }
}

function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
