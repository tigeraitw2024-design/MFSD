/**
 * Google Apps Script · 報名表單後端
 *
 * 分頁說明：
 *   報名名單 / 樞紐分析 / 經銷商存取記錄 … 既有專案的分頁（已有資料，勿動）
 *   課程報名表單 …………………………… 課程報名資料；自動建立在所有分頁「最後面」
 *      → 收到報名會自動寄「報名確認信」，並把寄信結果記在「寄信狀態」欄
 *
 * ════════ 部署步驟 ════════
 *   1. 把這整份貼到 Apps Script 的 Code.gs（覆蓋）
 *   2. 上方函式下拉選 setupCourseSheet → Run
 *      ↑ 因為新增了寄信功能，這次會跳「需要新權限」→ 請全部允許（授權 Gmail 寄信）
 *   3.（建議）函式選 testCourseEmail → Run → 會寄一封測試信到你自己信箱，確認版型
 *   4. Deploy → Manage deployments → 鉛筆 → Version: New version → Deploy
 *   5. 確認部署 URL 跟 index.html 的 SHEET_WEBHOOK 一致
 *
 *   ⚠️ 不要執行 setupSheets——那是舊專案用的，會重新排序分頁。
 *   ※ 連設定函式都不跑也行：第一筆課程報名進來時，logCourseSignup 會自動
 *      把「課程報名表單」分頁建在最後面並寫入＋寄信，全程不動其他分頁。
 */

const SHEET_ID = '1cua3wvQo747ecj7iXA5WFh17Rf-u2sIjTEyHeTh2erk';

const SHEET_SIGNUP = '報名名單';
const SHEET_PIVOT  = '樞紐分析';
const SHEET_DEALER = '經銷商存取記錄';
const SHEET_COURSE = '課程報名表單';

const SIGNUP_HEADERS = [
  '時間戳', '統編', '公司名稱', '勞工投保人數(200以下)', '行業別', '公司地址', '資格判定',
  '姓名', '角色', '手機', 'Email', '公司規模', 'AI 導入階段',
  '想優先改善的場景', '希望達成的目標', '預計導入時程', '備註'
];

const DEALER_HEADERS = ['時間戳', 'IP', '國家', '城市', '機構 / ISP', '瀏覽器'];

const COURSE_HEADERS = [
  '時間戳', '統編', '公司名稱', '資格判定', '報名梯次',
  '姓名', '職稱', '電話', 'Email', '年齡', '身分證字號', '寄信狀態', '公司人數'
];
const COURSE_MAIL_STATUS_COL = 12;   // 寄信狀態固定在第 12 欄

// ════════ 入口 ════════
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.type === 'dealer_access') return logDealerAccess(data);
    if (data.type === 'course_signup') return logCourseSignup(data);
    return logSignup(data);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'counts') return courseCounts();
  return ContentService.createTextOutput('TigerAI 報名表單 API 運作中 ✓');
}

// 回傳各梯次目前報名人數（只回傳彙總數字，不含任何個資）
// 課程網站用來顯示「剩餘名額」。前端呼叫：SHEET_WEBHOOK + '?action=counts'
function courseCounts() {
  const counts = {};
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_COURSE);
    if (sheet && sheet.getLastRow() > 1) {
      // 「報名梯次」= 第 5 欄
      const values = sheet.getRange(2, 5, sheet.getLastRow() - 1, 1).getValues();
      values.forEach(r => {
        const key = String(r[0] || '').trim();
        if (key) counts[key] = (counts[key] || 0) + 1;
      });
    }
  } catch (err) { /* 出錯就回空物件，前端會當作 0 */ }
  return ContentService.createTextOutput(JSON.stringify(counts))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════ 寫入：報名 ════════
function logSignup(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_SIGNUP);
  if (!sheet) sheet = ss.insertSheet(SHEET_SIGNUP);
  ensureHeaders(sheet, SIGNUP_HEADERS);
  sheet.appendRow([
    data.timestamp || new Date().toISOString(),
    data.taxId || '',
    data.companyName || '',
    data.laborInsurance || '',
    data.industry || '',
    data.address || '',
    data.eligibility || '',
    data.name || '',
    data.role || '',
    data.phone || '',
    data.email || '',
    data.companySize || '',
    data.aiStage || '',
    data.priorityScenarios || '',
    data.aiGoals || '',
    data.timeline || '',
    data.notes || ''
  ]);
  return ok();
}

// ════════ 寫入：經銷商存取 ════════
function logDealerAccess(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_DEALER);
  if (!sheet) sheet = ss.insertSheet(SHEET_DEALER);
  ensureHeaders(sheet, DEALER_HEADERS);
  sheet.appendRow([
    data.timestamp || new Date().toISOString(),
    data.ip || '',
    data.country || '',
    data.city || '',
    data.org || '',
    data.userAgent || ''
  ]);
  return ok();
}

// ════════ 寫入：課程報名表單（寫資料 + 寄確認信 + 記錄寄信狀態）════════
function logCourseSignup(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_COURSE);
  if (!sheet) sheet = ss.insertSheet(SHEET_COURSE, ss.getSheets().length);  // 找不到就建在最後面
  ensureHeaders(sheet, COURSE_HEADERS);
  sheet.appendRow([
    data.timestamp || new Date().toISOString(),
    data.taxId || '',
    data.companyName || '',
    data.eligibility || '',
    data.cohort || '',
    data.name || '',
    data.jobTitle || '',
    data.phone || '',
    data.email || '',
    data.age || '',
    data.nationalId || '',
    '',                       // 寄信狀態，下面寄完信再填回
    data.companySize || ''    // 公司人數（最後一欄）
  ]);
  const row = sheet.getLastRow();

  // 寄送報名確認信，並把結果寫回「寄信狀態」欄
  let mailStatus;
  try {
    sendCourseConfirmEmail(data);
    mailStatus = '✅ 已寄出 ' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
  } catch (err) {
    mailStatus = '❌ 寄信失敗：' + err.message;
  }
  sheet.getRange(row, COURSE_MAIL_STATUS_COL).setValue(mailStatus);

  return ok();
}

// ════════ 寄送：報名確認信 ════════
function sendCourseConfirmEmail(data) {
  const to = String(data.email || '').trim();
  if (!to) throw new Error('無 Email，未寄送');

  // 「報名梯次」存的是「縣市｜日期 時間｜地點」整串，拆開來顯示
  const parts = String(data.cohort || '').split('｜');
  const city     = (parts[0] || '—').trim();
  const datetime = (parts[1] || '—').trim();
  const place    = (parts[2] || '—').trim();

  const subject = '【報名確認】製造業銷售流程導入 AI 工具課程';
  const body =
    (data.name || '') + ' ' + (data.jobTitle || '') + ' 您好，\n\n' +
    '感謝您報名「製造業銷售流程導入 AI 工具」課程，\n' +
    '經銷商業績追蹤 ・ 銷售週報自動產出 ・ 客服訊息自動分類，\n' +
    '我們已收到您的報名資料。\n\n' +
    '▌您的報名資訊\n' +
    '　公司名稱：' + (data.companyName || '') + '\n' +
    '　公司統編：' + (data.taxId || '') + '\n' +
    '　報名學員：' + (data.name || '') + ' / ' + (data.jobTitle || '') + '\n' +
    '　聯絡電話：' + (data.phone || '') + '\n' +
    '　報名梯次：\n' +
    '　　．開課縣市：' + city + '\n' +
    '　　．上課時間：' + datetime + '\n' +
    '　　．上課地點：' + place + '\n\n' +
    '▌課程資訊\n' +
    '　．單日 6 小時實戰課程\n' +
    '　．結訓帶 3 條自動化流程回公司直接用：\n' +
    '　　1. 週一早上 8 點自動寄週報\n' +
    '　　2. 客戶 LINE 訊息 30 秒自動分流\n' +
    '　　3. 大客戶斷單隔天就跳通知\n' +
    '　．全部用你公司現在已經在用的 Sheets、LINE、Email 串起來\n\n' +
    '▌行前提醒\n' +
    '　．請攜帶個人筆電，以便現場實作\n' +
    '　．課程當天請提早 10 分鐘報到\n' +
    '　．如需改期或取消，請於開課 3 日前來信告知\n\n' +
    '如有任何問題，歡迎隨時與我們聯繫，期待課堂上見！\n\n' +
    '──────────────────\n' +
    '虎智科技 TigerAI\n' +
    '業務聯絡窗口｜AI 諮詢顧問 Evan Chi 紀如鴻\n' +
    'Email：evanchi@tigerai.tw\n' +
    '電話：886-960021437\n' +
    'LINE ID：evanvchi\n';

  MailApp.sendEmail({ to: to, subject: subject, body: body, name: '虎智科技 TigerAI' });
}

// 測試用：手動執行，寄一封範例確認信到你自己的信箱（確認版型用）
function testCourseEmail() {
  sendCourseConfirmEmail({
    name: '王大明', jobTitle: '生產部經理',
    companyName: '測試股份有限公司', taxId: '12345678',
    phone: '0912345678', email: Session.getActiveUser().getEmail(),
    cohort: '臺北市｜2026/6/5 10:00-17:00｜IEAT會議中心臺北市中山區松江路350號'
  });
  Logger.log('已寄測試信至 ' + Session.getActiveUser().getEmail());
}

function ok() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════ 工具：標題列同步 ════════
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
    sheet.getRange(1, 1, 1, width)
      .setFontWeight('bold')
      .setBackground('#0D1B2A')
      .setFontColor('#00E5CC');
    sheet.setFrozenRows(1);
  }
}

// ════════════════════════════════════════════════════════════════
// 課程報名表單 · 安全建立分頁
// 只建立「課程報名表單」這一個分頁並設定標題列；
// 不排序、不碰任何其他分頁、不重建樞紐。第一次使用前手動跑一次即可（不跑也行）。
// ════════════════════════════════════════════════════════════════
function setupCourseSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_COURSE);
  const created = !sheet;
  if (!sheet) sheet = ss.insertSheet(SHEET_COURSE, ss.getSheets().length);  // 新分頁建在最後面
  ensureHeaders(sheet, COURSE_HEADERS);
  const msg = (created ? '✅ 已建立分頁「' : '✅ 已確認分頁「') + SHEET_COURSE + '」並設定好標題列（位於最後一個分頁）';
  Logger.log(msg);   // 結果看下方「執行記錄」即可
}

// ════════════════════════════════════════════════════════════════
// ⚠️ 舊專案用：整理分頁順序 + 重建樞紐（會重新排序分頁，課程報名專案請勿執行）
// ════════════════════════════════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // 確保三個分頁都存在
  let signup = ss.getSheetByName(SHEET_SIGNUP) || ss.insertSheet(SHEET_SIGNUP);
  ensureHeaders(signup, SIGNUP_HEADERS);

  let pivot = ss.getSheetByName(SHEET_PIVOT) || ss.insertSheet(SHEET_PIVOT);
  rebuildPivot(pivot);

  let dealer = ss.getSheetByName(SHEET_DEALER) || ss.insertSheet(SHEET_DEALER);
  ensureHeaders(dealer, DEALER_HEADERS);

  // 重新排序：報名名單 (1) → 樞紐分析 (2) → 經銷商存取記錄 (3)
  // ※ 不碰「課程報名表單 / 梯次上架資訊」等其他分頁，避免動到既有資料順序
  signup.activate();
  ss.moveActiveSheet(1);
  pivot.activate();
  ss.moveActiveSheet(2);
  dealer.activate();
  ss.moveActiveSheet(3);

  signup.activate();

  const msg = '✅ 已完成：1️⃣ 報名名單 / 2️⃣ 樞紐分析 / 3️⃣ 經銷商存取記錄';
  Logger.log(msg);   // 結果看下方「執行記錄」即可
}

// ════════════════════════════════════════════════════════════════
// 樞紐分析公式重建
// ════════════════════════════════════════════════════════════════
function rebuildPivot(sheet) {
  sheet.clear();
  const SRC = `'${SHEET_SIGNUP}'`;

  const rows = [
    ['📊 報名樞紐分析（自動更新）', '', ''],
    ['更新時間', `=TEXT(NOW(),"yyyy-mm-dd hh:mm")`, ''],
    ['', '', ''],

    ['📋 總體統計', '人數', ''],
    ['總報名數', `=COUNTA(${SRC}!B2:B)`, ''],
    ['', '', ''],

    ['🏆 資格判定分布', '人數', ''],
    ['🎯 完全符合 (A+B)', `=COUNTIF(${SRC}!G:G, "green_full")`, ''],
    ['✅ 部分符合 (僅 B)', `=COUNTIF(${SRC}!G:G, "green_partial")`, ''],
    ['⚠️ 部分符合（狀態/資本問題）', `=COUNTIF(${SRC}!G:G, "yellow")`, ''],
    ['❌ 暫不符合', `=COUNTIF(${SRC}!G:G, "red")`, ''],
    ['', '', ''],

    ['👤 聯絡人角色', '人數', ''],
    ['資訊主管', `=COUNTIF(${SRC}!I:I, "資訊主管")`, ''],
    ['數位轉型負責人', `=COUNTIF(${SRC}!I:I, "數位轉型負責人")`, ''],
    ['部門主管', `=COUNTIF(${SRC}!I:I, "部門主管")`, ''],
    ['經營管理層', `=COUNTIF(${SRC}!I:I, "經營管理層")`, ''],
    ['專案窗口', `=COUNTIF(${SRC}!I:I, "專案窗口")`, ''],
    ['其他', `=COUNTIF(${SRC}!I:I, "其他")`, ''],
    ['', '', ''],

    ['🏢 公司規模', '人數', ''],
    ['1~30 人', `=COUNTIF(${SRC}!L:L, "1~30人")`, ''],
    ['31~100 人', `=COUNTIF(${SRC}!L:L, "31~100人")`, ''],
    ['101~199 人', `=COUNTIF(${SRC}!L:L, "101~199人")`, ''],
    ['200 人以上', `=COUNTIF(${SRC}!L:L, "200人以上")`, ''],
    ['', '', ''],

    ['🤖 AI 導入階段', '人數', ''],
    ['有興趣，尚未開始', `=COUNTIF(${SRC}!M:M, "有興趣，尚未開始")`, ''],
    ['已有零星試用', `=COUNTIF(${SRC}!M:M, "已有零星試用")`, ''],
    ['已有 PoC', `=COUNTIF(${SRC}!M:M, "已有PoC")`, ''],
    ['已有正式導入', `=COUNTIF(${SRC}!M:M, "已有正式導入")`, ''],
    ['已跨部門推進', `=COUNTIF(${SRC}!M:M, "已跨部門推進")`, ''],
    ['其他', `=COUNTIF(${SRC}!M:M, "其他")`, ''],
    ['', '', ''],

    ['⏱ 預計導入時程', '人數', ''],
    ['立即啟動', `=COUNTIF(${SRC}!P:P, "立即啟動")`, ''],
    ['3 個月內', `=COUNTIF(${SRC}!P:P, "3個月內")`, ''],
    ['6 個月內', `=COUNTIF(${SRC}!P:P, "6個月內")`, ''],
    ['年度規劃中', `=COUNTIF(${SRC}!P:P, "年度規劃中")`, ''],
    ['', '', ''],

    ['🎯 想優先改善的場景（多選統計）', '次數', ''],
    ['報價與接單', `=COUNTIF(${SRC}!N:N, "*報價與接單*")`, ''],
    ['自動報表', `=COUNTIF(${SRC}!N:N, "*自動報表*")`, ''],
    ['內部知識庫', `=COUNTIF(${SRC}!N:N, "*內部知識庫*")`, ''],
    ['文件處理', `=COUNTIF(${SRC}!N:N, "*文件處理*")`, ''],
    ['客服助理', `=COUNTIF(${SRC}!N:N, "*客服助理*")`, ''],
    ['跨系統流程', `=COUNTIF(${SRC}!N:N, "*跨系統流程*")`, ''],
    ['自動通知', `=COUNTIF(${SRC}!N:N, "*自動通知*")`, ''],
    ['內控簽核', `=COUNTIF(${SRC}!N:N, "*內控簽核*")`, ''],
    ['營運分析', `=COUNTIF(${SRC}!N:N, "*營運分析*")`, ''],
    ['地端 AI', `=COUNTIF(${SRC}!N:N, "*地端AI*")`, ''],
    ['其他', `=COUNTIF(${SRC}!N:N, "*其他*")`, ''],
    ['', '', ''],

    ['🚀 希望達成的目標（多選統計）', '次數', ''],
    ['降低人工時間', `=COUNTIF(${SRC}!O:O, "*降低人工時間*")`, ''],
    ['評估地端部署', `=COUNTIF(${SRC}!O:O, "*評估地端部署*")`, ''],
    ['整合分散流程', `=COUNTIF(${SRC}!O:O, "*整合分散流程*")`, ''],
    ['提升服務效率', `=COUNTIF(${SRC}!O:O, "*提升服務效率*")`, ''],
    ['建立治理架構', `=COUNTIF(${SRC}!O:O, "*建立治理架構*")`, ''],
    ['改善知識流動', `=COUNTIF(${SRC}!O:O, "*改善知識流動*")`, ''],
    ['其他', `=COUNTIF(${SRC}!O:O, "*其他*")`, '']
  ];

  sheet.getRange(1, 1, rows.length, 3).setValues(rows);

  // 區塊標題列：column B 是「人數」或「次數」就視為小節標題
  rows.forEach((row, i) => {
    if (row[1] === '人數' || row[1] === '次數') {
      sheet.getRange(i + 1, 1, 1, 3)
        .setFontWeight('bold')
        .setBackground('#0D1B2A')
        .setFontColor('#00E5CC');
    }
  });

  // 第一列大標題：合併、置中、深藍底白字
  sheet.getRange(1, 1, 1, 3).merge()
    .setFontWeight('bold').setFontSize(14)
    .setBackground('#1B3A5C').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');

  // 凍結前兩列（標題 + 更新時間）
  sheet.setFrozenRows(2);

  // 欄寬
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 80);
}
