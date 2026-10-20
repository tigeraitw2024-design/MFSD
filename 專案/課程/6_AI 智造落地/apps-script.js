/**
 * Google Apps Script · AI 智造落地(在職菁英課程)報名後端
 * 課程:AI 智造落地實戰班：從找出痛點到產線升級的務實指南
 *
 * 這份程式「綁定」你開啟 Apps Script 時所在的那份 Google Sheet,
 * 不用再填 Sheet ID。
 *
 * 分頁:
 *   第 1 個分頁(gid=0)…… 梯次目錄:前端讀「發布到網路」的 CSV
 *       A 欄=開課縣市  B 欄=課程日期  C 欄=課程時間  D 欄=課程地點(第 1 列標題,資料從第 2 列起)
 *   AI智造落地報名表單 ……… 報名資料:doPost 自動建立、寫入 + 寄確認信
 *
 * ════════ 部署步驟 ════════
 *   1. 打開新 Sheet → 擴充功能 → Apps Script → 把 Code.gs 的預設內容全部刪掉,貼上這整份
 *   2. 存檔(Ctrl+S)
 *   3. 上方函式下拉選 setupCourseSheet → 執行
 *      → 會跳授權視窗,一路同意(需要 Sheet 與 Gmail 寄信權限)
 *      → 會自動建好「AI智造落地報名表單」分頁與欄位
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

const SHEET_COURSE = 'AI智造落地報名表單';

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
  return ContentService.createTextOutput('AI 智造落地報名 API 運作中 ✓');
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
  // ────── DIAGNOSTIC LOGGING(抓資料為什麼不見的 bug · 2026-10-14 加)──────
  console.log('[DIAG] === logCourseSignup 開始 ===');
  console.log('[DIAG] 時間:', new Date().toISOString());
  console.log('[DIAG] 收到報名:', JSON.stringify({ name: data.name, email: data.email, company: data.companyName }));

  const ss = getSS();
  console.log('[DIAG] Spreadsheet ID:', ss.getId());
  console.log('[DIAG] Spreadsheet URL:', ss.getUrl());
  console.log('[DIAG] Spreadsheet 名稱:', ss.getName());
  console.log('[DIAG] 所有分頁:', ss.getSheets().map(s => s.getName() + '(' + s.getLastRow() + ' 列)').join(' / '));

  let sheet = ss.getSheetByName(SHEET_COURSE);
  if (!sheet) {
    console.log('[DIAG] ⚠️ 找不到分頁 ' + SHEET_COURSE + ' · 建立新分頁');
    sheet = ss.insertSheet(SHEET_COURSE, ss.getSheets().length);
  }
  console.log('[DIAG] 目標分頁:', sheet.getName(), '· getLastRow(寫入前):', sheet.getLastRow());

  ensureHeaders(sheet, COURSE_HEADERS);
  console.log('[DIAG] ensureHeaders 完成 · getLastRow:', sheet.getLastRow());

  const rowData = [
    data.timestamp || new Date().toISOString(),
    data.name || '',
    data.jobTitle || '',
    data.email || '',
    data.gender || '',
    data.meal || '',
    data.taxId || '',
    data.companyName || '',
    data.industry || '',
    data.advisor || '',
    data.isAdvised || '',
    data.cohort || '',
    data.contactName || '',
    data.contactPhone || '',
    data.contactEmail || '',
    data.note || '',
    data.consent || '',
    ''
  ];

  try {
    sheet.appendRow(rowData);
    console.log('[DIAG] ✅ appendRow 成功');
  } catch (err) {
    console.log('[DIAG] ❌ appendRow 失敗:', err.toString(), err.stack);
    throw err;
  }

  const row = sheet.getLastRow();
  console.log('[DIAG] 寫入後 getLastRow:', row);
  // 驗證寫入 · 讀回那一列第 2 欄(姓名)看是不是剛剛寫的
  const verifyName = sheet.getRange(row, 2).getValue();
  console.log('[DIAG] 驗證讀回姓名:', verifyName, '預期:', data.name);
  if (verifyName !== data.name) {
    console.log('[DIAG] ⚠️⚠️⚠️ 寫入驗證失敗 · 讀回的姓名跟預期不一樣!');
  }

  let mailStatus;
  try {
    sendCourseConfirmEmail(data);
    mailStatus = '✅ 已寄出 ' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
    console.log('[DIAG] 寄信成功');
  } catch (err) {
    mailStatus = '❌ 寄信失敗:' + err.message;
    console.log('[DIAG] 寄信失敗:', err.toString());
  }
  sheet.getRange(row, COURSE_MAIL_STATUS_COL).setValue(mailStatus);
  console.log('[DIAG] 寄信狀態已填回 col', COURSE_MAIL_STATUS_COL);
  console.log('[DIAG] === logCourseSignup 結束 ===');

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

  const subject = '【報名確認】AI 智造落地實戰班：從找出痛點到產線升級的務實指南';
  const body =
    (data.name || '') + ' ' + (data.jobTitle || '') + ' 您好,\n\n' +
    '感謝您報名「AI 智造落地實戰班：從找出痛點到產線升級的務實指南」在職菁英課程,\n' +
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
    '．工研院  胡家寧  02-2341-4105 分機 350\n';

  MailApp.sendEmail({ to: to, subject: subject, body: body, name: '虎智科技 TigerAI' });
}

// ════════ 緊急:從 Gmail 寄件夾救回報名資料 ════════
// 用法:
//   1. 把這份 code 貼到 Apps Script 存檔
//   2. 函式下拉選 recoverFromGmailSent · 執行(第一次會要你授權 Gmail 讀取權限 · 同意)
//   3. 執行紀錄會告訴你補了幾筆
//   4. 看 Sheet · 補回的資料會在最後面
// 說明:
//   - 以「報名學員的 Email」當 dedup key · Sheet 已有就跳過
//   - Email 裡沒有的欄位(性別、聯絡人 E-Mail、報名留言、同意狀態)留空
//   - 時間戳用寄信時間代替(不是原本報名的 timestamp,但接近)
function recoverFromGmailSent() {
  const ss = getSS();
  let sheet = ss.getSheetByName(SHEET_COURSE);
  if (!sheet) throw new Error('找不到報名表單分頁 ' + SHEET_COURSE);
  ensureHeaders(sheet, COURSE_HEADERS);

  // 1. 收集現有 Sheet 的 Email(作為 dedup key)
  const existing = new Set();
  if (sheet.getLastRow() > 1) {
    const emailCol = sheet.getRange(2, 4, sheet.getLastRow() - 1, 1).getValues();   // D 欄 E-Mail
    emailCol.forEach(r => {
      const e = String(r[0] || '').trim().toLowerCase();
      if (e) existing.add(e);
    });
  }
  console.log('Sheet 現有 ' + existing.size + ' 筆報名(以 Email 計)');

  // 2. 搜尋 Gmail 寄件夾所有報名確認信(過去 60 天)
  const query = 'from:me subject:("【報名確認】AI 智造落地") newer_than:60d';
  const threads = GmailApp.search(query, 0, 200);
  console.log('Gmail 搜到 ' + threads.length + ' 封報名確認信 thread');

  const addedRows = [];
  const skipped = [];
  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const to = (msg.getTo() || '').toLowerCase().replace(/.*<|>.*/g, '').trim();
      if (!to) return;
      if (existing.has(to)) { skipped.push(to); return; }

      const body = msg.getPlainBody();
      const parsed = parseConfirmEmail(body);
      if (!parsed.name) { skipped.push(to + ' (無法 parse)'); return; }

      existing.add(to);
      addedRows.push([
        Utilities.formatDate(msg.getDate(), 'Asia/Taipei', "yyyy-MM-dd'T'HH:mm:ss"),
        parsed.name,                 // 中文姓名
        parsed.jobTitle,             // 職稱
        to,                          // E-Mail
        '',                          // 性別(信裡沒有 · 留空)
        parsed.meal,                 // 用餐選擇
        parsed.taxId,                // 統編
        parsed.companyName,          // 公司名稱
        parsed.industry,             // 產業別
        parsed.advisor,              // 輔導的單位
        parsed.isAdvised,            // 是否為受輔導廠商
        parsed.cohort,               // 報名梯次(拼回 city｜datetime｜place)
        parsed.contactName,          // 聯絡人姓名
        parsed.contactPhone,         // 聯絡人手機
        '',                          // 聯絡人 E-Mail(信裡沒有 · 留空)
        '',                          // 報名留言(信裡沒有 · 留空)
        '[Gmail 回補]',              // 同意狀態
        '✅ 已寄出 ' + Utilities.formatDate(msg.getDate(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm') + '(回補)'
      ]);
    });
  });

  // 3. 批次寫入
  if (addedRows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, addedRows.length, COURSE_HEADERS.length).setValues(addedRows);
  }
  console.log('✅ 補回 ' + addedRows.length + ' 筆 · 跳過 ' + skipped.length + ' 筆(已存在或無法 parse)');
  if (skipped.length > 0) console.log('跳過的:', skipped.slice(0, 20).join(', '));
  return { added: addedRows.length, skipped: skipped.length };
}

// 從信件 body 解析各欄位
function parseConfirmEmail(body) {
  const grab = (re) => {
    const m = body.match(re);
    return m ? m[1].trim() : '';
  };
  // 「報名學員:王大明 / 生產部經理」
  const studentLine = grab(/報名學員[::]\s*(.+)/);
  let name = '', jobTitle = '';
  if (studentLine) {
    const sp = studentLine.split(/\s*\/\s*/);
    name = (sp[0] || '').trim();
    jobTitle = (sp[1] || '').trim();
  }
  // 「公司聯絡窗口:XX / 09xx」
  const contactLine = grab(/公司聯絡窗口[::]\s*(.+)/);
  let contactName = '', contactPhone = '';
  if (contactLine) {
    const sp = contactLine.split(/\s*\/\s*/);
    contactName = (sp[0] || '').trim();
    contactPhone = (sp[1] || '').trim();
  }
  // 報名梯次拆城市/時間/地點 → 拼回 city｜datetime｜place
  const city     = grab(/開課縣市[::]\s*(.+)/);
  const datetime = grab(/上課時間[::]\s*(.+)/);
  const place    = grab(/上課地點[::]\s*(.+)/);
  const cohort = [city, datetime, place].filter(x => x && x !== '—').join('｜');

  return {
    name, jobTitle,
    companyName: grab(/公司名稱[::]\s*(.+)/),
    taxId:       grab(/公司統編[::]\s*(.+)/),
    industry:    grab(/產業別\s*[::]\s*(.+)/),
    meal:        grab(/用餐選擇[::]\s*(.+)/),
    advisor:     grab(/輔導單位[::]\s*(.+)/),
    isAdvised:   grab(/是否受輔導[::]\s*(.+)/),
    cohort,
    contactName, contactPhone
  };
}

// ════════ 測試工具 ════════

// 貼完程式先手動跑一次:建好「AI智造落地報名表單」分頁並補齊欄位、順便完成授權
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
    advisor: '工研院電光所', isAdvised: '是',
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
