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

const SHEET_ID = '1GCSnwhSErrTHAZYUdztsdBHEtFlLnFm1H3biuyCnNC4';
const SHEET_COURSE = 'AI工作流程報名表單';
const SHEET_COHORT = '梯次目錄';
const SHEET_BACKUP = '_備份_AI工作流程報名表單';   // 隱藏備份分頁 · 用於防誤刪
const COHORT_CAPACITY_DEFAULT = 30;   // 預設每梯次名額(梯次沒填就用這個)

// 2026-10-14 改為「次產業輔導團在職菁英培訓」版 · 欄位對齊 course-enterprise
const COURSE_HEADERS = [
  '時間戳',
  '中文姓名', '職稱', 'E-Mail', '性別', '用餐選擇',
  '統編', '公司名稱', '產業別',
  '輔導的單位', '是否為受輔導廠商',
  '報名梯次',
  '聯絡人姓名', '聯絡人手機', '聯絡人E-Mail',
  '報名留言', '同意狀態',
  '寄信狀態',
  '單場講師',   // col 19 · 空=雙講師 · 有值=單場(Sheet 填什麼帶什麼)
  '單場主題',   // col 20 · 單場的課程主題(Sheet 帶)
  '錄取狀態',   // col 21 · 正取 / 候補(後端自動判斷)
  '取消 Token'  // col 22 · 隱式 UUID · 給信件「取消報名」連結用
];
const COURSE_MAIL_STATUS_COL = 18;
const COURSE_ADMIT_STATUS_COL = 21;
const COURSE_CANCEL_TOKEN_COL = 22;
// 報名梯次在 col 12(courseCounts 要用)
const COURSE_COHORT_COL = 12;

// 梯次目錄 Sheet 欄位
//   單場講師/主題 → 空白=雙講師場;有填=單場專場
//   交通說明/停車說明 → 選填 · 每個場次填自己的 · 空白則信件不顯示該行
//   名額 → 選填 · 每個場次自己填 · 空白則用預設 30
const COHORT_HEADERS = ['開課縣市', '課程日期', '課程時間', '課程地點', '單場講師', '單場主題', '交通', '停車', '名額'];
// 範例場次(setupAll 第一次跑會塞這 1 列 · 之後你直接改 Sheet)
const COHORT_SEED = [
  ['嘉義', '2026/11/9 ~ 11/17(四日)', '每日 09:00-17:30 · 共 30h', '嘉義市西區國華街 245 號 5 樓', '', '',
    '嘉義火車站步行約 12 分鐘(約 800 公尺)',
    '大樓 3、4 樓設有停車場(費用依現場公告)',
    30]
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
  if (e && e.parameter) {
    if (e.parameter.action === 'counts') return courseCounts();
    if (e.parameter.action === 'cancel' && e.parameter.token) return handleCancel(e.parameter.token);
  }
  return ContentService.createTextOutput('AI 工作流程 · 嘉義班 報名 API 運作中 ✓');
}

// ════════ 處理取消報名 ════════
// 使用者點 Email 裡的連結(?action=cancel&token=XXX)· 回傳 HTML 頁面
function handleCancel(token) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_COURSE);
    if (!sheet || sheet.getLastRow() < 2) return cancelPage('error', '找不到報名資料');

    // 掃 Token 欄找對應列
    const tokens = sheet.getRange(2, COURSE_CANCEL_TOKEN_COL, sheet.getLastRow() - 1, 1).getValues();
    let targetRow = -1;
    for (let i = 0; i < tokens.length; i++) {
      if (String(tokens[i][0] || '').trim() === token) { targetRow = i + 2; break; }
    }
    if (targetRow === -1) return cancelPage('error', '取消連結無效 · 可能已取消過或連結失效');

    // 讀對應列資料
    const row = sheet.getRange(targetRow, 1, 1, COURSE_HEADERS.length).getValues()[0];
    const name = row[1], email = row[3], cohort = row[COURSE_COHORT_COL - 1];
    const currentAdmit = String(row[COURSE_ADMIT_STATUS_COL - 1] || '').trim();

    // 若已取消過,擋
    if (currentAdmit.indexOf('已取消') === 0) {
      return cancelPage('ok', '您已經取消過報名了 · 無需再次取消', { name, cohort });
    }

    // 標記為已取消
    const cancelStamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
    sheet.getRange(targetRow, COURSE_ADMIT_STATUS_COL).setValue('已取消 ' + cancelStamp + ' (原:' + currentAdmit + ')');

    // 寄確認信給對方
    try {
      const parts = String(cohort || '').split('｜');
      const city = (parts[0] || '').trim();
      const datetime = (parts[1] || '').trim();
      MailApp.sendEmail({
        to: email,
        subject: '【取消確認】AI 工作流程 · 嘉義班 已取消您的報名',
        body: name + ' 您好,\n\n' +
          '我們已收到您的取消通知,您報名的以下場次已取消:\n\n' +
          '　．開課縣市:' + city + '\n' +
          '　．上課時間:' + datetime + '\n\n' +
          '取消時間:' + cancelStamp + '\n\n' +
          '若為誤點取消 · 或改期後想重新報名 · 請直接聯繫:\n' +
          '　．虎智科技  紀先生  02-66058192\n' +
          '　．郭哿埼  0972-629215\n\n' +
          '感謝您曾支持本課程 · 期待未來再相見。\n',
        name: '虎智科技 TigerAI'
      });
    } catch (err) { /* 寄信失敗不中斷 */ }

    return cancelPage('ok', '已成功取消您的報名', { name, cohort, cancelStamp });
  } catch (err) {
    return cancelPage('error', '取消失敗 · 錯誤:' + err.toString());
  }
}

function cancelPage(status, message, meta) {
  meta = meta || {};
  const isOk = status === 'ok';
  const bg = isOk ? '#EEF4EC' : '#FEE2E2';
  const barColor = isOk ? '#4D9B6A' : '#DC2626';
  const emoji = isOk ? '✓' : '⚠';
  const title = isOk ? '已成功取消' : '取消失敗';

  const metaBlock = meta.name ? (
    '<div style="background:#fff;border:1px solid rgba(0,0,0,0.08);border-radius:12px;padding:20px;margin-top:24px;text-align:left;">' +
      '<div style="font-size:11px;letter-spacing:0.22em;color:#6B7A72;font-weight:700;margin-bottom:10px;">取消內容</div>' +
      '<div style="font-size:15px;color:#1F3A2B;line-height:1.9;">' +
        '<div><strong>姓名</strong>　' + (meta.name || '') + '</div>' +
        (meta.cohort ? '<div><strong>場次</strong>　' + meta.cohort + '</div>' : '') +
        (meta.cancelStamp ? '<div><strong>時間</strong>　' + meta.cancelStamp + '</div>' : '') +
      '</div>' +
    '</div>'
  ) : '';

  const html =
    '<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<title>' + title + ' · AI 工作流程</title>' +
    '<style>body{margin:0;padding:0;background:' + bg + ';font-family:"Noto Sans TC","PingFang TC","Microsoft JhengHei",sans-serif;min-height:100vh;display:flex;align-items:center;justify-content:center;}' +
    '.card{background:#fff;max-width:480px;width:92%;padding:48px 32px 40px;border-radius:20px;text-align:center;box-shadow:0 20px 60px -20px rgba(31,58,43,0.25);}' +
    '.icon{width:72px;height:72px;border-radius:50%;background:' + barColor + ';color:#fff;font-size:40px;line-height:72px;display:inline-block;margin-bottom:20px;font-weight:700;}' +
    '.title{font-size:26px;font-weight:900;color:#1F3A2B;margin:0 0 12px;}' +
    '.msg{color:#3E6B52;font-size:16px;line-height:1.8;margin:0 0 8px;}' +
    '.note{color:#9EACA3;font-size:12px;letter-spacing:0.18em;margin-top:28px;font-weight:700;}' +
    '</style></head><body><div class="card">' +
    '<div class="icon">' + emoji + '</div>' +
    '<h1 class="title">' + title + '</h1>' +
    '<p class="msg">' + message + '</p>' +
    metaBlock +
    '<div class="note">虎智科技 TigerAI</div>' +
    '</div></body></html>';

  return HtmlService.createHtmlOutput(html).setTitle(title + ' · AI 工作流程');
}

// 從梯次目錄抓這場的 meta(capacity / 交通 / 停車)
// cohortString 格式:'城市｜日期 時間｜地點'
function getCohortMeta(cohortString) {
  const fallback = { capacity: COHORT_CAPACITY_DEFAULT, transportInfo: '', parkingInfo: '' };
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const cohortSheet = ss.getSheetByName(SHEET_COHORT);
    if (!cohortSheet || cohortSheet.getLastRow() < 2) return fallback;
    const rows = cohortSheet.getRange(2, 1, cohortSheet.getLastRow() - 1, COHORT_HEADERS.length).getValues();
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const rowLabel = [String(r[0] || '').trim(),
                        [String(r[1] || '').trim(), String(r[2] || '').trim()].filter(Boolean).join(' '),
                        String(r[3] || '').trim()].filter(Boolean).join('｜');
      if (rowLabel === cohortString) {
        return {
          capacity: parseInt(r[8], 10) || COHORT_CAPACITY_DEFAULT,  // col I 名額
          transportInfo: String(r[6] || '').trim(),                 // col G 交通
          parkingInfo:   String(r[7] || '').trim()                  // col H 停車
        };
      }
    }
  } catch (err) { /* 查不到就用 fallback */ }
  return fallback;
}

// 回傳各場次目前報名人數
function courseCounts() {
  const counts = {};
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_COURSE);
    if (sheet && sheet.getLastRow() > 1) {
      // 報名梯次 = 第 12 欄
      const values = sheet.getRange(2, COURSE_COHORT_COL, sheet.getLastRow() - 1, 1).getValues();
      values.forEach(r => {
        const key = String(r[0] || '').trim();
        if (key) counts[key] = (counts[key] || 0) + 1;
      });
    }
  } catch (err) { /* 出錯回空 */ }
  return ContentService.createTextOutput(JSON.stringify(counts))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════ 寫入:報名 + 寄確認信(含 3 層備份)════════
function logCourseSignup(data) {
  // 【第 1 層】console.log 原始資料 · 絕對不可刪
  // Google Cloud Logging 保留 ≥ 7 天 · Sheet 全毀也救得回
  console.log('[SIGNUP_RAW]', new Date().toISOString(), JSON.stringify(data));

  const ss = SpreadsheetApp.openById(SHEET_ID);

  // 【第 3 層】DIAG log · 診斷寫入狀況
  console.log('[DIAG] Spreadsheet:', ss.getName(), '(' + ss.getId() + ')');
  console.log('[DIAG] 所有分頁:', ss.getSheets().map(s => s.getName() + '(' + s.getLastRow() + ')').join(' / '));

  let sheet = ss.getSheetByName(SHEET_COURSE);
  if (!sheet) sheet = ss.insertSheet(SHEET_COURSE, ss.getSheets().length);
  ensureHeaders(sheet, COURSE_HEADERS);

  // 產一組取消 token(給 Email 連結用)
  const cancelToken = Utilities.getUuid();
  data._cancelToken = cancelToken;

  const rowData = [
    data.timestamp || new Date().toISOString(),
    // 學員 (2-6)
    data.name || '', data.jobTitle || '', data.email || '', data.gender || '', data.meal || '',
    // 公司 (7-9)
    data.taxId || '', data.companyName || '', data.industry || '',
    // 輔導 (10-11)
    data.advisor || '', data.isAdvised || '',
    // 梯次 (12)
    data.cohort || '',
    // 聯絡窗口 (13-15)
    data.contactName || '', data.contactPhone || '', data.contactEmail || '',
    // 其他 (16-17)
    data.note || '', data.consent || '',
    // 系統 (18-22)
    '',                            // col 18 · 寄信狀態(下方 setValue 填)
    data.soloInstructor || '',     // col 19 · 單場講師
    data.soloTopic || '',          // col 20 · 單場主題
    '',                            // col 21 · 錄取狀態(下方 setValue 填)
    cancelToken                    // col 22 · 取消 token
  ];

  console.log('[DIAG] 主表 getLastRow 寫入前:', sheet.getLastRow());
  sheet.appendRow(rowData);
  const row = sheet.getLastRow();
  console.log('[DIAG] 主表 getLastRow 寫入後:', row);

  // 寫入驗證:讀回看是不是剛剛寫的
  const verifyName = sheet.getRange(row, 2).getValue();
  if (verifyName !== (data.name || '')) {
    console.log('[DIAG] ⚠️⚠️⚠️ 寫入驗證失敗 · 讀回姓名:', verifyName, '預期:', data.name);
  }

  // 【錄取判定】寫入後 · 查「同一場次」目前有幾人(含我自己)· 第 N 位 > 名額 就是候補
  const myCohort = (data.cohort || '').trim();
  const cohortMeta = getCohortMeta(myCohort);
  const capacity = cohortMeta.capacity;
  let positionInCohort = 1;
  if (myCohort && sheet.getLastRow() > 1) {
    const allCohorts = sheet.getRange(2, COURSE_COHORT_COL, sheet.getLastRow() - 1, 1).getValues();
    positionInCohort = allCohorts.filter(r => String(r[0] || '').trim() === myCohort).length;
  }
  const isWaitlist = positionInCohort > capacity;
  const admitStatus = isWaitlist ? '候補 (第 ' + (positionInCohort - capacity) + ' 位)' : '正取 (第 ' + positionInCohort + ' 位)';
  sheet.getRange(row, COURSE_ADMIT_STATUS_COL).setValue(admitStatus);
  console.log('[DIAG] 錄取狀態:', admitStatus, '(同梯次第', positionInCohort, '位 / 名額', capacity, ')');

  // 【第 2 層】隱藏備份分頁 · 雙寫(主表萬一被誤刪 · 這份還在)
  try {
    let backup = ss.getSheetByName(SHEET_BACKUP);
    if (!backup) {
      backup = ss.insertSheet(SHEET_BACKUP, ss.getSheets().length);
      ensureHeaders(backup, COURSE_HEADERS);
      backup.hideSheet();
      const protection = backup.protect().setDescription('備份用 · 禁止編輯');
      protection.removeEditors(protection.getEditors());
      if (protection.canDomainEdit()) protection.setDomainEdit(false);
    }
    const backupRowData = rowData.slice();
    backupRowData[COURSE_ADMIT_STATUS_COL - 1] = admitStatus;  // 備份也記錄錄取狀態
    backup.appendRow(backupRowData);
    console.log('[DIAG] 備份分頁已寫入 · getLastRow:', backup.getLastRow());
  } catch (err) {
    console.log('[DIAG] ❌ 備份失敗(主表仍成功):', err.toString());
  }

  // 寄確認信(傳 isWaitlist 進去 · 決定要不要加候補提醒)
  let mailStatus;
  try {
    sendCourseConfirmEmail(data, isWaitlist, positionInCohort);
    mailStatus = (isWaitlist ? '[候補] ' : '') + '✅ 已寄出 ' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
  } catch (err) {
    mailStatus = '❌ 寄信失敗:' + err.message;
  }
  sheet.getRange(row, COURSE_MAIL_STATUS_COL).setValue(mailStatus);

  return jsonResp({ ok: true });
}

// ════════ 救援 · 從 Gmail 寄件夾回補資料 ════════
// 執行方式:Apps Script 編輯器函式下拉選 recoverFromGmailSent → ▶️ 執行
// 以 Email 為 dedup key · Sheet 已有則跳過
function recoverFromGmailSent() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_COURSE);
  if (!sheet) throw new Error('找不到報名表單分頁 ' + SHEET_COURSE);
  ensureHeaders(sheet, COURSE_HEADERS);

  const existing = new Set();
  if (sheet.getLastRow() > 1) {
    const emailCol = sheet.getRange(2, 4, sheet.getLastRow() - 1, 1).getValues();
    emailCol.forEach(r => {
      const e = String(r[0] || '').trim().toLowerCase();
      if (e) existing.add(e);
    });
  }
  console.log('Sheet 現有 ' + existing.size + ' 筆(以 Email 計)');

  const query = 'from:me subject:("【報名確認】AI 工作流程") newer_than:90d';
  const threads = GmailApp.search(query, 0, 200);
  console.log('Gmail 搜到 ' + threads.length + ' 封確認信');

  const addedRows = [];
  const skipped = [];
  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const to = (msg.getTo() || '').toLowerCase().replace(/.*<|>.*/g, '').trim();
      if (!to) return;
      if (existing.has(to)) { skipped.push(to); return; }
      const body = msg.getPlainBody();
      const parsed = parseConfirmEmail(body);
      if (!parsed.name) { skipped.push(to + '(parse 失敗)'); return; }
      existing.add(to);
      addedRows.push([
        Utilities.formatDate(msg.getDate(), 'Asia/Taipei', "yyyy-MM-dd'T'HH:mm:ss"),
        // 學員
        parsed.name, parsed.jobTitle, to, '', parsed.meal,
        // 公司(回補信件沒有 · 填已知欄位)
        parsed.taxId, parsed.companyName, '',
        // 輔導(回補信件沒有)
        '', '',
        // 梯次
        parsed.cohort,
        // 聯絡窗口(回補信件沒有)
        '', '', '',
        // 其他
        '', '[Gmail 回補]',
        // 系統
        '✅ 回補 ' + Utilities.formatDate(msg.getDate(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm'),
        parsed.soloInstructor, parsed.soloTopic,
        '[Gmail 回補 · 未判定]',
        ''                         // 取消 Token · 回補不給(不能取消歷史資料)
      ]);
    });
  });

  if (addedRows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, addedRows.length, COURSE_HEADERS.length).setValues(addedRows);
  }
  console.log('✅ 補回 ' + addedRows.length + ' 筆 · 跳過 ' + skipped.length + ' 筆');
  return { added: addedRows.length, skipped: skipped.length };
}

function parseConfirmEmail(body) {
  const grab = (re) => { const m = body.match(re); return m ? m[1].trim() : ''; };
  const studentLine = grab(/報名學員[::]\s*(.+)/);
  let name = '', jobTitle = '';
  if (studentLine) {
    const sp = studentLine.split(/\s*\/\s*/);
    name = (sp[0] || '').trim();
    jobTitle = (sp[1] || '').trim();
  }
  const city     = grab(/開課縣市[::]\s*(.+)/);
  const datetime = grab(/上課時間[::]\s*(.+)/);
  const place    = grab(/上課地點[::]\s*(.+)/);
  const cohort = [city, datetime, place].filter(x => x && x !== '—').join('｜');
  const soloRaw = grab(/講師模式[::]\s*(.+)/);
  let soloInstructor = '', soloTopic = '';
  if (soloRaw) {
    const sp = soloRaw.split(/\s*·\s*/);
    soloInstructor = (sp[0] || '').trim();
    soloTopic = (sp[1] || '').trim();
  }
  return {
    name, jobTitle,
    phone:       grab(/手機\s*[::]\s*(.+)/),
    companyName: grab(/公司名稱[::]\s*(.+)/),
    taxId:       grab(/公司統編[::]\s*(.+)/),
    referral:    grab(/邀請單位[::]\s*(.+)/),
    meal:        grab(/用餐選擇[::]\s*(.+)/),
    cohort, soloInstructor, soloTopic
  };
}

// ════════ 寄送:報名確認信 ════════
// 兩種模式:雙講師(現行寫死內容) · 通用單場(講師名/主題全由 Sheet 帶,套到任何講師)
// isWaitlist = true 時信件主旨加「【候補】」· 內文加候補提醒 block
function sendCourseConfirmEmail(data, isWaitlist, positionInCohort) {
  const to = String(data.email || '').trim();
  if (!to) throw new Error('無 Email');

  const parts = String(data.cohort || '').split('｜');
  const city     = (parts[0] || '—').trim();
  const datetime = (parts[1] || '—').trim();
  const place    = (parts[2] || '—').trim();
  const solo     = String(data.soloInstructor || '').trim();   // 有值 = 單場
  const topic    = String(data.soloTopic || '').trim() || '單場專題';

  // 從梯次目錄抓這場的 meta(名額/交通/停車)
  const cohortMeta = getCohortMeta(data.cohort);
  const transportInfo = cohortMeta.transportInfo;
  const parkingInfo = cohortMeta.parkingInfo;
  const capacity = cohortMeta.capacity;

  // 候補提醒 banner(放信件最上方 · 很醒目 · 名額跟著 Sheet 的設定走)
  const waitlistBanner = isWaitlist ? (
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '📋 您為本梯次候補(第 ' + (positionInCohort - capacity) + ' 位候補)\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    '本梯次已達 ' + capacity + ' 人上限 · 您為候補身分。\n' +
    '原則上您仍可以依預定時間到場上課。\n' +
    '若因名額控管無法參加,我們會「個別來信通知」您,請留意信箱。\n' +
    '若開課前未收到我們另行通知,代表候補成功 · 請依下方資訊按時到場。\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n'
  ) : '';

  let subject, headline, courseBlock;
  if (solo) {
    // ── 通用單場版 · 講師/主題從 Sheet 拉,不 hardcode ──
    subject = (isWaitlist ? '【候補】' : '【報名確認】') + 'AI 工作流程 · ' + solo + ' 單場 · ' + topic;
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

  // 取消報名連結(每筆報名有獨一無二 token)
  const cancelToken = data._cancelToken || '';
  const cancelUrl = cancelToken
    ? (ScriptApp.getService().getUrl() + '?action=cancel&token=' + encodeURIComponent(cancelToken))
    : '';

  // ────── 純文字版 body(給舊 email client fallback)──────
  const cancelBlockText = cancelUrl ? (
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '🔗 如需取消報名,請點以下連結(一次性 · 點了即生效):\n' +
    cancelUrl + '\n' +
    '點了會寄取消確認信給您 · 無需再回信告知\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n'
  ) : '';

  const body =
    waitlistBanner +
    (data.name || '') + ' ' + (data.jobTitle || '') + ' 您好,\n\n' +
    headline +
    '▌您的報名資訊\n' +
    '　公司名稱:' + (data.companyName || '') + '\n' +
    (data.taxId ? '　公司統編:' + data.taxId + '\n' : '') +
    '　產業別  :' + (data.industry || '') + '\n' +
    '　報名學員:' + (data.name || '') + ' / ' + (data.jobTitle || '') + '\n' +
    '　性別    :' + (data.gender || '') + '\n' +
    '　用餐選擇:' + (data.meal || '') + '\n' +
    '　輔導單位:' + (data.advisor || '') + '\n' +
    '　受輔導  :' + (data.isAdvised || '') + '\n' +
    '　報名梯次:\n' +
    '　　．開課縣市:' + city + '\n' +
    '　　．上課時間:' + datetime + '\n' +
    '　　．上課地點:' + place + '\n' +
    (solo ? '　　．講師模式:' + solo + ' · ' + topic + '\n' : '') +
    '　公司聯絡窗口:' + (data.contactName || '') + ' / ' + (data.contactPhone || '') + ' / ' + (data.contactEmail || '') + '\n' +
    '\n' +
    courseBlock +
    '▌行前提醒\n' +
    '　．請攜帶個人筆電(MacOS / Windows 皆可)\n' +
    '　．課程當天請提早 10 分鐘報到\n' +
    '　．開課前 3 天將發送行前通知信\n' +
    (transportInfo ? '　．交通:' + transportInfo + '\n' : '') +
    (parkingInfo   ? '　．停車:' + parkingInfo + '\n' : '') +
    '　．如需改期或取消,請於開課 3 日前來信告知\n\n' +
    cancelBlockText +
    '如有任何問題,歡迎隨時與我們聯繫,期待課堂上見!\n\n' +
    '──────────────────\n' +
    '本課程聯絡窗口\n' +
    '．虎智科技  紀先生  02-66058192\n' +
    '．郭哿埼  KE CHI  0972629215\n';

  // ────── HTML 版 body(現代 email client 看這個 · 有按鈕)──────
  const htmlBody = buildHtmlEmail({
    name: data.name || '', jobTitle: data.jobTitle || '',
    headline, courseBlock,
    city, datetime, place,
    data, solo, topic,
    transportInfo, parkingInfo,
    isWaitlist, positionInCohort, capacity,
    cancelUrl
  });

  MailApp.sendEmail({ to: to, subject: subject, body: body, htmlBody: htmlBody, name: '虎智科技 TigerAI' });
}

// ════════════════════════════════════════════════════════
// 虎智科技 TigerAI 標準 EDM Email 版型
// ────────────────────────────────────────────────────────
// 設計規範(其他課程 email 要延用):
//   · 深綠 brand header bar + 虎智 TigerAI logo
//   · Hero banner 大標題(主題名 + 課程名)
//   · 候補警示(橘底 · 若適用)
//   · 分區卡:報名資訊 / 梯次 / 課程 / 提醒(各自白底卡 · emoji icon)
//   · CTA 大按鈕 · 深綠 pill
//   · 深綠 footer bar + 聯絡窗口 + 版權
// Palette(與網站一致):
//   deep #1F3A2B · mid #3E6B52 · strong #4D9B6A
//   bg #EEF4EC · card #FFFFFF · card-tint #F6FAF4 · line #E0EBDC
//   text main #2C3A33 · text mid #6B7A72 · text light #9EACA3
// ════════════════════════════════════════════════════════
function buildHtmlEmail(opts) {
  const { name, jobTitle, headline, courseBlock, city, datetime, place,
          data, solo, topic, transportInfo, parkingInfo,
          isWaitlist, positionInCohort, capacity, cancelUrl } = opts;

  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  // ─── 色票 ───
  const C = {
    deep: '#1F3A2B', mid: '#3E6B52', strong: '#4D9B6A',
    bg: '#EEF4EC', card: '#FFFFFF', tint: '#F6FAF4', line: '#E0EBDC',
    textMain: '#2C3A33', textMid: '#6B7A72', textLight: '#9EACA3',
    warnBg: '#FFF4E6', warnBorder: '#F59E0B', warnText: '#B45309', warnDeep: '#78350F'
  };

  // ─── 共用小元件 ───
  const brandBar =
    '<div style="background:' + C.deep + ';padding:18px 28px;border-radius:16px 16px 0 0;text-align:left;">' +
      '<div style="color:#fff;font-size:16px;font-weight:900;letter-spacing:0.08em;">' +
        '<span style="color:' + C.strong + ';">⚡</span> 虎智科技 <span style="color:' + C.strong + ';">TigerAI</span>' +
      '</div>' +
    '</div>';

  const heroBanner =
    '<div style="background:' + C.tint + ';padding:28px 28px 24px;border-bottom:1px solid ' + C.line + ';">' +
      '<div style="font-size:11px;letter-spacing:0.3em;color:' + C.strong + ';font-weight:700;margin-bottom:8px;">' +
        (isWaitlist ? '📋 候補登記確認' : '✓ 報名已確認') +
      '</div>' +
      '<div style="font-size:22px;font-weight:900;color:' + C.deep + ';line-height:1.3;">AI 工作流程 · 嘉義班</div>' +
      '<div style="font-size:13px;color:' + C.mid + ';margin-top:6px;">從 AI 工具使用到智慧工作流程 · 四日 30 小時實體課程</div>' +
    '</div>';

  const waitlistCard = isWaitlist ? (
    '<div style="margin:20px 28px 0;background:' + C.warnBg + ';border-left:4px solid ' + C.warnBorder + ';border-radius:8px;padding:16px 20px;">' +
      '<div style="font-size:14px;font-weight:900;color:' + C.warnText + ';margin-bottom:8px;">📋 您為本梯次候補(第 ' + (positionInCohort - capacity) + ' 位候補)</div>' +
      '<div style="font-size:13px;color:' + C.warnDeep + ';line-height:1.75;">' +
        '本梯次已達 ' + capacity + ' 人上限。原則上您仍可依預定時間到場 · ' +
        '若因名額控管無法參加 · 我們會<strong>個別來信通知</strong>。' +
      '</div>' +
    '</div>'
  ) : '';

  const greeting =
    '<div style="padding:24px 28px 0;">' +
      '<div style="font-size:16px;color:' + C.deep + ';font-weight:700;margin-bottom:8px;">' +
        esc(name) + ' ' + esc(jobTitle) + ' 您好,' +
      '</div>' +
      '<div style="font-size:14px;color:' + C.textMain + ';line-height:1.9;">' +
        esc(headline).replace(/\n/g, '<br>').replace(/,/g, '，') +
      '</div>' +
    '</div>';

  // 分區卡 heading
  const sectionHead = (icon, label) =>
    '<div style="padding:0 28px;margin-top:28px;">' +
      '<div style="display:inline-block;background:' + C.deep + ';color:#fff;padding:6px 14px;border-radius:999px;font-size:11px;letter-spacing:0.22em;font-weight:700;">' +
        esc(icon) + ' &nbsp; ' + esc(label) +
      '</div>' +
    '</div>';

  const infoRow = (k, v) => v ? (
    '<tr>' +
      '<td style="padding:10px 16px;color:' + C.textMid + ';font-size:12px;letter-spacing:0.1em;white-space:nowrap;vertical-align:top;width:30%;border-bottom:1px solid ' + C.line + ';">' + esc(k) + '</td>' +
      '<td style="padding:10px 16px;color:' + C.deep + ';font-size:14px;line-height:1.7;border-bottom:1px solid ' + C.line + ';">' + esc(v) + '</td>' +
    '</tr>'
  ) : '';

  // ─── Card 1:報名資訊 ───
  const signupCard =
    sectionHead('👤', '您的報名資訊') +
    '<div style="padding:0 28px;margin-top:12px;">' +
      '<table style="width:100%;border-collapse:collapse;background:' + C.card + ';border:1px solid ' + C.line + ';border-radius:10px;overflow:hidden;">' +
        infoRow('公司名稱', data.companyName) +
        infoRow('公司統編', data.taxId) +
        infoRow('產業別', data.industry) +
        infoRow('報名學員', (data.name || '') + ' / ' + (data.jobTitle || '')) +
        infoRow('性別', data.gender) +
        infoRow('用餐選擇', data.meal) +
        infoRow('輔導單位', data.advisor) +
        infoRow('是否受輔導', data.isAdvised) +
        infoRow('聯絡窗口', (data.contactName || '') + ' / ' + (data.contactPhone || '') + ' / ' + (data.contactEmail || '')) +
      '</table>' +
    '</div>';

  // ─── Card 2:報名梯次 ───
  const cohortCard =
    sectionHead('📅', '報名梯次') +
    '<div style="padding:0 28px;margin-top:12px;">' +
      '<div style="background:' + C.tint + ';border:1px solid ' + C.line + ';border-radius:10px;padding:18px 20px;">' +
        '<div style="font-size:18px;font-weight:900;color:' + C.deep + ';margin-bottom:12px;">' + esc(city) + '</div>' +
        '<div style="font-size:14px;color:' + C.textMain + ';line-height:1.9;">' +
          '<div><span style="color:' + C.textMid + ';font-size:12px;letter-spacing:0.1em;margin-right:10px;">時間</span>' + esc(datetime) + '</div>' +
          '<div><span style="color:' + C.textMid + ';font-size:12px;letter-spacing:0.1em;margin-right:10px;">地點</span>' + esc(place) + '</div>' +
          (solo ? '<div style="margin-top:10px;padding:8px 12px;background:' + C.strong + ';color:#fff;border-radius:6px;font-size:12px;font-weight:700;display:inline-block;">★ 單場專題 · ' + esc(solo) + ' · ' + esc(topic) + '</div>' : '') +
        '</div>' +
      '</div>' +
    '</div>';

  // ─── Card 3:課程資訊 ───
  const courseContent = esc(courseBlock.replace(/^▌課程資訊\n/, '')).replace(/\n/g, '<br>');
  const courseCard =
    sectionHead('📚', '課程資訊') +
    '<div style="padding:0 28px;margin-top:12px;">' +
      '<div style="background:' + C.card + ';border:1px solid ' + C.line + ';border-radius:10px;padding:18px 20px;font-size:14px;color:' + C.textMain + ';line-height:1.9;">' +
        courseContent +
      '</div>' +
    '</div>';

  // ─── Card 4:行前提醒 ───
  const reminders = [
    { icon: '💻', text: '請攜帶個人筆電(MacOS / Windows 皆可)' },
    { icon: '⏰', text: '課程當天請提早 10 分鐘報到' },
    { icon: '📨', text: '開課前 3 天將發送行前通知信' },
    transportInfo ? { icon: '🚉', text: '交通 · ' + transportInfo } : null,
    parkingInfo   ? { icon: '🅿️', text: '停車 · ' + parkingInfo } : null,
    { icon: '📝', text: '如需改期或取消 · 可點下方按鈕 · 或於開課 3 日前來信告知' }
  ].filter(Boolean);
  const reminderList = reminders.map(r =>
    '<tr>' +
      '<td style="padding:10px 0 10px 4px;width:28px;font-size:16px;vertical-align:top;line-height:1.6;">' + r.icon + '</td>' +
      '<td style="padding:10px 0;color:' + C.textMain + ';font-size:14px;line-height:1.75;border-bottom:1px solid ' + C.line + ';">' + esc(r.text) + '</td>' +
    '</tr>'
  ).join('');
  const reminderCard =
    sectionHead('⚡', '行前提醒') +
    '<div style="padding:0 28px;margin-top:12px;">' +
      '<div style="background:' + C.card + ';border:1px solid ' + C.line + ';border-radius:10px;padding:4px 20px 12px;">' +
        '<table style="width:100%;border-collapse:collapse;">' + reminderList + '</table>' +
      '</div>' +
    '</div>';

  // ─── 取消按鈕 ───
  const cancelCard = cancelUrl ? (
    '<div style="padding:36px 28px 8px;text-align:center;">' +
      '<a href="' + esc(cancelUrl) + '" style="display:inline-block;background:' + C.deep + ';color:#fff!important;text-decoration:none;padding:15px 42px;border-radius:999px;font-weight:900;font-size:15px;letter-spacing:0.1em;box-shadow:0 10px 24px -8px rgba(31,58,43,0.45);">' +
        '✕ &nbsp; 取消報名' +
      '</a>' +
      '<div style="font-size:12px;color:' + C.textLight + ';margin-top:12px;letter-spacing:0.08em;line-height:1.7;">' +
        '點了會寄取消確認信給您 · 無需再回信告知 · 謝謝' +
      '</div>' +
    '</div>'
  ) : '';

  // ─── Footer ───
  const footer =
    '<div style="background:' + C.deep + ';padding:28px;border-radius:0 0 16px 16px;margin-top:36px;">' +
      '<div style="font-size:11px;letter-spacing:0.3em;color:' + C.strong + ';font-weight:700;margin-bottom:14px;text-align:center;">本課程聯絡窗口</div>' +
      '<div style="font-size:14px;color:#fff;line-height:2;text-align:center;">' +
        '<div><strong>虎智科技 · 紀先生</strong> &nbsp;<span style="color:#B6D4BC;font-family:monospace;">02-66058192</span></div>' +
        '<div><strong>郭哿埼 KE CHI</strong> &nbsp;<span style="color:#B6D4BC;font-family:monospace;">0972-629-215</span></div>' +
      '</div>' +
    '</div>';

  // ─── 組合 ───
  return (
    '<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>' + (isWaitlist ? '候補登記確認' : '報名確認') + '</title></head>' +
    '<body style="margin:0;padding:0;background:' + C.bg + ';font-family:\'Noto Sans TC\',\'PingFang TC\',\'Microsoft JhengHei\',sans-serif;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%;background:' + C.bg + ';padding:28px 16px;">' +
      '<tr><td align="center">' +
        '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="max-width:640px;width:100%;background:' + C.card + ';border-radius:16px;overflow:hidden;box-shadow:0 20px 60px -24px rgba(31,58,43,0.25);">' +
          '<tr><td>' + brandBar + '</td></tr>' +
          '<tr><td>' + heroBanner + '</td></tr>' +
          (waitlistCard ? '<tr><td>' + waitlistCard + '</td></tr>' : '') +
          '<tr><td>' + greeting + '</td></tr>' +
          '<tr><td>' + signupCard + '</td></tr>' +
          '<tr><td>' + cohortCard + '</td></tr>' +
          '<tr><td>' + courseCard + '</td></tr>' +
          '<tr><td>' + reminderCard + '</td></tr>' +
          '<tr><td>' + cancelCard + '</td></tr>' +
          '<tr><td>' + footer + '</td></tr>' +
        '</table>' +
        '<div style="font-size:11px;color:' + C.textLight + ';margin-top:16px;letter-spacing:0.2em;">© 虎智科技 TigerAI</div>' +
      '</td></tr>' +
    '</table>' +
    '</body></html>'
  );
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

// 從 Sheet 的梯次目錄抓第一筆 · 動態組測試用 cohort 字串
function getFirstCohortForTest() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_COHORT);
    if (!sheet || sheet.getLastRow() < 2) return '嘉義｜待設定｜待設定';
    const row = sheet.getRange(2, 1, 1, COHORT_HEADERS.length).getValues()[0];
    const city = (row[0] || '').trim();
    const date = (row[1] || '').trim();
    const time = (row[2] || '').trim();
    const place = (row[3] || '').trim();
    const solo = (row[4] || '').trim();
    const topic = (row[5] || '').trim();
    const cohortStr = [city, [date, time].filter(Boolean).join(' '), place].filter(Boolean).join('｜');
    return { cohort: cohortStr, soloInstructor: solo, soloTopic: topic };
  } catch (err) {
    Logger.log('⚠️ 讀 Sheet 梯次失敗 · 用 fallback:' + err.toString());
    return { cohort: '嘉義｜待設定｜待設定', soloInstructor: '', soloTopic: '' };
  }
}

// 測試用資料(所有 testCourseEmail* 共用 · 新欄位全部塞入)
function _buildTestData(c) {
  return {
    name: '王大明', jobTitle: '生產部經理',
    email: Session.getActiveUser().getEmail(),
    gender: '男', meal: '葷',
    taxId: '12345678', companyName: '測試電子股份有限公司', industry: '電子零組件業',
    advisor: '工研院綠能所', isAdvised: '是',
    cohort: c.cohort,
    contactName: '李經理', contactPhone: '0912345678', contactEmail: Session.getActiveUser().getEmail(),
    note: '素食一份',
    _cancelToken: 'TEST-' + new Date().getTime()
  };
}

function testCourseEmail() {
  const c = getFirstCohortForTest();
  sendCourseConfirmEmail(_buildTestData(c), false, 1);
  Logger.log('已寄【正取版】測試信至 ' + Session.getActiveUser().getEmail());
}

// 候補版測試(第 31 位 · 應該會出現候補 banner 跟取消連結)
function testCourseEmailWaitlist() {
  const c = getFirstCohortForTest();
  sendCourseConfirmEmail(_buildTestData(c), true, 31);
  Logger.log('已寄【候補版】測試信至 ' + Session.getActiveUser().getEmail());
}

// 通用單場版測試 · 從 Sheet 第一筆有「單場講師」值的梯次抓;沒有就用假的
function testCourseEmailSolo() {
  const c = getFirstCohortForTest();
  const data = _buildTestData(c);
  data.soloInstructor = c.soloInstructor || '張惟翔 · AI 增幅者';
  data.soloTopic = c.soloTopic || '提示詞工程 + Agent 建置單日工作坊';
  data.note = '';
  sendCourseConfirmEmail(data, false, 1);
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
