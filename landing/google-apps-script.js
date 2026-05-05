/**
 * Google Apps Script · 報名表單後端
 *
 * 三個分頁（順序固定）：
 *   1. 報名名單 — 表單投遞原始資料
 *   2. 樞紐分析 — 自動匯總統計（用 COUNTIF 公式，會即時更新）
 *   3. 經銷商存取記錄 — 經銷商工具進入紀錄（IP / 國家 / 城市 / 機構 / 瀏覽器）
 *
 * ════════ 部署步驟 ════════
 *   1. 把這整份貼到 Apps Script 的 Code.gs（覆蓋）
 *   2. 上方 Run 下拉選單 → 選 setupSheets → 點 Run
 *      ↑ 一次性手動執行，會建立三個分頁、排序、寫入樞紐公式
 *   3. Deploy → Manage deployments → 鉛筆 → Version: New version → Deploy
 *   4. 確認部署 URL 跟 index.html 的 SHEET_WEBHOOK 一致
 */

const SHEET_ID = '1cua3wvQo747ecj7iXA5WFh17Rf-u2sIjTEyHeTh2erk';

const SHEET_SIGNUP = '報名名單';
const SHEET_PIVOT  = '樞紐分析';
const SHEET_DEALER = '經銷商存取記錄';

const SIGNUP_HEADERS = [
  '時間戳', '統編', '公司名稱', '資本額', '行業別', '公司地址', '資格判定',
  '姓名', '角色', '手機', 'Email', '公司規模', 'AI 導入階段',
  '想優先改善的場景', '希望達成的目標', '預計導入時程', '備註'
];

const DEALER_HEADERS = ['時間戳', 'IP', '國家', '城市', '機構 / ISP', '瀏覽器'];

// ════════ 入口 ════════
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.type === 'dealer_access') return logDealerAccess(data);
    return logSignup(data);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput('TigerAI 報名表單 API 運作中 ✓');
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
    data.capital || '',
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
// 一次性設定（手動執行，整理三個分頁順序 + 重建樞紐）
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
  signup.activate();
  ss.moveActiveSheet(1);
  pivot.activate();
  ss.moveActiveSheet(2);
  dealer.activate();
  ss.moveActiveSheet(3);

  signup.activate();

  SpreadsheetApp.getUi().alert(
    '✅ 已完成：\n' +
    '1️⃣ 報名名單\n' +
    '2️⃣ 樞紐分析（自動以公式更新）\n' +
    '3️⃣ 經銷商存取記錄'
  );
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
    ['50 人以下', `=COUNTIF(${SRC}!L:L, "50人以下")`, ''],
    ['51-200 人', `=COUNTIF(${SRC}!L:L, "51-200人")`, ''],
    ['201-1000 人', `=COUNTIF(${SRC}!L:L, "201-1000人")`, ''],
    ['1000 人以上', `=COUNTIF(${SRC}!L:L, "1000人以上")`, ''],
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
