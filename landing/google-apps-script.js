/**
 * Google Apps Script · 報名表單後端
 *
 * ════════ 部署步驟 ════════
 * 1. 到 Google Drive 新建一個 Google Sheet，取名如「虎智補助報名名單」
 * 2. Sheet 第一列貼上欄位標題（複製下方 HEADERS 陣列的值，用 Tab 分隔）
 * 3. 在 Sheet 的「擴充功能」→「Apps Script」
 * 4. 把這整個檔案的內容貼進去（覆蓋預設的 Code.gs）
 * 5. 把下方 SHEET_ID 換成你 Sheet 的 ID（URL 中 /d/ 後面那段）
 * 6. 儲存，點「部署」→「新增部署作業」
 *    - 類型：網頁應用程式
 *    - 執行身分：我
 *    - 存取權：所有人
 * 7. 授權後會拿到「網頁應用程式 URL」，複製此 URL
 * 8. 把 URL 貼到 index.html 的 CONFIG.SHEET_WEBHOOK
 * 9. 完成！
 */

const SHEET_ID = '1cua3wvQo747ecj7iXA5WFh17Rf-u2sIjTEyHeTh2erk';
const SHEET_NAME = '報名名單'; // 工作表名稱（分頁名，非 Sheet 檔名）

const HEADERS = [
  '時間戳',
  '統編',
  '公司名稱',
  '資本額',
  '行業別',
  '公司地址',
  '資格判定',
  '姓名',
  '角色',
  '手機',
  'Email',
  '公司規模',
  'AI 導入階段',
  '想優先改善的場景',
  '希望達成的目標',
  '預計導入時程',
  '備註'
];

function doPost(e) {
  try {
    // 解析來自前端的 JSON
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // 每次都確保 row 1 標題跟 HEADERS 一致；欄位增減或改名都會自動對齊
    ensureHeaders(sheet);

    // 組出一列資料
    const row = [
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
    ];
    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput('TigerAI 報名表單 API 運作中 ✓');
}

/**
 * 確保第 1 列標題跟 HEADERS 陣列同步。
 * 若長度不同、任一欄不同，就整列覆寫並重刷樣式 + 凍結首列。
 */
function ensureHeaders(sheet) {
  const width = HEADERS.length;
  const lastCol = sheet.getLastColumn();
  let needsUpdate = false;

  if (lastCol < width) {
    needsUpdate = true;
  } else {
    const current = sheet.getRange(1, 1, 1, width).getValues()[0];
    needsUpdate = HEADERS.some((h, i) => current[i] !== h);
  }

  if (needsUpdate) {
    sheet.getRange(1, 1, 1, width).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, width)
      .setFontWeight('bold')
      .setBackground('#0D1B2A')
      .setFontColor('#00E5CC');
    sheet.setFrozenRows(1);
  }
}
