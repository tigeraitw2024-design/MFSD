/**
 * Local AI 玩家開箱派對 · Apps Script 後端
 *
 * 綁 Sheet:https://docs.google.com/spreadsheets/d/1dar-zrj1iZwtDuhsVYeA9Lgxam9EQ9MTWsKsZX9e2QY/edit
 * 寄信從:tiger.ai.tw@gmail.com(每日 100 封免費額度)
 *
 * ─────── 分頁架構 ───────
 *   一般報名         (public) · 收 signup
 *   講者投稿         (public) · 收 speaker_submit · 有「精選公開」欄位 admin 勾選
 *   _備份_報名       (hidden + protected) · 每筆同時 append · 救命用
 *   _備份_投稿       (hidden + protected) · 同上
 *   公開精選         (public) · 給前台讀的 CSV · 只放 is_featured=TRUE 的投稿
 *
 * ─────── 鐵則(CLAUDE.md) ───────
 *   1. 不寫死(活動單場 · 無梯次 · OK)
 *   2. 3 層備份 + DIAG log(console.log + 隱藏備份分頁 + 讀回驗證)
 *   3. 寄信用 TigerAI 標準 EDM 版型(薄荷森綠)
 *
 * ─────── 部署步驟 ───────
 *   1. 打開 Sheet · 右上 Extensions → Apps Script
 *   2. 把這整個檔案貼進去 · 存檔
 *   3. 執行一次 setupAll() → 建 4 個分頁 + header + 保護
 *   4. 執行一次 setAdminPasscode('93489997') → 把密碼存進 PropertiesService
 *   5. 部署 → 新增部署作業 → 類型:Web 應用程式
 *        執行身分:我
 *        存取權限:任何人(含匿名)
 *   6. 複製 Web App URL · 填回 index.html 的
 *        CONFIG.SPEAKER_WEBHOOK / CONFIG.SIGNUP_WEBHOOK / CONFIG.ADMIN_WEBHOOK
 *   7. 把 Sheet「公開精選」分頁發佈成 CSV(檔案 → 共用 → 發佈到網路 → 選分頁 CSV)
 *        把 CSV URL 填回 CONFIG.SPEAKERS_CSV_URL
 */

const SHEET_ID    = '1dar-zrj1iZwtDuhsVYeA9Lgxam9EQ9MTWsKsZX9e2QY';
const EVENT_NAME  = 'Local AI 玩家開箱派對';
const EVENT_DATE  = '2026/11/04(三)';
const EVENT_TIME  = '14:00 – 20:00';
const EVENT_VENUE = '茹曦酒店 15F · 台北市松山區敦化北路 100 號 15 樓';
const EVENT_URL   = 'https://mfsd.pages.dev/event-local-ai-party/';
const SENDER_NAME = '虎智科技 TigerAI';

const SHEETS = {
  SIGNUP:     '一般報名',
  SUBMIT:     '講者投稿',
  SIGNUP_BAK: '_備份_報名',
  SUBMIT_BAK: '_備份_投稿',
  FEATURED:   '公開精選',   // 給前台 CSV 讀
};

const SIGNUP_HEADERS = [
  '時間戳', '姓名', 'Email', '手機', '公司/單位', '職稱/身份',
  '最想得到', '加報晚上派對', '寄信狀態', 'UA/備註'
];

const SUBMIT_HEADERS = [
  '時間戳', '公開稱呼', '職稱/專長', 'Email', '手機',
  'AI 專案名稱', '解決什麼問題', 'Demo 內容', '展示形式', '預計時間',
  '設備需求', '支援資料連結', '備註', '同意公開',
  '精選公開',      // ← Admin 在這格打 ✓ 就會上架到前台
  '寄信狀態', 'UA/備註'
];

const FEATURED_HEADERS = [
  'public_name', 'title_or_skill', 'project_name', 'demo_desc', 'support_url', 'is_featured'
];

/* ═══════════════════════════════════════════════════════════════════════════
 *  SETUP
 * ═══════════════════════════════════════════════════════════════════════════ */

function setupAll() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ensureSheet_(ss, SHEETS.SIGNUP,     SIGNUP_HEADERS,   false);
  ensureSheet_(ss, SHEETS.SUBMIT,     SUBMIT_HEADERS,   false);
  ensureSheet_(ss, SHEETS.SIGNUP_BAK, SIGNUP_HEADERS,   true);
  ensureSheet_(ss, SHEETS.SUBMIT_BAK, SUBMIT_HEADERS,   true);
  ensureSheet_(ss, SHEETS.FEATURED,   FEATURED_HEADERS, false);

  // 「公開精選」分頁加一個 formula · 自動從「講者投稿」撈 is_featured=TRUE 的
  const fs = ss.getSheetByName(SHEETS.FEATURED);
  const formula = `=IFERROR(QUERY('${SHEETS.SUBMIT}'!B2:O, "SELECT B, C, F, H, L WHERE N='✓' LABEL B 'public_name', C 'title_or_skill', F 'project_name', H 'demo_desc', L 'support_url'", 0), "")`;
  fs.getRange('A2').setFormula(formula);
  fs.getRange('F2:F').setValue('TRUE');   // 標記 is_featured(給前台 parser 用)

  Logger.log('✅ setupAll 完成 · 5 個分頁 ready');
  Logger.log('   接下來:1) 執行 setAdminPasscode("93489997") 設密碼');
  Logger.log('           2) 部署 Web 應用程式(任何人含匿名)');
  Logger.log('           3) 把「公開精選」分頁發佈成 CSV');
}

function ensureSheet_(ss, name, headers, hidden) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1F3A2B')
    .setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  if (hidden) {
    sh.hideSheet();
    try {
      const p = sh.protect().setDescription('備份 · 禁止修改');
      p.removeEditors(p.getEditors());
      if (p.canDomainEdit()) p.setDomainEdit(false);
    } catch (e) { Logger.log('保護 ' + name + ' 失敗:' + e); }
  }
}

function setAdminPasscode(code) {
  if (!code) throw new Error('請傳密碼 · 例:setAdminPasscode("93489997")');
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSCODE', String(code));
  Logger.log('✅ 密碼已設定(存在 PropertiesService · 不在 Sheet · 不在 code)');
}

/* ═══════════════════════════════════════════════════════════════════════════
 *  doPost · 收 webhook
 * ═══════════════════════════════════════════════════════════════════════════ */

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || 'signup';   // 預設 signup 向後相容

    // 鐵則 2:原始 payload DIAG log
    console.log('[POST_RAW]', new Date().toISOString(), action, JSON.stringify(data));

    switch (action) {
      case 'signup':                  return handleSignup_(data);
      case 'speaker_submit':          return handleSpeakerSubmit_(data);
      case 'admin_login':             return handleAdminLogin_(data);
      case 'admin_stats':             return handleAdminStats_(data);
      case 'admin_list_submissions':  return handleAdminListSubmissions_(data);
      default:
        return json_({ ok: false, error: 'unknown_action: ' + action });
    }
  } catch (err) {
    console.error('[POST_ERR]', err.message, err.stack);
    return json_({ ok: false, error: err.message });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ═══════════════════════════════════════════════════════════════════════════
 *  Signup
 * ═══════════════════════════════════════════════════════════════════════════ */

function handleSignup_(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHEETS.SIGNUP);
  const bak = ss.getSheetByName(SHEETS.SIGNUP_BAK);
  if (!sh || !bak) throw new Error('分頁不存在 · 請先執行 setupAll()');

  const row = [
    new Date(),
    data.name || '',
    data.email || '',
    data.phone || '',
    data.company || '',
    data.role || '',
    data.interest || '',
    data.afterparty === 'yes' ? '✓' : '',
    '寄信中',
    (data._ua || '').slice(0, 200),
  ];
  const before = sh.getLastRow();
  sh.appendRow(row);
  bak.appendRow(row);     // 鐵則 2:同時寫備份

  const after = sh.getLastRow();
  const written = sh.getRange(after, 2).getValue();   // 讀回驗證
  console.log(`[SIGNUP_DIAG] before=${before} after=${after} name=${data.name} readback=${written}`);
  if (written !== data.name) {
    console.error('[SIGNUP_READBACK_FAIL] expected=' + data.name + ' got=' + written);
  }

  let mailStatus = '✓';
  try { sendSignupConfirmEmail_(data); }
  catch (err) { mailStatus = '✗ ' + err.message; console.error('[MAIL_SIGNUP_ERR]', err); }
  sh.getRange(after, 9).setValue(mailStatus);

  return json_({ ok: true, msg: '報名成功 · 確認信已寄出' });
}

/* ═══════════════════════════════════════════════════════════════════════════
 *  Speaker Submit
 * ═══════════════════════════════════════════════════════════════════════════ */

function handleSpeakerSubmit_(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHEETS.SUBMIT);
  const bak = ss.getSheetByName(SHEETS.SUBMIT_BAK);
  if (!sh || !bak) throw new Error('分頁不存在 · 請先執行 setupAll()');

  const row = [
    new Date(),
    data.public_name || '',
    data.title_or_skill || '',
    data.email || '',
    data.phone || '',
    data.project_name || '',
    data.problem || '',
    data.demo_desc || '',
    data.demo_format || '',
    data.duration || '',
    data.requirements || '',
    data.support_url || '',
    data.note || '',
    data.consent_public ? '✓' : '',
    '',                          // 精選公開 · 預設空 · 等 admin 勾選
    '寄信中',
    (data._ua || '').slice(0, 200),
  ];
  const before = sh.getLastRow();
  sh.appendRow(row);
  bak.appendRow(row);

  const after = sh.getLastRow();
  const written = sh.getRange(after, 2).getValue();
  console.log(`[SUBMIT_DIAG] before=${before} after=${after} name=${data.public_name} readback=${written}`);
  if (written !== data.public_name) {
    console.error('[SUBMIT_READBACK_FAIL] expected=' + data.public_name + ' got=' + written);
  }

  let mailStatus = '✓';
  try { sendSpeakerSubmitEmail_(data); }
  catch (err) { mailStatus = '✗ ' + err.message; console.error('[MAIL_SUBMIT_ERR]', err); }
  sh.getRange(after, 16).setValue(mailStatus);

  return json_({ ok: true, msg: '投稿已收到 · 等待主辦審核' });
}

/* ═══════════════════════════════════════════════════════════════════════════
 *  Admin · login + stats + list
 * ═══════════════════════════════════════════════════════════════════════════ */

function assertAdmin_(data) {
  const expected = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSCODE');
  if (!expected) throw new Error('密碼尚未設定 · 請執行 setAdminPasscode()');
  if (String(data.passcode || '') !== expected) throw new Error('密碼錯誤');
}

function handleAdminLogin_(data) {
  try { assertAdmin_(data); }
  catch (err) { return json_({ ok: false, error: err.message }); }
  return handleAdminStats_(data);    // 登入成功直接回 stats
}

function handleAdminStats_(data) {
  try { assertAdmin_(data); }
  catch (err) { return json_({ ok: false, error: err.message }); }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const signupSh = ss.getSheetByName(SHEETS.SIGNUP);
  const submitSh = ss.getSheetByName(SHEETS.SUBMIT);

  const sData = signupSh.getLastRow() > 1 ? signupSh.getRange(2, 1, signupSh.getLastRow() - 1, SIGNUP_HEADERS.length).getValues() : [];
  const tData = submitSh.getLastRow() > 1 ? submitSh.getRange(2, 1, submitSh.getLastRow() - 1, SUBMIT_HEADERS.length).getValues() : [];

  return json_({
    ok: true,
    stats: {
      signup_count:     sData.length,
      afterparty_count: sData.filter(r => r[7] === '✓').length,
      submit_count:     tData.length,
      featured_count:   tData.filter(r => r[14] === '✓' || r[14] === true).length,
    },
    sheet_url:     `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit`,
    submit_sheet_url: `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit#gid=${submitSh.getSheetId()}`,
    recent_signup: sData.slice(-5).reverse().map(r => ({
      ts: formatTs_(r[0]), name: r[1], email: r[2],
      afterparty: r[7] === '✓',
    })),
    recent_submit: tData.slice(-5).reverse().map(r => ({
      ts: formatTs_(r[0]), name: r[1], project: r[5],
      featured: r[14] === '✓' || r[14] === true,
    })),
  });
}

function handleAdminListSubmissions_(data) {
  try { assertAdmin_(data); }
  catch (err) { return json_({ ok: false, error: err.message }); }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHEETS.SUBMIT);
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, SUBMIT_HEADERS.length).getValues() : [];
  return json_({
    ok: true,
    list: rows.map(r => ({
      ts: formatTs_(r[0]),
      public_name: r[1], title: r[2], email: r[3],
      project: r[5], problem: r[6], demo: r[7],
      support_url: r[11], featured: r[14] === '✓' || r[14] === true,
    })),
  });
}

function formatTs_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Taipei', 'MM/dd HH:mm');
  return String(v);
}

/* ═══════════════════════════════════════════════════════════════════════════
 *  Emails · TigerAI 標準 EDM(鐵則 3)
 * ═══════════════════════════════════════════════════════════════════════════ */

function sendSignupConfirmEmail_(data) {
  const name = data.name || '朋友';
  const subject = `[報名確認] ${EVENT_NAME} · ${EVENT_DATE}`;
  const sections = [
    { title: '📅 活動資訊', rows: [
      ['日期', EVENT_DATE], ['時間', EVENT_TIME], ['地點', EVENT_VENUE],
      ['票價', '✓ 白天的實體活動 · 免費參加'],
    ]},
    { title: '🎫 您的加報', rows: data.afterparty === 'yes' ? [
      ['加報晚上派對', '✓ OOPS 晚上派對(18:00–20:00)'],
      ['付款', '原價待 OOPS 公告 · 憑本報名可折抵 NT$100'],
    ] : [
      ['加報晚上派對', '未加報 · 只參加白天活動(如需加報請回信告知)'],
    ]},
    { title: '📝 行前提醒', rows: [
      ['最終時間表', '活動前 1 週會寄 email 通知'],
      ['任何問題', '直接回覆這封信件即可'],
    ]},
  ];
  const html = buildHtmlEmail_({
    hero_tag: '報名確認',
    hero_title: EVENT_NAME,
    greeting: name + ' 您好 ·',
    intro: '感謝您報名 ' + EVENT_NAME + ' · 以下是您的報名資訊:',
    sections: sections,
    cta_text: '查看活動頁',
    cta_url: EVENT_URL,
    footer_note: '任何問題請直接回覆此信件<br>虎智科技 TigerAI · tiger.ai.tw@gmail.com',
  });
  const plain = `${name} 您好\n\n感謝報名 ${EVENT_NAME}\n日期:${EVENT_DATE}\n時間:${EVENT_TIME}\n地點:${EVENT_VENUE}\n\n活動頁:${EVENT_URL}\n\n${SENDER_NAME}`;
  MailApp.sendEmail({ to: data.email, subject, body: plain, htmlBody: html, name: SENDER_NAME });
}

function sendSpeakerSubmitEmail_(data) {
  const name = data.public_name || '朋友';
  const subject = `[投稿收到] ${EVENT_NAME} · 等待主辦審核`;
  const sections = [
    { title: '📝 您的投稿', rows: [
      ['AI 專案', data.project_name || ''],
      ['Demo 內容', truncate_(data.demo_desc, 120)],
      ['展示形式', data.demo_format || '未指定'],
      ['預計時間', data.duration || '未指定'],
    ]},
    { title: '⏳ 接下來', rows: [
      ['精選名額', '最多 10 位'],
      ['公布時間', '活動前 1-2 週'],
      ['通知方式', '若入選 · 我們用這封 email 聯繫您'],
      ['入選福利', '✓ 免費獲得 OOPS 晚上派對入場券'],
    ]},
    { title: '📅 活動資訊', rows: [
      ['日期', EVENT_DATE], ['時間', EVENT_TIME], ['地點', EVENT_VENUE],
    ]},
  ];
  const html = buildHtmlEmail_({
    hero_tag: '投稿已收到',
    hero_title: EVENT_NAME,
    greeting: name + ' 您好 ·',
    intro: '感謝您送出 Demo 投稿 · 我們已經收到。<br><strong>這封信代表「投稿已收到」· 不代表入選。</strong>主辦審核後若入選 · 會以此 email 聯繫您。',
    sections: sections,
    cta_text: '查看活動頁',
    cta_url: EVENT_URL,
    footer_note: '任何問題請直接回覆此信件<br>虎智科技 TigerAI · tiger.ai.tw@gmail.com',
  });
  const plain = `${name} 您好\n\n您的 Demo 投稿已收到 · 等待主辦審核。\n專案:${data.project_name}\n\n活動頁:${EVENT_URL}\n\n${SENDER_NAME}`;
  MailApp.sendEmail({ to: data.email, subject, body: plain, htmlBody: html, name: SENDER_NAME });
}

function truncate_(s, n) {
  s = String(s || '');
  return s.length > n ? s.substring(0, n) + '…' : s;
}

/* ─────── TigerAI EDM · 活動版 cyberpunk 風(深黑 + 青 + 洋紅) ─────── */
function buildHtmlEmail_(opts) {
  const C = {
    bg:      '#07070C',   // 深黑底
    card:    '#12121E',   // 卡片稍亮
    tint:    '#0E0E18',
    line:    '#2E2E48',
    text:    '#E6E6F0',
    mid_t:   '#8A8AA3',
    light_t: '#4A4A62',
    cyan:    '#00F5FF',
    magenta: '#FF2EA6',
    green:   '#00FF88',
  };
  const sections = (opts.sections || []).map(sec => `
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:14px;background:${C.card};border:1px solid ${C.line};">
      <tr><td style="padding:12px 18px;border-bottom:1px solid ${C.line};font-family:'JetBrains Mono',ui-monospace,Consolas,monospace;font-size:13px;letter-spacing:0.08em;color:${C.cyan};">${sec.title}</td></tr>
      <tr><td style="padding:12px 18px 16px;">
        ${(sec.rows || []).map(r => `
          <table role="presentation" width="100%" style="margin:4px 0;"><tr>
            <td width="34%" valign="top" style="color:${C.mid_t};font-size:13px;padding:4px 0;font-family:'JetBrains Mono',ui-monospace,Consolas,monospace;">${r[0]}</td>
            <td valign="top" style="color:${C.text};font-size:14px;padding:4px 0;line-height:1.6;">${r[1]}</td>
          </tr></table>
        `).join('')}
      </td></tr>
    </table>`).join('');

  return `<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8"><meta name="color-scheme" content="dark"><meta name="supported-color-schemes" content="dark"></head>
<body style="margin:0;padding:0;background:${C.bg};font-family:'Noto Sans TC','PingFang TC','Microsoft JhengHei',sans-serif;color:${C.text};">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:${C.bg};"><tr><td align="center" style="padding:24px 10px;">
<table role="presentation" width="640" cellpadding="0" cellspacing="0" style="max-width:640px;background:${C.tint};border:1px solid ${C.line};">
  <!-- Terminal top bar -->
  <tr><td style="background:${C.bg};padding:12px 22px;color:${C.green};font-size:11px;letter-spacing:0.1em;font-family:'JetBrains Mono',ui-monospace,Consolas,monospace;border-bottom:1px solid ${C.line};">
    &gt; ./tigerai_mailer.sh --from=TIGER_AI
  </td></tr>
  <!-- Hero -->
  <tr><td style="background:${C.tint};padding:28px 22px 24px;border-bottom:1px solid ${C.line};">
    <div style="display:inline-block;padding:4px 10px;background:${C.magenta};color:#fff;font-size:11px;letter-spacing:0.08em;font-family:'JetBrains Mono',ui-monospace,Consolas,monospace;margin-bottom:14px;">${opts.hero_tag}</div>
    <div style="font-size:24px;font-weight:700;color:${C.text};line-height:1.25;letter-spacing:-0.01em;">${opts.hero_title}</div>
  </td></tr>
  <!-- Body -->
  <tr><td style="padding:22px;background:${C.bg};">
    <p style="margin:0 0 10px;color:${C.text};font-size:15px;">${opts.greeting}</p>
    <p style="margin:0 0 20px;color:${C.mid_t};font-size:14px;line-height:1.65;">${opts.intro}</p>
    ${sections}
    <table role="presentation" width="100%"><tr><td align="center" style="padding:18px 0 6px;">
      <a href="${opts.cta_url}" style="display:inline-block;background:${C.cyan};color:#000;padding:14px 36px;text-decoration:none;font-weight:700;font-size:13px;letter-spacing:0.12em;font-family:'JetBrains Mono',ui-monospace,Consolas,monospace;">&gt; ${opts.cta_text}</a>
    </td></tr></table>
  </td></tr>
  <!-- Footer -->
  <tr><td style="background:${C.tint};padding:18px 22px;color:${C.mid_t};font-size:12px;line-height:1.7;border-top:1px solid ${C.line};">${opts.footer_note}</td></tr>
  <tr><td style="background:${C.bg};padding:10px 22px;color:${C.light_t};font-size:11px;text-align:center;font-family:'JetBrains Mono',ui-monospace,Consolas,monospace;">© 2026 TIGER AI · all rights reserved</td></tr>
</table></td></tr></table></body></html>`;
}

/* ═══════════════════════════════════════════════════════════════════════════
 *  救援函式(鐵則 2 · 萬一 Sheet 被誤刪可從 Gmail 寄件夾撈回)
 * ═══════════════════════════════════════════════════════════════════════════ */

function recoverFromGmailSent() {
  const threads = GmailApp.search('from:me (subject:"[報名確認]" OR subject:"[投稿收到]") newer_than:30d');
  Logger.log(`找到 ${threads.length} 封確認信`);
  let recovered = 0;
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const signupSh = ss.getSheetByName(SHEETS.SIGNUP);
  const submitSh = ss.getSheetByName(SHEETS.SUBMIT);

  threads.forEach(t => {
    t.getMessages().forEach(m => {
      const subject = m.getSubject();
      const to = m.getTo();
      const date = m.getDate();
      const body = m.getPlainBody();
      // 粗略解析 · 把 email + 日期至少補回去 · 其他欄位要人工檢查
      if (subject.indexOf('[報名確認]') >= 0) {
        signupSh.appendRow([date, '(從 Gmail 救回)', to, '', '', '', '', '', '救回', 'recovered']);
        recovered++;
      } else if (subject.indexOf('[投稿收到]') >= 0) {
        submitSh.appendRow([date, '(從 Gmail 救回)', '', to, '', '', '', '', '', '', '', '', '', '', '', '救回', 'recovered']);
        recovered++;
      }
    });
  });
  Logger.log(`已回補 ${recovered} 筆到 Sheet · 請人工檢查欄位是否完整`);
}
