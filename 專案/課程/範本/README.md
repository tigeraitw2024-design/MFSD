# [這個專案叫什麼] — 素材資料

這個資料夾裝的是這個專案的所有原始素材。

---

## 你只要填 2 份

| 檔案 | 說明 |
|---|---|
| `1. 素材資料/課程.md` | 課程全部資訊(名字、對象、3 個成果、講師、梯次、聯絡) |
| `1. 素材資料/設計風格.md` | 主色、質感關鍵字、想給訪客的感覺 |

其他細節(統編驗證規則、確認信文案、章節微調 …)**先不用管**,建站過程中 Claude 會問你,或你隨時說要調整。

## 素材放這裡(選填)

`1. 素材資料/素材/` 放:
- `logo.png` · `hero.png` · `og.png` · `favicon.png`

沒放的話 Claude 會用 SVG / 文字型預設。

## 課程現場照片(輪播 pstrip)

**不放在這裡** · 走 Google Drive 公開資料夾 · **全部課程共用一份 manifest**。

- 清單檔:`_shared/media_manifest.json`
- 產生 script:`專案/sync_media_manifest.py`
- 設定在 script 開頭的 `DRIVE_FOLDER_URLS = [...]`(可以放多個 Drive 連結)
- 新增 Drive 連結:編輯那個 list · 跑 `python 專案/sync_media_manifest.py` · commit + push
- **自動排程**:GitHub Actions `.github/workflows/sync-media-manifest.yml` 每週一台北 03:00 自動跑
- 想要立即更新:到 GitHub Actions 頁手動 Run workflow · 或跟 Claude 說「幫我重抓 Drive 照片」

7 個 course 的 `index.html` 都已經載 `/_shared/pstrip.js` · 吃同一份 manifest · 一改全部同步。
新建 course 複製現有的 pstrip 區塊即可(空容器 `<div class="pstrip" data-pstrip-auto></div>` + `<script src="/_shared/pstrip.js" defer></script>`)。

## 填完後

對 Claude 說「跑 [這個專案名]」就會開工。

## 產出會在哪裡

`2. 產出/` 資料夾底下(現在是空的,建完會出現 `index.html` + `assets/`)

---

# ⚠️ Apps Script 兩條鐵則(Claude 建課程 Apps Script 必遵守)

## 鐵則 1 · 範本不等於寫死 · 梯次資料全部放 Sheet

User 給的 docx / 範例 = **第一筆 seed data** · 不是唯一。課程可能有多個梯次(嘉義、新竹、台中…)· code **不能把任何場次專屬資訊寫死**。

**梯次目錄 Sheet 應有欄位**(以 `專案/課程/8_AI工作流程/apps-script.js` 為標準):
```
開課縣市 | 課程日期 | 課程時間 | 課程地點 | 單場講師 | 單場主題 | 交通 | 停車 | 名額
```

**寫 code 的 pattern:**
- 寫 `getCohortMeta(cohortString)` helper · 從 Sheet 動態查
- `logCourseSignup` 跟 `sendCourseConfirmEmail` 都呼叫這個 · 不要用常數
- 名額查不到 → fallback `COHORT_CAPACITY_DEFAULT = 30`
- 交通/停車查不到 → 信件該行**不顯示**(不要 fallback 到別場內容)

**反例(禁止)**:信件內把「嘉義火車站步行 12 分鐘」「大樓 3、4 樓停車場」寫死 → 新竹場寄信就會錯亂。

## 鐵則 2 · 報名流程必須有 3 層備份 + DIAG log

處理報名 / 寫 Sheet 的函式 **一律** 做這 3 件事(**不是 debug 才加 · 是 default**):

**第 1 層 · 原始 payload log**
```js
console.log('[SIGNUP_RAW]', new Date().toISOString(), JSON.stringify(data));
```
Google Cloud Logging ≥ 7 天 · Sheet 全毀也救得回。

**第 2 層 · 隱藏備份分頁**
在同一份 Spreadsheet 建 `_備份_<課程名>報名表單` 分頁:
- `hideSheet()` 隱藏
- `protect()` + `removeEditors()` 禁止編輯/刪除
- 寫主表 appendRow 時同一列資料也 appendRow 到備份分頁

**第 3 層 · DIAG log**
- Spreadsheet 名稱 / ID
- 寫入前後 `getLastRow()`
- 讀回驗證(讀寫入列的姓名 vs 預期)

**救援函式**
另外要寫 `recoverFromGmailSent()` · user 執行一次就能從 Gmail 寄件夾解析信件回補到 Sheet(以防萬一 Sheet 被誤刪)。

## 兩條鐵則都未遵守 → 已知會失火

- 2026-10-14 · course-smart 寄 50 封信但 Sheet 只有 30 列(事後查發現版本記錄顯示資料被誤刪 → **沒備份就救不回來**)
- 2026-10-14 · 8_AI工作流程 嘉義場 code 寫死地點 → 新竹場無法複用(用戶抓到:**要加梯次欄位讓 Sheet 自己填**)

## 鐵則 3 · 所有 email 用 TigerAI 標準 EDM 版型

虎智科技寄給客戶的 email(報名確認、候補、取消、行前通知等)**一律用 HTML EDM 版型** · 不能回去貼純文字。

**標準結構(從上到下):**
1. Brand bar(深綠底 + 虎智 TigerAI logo)
2. Hero banner(淡綠底 + eyebrow pill + 課程名)
3. 候補警示 card(if applicable · 橘底)
4. 問候語
5. 分區卡 Card:報名資訊 / 梯次 / 課程 / 行前提醒(各自白底 + emoji icon section-head)
6. CTA 大按鈕(深綠 pill · 例:取消報名)
7. Footer(深綠底 · 聯絡窗口)
8. 版權

**色票**(與網站同一套 · 薄荷森綠):
```
#1F3A2B deep  · #3E6B52 mid  · #4D9B6A strong
#EEF4EC bg    · #FFFFFF card · #F6FAF4 tint  · #E0EBDC line
#2C3A33 text main · #6B7A72 text mid · #9EACA3 text light
```

**技術規範:**
- `MailApp.sendEmail({ to, subject, body, htmlBody, name })` 同時傳純文字跟 HTML
- Table-based layout(email client 相容)
- inline styles only(不要 `<style>` tag)
- 字體:`'Noto Sans TC','PingFang TC','Microsoft JhengHei',sans-serif`
- 外層 max-width 640px 置中

細節見 memory [[tigerai-edm-standard-style]]。

## 參考實作

`專案/課程/8_AI工作流程/apps-script.js` — 三條鐵則都已落實的範本 · 包含 `buildHtmlEmail(opts)` 標準 EDM 函式
