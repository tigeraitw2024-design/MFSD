# 5_AI 升級計劃 — 素材資料

從 `2_在職菁英課程` 複製出來的專欄,課程內容完全相同,只換兩件事:

1. **主題名稱**:「AI 賦能 · 生產力躍升與製造業應用趨勢」→「從自動化到智動化：手把手帶您做出智慧製造AI 升級計畫」
2. **設計風格**:MIID 紅黑印刷風 → ZCOOL 站酷獎海報的橘黑蜂巢風(見 `1. 素材資料/設計風格.md`)

---

## 資料夾

| 檔案 | 說明 |
|---|---|
| `1. 素材資料/課程.md` | 課程全部資訊(跟在職菁英課程一樣,只改官方全名) |
| `1. 素材資料/設計風格.md` | 橘黑蜂巢風格規範 |
| `1. 素材資料/素材/` | 沿用在職菁英課程的 logo / OG / 課綱 docx |
| `2. 產出/index.html` | 產出的網站(同步一份到根目錄 `course-upgrade/` 供 Cloudflare 部署) |
| `apps-script.js` | 報名確認信(只改課程名稱) |

## 上線前要人工做的事

- `CONFIG.SHEET_WEBHOOK` 與 `COHORTS_CSV_URL` 已換成 AI 升級計劃專用 Sheet(2026-09-14)。梯次填在該 Sheet 第 1 個分頁(A 縣市 / B 日期 / C 時間 / D 地點),報名資料進「AI升級計劃報名表單」分頁。
- `assets/og-card.png` 是舊主題的分享卡,要重做一張新的(1200×630)換掉。
- 網址預設 `https://mfsd.pages.dev/course-upgrade/`,push 後 Cloudflare 會自動長出來。

## 「實戰現場」照片輪播怎麼更新

照片來源是 Google Drive 資料夾(設成「知道連結的任何人都可檢視」):
https://drive.google.com/drive/folders/1GhdzqTVQMvumURYvx6RYwHkhk8xroqw6

工具放在 `專案/sync_photos.py`(五個課程網站共用,會一次更新五站)。Drive 裡新增或刪除照片後,在專案根目錄執行一行:

```
python "專案/sync_photos.py"
```

它會自動:抓沒抓過的照片(Drive 刪掉的也會跟著移除)→ 跳過影片 → 依 EXIF 轉正、置中裁成六角比例、壓縮 → 輸出到兩個 `assets/photos/` → 更新 `manifest.json`。Drive 裡放什麼就輪播什麼,不做過濾,要拿掉的照片直接在 Drive 刪掉再重跑。
網頁載入時讀 `manifest.json` 自動長出輪播,不用改 HTML。跑完 git commit + push 就上線。

- 第一次要先 `pip install gdown pillow`;有 iPhone HEIC 再加 `pip install pillow-heif`。
- 原檔快取在 `專案/.photo_cache/`,已加進 .gitignore。
