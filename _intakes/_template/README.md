# [Client Slug] — Intake

這個資料夾裝的是「[客戶或組織名]」的網站資訊。
填完必要的 `.md`,對 Claude 說「跑套版 [slug]」就會開工。

## 檔案清單

| 檔案 | 必填? | 一句話說明 |
|---|---|---|
| `brand.md` | ✅ | 主色、字體氣質、視覺風格 |
| `course.md` | ✅ | 課程主體資訊 |
| `scenarios.md` | ✅ | 3 條自動化流程(或 3 個關鍵成果) |
| `contact.md` | ✅ | Footer 聯絡窗口 |
| `instructor.md` | ⭕ | 講師,不填會走「無講師版」 |
| `cohorts.csv` | ⭕ | 梯次表,或改貼 Google Sheet Publish CSV URL |
| `email-confirm.md` | ⭕ | 確認信自訂內容 |
| `eligibility.md` | ⭕ | g0v 統編驗證規則 |
| `assets/` | ⭕ | Logo、Hero、OG 分享卡 |

## 填完後

對 Claude 說:「跑套版 [slug]」
