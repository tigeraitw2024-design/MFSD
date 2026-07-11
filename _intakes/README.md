# _intakes — 新網站放檔區

未來要幫其他中心、協會、公司做課程網站時,把資料丟這裡,然後叫 Claude 就好。

## 怎麼用

1. **複製 `_template/` 為 `_intakes/[你要的英文 slug]/`**
   例如 `iot-association`、`ax-academy`、`nptu-workshop`

2. **打開新資料夾裡的 `.md` 檔,把內容填一填**
   - 必填:`brand.md` / `course.md` / `scenarios.md` / `contact.md`
   - 選填:其他有時間再補,沒填會走預設

3. **叫 Claude 開工**,任一句都行:
   - `跑套版 [slug]`
   - `幫 [slug] 建站`
   - `用 _intakes/[slug] 生一個網站`

Claude 會先讀完 intake、跟你確認缺什麼,同意後才動筆產出 `sites/[slug]/index.html`。

## 完整 SOP

https://claude.ai/code/artifact/61694393-7984-4819-aab9-1f8b1f298ad8 → 第 08 章

## 目前的實作範例

`course/index.html`(MFSD 課程網站)就是這個套版的**第一個**成品,可以參考它長什麼樣。
