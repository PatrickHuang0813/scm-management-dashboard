# SCM Control Room

一個不需要 Node.js 或資料庫的 SCM 庫存 dashboard。資料會存在該瀏覽器的 Local Storage，可匯入或匯出 Excel `.xlsx` 檔。

## Excel 欄位

第一列必須包含下列固定欄位：

`料號`、`品名`、`分類`、`供應商`、`倉庫`、`現有庫存`、`安全庫存`、`單價`、`狀態`、`更新日期`

可用頁面上的「下載 Excel 範本」取得正確格式。匯入後，資料會覆蓋目前畫面中的資料集；匯出會輸出目前篩選後的資料。

## 在本機使用

直接在瀏覽器開啟 `index.html` 即可。Excel 功能透過 SheetJS CDN 載入，因此首次使用需要網際網路連線。

## 部署

此專案是純靜態網站，將整個資料夾上傳至任一靜態網站服務即可：GitHub Pages、Azure Static Web Apps、Cloudflare Pages、Netlify 或公司內網 IIS。

GitHub Pages 最簡流程：建立 GitHub repository，推送此資料夾內容，在 repository 的 **Settings > Pages** 選取主要分支與 `/ (root)` 後發布。無伺服器端資料庫時，資料會保留在每位使用者的瀏覽器；多人共用資料需要第二版串接 API 與資料庫。
