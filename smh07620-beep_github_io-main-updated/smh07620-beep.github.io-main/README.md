# 2026 生化教育訓練考核系統

Flask + HTML/JavaScript 的線上考核與教育訓練投影片系統，涵蓋六大組別：
1 生化組、2 鏡檢組、3 血清組、4 血庫組、5 細菌組、6 血液組。

## 專案內容

- 頁面最上層可切換六大組別，各組別的投影片與考題彼此獨立
- **1 生化組**：沿用原有固定考題頁籤與題庫（不受其他組別異動影響）
  - 一致性與法定傳染病通報
  - c503 一般作業流程與異常訊號故障排除
  - Cobas b 211 異常訊號故障排除與 QC 設定
  - COVER C1 / Sebia 人員考區
- **其餘 5 組別（鏡檢/血清/血庫/細菌/血液）**：投影片上傳與考題頁籤／題目皆由管理者後台新增，可依需要自由建立
- 教育訓練 PPT/PPTX 線上瀏覽
- PPT/PPTX 上傳後自動轉 PDF/PNG
- 各頁籤考試、計分與合格判定
- 附件 1 Word 匯出
- CSV 成績匯出（含組別欄位）
- 考核成績集中儲存於 PostgreSQL（本機無 DATABASE_URL 時自動使用 SQLite）
- 成績後台以 ADMIN_KEY 保護

## GitHub

本專案可直接建立 GitHub Repository 後上傳。建議 Repository 根目錄直接放：

```text
app.py
requirements.txt
Dockerfile
render.yaml
Procfile
.gitignore
.env.example
static/
data/
```

不要上傳 `.env`、`uploads/`、`tmp_convert/` 或 Python 快取。

## 本機執行

Python 3.9+：

```bash
pip install -r requirements.txt
python app.py
```

開啟 `http://127.0.0.1:5000/`。

### LibreOffice

PPT/PPTX 自動轉圖功能需要 LibreOffice。

Windows 可設定：

```powershell
$env:SOFFICE_PATH="C:\\Program Files\\LibreOffice\\program\\soffice.exe"
python app.py
```

Linux 若 `soffice` 已在 PATH，可直接使用預設值。

## Docker / 雲端部署

專案已附 `Dockerfile`，會自動安裝 LibreOffice 與中文字型，並以 Gunicorn 啟動 Flask。

Render 等 Docker 部署平台可直接使用 `Dockerfile`。健康檢查網址：

```text
/health
```

## 重要說明

### GitHub Pages 不適用於本專案

GitHub Pages 只能提供靜態網站；本系統需要 Flask API、PPT 轉檔與檔案儲存，因此應使用 GitHub 作為程式碼儲存庫，再把專案部署到可執行 Python/Docker 的主機。

### 成績儲存

考核提交後會透過 Flask `/api/records` 寫入中央資料庫。Render 部署會自動建立 PostgreSQL，並把 `DATABASE_URL` 注入 Web Service；本機若沒有設定 `DATABASE_URL`，則使用 `data/exam_records.db` 的 SQLite。

管理後台的查詢、CSV 匯出與清空紀錄需要 `X-Admin-Key`，前端會在第一次開啟成績後台時要求輸入 `ADMIN_KEY`。請勿把金鑰寫進 GitHub。

### 管理者金鑰 (ADMIN_KEY)

目前預設金鑰為 `9544106`（寫在 `app.py` 內的環境變數預設值）。正式站建議改由 **Render Dashboard → Environment → ADMIN_KEY** 覆蓋這個預設值，避免使用寫在公開 GitHub Repository 裡的固定金鑰。本機開發可在 `.env` 檔設定同名變數覆蓋。

### Render 注意事項

- PostgreSQL 用於集中保存考核成績。
- Render Web Service 的本地檔案系統不應視為永久儲存；目前使用者上傳的 PPT/PPTX 仍屬伺服器本地檔案，若需要永久共享教材，建議後續改用物件儲存（例如 S3 相容服務）。
- PPT/PPTX 轉圖仍由 Docker 內的 LibreOffice + PyMuPDF 處理。
- `ADMIN_KEY` 建議在 Render Dashboard / Blueprint 建立時另外設定，不要沿用程式碼內的預設值。

## Word 匯出成績

Word 匯出資料現在同時提供 `score`、`evaluationScore`、`passingScore`、`correctCount`、`wrongCount`、`totalQuestions`、`status`、`result` 等欄位，供「附件1.docx」範本使用。這可避免範本使用「評核分數」欄位時因前端沒有傳值而顯示 0。


## 管理者教材後台

登入網頁右上角「管理後台」後，可新增 PPT/PPTX、設定教材名稱/說明/考題頁籤、啟用/停用與刪除。教材會在伺服器端轉成逐頁 PNG，並保存於 Render Persistent Disk 的 `/var/data/materials`。

Render 預設檔案系統是暫存的；Persistent Disk 可保留服務在重新部署/重啟後的檔案，但需要付費 Web Service。citeturn0search0turn0search1
