# 2026 生化教育訓練考核系統

Flask + HTML/JavaScript 的線上考核與教育訓練投影片系統。

## 專案內容

- 一致性與法定傳染病通報
- c503 一般作業流程與異常訊號故障排除
- Cobas b 211 異常訊號故障排除與 QC 設定
- COVER C1 / Sebia 人員考區
- 教育訓練 PPT/PPTX 線上瀏覽
- PPT/PPTX 上傳後自動轉 PDF/PNG
- 各頁籤考試、計分與合格判定
- 附件 1 Word 匯出
- CSV 成績匯出
- 本機瀏覽器暫存考核紀錄

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

目前前端考核紀錄仍使用瀏覽器 `localStorage`，因此不是多人共用的中央資料庫。若要正式讓整個科室多人使用，建議下一階段加入 SQLite/PostgreSQL + Flask API，將成績集中儲存。

### PPT 上傳權限

目前既有前端提供 PPT 上傳/刪除功能。若部署到公開網際網路，建議下一階段加入管理者登入或 API 權限控制，避免任何訪客任意上傳/刪除教材。

## Word 匯出成績

Word 匯出資料現在同時提供 `score`、`evaluationScore`、`passingScore`、`correctCount`、`wrongCount`、`totalQuestions`、`status`、`result` 等欄位，供「附件1.docx」範本使用。這可避免範本使用「評核分數」欄位時因前端沒有傳值而顯示 0。
