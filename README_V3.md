# 生化教育訓練考核系統 V3

本版本以 Flask + PostgreSQL/Supabase 為核心，支援：

- GitHub → Render Docker 部署
- Supabase PostgreSQL 成績集中保存
- 各組別獨立後台與教材/Word 模板
- 後台新增選擇題、問答題與題目圖片
- 每個頁籤設定抽題數，前台隨機抽題
- 選擇題自動評分、問答題人工評分欄位
- 各組 Word 模板上傳與匯出架構

## Render 環境變數

- `DATABASE_URL`: Supabase PostgreSQL connection string
- `ADMIN_KEY`: 後台管理密碼
- `MATERIAL_STORAGE`: `/var/data/materials`

## 部署

1. 將此資料夾內容上傳 GitHub（不要再包一層資料夾）。
2. Render 建立 Web Service，Runtime 選 Docker。
3. 在 Render Environment Variables 設定上述三個變數。
4. 若要永久保存上傳教材與模板，請在 Render 掛載 Persistent Disk 到 `/var/data`。

## 重要

Word 模板必須是 `.docx`。若使用 Render Free 且未掛載 Persistent Disk，上傳的教材、題目圖片與模板可能在重新部署後遺失；成績仍由 Supabase 保存。
