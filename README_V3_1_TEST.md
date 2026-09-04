# V3.1 實測修正版

## 本機驗證

```bash
python -m py_compile app.py
pip install -r requirements.txt
python app.py
```

健康檢查：`GET /health`

## Render 環境變數

- `DATABASE_URL`: Supabase PostgreSQL connection string
- `ADMIN_KEY`: 後台管理密碼
- `MATERIAL_STORAGE`: `/var/data/materials`

## 已修正

- 前台動態考題改由 `/api/quiz-questions/random` 抽題，不再每次載入全部題目。
- 前台保留 `questionType`、`imageUrl`，可依題型顯示圖片與問答欄位。
- 新增前台每次抽題數控制，預設 10 題。
- 保留各組 Word 模板、圖片上傳、Supabase 成績資料表。

## 注意

Render 若要永久保存教材、圖片與 Word 模板，請使用 Persistent Disk 掛載 `/var/data`；Supabase 只負責集中保存成績與題庫資料。
