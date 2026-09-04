# -*- coding: utf-8 -*-
"""
羅東聖母醫院 檢驗科生化組 - 2026生化教育訓練考核系統
後端服務 (Flask)

功能：
1. 提供靜態網頁 (static/index.html) 與已轉檔的簡報圖片。
2. 提供 /api/slides 給前端讀取目前所有簡報清單（內建 + 使用者上傳）。
3. 提供 /api/slides/upload：接收使用者上傳的 PPTX，
   自動呼叫 LibreOffice 轉成 PDF，再用 PyMuPDF 逐頁轉成 PNG 圖片，
   使前端可直接翻頁瀏覽，不需下載檔案。
4. 提供 /api/slides/<id> (DELETE)：刪除使用者上傳的簡報（內建簡報無法刪除）。

部署需求：
- Python 3.9+
- pip install -r requirements.txt
- 主機需安裝 LibreOffice（可執行 soffice 指令）。
  Windows 預設路徑通常在 "C:\\Program Files\\LibreOffice\\program\\soffice.exe"，
  若不在 PATH 中，請設定環境變數 SOFFICE_PATH 指向該執行檔完整路徑。

啟動方式：
    python app.py
    預設監聽 0.0.0.0:5000，可用瀏覽器開啟 http://<伺服器IP>:5000/
    正式環境建議搭配 gunicorn / waitress 等 WSGI 伺服器，並在前面加 Nginx/IIS 反向代理。
"""

import os
import json
import uuid
import shutil
import subprocess
import threading
import datetime
import sqlite3
from pathlib import Path

from flask import Flask, request, jsonify, send_from_directory, send_file, abort

try:
    import fitz  # PyMuPDF，用於 PDF -> PNG，跨平台不需額外安裝 poppler
except ImportError:
    fitz = None


# ---------------------------------------------------------------------------
# 路徑與基本設定
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
SLIDES_DIR = STATIC_DIR / "slides"
MATERIAL_STORAGE = Path(os.environ.get("MATERIAL_STORAGE", "/var/data/materials"))
# Render Persistent Disk 建議掛載 /var/data；本機若無此路徑權限則改用專案 uploads。
try:
    MATERIAL_STORAGE.mkdir(parents=True, exist_ok=True)
except OSError:
    MATERIAL_STORAGE = BASE_DIR / "uploads"
UPLOAD_DIR = MATERIAL_STORAGE / "ppt"
QUESTION_IMAGES_DIR = MATERIAL_STORAGE / "question_images"
UPLOADED_SLIDES_DIR = MATERIAL_STORAGE / "slides"
DOC_TEMPLATES_DIR = MATERIAL_STORAGE / "doc_templates"
DATA_DIR = BASE_DIR / "data"
META_FILE = DATA_DIR / "slides_meta.json"
TMP_DIR = BASE_DIR / "tmp_convert"

for d in (SLIDES_DIR, UPLOAD_DIR, UPLOADED_SLIDES_DIR, DOC_TEMPLATES_DIR, QUESTION_IMAGES_DIR, DATA_DIR, TMP_DIR):
    d.mkdir(parents=True, exist_ok=True)

# 如果 soffice 不在系統 PATH 中，可用環境變數指定完整路徑，例如：
#   set SOFFICE_PATH=C:\Program Files\LibreOffice\program\soffice.exe   (Windows)
#   export SOFFICE_PATH=/usr/bin/soffice                                 (Linux)
SOFFICE_BIN = os.environ.get("SOFFICE_PATH", "soffice")

ALLOWED_EXT = {".pptx", ".ppt"}
MAX_UPLOAD_MB = 80
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
# 預設管理者金鑰。正式站建議改由 Render Dashboard 環境變數 ADMIN_KEY 覆蓋，
# 不要把正式使用的金鑰提交到公開的 GitHub Repository。
ADMIN_KEY = os.environ.get("ADMIN_KEY", "9544106")
SQLITE_DB = DATA_DIR / "exam_records.db"

# ---------------------------------------------------------------------------
# 六大組別（頁面最上層分組）
# ---------------------------------------------------------------------------
GROUPS = {
    "grpBio": "1 生化組",
    "grpMicro": "2 鏡檢組",
    "grpSero": "3 血清組",
    "grpBB": "4 血庫組",
    "grpBact": "5 細菌組",
    "grpHema": "6 血液組",
}
DEFAULT_GROUP = "grpBio"

# 生化組沿用原本既有的固定考題頁籤／教材分類代碼。
CATEGORY_LABELS = {
    "subA1": "1-1 一致性與法定傳染病通報",
    "subA2": "1-2 c503一般作業流程與異常訊號故障排除",
    "subA3": "1-3 Cobas b 211異常訊號故障排除與QC設定",
    "zoneB": "2 COVER C1人員考區（Sebia）",
    "": "未分類 / 一般補充教材",
}

# LibreOffice headless 同一時間只能處理一份轉檔工作，避免多人同時上傳互相衝突
conversion_lock = threading.Lock()

app = Flask(__name__, static_folder=str(STATIC_DIR), static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024


# ---------------------------------------------------------------------------
# 考核成績資料庫
# ---------------------------------------------------------------------------
def _db_conn():
    """有 DATABASE_URL 時強制使用 PostgreSQL；本機未設定時才使用 SQLite。

    這樣 Render 若 PostgreSQL 暫時連線失敗，會直接顯示錯誤，避免悄悄寫進
    Render 的暫存 SQLite，造成「看起來有存檔、實際成績沒有進中央資料庫」的問題。
    """
    if DATABASE_URL:
        import psycopg
        from psycopg.rows import dict_row
        conn = psycopg.connect(DATABASE_URL, row_factory=dict_row, connect_timeout=10)
        conn.autocommit = True
        return conn, "postgres"
    conn = sqlite3.connect(str(SQLITE_DB), timeout=30)
    conn.row_factory = sqlite3.Row
    # 與 Postgres 分支的 conn.autocommit = True 對齊：SQLite 預設每個
    # INSERT/UPDATE/DELETE 需要顯式 commit() 才會真正寫入，若忘記 commit
    # 又直接 close()，寫入會被靜默 rollback。isolation_level=None 讓每個
    # SQL 陳述式即時自動提交，行為與 Postgres 分支一致。
    conn.isolation_level = None
    return conn, "sqlite"


def init_exam_db():
    conn, kind = _db_conn()
    try:
        if kind == "postgres":
            conn.execute("""
                CREATE TABLE IF NOT EXISTS exam_records (
                    id TEXT PRIMARY KEY,
                    created_at TEXT NOT NULL,
                    name TEXT NOT NULL,
                    emp_id TEXT NOT NULL,
                    role TEXT NOT NULL,
                    evaluator_name TEXT,
                    evaluator_title TEXT,
                    quiz_title TEXT NOT NULL,
                    score INTEGER NOT NULL,
                    status TEXT NOT NULL,
                    correct_count INTEGER NOT NULL DEFAULT 0,
                    wrong_count INTEGER NOT NULL DEFAULT 0,
                    answers_detail JSONB NOT NULL,
                    group_key TEXT NOT NULL DEFAULT 'grpBio'
                )
            """)
            try:
                conn.execute("ALTER TABLE exam_records ADD COLUMN IF NOT EXISTS group_key TEXT NOT NULL DEFAULT 'grpBio'")
            except Exception:
                pass
        else:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS exam_records (
                    id TEXT PRIMARY KEY,
                    created_at TEXT NOT NULL,
                    name TEXT NOT NULL,
                    emp_id TEXT NOT NULL,
                    role TEXT NOT NULL,
                    evaluator_name TEXT,
                    evaluator_title TEXT,
                    quiz_title TEXT NOT NULL,
                    score INTEGER NOT NULL,
                    status TEXT NOT NULL,
                    correct_count INTEGER NOT NULL DEFAULT 0,
                    wrong_count INTEGER NOT NULL DEFAULT 0,
                    answers_detail TEXT NOT NULL,
                    group_key TEXT NOT NULL DEFAULT 'grpBio'
                )
            """)
            existing_cols = {row[1] for row in conn.execute("PRAGMA table_info(exam_records)").fetchall()}
            if "group_key" not in existing_cols:
                conn.execute("ALTER TABLE exam_records ADD COLUMN group_key TEXT NOT NULL DEFAULT 'grpBio'")
    finally:
        conn.close()


def _record_to_dict(row):
    r = dict(row)
    raw = r.pop("answers_detail", "[]")
    if isinstance(raw, str):
        try:
            raw = json.loads(raw)
        except json.JSONDecodeError:
            raw = []
    r["answersDetail"] = raw
    r["timestamp"] = r.pop("created_at", "")
    r["empId"] = r.pop("emp_id", "")
    r["evaluatorName"] = r.pop("evaluator_name", "")
    r["evaluatorTitle"] = r.pop("evaluator_title", "")
    group = r.pop("group_key", "grpBio") or "grpBio"
    r["groupKey"] = group
    r["groupLabel"] = GROUPS.get(group, GROUPS["grpBio"])
    return r


def require_admin():
    supplied = request.headers.get("X-Admin-Key", "")
    if not ADMIN_KEY or supplied != ADMIN_KEY:
        return jsonify({"error": "未授權。請提供正確的管理者金鑰 ADMIN_KEY。"}), 401
    return None


def normalize_group(value):
    """把不合法或空白的組別代碼收斂為預設的生化組，避免髒資料。"""
    return value if value in GROUPS else DEFAULT_GROUP


init_exam_db()


def init_materials_db():
    conn, kind = _db_conn()
    try:
        if kind == "postgres":
            conn.execute("""
                CREATE TABLE IF NOT EXISTS materials (
                    id TEXT PRIMARY KEY,
                    filename TEXT NOT NULL,
                    title TEXT NOT NULL,
                    description TEXT NOT NULL DEFAULT '',
                    category TEXT NOT NULL DEFAULT '',
                    group_key TEXT NOT NULL DEFAULT 'grpBio',
                    folder TEXT NOT NULL,
                    page_count INTEGER NOT NULL DEFAULT 0,
                    date_added TEXT NOT NULL,
                    storage_filename TEXT NOT NULL,
                    active BOOLEAN NOT NULL DEFAULT TRUE
                )
            """)
            try:
                conn.execute("ALTER TABLE materials ADD COLUMN IF NOT EXISTS group_key TEXT NOT NULL DEFAULT 'grpBio'")
            except Exception:
                pass
        else:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS materials (
                    id TEXT PRIMARY KEY,
                    filename TEXT NOT NULL,
                    title TEXT NOT NULL,
                    description TEXT NOT NULL DEFAULT '',
                    category TEXT NOT NULL DEFAULT '',
                    group_key TEXT NOT NULL DEFAULT 'grpBio',
                    folder TEXT NOT NULL,
                    page_count INTEGER NOT NULL DEFAULT 0,
                    date_added TEXT NOT NULL,
                    storage_filename TEXT NOT NULL,
                    active INTEGER NOT NULL DEFAULT 1
                )
            """)
            existing_cols = {row[1] for row in conn.execute("PRAGMA table_info(materials)").fetchall()}
            if "group_key" not in existing_cols:
                conn.execute("ALTER TABLE materials ADD COLUMN group_key TEXT NOT NULL DEFAULT 'grpBio'")
    finally:
        conn.close()


def init_quiz_db():
    """考題頁籤（quiz_categories）與題目（quiz_questions）：
    生化組沿用前端內建的固定 4 個頁籤與題庫；其餘 5 組的考題頁籤與題目
    則完全由後台管理者透過這兩張表新增/編輯/刪除。"""
    conn, kind = _db_conn()
    try:
        if kind == "postgres":
            conn.execute("""
                CREATE TABLE IF NOT EXISTS quiz_categories (
                    id TEXT PRIMARY KEY,
                    group_key TEXT NOT NULL DEFAULT 'grpBio',
                    title TEXT NOT NULL,
                    description TEXT NOT NULL DEFAULT '',
                    sort_order INTEGER NOT NULL DEFAULT 0,
                    date_added TEXT NOT NULL,
                    active BOOLEAN NOT NULL DEFAULT TRUE
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS quiz_questions (
                    id TEXT PRIMARY KEY,
                    quiz_category_id TEXT NOT NULL,
                    tag TEXT NOT NULL DEFAULT '',
                    question TEXT NOT NULL,
                    question_type TEXT NOT NULL DEFAULT 'choice',
                    image_url TEXT NOT NULL DEFAULT '',
                    options JSONB NOT NULL,
                    correct INTEGER NOT NULL DEFAULT 0,
                    explanation TEXT NOT NULL DEFAULT '',
                    sort_order INTEGER NOT NULL DEFAULT 0
                )
            """)
        else:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS quiz_categories (
                    id TEXT PRIMARY KEY,
                    group_key TEXT NOT NULL DEFAULT 'grpBio',
                    title TEXT NOT NULL,
                    description TEXT NOT NULL DEFAULT '',
                    sort_order INTEGER NOT NULL DEFAULT 0,
                    date_added TEXT NOT NULL,
                    active INTEGER NOT NULL DEFAULT 1
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS quiz_questions (
                    id TEXT PRIMARY KEY,
                    quiz_category_id TEXT NOT NULL,
                    tag TEXT NOT NULL DEFAULT '',
                    question TEXT NOT NULL,
                    question_type TEXT NOT NULL DEFAULT 'choice',
                    image_url TEXT NOT NULL DEFAULT '',
                    options TEXT NOT NULL,
                    correct INTEGER NOT NULL DEFAULT 0,
                    explanation TEXT NOT NULL DEFAULT '',
                    sort_order INTEGER NOT NULL DEFAULT 0
                )
            """)
    finally:
        conn.close()


def material_row_to_dict(row):
    r = dict(row)
    r["isBuiltin"] = False
    r["is_builtin"] = False
    r["desc"] = r.pop("description", "")
    r["dateAdded"] = r.pop("date_added", "")
    r["pageCount"] = int(r.pop("page_count", 0) or 0)
    r["storageFilename"] = r.pop("storage_filename", "")
    r["active"] = bool(r.get("active", True))
    r["group"] = normalize_group(r.pop("group_key", DEFAULT_GROUP))
    return r


def list_uploaded_materials(include_inactive=False):
    conn, _ = _db_conn()
    try:
        sql = "SELECT * FROM materials"
        if not include_inactive:
            sql += " WHERE active = " + ("TRUE" if DATABASE_URL else "1")
        sql += " ORDER BY date_added DESC"
        return [material_row_to_dict(r) for r in conn.execute(sql).fetchall()]
    finally:
        conn.close()


def get_material(material_id):
    conn, _ = _db_conn()
    try:
        row = conn.execute("SELECT * FROM materials WHERE id = %s" % ("%s" if DATABASE_URL else "?"), (material_id,)).fetchone()
        return material_row_to_dict(row) if row else None
    finally:
        conn.close()


init_materials_db()
init_quiz_db()


# ---------------------------------------------------------------------------
# 考題頁籤 (quiz_categories) 與題目 (quiz_questions) 存取輔助函式
# ---------------------------------------------------------------------------
def quiz_category_row_to_dict(row):
    r = dict(row)
    r["desc"] = r.pop("description", "")
    r["dateAdded"] = r.pop("date_added", "")
    r["active"] = bool(r.get("active", True))
    r["group"] = normalize_group(r.pop("group_key", DEFAULT_GROUP))
    r["sortOrder"] = int(r.pop("sort_order", 0) or 0)
    return r


def quiz_question_row_to_dict(row):
    r = dict(row)
    raw_options = r.pop("options", "[]")
    if isinstance(raw_options, str):
        try:
            raw_options = json.loads(raw_options)
        except json.JSONDecodeError:
            raw_options = []
    r["options"] = raw_options
    r["questionType"] = r.pop("question_type", "choice")
    r["imageUrl"] = r.pop("image_url", "")
    r["quizCategoryId"] = r.pop("quiz_category_id", "")
    r["sortOrder"] = int(r.pop("sort_order", 0) or 0)
    return r


def list_quiz_categories(group_key=None, include_inactive=False):
    conn, kind = _db_conn()
    try:
        ph = "%s" if kind == "postgres" else "?"
        sql = "SELECT * FROM quiz_categories"
        params = []
        clauses = []
        if group_key:
            clauses.append(f"group_key = {ph}")
            params.append(group_key)
        if not include_inactive:
            clauses.append("active = " + ("TRUE" if kind == "postgres" else "1"))
        if clauses:
            sql += " WHERE " + " AND ".join(clauses)
        sql += " ORDER BY sort_order ASC, date_added ASC"
        return [quiz_category_row_to_dict(r) for r in conn.execute(sql, params).fetchall()]
    finally:
        conn.close()


def get_quiz_category(category_id):
    conn, kind = _db_conn()
    try:
        ph = "%s" if kind == "postgres" else "?"
        row = conn.execute(f"SELECT * FROM quiz_categories WHERE id = {ph}", (category_id,)).fetchone()
        return quiz_category_row_to_dict(row) if row else None
    finally:
        conn.close()


def list_quiz_questions(category_id):
    conn, kind = _db_conn()
    try:
        ph = "%s" if kind == "postgres" else "?"
        rows = conn.execute(
            f"SELECT * FROM quiz_questions WHERE quiz_category_id = {ph} ORDER BY sort_order ASC",
            (category_id,),
        ).fetchall()
        return [quiz_question_row_to_dict(r) for r in rows]
    finally:
        conn.close()


def get_quiz_question(question_id):
    conn, kind = _db_conn()
    try:
        ph = "%s" if kind == "postgres" else "?"
        row = conn.execute(f"SELECT * FROM quiz_questions WHERE id = {ph}", (question_id,)).fetchone()
        return quiz_question_row_to_dict(row) if row else None
    finally:
        conn.close()


def resolve_category_label(category, group_key):
    """依 category 代碼組出前端顯示用的頁籤名稱：
    生化組固定代碼查 CATEGORY_LABELS；其餘組別的代碼是 quiz_categories.id，查資料庫標題。"""
    if not category:
        return CATEGORY_LABELS[""]
    if category in CATEGORY_LABELS:
        return CATEGORY_LABELS[category]
    cat = get_quiz_category(category)
    if cat:
        return cat["title"]
    return CATEGORY_LABELS[""]


# ---------------------------------------------------------------------------
# 附件1 Word 匯出範本 (doc_templates) —— 每個組別可各自上傳一份空白 .docx 範本，
# 前端匯出時改用該組別的範本自動填入，取代原本每次都要在瀏覽器手動選檔的方式。
# 生化組沿用原本前端「選一次、瀏覽器暫存」的既有流程，不受此表影響。
# ---------------------------------------------------------------------------
def init_doc_templates_db():
    conn, kind = _db_conn()
    try:
        if kind == "postgres":
            conn.execute("""
                CREATE TABLE IF NOT EXISTS doc_templates (
                    group_key TEXT PRIMARY KEY,
                    filename TEXT NOT NULL,
                    storage_filename TEXT NOT NULL,
                    uploaded_at TEXT NOT NULL
                )
            """)
        else:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS doc_templates (
                    group_key TEXT PRIMARY KEY,
                    filename TEXT NOT NULL,
                    storage_filename TEXT NOT NULL,
                    uploaded_at TEXT NOT NULL
                )
            """)
    finally:
        conn.close()


def list_doc_templates():
    conn, kind = _db_conn()
    try:
        rows = conn.execute("SELECT * FROM doc_templates").fetchall()
        return {row["group_key"]: dict(row) for row in rows}
    finally:
        conn.close()


def get_doc_template_row(group_key):
    conn, kind = _db_conn()
    try:
        ph = "%s" if kind == "postgres" else "?"
        row = conn.execute(f"SELECT * FROM doc_templates WHERE group_key = {ph}", (group_key,)).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def save_doc_template_row(group_key, filename, storage_filename):
    conn, kind = _db_conn()
    try:
        uploaded_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        if kind == "postgres":
            conn.execute("""
                INSERT INTO doc_templates (group_key, filename, storage_filename, uploaded_at)
                VALUES (%s,%s,%s,%s)
                ON CONFLICT (group_key) DO UPDATE SET filename=EXCLUDED.filename,
                    storage_filename=EXCLUDED.storage_filename, uploaded_at=EXCLUDED.uploaded_at
            """, (group_key, filename, storage_filename, uploaded_at))
        else:
            conn.execute("""
                INSERT INTO doc_templates (group_key, filename, storage_filename, uploaded_at) VALUES (?,?,?,?)
                ON CONFLICT(group_key) DO UPDATE SET filename=excluded.filename,
                    storage_filename=excluded.storage_filename, uploaded_at=excluded.uploaded_at
            """, (group_key, filename, storage_filename, uploaded_at))
    finally:
        conn.close()


def delete_doc_template_row(group_key):
    conn, kind = _db_conn()
    try:
        ph = "%s" if kind == "postgres" else "?"
        conn.execute(f"DELETE FROM doc_templates WHERE group_key = {ph}", (group_key,))
    finally:
        conn.close()


init_doc_templates_db()


# ----------------------------------------------------------------------------
# 簡報中繼資料 (slides_meta.json) 存取
# ---------------------------------------------------------------------------
def load_meta():
    if META_FILE.exists():
        return json.loads(META_FILE.read_text(encoding="utf-8"))
    return []


def save_meta(meta):
    META_FILE.write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def init_builtin_meta():
    """第一次啟動時，寫入 4 份內建簡報的中繼資料（圖片需已存在於 static/slides/ 底下）。"""
    if META_FILE.exists():
        return
    builtin = [
        {
            "id": "builtin-subA1",
            "filename": "2026_生化教育訓練.pptx",
            "title": "2026_生化教育訓練.pptx",
            "desc": "涵蓋一致性作業、異常檢體處理與法定傳染病通報流程之總論教材。",
            "category": "subA1",
            "group": "grpBio",
            "folder": "subA1",
            "pageCount": 16,
            "isBuiltin": True,
            "dateAdded": "",
        },
        {
            "id": "builtin-subA2",
            "filename": "2026生化組教育訓練_c503_操作與維護.pptx",
            "title": "2026生化組教育訓練_c503_操作與維護.pptx",
            "desc": "Cobas pro c503 一般作業流程、日常保養與異常訊號故障排除教材。",
            "category": "subA2",
            "group": "grpBio",
            "folder": "subA2",
            "pageCount": 13,
            "isBuiltin": True,
            "dateAdded": "",
        },
        {
            "id": "builtin-subA3",
            "filename": "Cobas_b_211_儀器教育訓練.pptx",
            "title": "Cobas_b_211_儀器教育訓練.pptx",
            "desc": "Cobas b 211 異常訊號故障排除、管路潤濕與 QC 批號設定教材。",
            "category": "subA3",
            "group": "grpBio",
            "folder": "subA3",
            "pageCount": 10,
            "isBuiltin": True,
            "dateAdded": "",
        },
        {
            "id": "builtin-zoneB",
            "filename": "Sebia_簡易故障排除指南.pptx",
            "title": "Sebia_簡易故障排除指南.pptx",
            "desc": "Sebia 儀器 Cup/Rack 卡住、管路異常與各類警報之簡易故障排除教材。",
            "category": "zoneB",
            "group": "grpBio",
            "folder": "zoneB",
            "pageCount": 7,
            "isBuiltin": True,
            "dateAdded": "",
        },
    ]
    save_meta(builtin)


init_builtin_meta()


# ---------------------------------------------------------------------------
# 核心功能：PPTX -> PDF -> 逐頁 PNG
# ---------------------------------------------------------------------------
def convert_pptx_to_images(pptx_path: Path, out_folder: Path) -> int:
    """
    將 pptx_path 轉換成一系列 PNG 圖片，存到 out_folder（每頁一張，slide-01.png、slide-02.png...）。
    回傳頁數。若轉檔失敗會拋出例外。
    """
    if fitz is None:
        raise RuntimeError("伺服器缺少 PyMuPDF 套件，請先執行 pip install -r requirements.txt")

    out_folder.mkdir(parents=True, exist_ok=True)

    with conversion_lock:
        profile_dir = TMP_DIR / f"profile-{uuid.uuid4().hex}"
        pdf_tmp_dir = TMP_DIR / f"pdf-{uuid.uuid4().hex}"
        profile_dir.mkdir(parents=True, exist_ok=True)
        pdf_tmp_dir.mkdir(parents=True, exist_ok=True)

        try:
            # 1) 用 LibreOffice 無頭模式將 pptx 轉成 pdf
            #    -env:UserInstallation 讓每次轉檔使用獨立設定檔，避免併發時互相鎖死
            cmd = [
                SOFFICE_BIN,
                "--headless",
                "--norestore",
                f"-env:UserInstallation=file:///{profile_dir.as_posix()}",
                "--convert-to",
                "pdf",
                "--outdir",
                str(pdf_tmp_dir),
                str(pptx_path),
            ]
            result = subprocess.run(
                cmd, capture_output=True, timeout=180
            )
            if result.returncode != 0:
                stderr = result.stderr.decode(errors="ignore")
                raise RuntimeError(f"LibreOffice 轉檔失敗：{stderr[:500]}")

            pdf_files = list(pdf_tmp_dir.glob("*.pdf"))
            if not pdf_files:
                raise RuntimeError("找不到轉檔後產生的 PDF 檔案，請確認 LibreOffice 是否正確安裝。")
            pdf_path = pdf_files[0]

            # 2) 用 PyMuPDF 把 PDF 每一頁轉成 PNG（約 110 DPI，畫質與檔案大小平衡）
            doc = fitz.open(str(pdf_path))
            zoom = 110 / 72.0
            mat = fitz.Matrix(zoom, zoom)
            page_count = doc.page_count
            for i in range(page_count):
                page = doc.load_page(i)
                pix = page.get_pixmap(matrix=mat)
                img_path = out_folder / f"slide-{i + 1:02d}.png"
                pix.save(str(img_path))
            doc.close()

            if page_count == 0:
                raise RuntimeError("簡報頁數為 0，請確認檔案內容是否正確。")

            return page_count
        finally:
            shutil.rmtree(profile_dir, ignore_errors=True)
            shutil.rmtree(pdf_tmp_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# API 路由
# ---------------------------------------------------------------------------
@app.get("/api/slides")
def api_list_slides():
    builtin = []
    for m in load_meta():
        if not m.get("isBuiltin"):
            continue
        group = normalize_group(m.get("group", DEFAULT_GROUP))
        builtin.append({
            **m,
            "group": group,
            "categoryLabel": resolve_category_label(m.get("category", ""), group),
            "imageFolder": f"slides/{m['folder']}",
            "downloadUrl": f"/download/{m['id']}",
        })
    uploaded = []
    for m in list_uploaded_materials(False):
        uploaded.append({
            **m,
            "categoryLabel": resolve_category_label(m.get("category", ""), m.get("group", DEFAULT_GROUP)),
            "imageFolder": f"uploaded-slides/{m['folder']}",
            "downloadUrl": f"/download/{m['id']}",
        })
    return jsonify(builtin + uploaded)


@app.get("/api/slides/admin")
def api_admin_slides():
    denied = require_admin()
    if denied:
        return denied
    items = []
    for m in load_meta():
        if m.get("isBuiltin"):
            group = normalize_group(m.get("group", DEFAULT_GROUP))
            items.append({**m, "group": group, "categoryLabel": resolve_category_label(m.get("category", ""), group)})
    for m in list_uploaded_materials(True):
        items.append({**m, "categoryLabel": resolve_category_label(m.get("category", ""), m.get("group", DEFAULT_GROUP))})
    return jsonify(items)


@app.get("/api/groups")
def api_list_groups():
    return jsonify([{"key": k, "label": v} for k, v in GROUPS.items()])


@app.get("/uploaded-slides/<folder>/<path:filename>")
def uploaded_slide_image(folder, filename):
    # 僅允許以資料庫登記的 folder 存取，避免路徑穿越。
    if Path(folder).name != folder or Path(filename).name != filename:
        abort(404)
    entry = get_material(folder)
    if not entry or not entry.get("active"):
        abort(404)
    return send_from_directory(UPLOADED_SLIDES_DIR / folder, filename)


@app.get("/download/<slide_id>")
def download_slide(slide_id):
    meta = load_meta()
    entry = next((m for m in meta if m["id"] == slide_id), None)
    if entry and entry.get("isBuiltin"):
        return send_from_directory(STATIC_DIR, entry["filename"], as_attachment=True, download_name=entry["filename"])
    entry = get_material(slide_id)
    if not entry:
        abort(404)
    path = UPLOAD_DIR / entry["id"] / entry["storageFilename"]
    if not path.exists():
        abort(404)
    return send_file(path, as_attachment=True, download_name=entry["filename"])


@app.post("/api/slides/upload")
def api_upload_slide():
    denied = require_admin()
    if denied:
        return denied
    if "file" not in request.files:
        return jsonify({"error": "未收到檔案"}), 400

    file = request.files["file"]
    group = normalize_group(request.form.get("group", DEFAULT_GROUP))
    category = request.form.get("category", "")
    # 生化組沿用固定頁籤代碼；其餘組別的 category 代碼是 quiz_categories.id，
    # 只要能在該表查到（且屬於同一組別）即視為合法，否則歸為未分類。
    if group == DEFAULT_GROUP:
        if category not in CATEGORY_LABELS:
            category = ""
    else:
        cat = get_quiz_category(category) if category else None
        if not cat or cat["group"] != group:
            category = ""
    title = request.form.get("title", "").strip()
    desc = request.form.get("desc", "").strip()

    original_name = Path(file.filename or "untitled.pptx").name
    ext = Path(original_name).suffix.lower()
    if ext not in ALLOWED_EXT:
        return jsonify({"error": "僅支援 .pptx / .ppt 檔案"}), 400

    slide_id = f"upload-{uuid.uuid4().hex[:12]}"
    material_dir = UPLOAD_DIR / slide_id
    out_folder = UPLOADED_SLIDES_DIR / slide_id
    material_dir.mkdir(parents=True, exist_ok=True)
    out_folder.mkdir(parents=True, exist_ok=True)
    saved_path = material_dir / f"source{ext}"
    file.save(str(saved_path))

    try:
        page_count = convert_pptx_to_images(saved_path, out_folder)
    except Exception as e:
        shutil.rmtree(material_dir, ignore_errors=True)
        shutil.rmtree(out_folder, ignore_errors=True)
        return jsonify({"error": f"轉檔失敗：{e}"}), 500

    entry = {
        "id": slide_id,
        "filename": original_name,
        "title": title or original_name,
        "description": desc or "管理者上傳之教育訓練補充教材",
        "category": category,
        "group_key": group,
        "folder": slide_id,
        "page_count": page_count,
        "date_added": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "storage_filename": saved_path.name,
        "active": True,
    }
    conn, kind = _db_conn()
    try:
        if kind == "postgres":
            conn.execute("""INSERT INTO materials (id,filename,title,description,category,group_key,folder,page_count,date_added,storage_filename,active) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""", tuple(entry.values()))
        else:
            conn.execute("""INSERT INTO materials (id,filename,title,description,category,group_key,folder,page_count,date_added,storage_filename,active) VALUES (?,?,?,?,?,?,?,?,?,?,?)""", tuple(entry.values()))
    except Exception:
        shutil.rmtree(material_dir, ignore_errors=True)
        shutil.rmtree(out_folder, ignore_errors=True)
        raise
    finally:
        conn.close()

    return jsonify({
        "id": slide_id, "filename": original_name, "title": entry["title"], "desc": entry["description"],
        "category": category, "group": group, "folder": slide_id, "pageCount": page_count, "isBuiltin": False,
        "dateAdded": entry["date_added"], "active": True,
        "categoryLabel": resolve_category_label(category, group),
        "imageFolder": f"uploaded-slides/{slide_id}", "downloadUrl": f"/download/{slide_id}"
    })


@app.patch("/api/slides/<slide_id>")
def api_update_slide(slide_id):
    denied = require_admin()
    if denied:
        return denied
    entry = get_material(slide_id)
    if not entry:
        return jsonify({"error": "找不到可編輯的上傳教材"}), 404
    data = request.get_json(silent=True) or {}
    title = str(data.get("title", entry["title"])).strip()[:255]
    desc = str(data.get("desc", entry.get("desc", ""))).strip()[:1000]
    group = normalize_group(str(data.get("group", entry.get("group", DEFAULT_GROUP))))
    category = str(data.get("category", entry.get("category", "")))
    active = bool(data.get("active", entry.get("active", True)))
    if group == DEFAULT_GROUP:
        if category not in CATEGORY_LABELS:
            category = ""
    else:
        cat = get_quiz_category(category) if category else None
        if not cat or cat["group"] != group:
            category = ""
    conn, kind = _db_conn()
    try:
        if kind == "postgres":
            conn.execute("UPDATE materials SET title=%s, description=%s, category=%s, group_key=%s, active=%s WHERE id=%s", (title, desc, category, group, active, slide_id))
        else:
            conn.execute("UPDATE materials SET title=?, description=?, category=?, group_key=?, active=? WHERE id=?", (title, desc, category, group, int(active), slide_id))
    finally:
        conn.close()
    return jsonify({"ok": True})


@app.delete("/api/slides/<slide_id>")
def api_delete_slide(slide_id):
    denied = require_admin()
    if denied:
        return denied
    entry = get_material(slide_id)
    if not entry:
        return jsonify({"error": "內建教材不能從後台刪除，或找不到此教材"}), 404
    shutil.rmtree(UPLOAD_DIR / entry["id"], ignore_errors=True)
    shutil.rmtree(UPLOADED_SLIDES_DIR / entry["folder"], ignore_errors=True)
    conn, kind = _db_conn()
    try:
        if kind == "postgres":
            conn.execute("DELETE FROM materials WHERE id=%s", (slide_id,))
        else:
            conn.execute("DELETE FROM materials WHERE id=?", (slide_id,))
    finally:
        conn.close()
    return jsonify({"ok": True})


# ---------------------------------------------------------------------------
# 考題頁籤 (quiz_categories) API —— 生化組以外的組別由後台管理者建立考題頁籤
# ---------------------------------------------------------------------------
@app.get("/api/quiz-categories")
def api_list_quiz_categories():
    group = request.args.get("group", "")
    group = group if group in GROUPS else None
    return jsonify(list_quiz_categories(group_key=group, include_inactive=False))


@app.get("/api/quiz-categories/admin")
def api_admin_list_quiz_categories():
    denied = require_admin()
    if denied:
        return denied
    group = request.args.get("group", "")
    group = group if group in GROUPS else None
    return jsonify(list_quiz_categories(group_key=group, include_inactive=True))


@app.post("/api/quiz-categories")
def api_create_quiz_category():
    denied = require_admin()
    if denied:
        return denied
    data = request.get_json(silent=True) or {}
    group = normalize_group(str(data.get("group", DEFAULT_GROUP)))
    title = str(data.get("title", "")).strip()[:255]
    desc = str(data.get("desc", "")).strip()[:1000]
    if not title:
        return jsonify({"error": "請輸入頁籤名稱"}), 400
    cat_id = f"cat-{uuid.uuid4().hex[:12]}"
    conn, kind = _db_conn()
    try:
        existing = conn.execute(
            f"SELECT COALESCE(MAX(sort_order), -1) AS m FROM quiz_categories WHERE group_key = {'%s' if kind == 'postgres' else '?'}",
            (group,),
        ).fetchone()
        next_order = (existing["m"] if isinstance(existing, dict) else existing[0]) + 1
        date_added = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        if kind == "postgres":
            conn.execute(
                "INSERT INTO quiz_categories (id,group_key,title,description,sort_order,date_added,active) VALUES (%s,%s,%s,%s,%s,%s,%s)",
                (cat_id, group, title, desc, next_order, date_added, True),
            )
        else:
            conn.execute(
                "INSERT INTO quiz_categories (id,group_key,title,description,sort_order,date_added,active) VALUES (?,?,?,?,?,?,?)",
                (cat_id, group, title, desc, next_order, date_added, 1),
            )
    finally:
        conn.close()
    return jsonify(get_quiz_category(cat_id))


@app.patch("/api/quiz-categories/<category_id>")
def api_update_quiz_category(category_id):
    denied = require_admin()
    if denied:
        return denied
    entry = get_quiz_category(category_id)
    if not entry:
        return jsonify({"error": "找不到此考題頁籤"}), 404
    data = request.get_json(silent=True) or {}
    title = str(data.get("title", entry["title"])).strip()[:255]
    desc = str(data.get("desc", entry.get("desc", ""))).strip()[:1000]
    active = bool(data.get("active", entry.get("active", True)))
    conn, kind = _db_conn()
    try:
        if kind == "postgres":
            conn.execute("UPDATE quiz_categories SET title=%s, description=%s, active=%s WHERE id=%s", (title, desc, active, category_id))
        else:
            conn.execute("UPDATE quiz_categories SET title=?, description=?, active=? WHERE id=?", (title, desc, int(active), category_id))
    finally:
        conn.close()
    return jsonify({"ok": True})


@app.delete("/api/quiz-categories/<category_id>")
def api_delete_quiz_category(category_id):
    denied = require_admin()
    if denied:
        return denied
    entry = get_quiz_category(category_id)
    if not entry:
        return jsonify({"error": "找不到此考題頁籤"}), 404
    conn, kind = _db_conn()
    try:
        if kind == "postgres":
            conn.execute("DELETE FROM quiz_questions WHERE quiz_category_id=%s", (category_id,))
            conn.execute("DELETE FROM quiz_categories WHERE id=%s", (category_id,))
            conn.execute("UPDATE materials SET category='' WHERE category=%s", (category_id,))
        else:
            conn.execute("DELETE FROM quiz_questions WHERE quiz_category_id=?", (category_id,))
            conn.execute("DELETE FROM quiz_categories WHERE id=?", (category_id,))
            conn.execute("UPDATE materials SET category='' WHERE category=?", (category_id,))
    finally:
        conn.close()
    return jsonify({"ok": True})


# ---------------------------------------------------------------------------
# 考題 (quiz_questions) API —— 每一題屬於某一個 quiz_category
# ---------------------------------------------------------------------------
@app.post("/api/quiz-question-images")
def api_upload_question_image():
    denied = require_admin()
    if denied: return denied
    if "file" not in request.files: return jsonify({"error":"缺少圖片檔案"}), 400
    f=request.files["file"]; ext=Path(f.filename or "").suffix.lower()
    if ext not in {".png",".jpg",".jpeg",".gif",".webp"}: return jsonify({"error":"僅接受 PNG/JPG/GIF/WEBP"}),400
    name=f"{uuid.uuid4().hex}{ext}"; f.save(str(QUESTION_IMAGES_DIR/name))
    return jsonify({"url":f"/question-images/{name}"})

@app.get("/question-images/<path:name>")
def question_image(name):
    return send_from_directory(str(QUESTION_IMAGES_DIR), name)

@app.get("/api/quiz-questions/random")
def api_random_quiz_questions():
    import random
    category_id = request.args.get("category", "")
    try: count = max(1, int(request.args.get("count", "10")))
    except ValueError: count = 10
    qs = list_quiz_questions(category_id) if category_id else []
    random.shuffle(qs)
    return jsonify(qs[:count])


@app.get("/api/quiz-questions")
def api_list_quiz_questions():
    category_id = request.args.get("category", "")
    if not category_id:
        return jsonify({"error": "缺少 category 參數"}), 400
    return jsonify(list_quiz_questions(category_id))


@app.post("/api/quiz-questions")
def api_create_quiz_question():
    denied = require_admin()
    if denied:
        return denied
    data = request.get_json(silent=True) or {}
    category_id = str(data.get("quizCategoryId", "")).strip()
    cat = get_quiz_category(category_id)
    if not cat:
        return jsonify({"error": "找不到對應的考題頁籤，請先建立頁籤"}), 400
    question = str(data.get("question", "")).strip()
    question_type = str(data.get("questionType", "choice")).lower()
    if question_type not in ("choice", "essay"):
        question_type = "choice"
    image_url = str(data.get("imageUrl", "")).strip()[:1000]
    options = data.get("options", [])
    if question_type == "essay":
        options = []
    tag = str(data.get("tag", "")).strip()[:100] or "一般"
    explanation = str(data.get("explanation", "")).strip()
    try:
        correct = int(data.get("correct", 0))
    except (TypeError, ValueError):
        correct = 0
    if not question or (question_type == "choice" and (not isinstance(options, list) or len(options) < 2)):
        return jsonify({"error": "請輸入題目與至少 2 個選項"}), 400
    options = [str(o).strip() for o in options][:6]
    correct = max(0, min(len(options) - 1, correct))
    q_id = f"q-{uuid.uuid4().hex[:12]}"
    conn, kind = _db_conn()
    try:
        existing = conn.execute(
            f"SELECT COALESCE(MAX(sort_order), -1) AS m FROM quiz_questions WHERE quiz_category_id = {'%s' if kind == 'postgres' else '?'}",
            (category_id,),
        ).fetchone()
        next_order = (existing["m"] if isinstance(existing, dict) else existing[0]) + 1
        if kind == "postgres":
            conn.execute(
                "INSERT INTO quiz_questions (id,quiz_category_id,tag,question,question_type,image_url,options,correct,explanation,sort_order) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                (q_id, category_id, tag, question, question_type, image_url, json.dumps(options, ensure_ascii=False), correct, explanation, next_order),
            )
        else:
            conn.execute(
                "INSERT INTO quiz_questions (id,quiz_category_id,tag,question,question_type,image_url,options,correct,explanation,sort_order) VALUES (?,?,?,?,?,?,?,?,?,?)",
                (q_id, category_id, tag, question, question_type, image_url, json.dumps(options, ensure_ascii=False), correct, explanation, next_order),
            )
    finally:
        conn.close()
    return jsonify(get_quiz_question(q_id))


@app.patch("/api/quiz-questions/<question_id>")
def api_update_quiz_question(question_id):
    denied = require_admin()
    if denied:
        return denied
    entry = get_quiz_question(question_id)
    if not entry:
        return jsonify({"error": "找不到此題目"}), 404
    data = request.get_json(silent=True) or {}
    question = str(data.get("question", entry["question"])).strip()
    question_type = str(data.get("questionType", entry.get("questionType", "choice"))).lower()
    if question_type not in ("choice", "essay"): question_type = "choice"
    image_url = str(data.get("imageUrl", entry.get("imageUrl", ""))).strip()[:1000]
    options = data.get("options", entry["options"])
    if question_type == "essay": options = []
    if question_type == "choice" and (not isinstance(options, list) or len(options) < 2):
        return jsonify({"error": "選項至少需要 2 個"}), 400
    options = [str(o).strip() for o in options][:6]
    tag = str(data.get("tag", entry.get("tag", ""))).strip()[:100] or "一般"
    explanation = str(data.get("explanation", entry.get("explanation", ""))).strip()
    try:
        correct = int(data.get("correct", entry.get("correct", 0)))
    except (TypeError, ValueError):
        correct = entry.get("correct", 0)
    correct = max(0, min(len(options) - 1, correct))
    conn, kind = _db_conn()
    try:
        if kind == "postgres":
            conn.execute(
                "UPDATE quiz_questions SET tag=%s, question=%s, question_type=%s, image_url=%s, options=%s, correct=%s, explanation=%s WHERE id=%s",
                (tag, question, question_type, image_url, json.dumps(options, ensure_ascii=False), correct, explanation, question_id),
            )
        else:
            conn.execute(
                "UPDATE quiz_questions SET tag=?, question=?, question_type=?, image_url=?, options=?, correct=?, explanation=? WHERE id=?",
                (tag, question, question_type, image_url, json.dumps(options, ensure_ascii=False), correct, explanation, question_id),
            )
    finally:
        conn.close()
    return jsonify({"ok": True})


@app.delete("/api/quiz-questions/<question_id>")
def api_delete_quiz_question(question_id):
    denied = require_admin()
    if denied:
        return denied
    entry = get_quiz_question(question_id)
    if not entry:
        return jsonify({"error": "找不到此題目"}), 404
    conn, kind = _db_conn()
    try:
        if kind == "postgres":
            conn.execute("DELETE FROM quiz_questions WHERE id=%s", (question_id,))
        else:
            conn.execute("DELETE FROM quiz_questions WHERE id=?", (question_id,))
    finally:
        conn.close()
    return jsonify({"ok": True})


# ---------------------------------------------------------------------------
# 附件1 Word 匯出範本 (doc_templates) API
# ---------------------------------------------------------------------------
DOC_TEMPLATE_ALLOWED_EXT = {".docx"}
MAX_DOC_TEMPLATE_MB = 20


@app.get("/api/doc-templates")
def api_list_doc_templates():
    """公開：各組別是否已設定範本（前端匯出前用來判斷要不要先提示管理者上傳）。"""
    rows = list_doc_templates()
    return jsonify([
        {
            "group": g,
            "label": label,
            "exists": g in rows,
            "filename": rows.get(g, {}).get("filename", ""),
            "uploadedAt": rows.get(g, {}).get("uploaded_at", ""),
        }
        for g, label in GROUPS.items()
    ])


@app.post("/api/doc-templates/<group_key>")
def api_upload_doc_template(group_key):
    denied = require_admin()
    if denied:
        return denied
    if group_key not in GROUPS:
        return jsonify({"error": "無效的組別代碼"}), 400
    if "file" not in request.files:
        return jsonify({"error": "缺少檔案"}), 400
    file = request.files["file"]
    original_name = file.filename or "附件1.docx"
    ext = Path(original_name).suffix.lower()
    if ext not in DOC_TEMPLATE_ALLOWED_EXT:
        return jsonify({"error": "僅接受 .docx 檔案"}), 400
    file.seek(0, os.SEEK_END)
    size_mb = file.tell() / (1024 * 1024)
    file.seek(0)
    if size_mb > MAX_DOC_TEMPLATE_MB:
        return jsonify({"error": f"檔案過大，上限 {MAX_DOC_TEMPLATE_MB}MB"}), 400
    storage_filename = f"{group_key}.docx"
    file.save(str(DOC_TEMPLATES_DIR / storage_filename))
    save_doc_template_row(group_key, original_name, storage_filename)
    return jsonify({"ok": True, "group": group_key, "filename": original_name})


@app.get("/api/doc-templates/<group_key>/download")
def api_download_doc_template(group_key):
    """公開：前端用 fetch + docxtemplater 取範本二進位內容並在瀏覽器端填入匯出。"""
    row = get_doc_template_row(group_key)
    if not row:
        return jsonify({"error": "此組別尚未上傳 Word 匯出範本"}), 404
    path = DOC_TEMPLATES_DIR / row["storage_filename"]
    if not path.exists():
        return jsonify({"error": "範本檔案遺失，請管理者重新上傳"}), 404
    return send_file(path, as_attachment=False, download_name=row["filename"])


@app.delete("/api/doc-templates/<group_key>")
def api_delete_doc_template(group_key):
    denied = require_admin()
    if denied:
        return denied
    row = get_doc_template_row(group_key)
    if row:
        path = DOC_TEMPLATES_DIR / row["storage_filename"]
        if path.exists():
            path.unlink()
        delete_doc_template_row(group_key)
    return jsonify({"ok": True})


@app.post("/api/records")
def api_create_record():
    data = request.get_json(silent=True) or {}
    required = ["id", "name", "empId", "role", "quizTitle", "score", "status"]
    missing = [k for k in required if data.get(k) in (None, "")]
    if missing:
        return jsonify({"error": f"缺少欄位：{', '.join(missing)}"}), 400

    try:
        score = int(data.get("score", 0))
        correct_count = int(data.get("correctCount", 0))
        wrong_count = int(data.get("wrongCount", 0))
    except (TypeError, ValueError):
        return jsonify({"error": "分數或題數格式錯誤"}), 400
    score = max(0, min(100, score))
    record_id = str(data["id"])[:100]
    created_at = datetime.datetime.now(datetime.timezone.utc).isoformat()
    answers = data.get("answersDetail", [])
    group_key = normalize_group(str(data.get("groupKey", DEFAULT_GROUP)))

    conn, kind = _db_conn()
    try:
        if kind == "postgres":
            conn.execute("""
                INSERT INTO exam_records
                (id, created_at, name, emp_id, role, evaluator_name, evaluator_title, quiz_title, score, status, correct_count, wrong_count, answers_detail, group_key)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                ON CONFLICT (id) DO NOTHING
            """, (record_id, created_at, str(data["name"])[:100], str(data["empId"])[:100],
                  str(data["role"])[:100], str(data.get("evaluatorName", ""))[:100],
                  str(data.get("evaluatorTitle", ""))[:100], str(data["quizTitle"])[:255], score,
                  str(data["status"])[:30], correct_count, wrong_count, json.dumps(answers, ensure_ascii=False),
                  group_key))
        else:
            conn.execute("""
                INSERT OR IGNORE INTO exam_records
                (id, created_at, name, emp_id, role, evaluator_name, evaluator_title, quiz_title, score, status, correct_count, wrong_count, answers_detail, group_key)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (record_id, created_at, str(data["name"])[:100], str(data["empId"])[:100],
                  str(data["role"])[:100], str(data.get("evaluatorName", ""))[:100],
                  str(data.get("evaluatorTitle", ""))[:100], str(data["quizTitle"])[:255], score,
                  str(data["status"])[:30], correct_count, wrong_count, json.dumps(answers, ensure_ascii=False),
                  group_key))
        return jsonify({"ok": True, "id": record_id})
    finally:
        conn.close()


@app.get("/api/records")
def api_list_records():
    denied = require_admin()
    if denied:
        return denied
    conn, _ = _db_conn()
    try:
        rows = conn.execute("SELECT * FROM exam_records ORDER BY created_at DESC").fetchall()
        return jsonify([_record_to_dict(r) for r in rows])
    finally:
        conn.close()


@app.delete("/api/records")
def api_clear_records():
    denied = require_admin()
    if denied:
        return denied
    conn, _ = _db_conn()
    try:
        conn.execute("DELETE FROM exam_records")
        return jsonify({"ok": True})
    finally:
        conn.close()


@app.get("/health")
def health():
    return jsonify({"ok": True, "service": "biochemical-training-exam"})


@app.get("/")
def index():
    return send_from_directory(STATIC_DIR, "index.html")


if __name__ == "__main__":
    # debug=False：正式上線請關閉 debug；如需開發除錯可暫時改 True
    app.run(host="0.0.0.0", port=5000, debug=False)
