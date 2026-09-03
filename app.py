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
UPLOAD_DIR = BASE_DIR / "uploads"
DATA_DIR = BASE_DIR / "data"
META_FILE = DATA_DIR / "slides_meta.json"
TMP_DIR = BASE_DIR / "tmp_convert"

for d in (SLIDES_DIR, UPLOAD_DIR, DATA_DIR, TMP_DIR):
    d.mkdir(parents=True, exist_ok=True)

# 如果 soffice 不在系統 PATH 中，可用環境變數指定完整路徑，例如：
#   set SOFFICE_PATH=C:\Program Files\LibreOffice\program\soffice.exe   (Windows)
#   export SOFFICE_PATH=/usr/bin/soffice                                 (Linux)
SOFFICE_BIN = os.environ.get("SOFFICE_PATH", "soffice")

ALLOWED_EXT = {".pptx", ".ppt"}
MAX_UPLOAD_MB = 80

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
    meta = load_meta()
    result = []
    for m in meta:
        result.append(
            {
                **m,
                "categoryLabel": CATEGORY_LABELS.get(m.get("category", ""), CATEGORY_LABELS[""]),
                "imageFolder": f"slides/{m['folder']}",
                "downloadUrl": f"/download/{m['id']}",
            }
        )
    # 內建排前面，上傳的依上傳時間新到舊排列
    result.sort(key=lambda x: (not x.get("isBuiltin"), x.get("dateAdded", "")), reverse=False)
    return jsonify(result)


@app.get("/download/<slide_id>")
def download_slide(slide_id):
    meta = load_meta()
    entry = next((m for m in meta if m["id"] == slide_id), None)
    if not entry:
        abort(404)

    if entry.get("isBuiltin"):
        return send_from_directory(
            STATIC_DIR, entry["filename"], as_attachment=True, download_name=entry["filename"]
        )

    for ext in (".pptx", ".ppt"):
        candidate = UPLOAD_DIR / f"{entry['id']}{ext}"
        if candidate.exists():
            return send_file(candidate, as_attachment=True, download_name=entry["filename"])

    abort(404)


@app.post("/api/slides/upload")
def api_upload_slide():
    if "file" not in request.files:
        return jsonify({"error": "未收到檔案"}), 400

    file = request.files["file"]
    category = request.form.get("category", "")
    if category not in CATEGORY_LABELS:
        category = ""

    original_name = file.filename or "untitled.pptx"
    ext = Path(original_name).suffix.lower()
    if ext not in ALLOWED_EXT:
        return jsonify({"error": "僅支援 .pptx / .ppt 檔案"}), 400

    slide_id = f"upload-{uuid.uuid4().hex[:10]}"
    saved_pptx_path = UPLOAD_DIR / f"{slide_id}{ext}"
    file.save(str(saved_pptx_path))

    out_folder = SLIDES_DIR / slide_id
    try:
        page_count = convert_pptx_to_images(saved_pptx_path, out_folder)
    except Exception as e:  # noqa: BLE001
        shutil.rmtree(out_folder, ignore_errors=True)
        saved_pptx_path.unlink(missing_ok=True)
        return jsonify({"error": f"轉檔失敗：{e}"}), 500

    entry = {
        "id": slide_id,
        "filename": original_name,
        "title": original_name,
        "desc": "使用者自行上傳之補充教材",
        "category": category,
        "folder": slide_id,
        "pageCount": page_count,
        "isBuiltin": False,
        "dateAdded": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    meta = load_meta()
    meta.append(entry)
    save_meta(meta)

    return jsonify(
        {
            **entry,
            "categoryLabel": CATEGORY_LABELS.get(category, CATEGORY_LABELS[""]),
            "imageFolder": f"slides/{slide_id}",
            "downloadUrl": f"/download/{slide_id}",
        }
    )


@app.delete("/api/slides/<slide_id>")
def api_delete_slide(slide_id):
    meta = load_meta()
    entry = next((m for m in meta if m["id"] == slide_id), None)
    if not entry:
        return jsonify({"error": "找不到此簡報"}), 404
    if entry.get("isBuiltin"):
        return jsonify({"error": "內建簡報無法刪除"}), 403

    shutil.rmtree(SLIDES_DIR / entry["folder"], ignore_errors=True)
    for ext in (".pptx", ".ppt"):
        (UPLOAD_DIR / f"{entry['id']}{ext}").unlink(missing_ok=True)

    meta = [m for m in meta if m["id"] != slide_id]
    save_meta(meta)
    return jsonify({"ok": True})


@app.get("/health")
def health():
    return jsonify({"ok": True, "service": "biochemical-training-exam"})


@app.get("/")
def index():
    return send_from_directory(STATIC_DIR, "index.html")


if __name__ == "__main__":
    # debug=False：正式上線請關閉 debug；如需開發除錯可暫時改 True
    app.run(host="0.0.0.0", port=5000, debug=False)
