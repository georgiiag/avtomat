import os
import time

from flask import Flask, render_template, jsonify, request, Response, redirect, url_for
from werkzeug.utils import secure_filename

app = Flask(__name__)

@app.after_request
def add_cors_headers(resp: Response):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Methods"] = "GET,POST,OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return resp

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/index.html')
def index_html():
    return render_template('index.html')

@app.route('/tp_constructor')
@app.route('/tp_constructor.html')
def tp_constructor():
    return render_template('tp_constructor.html')

@app.route('/recognizer')
@app.route('/recognizer.html')
def recognizer():
    return render_template('recognizer.html')

UPLOAD_DIR = os.path.join(app.root_path, "static", "shared_docs")
SNAPSHOT_DIR = os.path.join(app.root_path, "static", "shared_docs", "snapshots")
ATTACHMENTS_DIR = os.path.join(app.root_path, "static", "attachments")

def ensure_upload_dir():
    try:
        os.makedirs(UPLOAD_DIR, exist_ok=True)
    except Exception:
        pass

def ensure_snapshot_dir():
    try:
        os.makedirs(SNAPSHOT_DIR, exist_ok=True)
    except Exception:
        pass

def ensure_attachments_dir(subdir=None):
    try:
        os.makedirs(ATTACHMENTS_DIR, exist_ok=True)
        if subdir:
            os.makedirs(os.path.join(ATTACHMENTS_DIR, str(subdir)), exist_ok=True)
    except Exception:
        pass

def docx_to_simple_html(path: str) -> str:
    try:
        from docx import Document
    except Exception:
        return "<div style='color:#b13b2e'>python-docx не установлен</div>"
    try:
        doc = Document(path)
    except Exception as e:
        return f"<div style='color:#b13b2e'>Не удалось открыть DOCX: {e}</div>"

    parts = []
    for p in getattr(doc, "paragraphs", []):
        txt = (p.text or "").strip()
        if not txt:
            continue
        parts.append(f"<p>{txt}</p>")

    for t in getattr(doc, "tables", []):
        try:
            rows = []
            for r in t.rows:
                cells = []
                for c in r.cells:
                    cells.append(f"<td>{(c.text or '').strip()}</td>")
                rows.append("<tr>" + "".join(cells) + "</tr>")
            parts.append("<table>" + "".join(rows) + "</table>")
        except Exception:
            continue

    if not parts:
        return "<div style='color:#667085'>Документ пустой</div>"
    return "\n".join(parts)

@app.route('/share', methods=['GET', 'POST', 'OPTIONS'])
def share_doc():
    if request.method == "OPTIONS":
        return ("", 204)

    ensure_upload_dir()

    if request.method == "POST":
        if 'file' not in request.files:
            return jsonify({"ok": False, "error": "file is required"}), 400
        f = request.files['file']
        if not f or not f.filename:
            return jsonify({"ok": False, "error": "empty file"}), 400
        filename = secure_filename(f.filename)
        if not filename.lower().endswith(".docx"):
            return jsonify({"ok": False, "error": "only .docx supported"}), 400

        doc_id = f"doc-{int(__import__('time').time())}-{os.urandom(3).hex()}"
        save_name = f"{doc_id}__{filename}"
        path = os.path.join(UPLOAD_DIR, save_name)
        f.save(path)
        return redirect(url_for("view_shared_doc", doc_id=doc_id))

    files = []
    try:
        for name in os.listdir(UPLOAD_DIR):
            if "__" not in name or not name.lower().endswith(".docx"):
                continue
            doc_id, orig = name.split("__", 1)
            files.append((doc_id, orig, name))
        files.sort(key=lambda x: x[0], reverse=True)
    except Exception:
        files = []

    items = "\n".join([f"<li><a href='/share/{doc_id}'>{orig}</a></li>" for doc_id, orig, _ in files[:50]]) or "<li style='color:#667085'>Пока нет файлов</li>"

    return f"""<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Общий доступ — DOCX</title>
  <style>
    body {{ font-family: Arial, sans-serif; background:#f0f2f5; margin:0; color:#111; }}
    .bar {{ background:#2c3e50; color:#fff; padding:14px 18px; display:flex; justify-content:space-between; }}
    .wrap {{ max-width:1000px; margin:0 auto; padding:18px; }}
    .card {{ background:#fff; border-radius:12px; padding:16px; box-shadow:0 4px 6px rgba(0,0,0,0.05); margin-bottom:14px; }}
    .btn {{ border:0; background:#3498db; color:#fff; padding:10px 14px; border-radius:10px; cursor:pointer; font-weight:600; }}
    input[type=file] {{ padding:8px; border:1px solid #dfe6e9; border-radius:10px; background:#fff; }}
    ul {{ margin:0; padding-left:18px; }}
    a {{ color:#2c3e50; }}
  </style>
</head>
<body>
  <div class="bar">
    <div>Общий доступ: загрузка DOCX для просмотра</div>
    <div><a href="/tp_constructor" style="color:#fff; text-decoration:none; opacity:0.9;">Конструктор ТП</a></div>
  </div>
  <div class="wrap">
    <div class="card">
      <form method="post" enctype="multipart/form-data">
        <div style="font-weight:700; margin-bottom:8px;">Загрузить DOCX</div>
        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <input name="file" type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document">
          <button class="btn" type="submit">Загрузить</button>
        </div>
      </form>
    </div>
    <div class="card">
      <div style="font-weight:700; margin-bottom:8px;">Последние файлы</div>
      <ul>{items}</ul>
    </div>
  </div>
</body>
</html>"""

@app.route('/share/<doc_id>', methods=['GET', 'OPTIONS'])
def view_shared_doc(doc_id: str):
    if request.method == "OPTIONS":
        return ("", 204)
    ensure_upload_dir()
    target = None
    for name in os.listdir(UPLOAD_DIR):
        if name.startswith(f"{doc_id}__") and name.lower().endswith(".docx"):
            target = name
            break
    if not target:
        return "Файл не найден", 404

    path = os.path.join(UPLOAD_DIR, target)
    orig = target.split("__", 1)[1] if "__" in target else target
    html = docx_to_simple_html(path)
    download_url = f"/static/shared_docs/{target}"

    return f"""<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{orig}</title>
  <style>
    body {{ font-family: Arial, sans-serif; background:#f0f2f5; margin:0; color:#111; }}
    .bar {{ background:#2c3e50; color:#fff; padding:14px 18px; display:flex; justify-content:space-between; gap:12px; flex-wrap:wrap; }}
    .wrap {{ max-width:1100px; margin:0 auto; padding:18px; }}
    .card {{ background:#fff; border-radius:12px; padding:16px; box-shadow:0 4px 6px rgba(0,0,0,0.05); }}
    a {{ color:#fff; text-decoration:none; opacity:0.9; }}
    .doc a {{ color:#2c3e50; text-decoration:underline; }}
    table {{ border-collapse:collapse; width:100%; margin:10px 0; }}
    td, th {{ border:1px solid #e6eaee; padding:8px; vertical-align:top; }}
  </style>
</head>
<body>
  <div class="bar">
    <div>{orig}</div>
    <div style="display:flex; gap:12px; align-items:center; flex-wrap:wrap;">
      <a href="/share">Назад</a>
      <a href="{download_url}">Скачать DOCX</a>
    </div>
  </div>
  <div class="wrap">
    <div class="card doc">{html}</div>
  </div>
</body>
</html>"""

@app.route('/api/save_tp_html', methods=['POST', 'OPTIONS'])
def save_tp_html():
    if request.method == "OPTIONS":
        return ("", 204)

    ensure_snapshot_dir()

    if 'file' not in request.files:
        return jsonify({"ok": False, "error": "file is required"}), 400
    f = request.files['file']
    if not f or not f.filename:
        return jsonify({"ok": False, "error": "empty file"}), 400

    filename = secure_filename(f.filename)
    if not filename.lower().endswith(".html"):
        filename = filename + ".html"

    doc_id = f"tp-{int(__import__('time').time())}-{os.urandom(3).hex()}"
    save_name = f"{doc_id}__{filename}"
    path = os.path.join(SNAPSHOT_DIR, save_name)
    try:
        f.save(path)
    except Exception as e:
        return jsonify({"ok": False, "error": f"failed to save: {e}"}), 500

    url = f"/static/shared_docs/snapshots/{save_name}"
    return jsonify({"ok": True, "url": url, "id": doc_id, "name": filename})

@app.route('/api/upload_attachment', methods=['POST', 'OPTIONS'])
def upload_attachment():
    if request.method == "OPTIONS":
        return ("", 204)

    if 'file' not in request.files:
        return jsonify({"ok": False, "error": "file is required"}), 400
    f = request.files['file']
    if not f or not f.filename:
        return jsonify({"ok": False, "error": "empty file"}), 400

    raw_filename = str(f.filename or "")
    raw_ext = os.path.splitext(raw_filename)[1].lower().lstrip(".")

    filename = secure_filename(raw_filename)
    ext = os.path.splitext(filename)[1].lower().lstrip(".")

    mimetype = str(getattr(f, "mimetype", "") or "").lower().strip()
    mime_to_ext = {
        "image/jpeg": "jpg",
        "image/jpg": "jpg",
        "image/png": "png",
        "image/gif": "gif",
        "image/webp": "webp",
        "image/bmp": "bmp",
        "image/tiff": "tiff",
        "image/svg+xml": "svg",
        "image/heic": "heic",
        "image/heif": "heif",
        "image/avif": "avif",
    }

    if not ext:
        ext = raw_ext or mime_to_ext.get(mimetype, "")
        if ext:
            filename = (filename or "file") + "." + ext
        elif not filename:
            filename = "file.bin"
            ext = "bin"

    if not filename:
        filename = f"file.{ext or 'bin'}"
        ext = ext or "bin"

    blocked = {
        "html", "htm", "js", "mjs", "css",
        "php", "py", "ps1", "bat", "cmd", "vbs",
        "exe", "dll", "com", "scr", "msi", "jar",
    }
    if ext in blocked and not mimetype.startswith("image/"):
        return jsonify({"ok": False, "error": f"unsupported file type: .{ext}"}), 400

    subdir = time.strftime("%Y%m%d", time.localtime())
    ensure_attachments_dir(subdir)

    file_id = f"att-{int(time.time())}-{os.urandom(3).hex()}"
    save_name = f"{file_id}__{filename}"
    path = os.path.join(ATTACHMENTS_DIR, subdir, save_name)
    try:
        f.save(path)
    except Exception as e:
        return jsonify({"ok": False, "error": f"failed to save: {e}"}), 500

    try:
        size = os.path.getsize(path)
    except Exception:
        size = None

    url = f"/static/attachments/{subdir}/{save_name}"
    return jsonify({"ok": True, "url": url, "name": filename, "ext": ext, "size": size})

@app.route('/api/parse_bom_docx', methods=['POST'])
def parse_bom_docx():
    if 'file' not in request.files:
        return jsonify({"ok": False, "error": "file is required"}), 400

    f = request.files['file']
    if not f or not f.filename:
        return jsonify({"ok": False, "error": "empty file"}), 400

    filename = (f.filename or "").lower()
    if os.path.basename(filename).startswith("~$"):
        return jsonify({"ok": False, "error": "Это временный файл MS Word (~$...). Закройте документ в Word и выберите исходный .doc файл."}), 400
    is_docx = filename.endswith(".docx")
    is_doc = filename.endswith(".doc")
    if not (is_docx or is_doc):
        return jsonify({"ok": False, "error": "only .doc/.docx supported"}), 400

    try:
        from docx import Document
    except Exception:
        return jsonify({"ok": False, "error": "python-docx is not installed"}), 500

    if is_docx:
        try:
            doc = Document(f.stream)
        except Exception as e:
            return jsonify({"ok": False, "error": f"failed to read docx: {e}"}), 400
    else:
        import shutil
        import subprocess
        import tempfile

        tmp_dir = tempfile.mkdtemp(prefix="tp_ctor_doc_")
        try:
            in_name = secure_filename(os.path.basename(f.filename)) or "spec.doc"
            if not in_name.lower().endswith(".doc"):
                in_name = f"{in_name}.doc"
            in_path = os.path.join(tmp_dir, in_name)
            try:
                f.save(in_path)
            except Exception as e:
                return jsonify({"ok": False, "error": f"failed to save .doc: {e}"}), 500

            out_path = os.path.splitext(in_path)[0] + ".docx"

            soffice = shutil.which("soffice") or shutil.which("libreoffice")
            if not soffice:
                return jsonify({"ok": False, "error": "Нужен конвертер .doc→.docx (LibreOffice/soffice). Используйте DOCX."}), 400

            try:
                subprocess.run(
                    [soffice, "--headless", "--nologo", "--convert-to", "docx", "--outdir", tmp_dir, in_path],
                    check=True,
                    capture_output=True,
                    text=True,
                    timeout=60,
                )
            except subprocess.TimeoutExpired:
                return jsonify({"ok": False, "error": "DOC conversion timeout (soffice)"}), 500
            except Exception as e:
                return jsonify({"ok": False, "error": f"DOC conversion failed (soffice): {e}"}), 500

            if not os.path.exists(out_path):
                found = [x for x in os.listdir(tmp_dir) if x.lower().endswith(".docx")]
                if found:
                    out_path = os.path.join(tmp_dir, found[0])
                else:
                    return jsonify({"ok": False, "error": "DOC conversion produced no .docx"}), 500

            try:
                doc = Document(out_path)
            except Exception as e:
                return jsonify({"ok": False, "error": f"failed to read converted docx: {e}"}), 400
        finally:
            try:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            except Exception:
                pass

    synonyms = {
        "level": ["уровень", "lvl", "level"],
        "code": ["код", "обозначение", "позиция", "поз", "артикул", "code"],
        "name": ["наименование", "номенклатура", "деталь", "name"],
        "qty": ["кол-во", "количество", "qty", "count"],
        "unit": ["ед", "ед.изм", "ед. изм", "единица", "unit"],
    }

    def norm(s: str) -> str:
        return " ".join((s or "").lower().replace("\xa0", " ").split()).strip()

    def find_col_index(headers, keys):
        for i, h in enumerate(headers):
            for k in keys:
                if k in h:
                    return i
        return -1

    def cell_indent_pt(cell) -> float:
        try:
            for p in getattr(cell, "paragraphs", []):
                if not (p.text or "").strip():
                    continue
                ind = getattr(getattr(p, "paragraph_format", None), "left_indent", None)
                if ind is None:
                    continue
                pt = getattr(ind, "pt", None)
                if pt is None:
                    continue
                return float(pt)
        except Exception:
            return 0.0
        return 0.0

    def infer_levels_from_indents(rows, code_idx: int, name_idx: int):
        indents = []
        for r in rows:
            ind = max(r.get("_indent_code", 0.0), r.get("_indent_name", 0.0))
            if ind > 0:
                indents.append(round(ind * 2) / 2.0)
        uniq = sorted(set(indents))
        if len(uniq) <= 1:
            return False

        for r in rows:
            ind = max(r.get("_indent_code", 0.0), r.get("_indent_name", 0.0))
            ind = round(ind * 2) / 2.0
            nearest = min(uniq, key=lambda x: abs(x - ind))
            r["level"] = uniq.index(nearest)
        return True

    def infer_levels_from_designation(rows):
        def depth(code: str) -> int:
            c = (code or "").strip()
            if not c:
                return 0
            if c.count(".") >= 1:
                return len([x for x in c.split(".") if x != ""])
            return 0

        depths = [depth(r.get("code")) for r in rows]
        non_zero = [d for d in depths if d > 0]
        if not non_zero:
            for r in rows:
                r["level"] = 0
            return

        base = min(non_zero)
        for r, d in zip(rows, depths):
            r["level"] = max(0, d - base) if d > 0 else 0

    def extract_product_name(document):
        lines = []
        for p in getattr(document, "paragraphs", []):
            raw = (p.text or "").replace("\xa0", " ").strip()
            if not raw:
                continue
            n = norm(raw)
            if not n:
                continue
            lines.append((raw, n))

        if not lines:
            return ""

        for raw, n in lines:
            if "спецификац" in n:
                cleaned = raw
                lower = n
                idx = lower.find("спецификац")
                if idx >= 0:
                    cleaned = raw[idx:]
                cleaned = cleaned.replace("СПЕЦИФИКАЦИЯ", "").replace("Спецификация", "").replace("спецификация", "").strip(" :—-")
                if cleaned:
                    return cleaned

        return lines[0][0]

    def extract_product_designation(document):
        import re

        lines = []
        for p in getattr(document, "paragraphs", []):
            raw = (p.text or "").replace("\xa0", " ").strip()
            if not raw:
                continue
            n = norm(raw)
            if not n:
                continue
            lines.append((raw, n))

        if not lines:
            return ""

        def after_sep(raw_line: str):
            for sep in (":", "—", "-"):
                if sep in raw_line:
                    right = raw_line.split(sep, 1)[1].strip()
                    if right:
                        return right
            return ""

        for raw, n in lines:
            if "обознач" in n or "шифр" in n or "код издел" in n or "обозн." in n:
                val = after_sep(raw)
                if val:
                    return val

        code_re = re.compile(r"([A-Za-zА-Яа-я0-9]{2,}(?:[./\\-][A-Za-zА-Яа-я0-9]{1,}){1,}[A-Za-zА-Яа-я0-9./\\-]{0,})")
        best = ""
        for raw, n in lines[:12]:
            if "спецификац" in n:
                continue
            for m in code_re.findall(raw):
                cand = (m or "").strip().strip(",.;")
                if len(cand) >= 4 and len(cand) > len(best):
                    best = cand
        return best

    best = None
    best_score = -1
    best_rows = None
    best_map = None
    best_prefix_text = ""

    for table in doc.tables:
        matrix_norm = []
        matrix_raw = []
        matrix_indent = []
        for row in table.rows:
            row_norm = []
            row_raw = []
            row_indent = []
            for cell in row.cells:
                raw_text = (cell.text or "").replace("\xa0", " ")
                row_raw.append(raw_text.strip())
                row_norm.append(norm(raw_text))
                row_indent.append(cell_indent_pt(cell))
            matrix_norm.append(row_norm)
            matrix_raw.append(row_raw)
            matrix_indent.append(row_indent)

        if not matrix_norm:
            continue

        header_row_index = 0
        header_map = None
        score = -1

        for candidate in range(min(3, len(matrix_norm))):
            headers = matrix_norm[candidate]
            m = {
                "level": find_col_index(headers, synonyms["level"]),
                "code": find_col_index(headers, synonyms["code"]),
                "name": find_col_index(headers, synonyms["name"]),
                "qty": find_col_index(headers, synonyms["qty"]),
                "unit": find_col_index(headers, synonyms["unit"]),
            }

            found = sum(1 for v in m.values() if v >= 0)
            if m["code"] >= 0:
                found += 2
            if m["name"] >= 0:
                found += 2

            if found > score:
                score = found
                header_row_index = candidate
                header_map = m

        if not header_map:
            continue

        parsed_rows = []
        for r_raw, r_indent in zip(matrix_raw[header_row_index + 1:], matrix_indent[header_row_index + 1:]):
            def get(idx):
                return r_raw[idx] if idx is not None and idx >= 0 and idx < len(r_raw) else ""

            def get_indent(idx):
                return float(r_indent[idx]) if idx is not None and idx >= 0 and idx < len(r_indent) else 0.0

            code = get(header_map["code"]).strip()
            name = get(header_map["name"]).strip()
            qty = get(header_map["qty"]).strip()
            unit = get(header_map["unit"]).strip()
            level_raw = get(header_map["level"]).strip()

            if not code and not name:
                continue

            level = 0
            if header_map["level"] is not None and header_map["level"] >= 0 and level_raw and level_raw.isdigit():
                level = int(level_raw)

            qty_num = None
            if qty:
                try:
                    qty_num = float(qty.replace(",", "."))
                except Exception:
                    qty_num = None

            parsed_rows.append({
                "level": level,
                "code": code,
                "name": name,
                "qty": qty_num if qty_num is not None else qty,
                "unit": unit,
                "_indent_code": get_indent(header_map["code"]),
                "_indent_name": get_indent(header_map["name"]),
            })

        final_score = score * 10 + len(parsed_rows)
        if final_score > best_score and parsed_rows:
            best_score = final_score
            best = table
            best_rows = parsed_rows
            best_map = header_map
            prefix_cells = []
            for rr in matrix_raw[:header_row_index]:
                for c in rr:
                    cc = (c or "").strip()
                    if cc:
                        prefix_cells.append(cc)
            best_prefix_text = " ".join(prefix_cells).strip()

    if not best_rows:
        return jsonify({"ok": False, "error": "spec table not found"}), 404

    if best_map and (best_map.get("level", -1) is None or best_map.get("level", -1) < 0):
        inferred = infer_levels_from_indents(best_rows, best_map.get("code", -1), best_map.get("name", -1))
        if not inferred:
            infer_levels_from_designation(best_rows)

    for r in best_rows:
        r.pop("_indent_code", None)
        r.pop("_indent_name", None)

    product_name = extract_product_name(doc)
    product_designation = extract_product_designation(doc)

    if not product_designation and best_prefix_text:
        import re

        raw = best_prefix_text
        n = norm(raw)

        def after_sep(raw_line: str):
            for sep in (":", "—", "-"):
                if sep in raw_line:
                    right = raw_line.split(sep, 1)[1].strip()
                    if right:
                        return right
            return ""

        if "обознач" in n or "шифр" in n or "код издел" in n or "обозн." in n:
            val = after_sep(raw)
            if val:
                product_designation = val
        if not product_designation:
            code_re = re.compile(r"([A-Za-zА-Яа-я0-9]{2,}(?:[./\\-][A-Za-zА-Яа-я0-9]{1,}){1,}[A-Za-zА-Яа-я0-9./\\-]{0,})")
            best = ""
            for m in code_re.findall(raw):
                cand = (m or "").strip().strip(",.;")
                if len(cand) >= 4 and len(cand) > len(best):
                    best = cand
            product_designation = best

    if not product_name and best_prefix_text:
        raw = best_prefix_text
        cleaned = raw
        cleaned_norm = norm(raw)
        if "спецификац" in cleaned_norm:
            idx = cleaned_norm.find("спецификац")
            if idx >= 0:
                cleaned = raw[idx:]
        cleaned = cleaned.replace("СПЕЦИФИКАЦИЯ", "").replace("Спецификация", "").replace("спецификация", "").strip(" :—-")
        if cleaned:
            product_name = cleaned

    if not product_designation and best_rows:
        product_designation = (best_rows[0].get("code") or "").strip()
    if not product_name and best_rows:
        product_name = (best_rows[0].get("name") or "").strip()
    return jsonify({"ok": True, "rows": best_rows, "columns": best_map, "product_name": product_name, "product_designation": product_designation})

@app.route('/api/status')
def get_status():
    # Базовая заглушка для статуса автомата
    return jsonify({
        "status": "online",
        "message": "Система готова к работе",
        "process": "Ожидание команды"
    })

if __name__ == '__main__':
    host = os.getenv("HOST", "0.0.0.0")
    preferred_port = int(os.getenv("PORT", "5000"))
    candidate_ports = [preferred_port, 8000, 8080, 5000]

    last_error = None
    for port in candidate_ports:
        try:
            app.run(debug=True, host=host, port=port, use_reloader=False)
            break
        except OSError as e:
            last_error = e
            continue

    if last_error:
        raise last_error
