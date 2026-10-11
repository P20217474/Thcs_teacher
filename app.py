import subprocess
import sys
import os
import json
import base64
import urllib.request
from datetime import datetime, timedelta


def _ensure(pkg, mod=None):
    mod = mod or pkg
    try:
        __import__(mod)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "--quiet"])


_ensure("streamlit")
_ensure("python-docx", "docx")
_ensure("pypdf", "pypdf")

import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import csv

try:
    from pypdf import PdfReader
    HAS_PYPDF = True
except Exception:
    HAS_PYPDF = False

try:
    from PIL import Image
    HAS_PIL = True
except Exception:
    HAS_PIL = False

try:
    import google.generativeai as genai_vision
    HAS_GEMINI_VISION = True
except Exception:
    HAS_GEMINI_VISION = False

st.set_page_config(page_title="Trợ lý Giáo viên THCS", page_icon="🧠", layout="wide")
st.title("🧠 Trợ lý Giáo viên THCS")
st.caption("Bộ não tự động • Bám mẫu giáo án • Lịch báo giảng • Xuất Word")


def cfg(key, default=""):
    try:
        val = st.secrets.get(key, default)
        return val if val else default
    except Exception:
        return default


PROVIDER_KEYS = {
    "GEMINI_API_KEY": cfg("GEMINI_API_KEY", ""),
    "GROQ_API_KEY": cfg("GROQ_API_KEY", ""),
    "OPENROUTER_API_KEY": cfg("OPENROUTER_API_KEY", ""),
    "NVIDIA_API_KEY": cfg("NVIDIA_API_KEY", ""),
}
for _k, _v in PROVIDER_KEYS.items():
    if _v:
        os.environ[_k] = _v
available_providers = [_k for _k, _v in PROVIDER_KEYS.items() if _v]
GEMINI_KEY = PROVIDER_KEYS.get("GEMINI_API_KEY", "")

PREFERRED_MODELS = {
    "GEMINI_API_KEY": ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-flash-latest"],
    "GROQ_API_KEY": ["llama-3.3-70b-versatile", "llama-3.1-70b-versatile"],
    "OPENROUTER_API_KEY": ["google/gemini-2.0-flash-exp:free"],
    "NVIDIA_API_KEY": ["meta/llama-3.3-70b-instruct"],
}

GITHUB_TOKEN = cfg("GITHUB_TOKEN", "")
GITHUB_REPO = cfg("GITHUB_REPO", "P20217474/Thcs_teacher")
GITHUB_BRANCH = cfg("GITHUB_BRANCH", "main")


# ============ GITHUB API ============
def gh_api(method, path, data=None):
    if not GITHUB_TOKEN:
        return {"error": "Chưa có GITHUB_TOKEN"}
    url = "https://api.github.com" + path
    headers = {
        "Authorization": "token " + GITHUB_TOKEN,
        "Accept": "application/vnd.github+json",
        "User-Agent": "agent-thcs",
    }
    body = json.dumps(data).encode("utf-8") if data else None
    req = urllib.request.Request(url, data=body, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        return {"error": str(e)}


def gh_get(path):
    res = gh_api("GET", f"/repos/{GITHUB_REPO}/contents/{path}?ref={GITHUB_BRANCH}")
    if not res or "error" in res or "content" not in res:
        return (None, None)
    try:
        return (base64.b64decode(res["content"]).decode("utf-8"), res["sha"])
    except Exception:
        return (None, None)


def gh_save(path, content, message="Update"):
    if not GITHUB_TOKEN:
        return (False, "Chưa có GITHUB_TOKEN")
    old, sha = gh_get(path)
    data = {
        "message": message,
        "content": base64.b64encode(content.encode("utf-8")).decode("ascii"),
        "branch": GITHUB_BRANCH,
    }
    if sha:
        data["sha"] = sha
    res = gh_api("PUT", f"/repos/{GITHUB_REPO}/contents/{path}", data)
    if res is None:
        return (False, "Không gọi được API")
    if isinstance(res, dict) and "error" in res:
        return (False, "Lỗi: " + str(res["error"]))
    if isinstance(res, dict) and "content" in res:
        return (True, "Đã lưu: " + path)
    return (False, "Phản hồi không hợp lệ")


def gh_delete(path):
    if not GITHUB_TOKEN:
        return (False, "Chưa có GITHUB_TOKEN")
    old, sha = gh_get(path)
    if not sha:
        return (False, "File không tồn tại")
    res = gh_api(
        "DELETE",
        f"/repos/{GITHUB_REPO}/contents/{path}",
        {"message": "Delete " + path, "sha": sha, "branch": GITHUB_BRANCH},
    )
    if isinstance(res, dict) and "error" in res:
        return (False, "Lỗi: " + str(res["error"]))
    if isinstance(res, dict) and ("content" in res or "commit" in res):
        return (True, "Đã xóa: " + path)
    return (False, "Phản hồi không hợp lệ")


def gh_list(folder):
    res = gh_api("GET", f"/repos/{GITHUB_REPO}/contents/{folder}?ref={GITHUB_BRANCH}")
    if not res or "error" in res or not isinstance(res, list):
        return []
    return [item["name"] for item in res if item["type"] == "file"]


def gh_test():
    if not GITHUB_TOKEN:
        return (False, "Thiếu GITHUB_TOKEN")
    res = gh_api("GET", "/user")
    if res is None:
        return (False, "Không gọi được API")
    if isinstance(res, dict) and "error" in res:
        return (False, "Lỗi: " + res["error"])
    if isinstance(res, dict) and "login" in res:
        login = res["login"]
        res2 = gh_api("GET", f"/repos/{GITHUB_REPO}")
        if isinstance(res2, dict) and "permissions" in res2:
            perms = res2["permissions"]
            if perms.get("push") or perms.get("admin"):
                return (True, "OK — " + login + " | Có quyền ghi")
            return (False, "Token thiếu quyền 'repo'")
        return (True, "OK — " + login)
    return (False, "Token không hợp lệ")


# ============ BỘ NHỚ ============
def load_memory():
    content, _ = gh_get("data/memory.json")
    if content:
        try:
            return json.loads(content)
        except Exception:
            pass
    return {"teacher_name": "", "school": "", "default_subject": "Toan", "default_grade": "6"}


def save_memory(memory):
    memory["last_updated"] = datetime.now().strftime("%Y-%m-%d")
    return gh_save("data/memory.json", json.dumps(memory, ensure_ascii=False, indent=2), "Update memory")


def load_folder(folder):
    files = gh_list(f"data/{folder}")
    out = {}
    for f in files:
        c, _ = gh_get(f"data/{folder}/{f}")
        if c:
            out[f] = c
    return out


def load_templates():
    return load_folder("templates")


def load_sources():
    return load_folder("sources")


def load_ppct():
    return load_folder("ppct")


def load_tkb():
    return load_folder("tkb")


def save_history(task, topic, content):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe = re.sub(r"[^a-zA-Z0-9]", "_", topic)[:30]
    fname = f"data/history/{ts}_{safe}.md"
    md = f"# {task}\n\n**Chủ đề:** {topic}\n\n---\n\n{content}"
    return gh_save(fname, md, f"History {topic}")


# ============ OCR ẢNH ============
def ocr_image_with_gemini(data, fname):
    if not GEMINI_KEY:
        return "[" + fname + ": cần GEMINI_API_KEY]"
    if not HAS_GEMINI_VISION:
        return "[" + fname + ": chưa cài google-generativeai]"
    if not HAS_PIL:
        return "[" + fname + ": chưa cài pillow]"
    try:
        genai_vision.configure(api_key=GEMINI_KEY)
        img = Image.open(io.BytesIO(data))
        models_to_try = ["gemini-2.0-flash", "gemini-2.5-flash", "gemini-flash-latest"]
        last_err = None
        for model_name in models_to_try:
            try:
                model = genai_vision.GenerativeModel(model_name)
                prompt = (
                    "Đây là ảnh trang sách giáo khoa. Hãy đọc TOÀN BỘ chữ trong ảnh "
                    "và trả về chính xác nội dung văn bản. Giữ nguyên cấu trúc: tiêu đề, "
                    "số thứ tự, đề mục, công thức toán. Với toán, dùng Unicode: x², √x, ≤, π. "
                    "Chỉ trả về nội dung text, không thêm mô tả."
                )
                response = model.generate_content([prompt, img])
                text = response.text.strip()
                if len(text) < 20:
                    return "[" + fname + ": AI không đọc được text]"
                return text
            except Exception as e:
                last_err = e
                continue
        return "[" + fname + ": Lỗi OCR — " + str(last_err) + "]"
    except Exception as e:
        return "[" + fname + ": Lỗi OCR — " + str(e) + "]"


# ============ ĐỌC FILE ============
def _read_docx_safe(data):
    try:
        d = Document(io.BytesIO(data))
        out = []
        for p in d.paragraphs:
            if p.text.strip():
                out.append(p.text)
        for t in d.tables:
            out.append("")
            for row in t.rows:
                cells = [c.text.strip() for c in row.cells]
                out.append("| " + " | ".join(cells) + " |")
            out.append("")
        return "\n".join(out)
    except Exception as e:
        return "[Lỗi đọc docx: " + str(e) + "]"


def _read_doc_old(data):
    return "[File .doc cũ. Mở bằng Word/WPS → Save As .docx → upload lại.]"


def _read_pdf_safe(data):
    if not HAS_PYPDF:
        return "[Chưa cài pypdf]"
    try:
        reader = PdfReader(io.BytesIO(data))
        if len(reader.pages) == 0:
            return "[PDF rỗng]"
        texts = []
        for page in reader.pages:
            try:
                t = ""
                try:
                    t = page.extract_text(extraction_mode="layout") or ""
                except TypeError:
                    t = page.extract_text() or ""
                if t.strip():
                    texts.append(t)
            except Exception:
                continue
        full = "\n".join(texts)
        if len(full.strip()) < 50:
            return "[PDF scan, không có text. Dùng OCR ảnh.]\n" + full
        return full
    except Exception as e:
        return "[Lỗi đọc pdf: " + str(e) + "]"


def _read_excel_safe(data):
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        out = []
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            rows = []
            for row in ws.iter_rows(values_only=True):
                rows.append([str(c) if c is not None else "" for c in row])
            if not rows:
                continue
            rows = [r for r in rows if any(c.strip() for c in r)]
            if not rows:
                continue
            max_cols = max(len(r) for r in rows)
            rows = [r + [""] * (max_cols - len(r)) for r in rows]
            out.append("### Sheet: " + sheet_name)
            out.append("| " + " | ".join(rows[0]) + " |")
            out.append("|" + "---|" * max_cols)
            for r in rows[1:]:
                out.append("| " + " | ".join(r) + " |")
            out.append("")
        return "\n".join(out)
    except Exception as e:
        return "[Lỗi đọc Excel: " + str(e) + "]"


def _read_csv_safe(data):
    try:
        text = None
        for enc in ["utf-8", "utf-8-sig", "latin-1", "cp1258"]:
            try:
                text = data.decode(enc)
                break
            except Exception:
                continue
        if text is None:
            text = data.decode("utf-8", errors="ignore")
        sample = text[:2000]
        delimiter = ","
        if sample.count("|") > sample.count(","):
            delimiter = "|"
        elif sample.count(";") > sample.count(","):
            delimiter = ";"
        elif sample.count("\t") > sample.count(","):
            delimiter = "\t"
        rows = list(csv.reader(io.StringIO(text), delimiter=delimiter))
        if not rows:
            return "[CSV rỗng]"
        max_cols = max(len(r) for r in rows)
        rows = [r + [""] * (max_cols - len(r)) for r in rows]
        out = ["| " + " | ".join(rows[0]) + " |", "|" + "---|" * max_cols]
        for row in rows[1:]:
            out.append("| " + " | ".join(row) + " |")
        return "\n".join(out)
    except Exception as e:
        return "[Lỗi đọc csv: " + str(e) + "]"


def _read_txt_safe(data):
    for enc in ["utf-8", "utf-8-sig", "latin-1", "cp1258"]:
        try:
            return data.decode(enc)
        except Exception:
            continue
    return data.decode("utf-8", errors="ignore")


def read_file(file, use_ocr=True):
    if file is None:
        return ("", "", 0)
    try:
        data = file.read()
    except Exception as e:
        return (file.name, "[Lỗi: " + str(e) + "]", 0)
    if not data:
        return (file.name, "[File rỗng]", 0)

    name = file.name
    nl = name.lower()
    magic = data[:8]
    text = ""

    if nl.endswith((".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp")):
        if use_ocr:
            text = ocr_image_with_gemini(data, name)
        else:
            text = "[" + name + ": ảnh]"
        return (name, text, len(text))

    if magic[:4] == b"%PDF":
        text = _read_pdf_safe(data)
    elif magic[:4] == b"PK\x03\x04":
        if nl.endswith((".xlsx", ".xlsm")):
            text = _read_excel_safe(data)
        else:
            text = _read_docx_safe(data)
    elif magic[:4] == b"\xd0\xcf\x11\xe0":
        text = _read_doc_old(data)
    else:
        if nl.endswith(".docx"):
            text = _read_docx_safe(data)
        elif nl.endswith(".doc"):
            text = _read_doc_old(data)
        elif nl.endswith(".pdf"):
            text = _read_pdf_safe(data)
        elif nl.endswith((".xlsx", ".xls", ".xlsm")):
            text = _read_excel_safe(data)
        elif nl.endswith(".csv"):
            text = _read_csv_safe(data)
        elif nl.endswith((".txt", ".md", ".log", ".json")):
            text = _read_txt_safe(data)
        else:
            try:
                decoded = _read_txt_safe(data)
                printable = sum(1 for c in decoded[:500] if c.isprintable() or c in "\n\r\t")
                text = decoded if printable / max(len(decoded[:500]), 1) > 0.8 else "[Định dạng không hỗ trợ]"
            except Exception:
                text = "[Không đọc được file]"
    return (name, text, len(text))


# ============ SESSION STATE ============
if "memory" not in st.session_state:
    st.session_state["memory"] = load_memory()
if "templates" not in st.session_state:
    st.session_state["templates"] = load_templates()
if "sources" not in st.session_state:
    st.session_state["sources"] = load_sources()
if "ppct" not in st.session_state:
    st.session_state["ppct"] = load_ppct()
if "tkb" not in st.session_state:
    st.session_state["tkb"] = load_tkb()
if "draft" not in st.session_state:
    st.session_state["draft"] = ""
if "last_topic" not in st.session_state:
    st.session_state["last_topic"] = ""
if "weekly_results" not in st.session_state:
    st.session_state["weekly_results"] = []
if "lich_bao_giang" not in st.session_state:
    st.session_state["lich_bao_giang"] = []
if "start_date_val" not in st.session_state:
    st.session_state["start_date_val"] = datetime.now().date()


# ============ SIDEBAR ============
with st.sidebar:
    st.header("🧠 Bộ não")
    mem = st.session_state["memory"]

    if GITHUB_TOKEN:
        st.success("💾 Đã có GITHUB_TOKEN")
        st.caption("Repo: " + GITHUB_REPO)
        if st.button("🧪 Test GitHub", use_container_width=True):
            ok, msg = gh_test()
            if ok:
                st.success("✅ " + msg)
            else:
                st.error("❌ " + msg)
    else:
        st.error("❌ Chưa có GITHUB_TOKEN")

    if GEMINI_KEY:
        st.success("🔍 Gemini Vision sẵn sàng")
    else:
        st.warning("⚠️ Chưa có GEMINI_API_KEY")

    mem["teacher_name"] = st.text_input("Giáo viên", value=mem.get("teacher_name", ""))
    mem["school"] = st.text_input("Trường", value=mem.get("school", ""))
    mem["default_subject"] = st.text_input("Môn mặc định", value=mem.get("default_subject", "Toan"))
    mem["default_grade"] = st.text_input("Khối mặc định", value=mem.get("default_grade", "6"))

    if st.button("💾 Lưu vào bộ nhớ", use_container_width=True):
        ok, msg = save_memory(mem)
        if ok:
            st.success("✅ " + msg)
        else:
            st.error("❌ " + msg)

    st.divider()
    st.write("Mẫu: " + str(len(st.session_state["templates"])) + " file")
    st.write("Nguồn: " + str(len(st.session_state["sources"])) + " file")
    st.write("PPCT: " + str(len(st.session_state["ppct"])) + " file")
    st.write("TKB: " + str(len(st.session_state["tkb"])) + " file")
    if st.button("🔄 Tải lại", use_container_width=True):
        st.session_state["templates"] = load_templates()
        st.session_state["sources"] = load_sources()
        st.session_state["ppct"] = load_ppct()
        st.session_state["tkb"] = load_tkb()
        st.rerun()
    st.write("AI: " + str(len(available_providers)) + " nhà cung cấp")


# ============ GỌI AI ============
def call_ai(prompt, system=""):
    import freelm
    llm = freelm.FreeLLM.from_env()
    msgs = []
    if system:
        msgs.append({"role": "system", "content": system})
    msgs.append({"role": "user", "content": prompt})
    last_err = None
    for pk in available_providers:
        for mn in PREFERRED_MODELS.get(pk, []):
            try:
                resp = llm.chat(msgs, model=mn)
                text = getattr(resp, "text", None) or getattr(resp, "content", "")
                if text and len(text.strip()) > 50:
                    return text
            except Exception as e:
                last_err = e
                continue
    try:
        resp = llm.chat(msgs, model="auto")
        return getattr(resp, "text", None) or getattr(resp, "content", "")
    except Exception as e:
        if last_err:
            raise last_err
        raise e


# ============ UPLOAD ============
tab_up, tab_kho = st.tabs(["📤 Upload mẫu/nguồn/PPCT/TKB", "📂 Kho tri thức"])

with tab_up:
    st.info("💡 MẪU = định dạng bài soạn. NGUỒN = nội dung SGK. PPCT/TKB = lịch dạy. Ảnh sẽ được Gemini Vision đọc.")
    up_type = st.radio(
        "Loại",
        ["Mẫu (định dạng)", "Nguồn (nội dung)", "PPCT", "TKB"],
        horizontal=True,
    )
    up_name = st.text_input("Tên file", placeholder="mau_bai_soan")
    up_files = st.file_uploader(
        "Chọn file",
        type=["docx", "pdf", "txt", "md", "csv", "xlsx", "xls", "doc",
              "png", "jpg", "jpeg", "webp", "bmp"],
        accept_multiple_files=True, key="up_files",
    )
    up_paste = st.text_area("Hoặc dán nội dung", height=120, key="up_paste")

    if up_files:
        st.markdown("**📋 Trạng thái:**")
        for f in up_files:
            nl = f.name.lower()
            if nl.endswith((".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif")):
                st.write("🖼️ " + f.name + " — ảnh, sẽ OCR")
            elif nl.endswith(".pdf"):
                st.write("📄 " + f.name + " — PDF")
            else:
                st.write("📎 " + f.name)

    if st.button("💾 Lưu lên GitHub", type="primary", use_container_width=True):
        ok0, msg0 = gh_test()
        if not ok0:
            st.error("❌ " + msg0)
        else:
            st.success("✅ " + msg0)
            parts = []
            if up_paste.strip():
                parts.append(up_paste.strip())
            for f in (up_files or []):
                nl = f.name.lower()
                if nl.endswith((".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif")):
                    with st.spinner("🔍 AI đọc ảnh " + f.name + "..."):
                        data = f.read()
                        txt = ocr_image_with_gemini(data, f.name)
                    parts.append("--- " + f.name + " (OCR) ---\n" + txt)
                    st.info("📖 " + f.name + ": " + str(len(txt)) + " ký tự")
                else:
                    _, txt, _ = read_file(f, use_ocr=False)
                    parts.append("--- " + f.name + " ---\n" + txt)

            if not parts:
                st.warning("Không có nội dung.")
            else:
                full = "\n\n".join(parts)
                safe = re.sub(r"[^a-zA-Z0-9_\-]", "_", up_name.strip())[:40]
                if not safe:
                    safe = "file_" + datetime.now().strftime("%Y%m%d_%H%M%S")
                if "Mẫu" in up_type:
                    folder = "templates"
                elif "PPCT" in up_type:
                    folder = "ppct"
                elif "TKB" in up_type:
                    folder = "tkb"
                else:
                    folder = "sources"
                path = f"data/{folder}/{safe}.md"
                with st.spinner("💾 Đang lưu..."):
                    ok, msg = gh_save(path, full, f"Upload {safe}")
                if ok:
                    st.success("✅ " + msg)
                    st.session_state["templates"] = load_templates()
                    st.session_state["sources"] = load_sources()
                    st.session_state["ppct"] = load_ppct()
                    st.session_state["tkb"] = load_tkb()
                    st.rerun()
                else:
                    st.error("❌ " + msg)

with tab_kho:
    for label, key, folder in [
        ("📐 Mẫu", "templates", "templates"),
        ("📚 Nguồn", "sources", "sources"),
        ("📗 PPCT", "ppct", "ppct"),
        ("📘 TKB", "tkb", "tkb"),
    ]:
        st.subheader(label + " đã lưu")
        data_dict = st.session_state[key]
        if data_dict:
            for name in list(data_dict.keys()):
                col1, col2 = st.columns([5, 1])
                with col1:
                    with st.expander("📄 " + name):
                        content = data_dict[name]
                        st.caption("Kích thước: " + str(len(content)) + " ký tự")
                        st.text(content[:2000])
                with col2:
                    if st.button("🗑️", key=f"del_{key}_{name}"):
                        with st.spinner("Đang xóa..."):
                            ok, msg = gh_delete(f"data/{folder}/" + name)
                        if ok:
                            st.session_state[key] = load_folder(folder)
                            st.rerun()
            if st.button("🗑️ Xóa TẤT CẢ " + label, key=f"del_all_{key}"):
                for name in list(data_dict.keys()):
                    gh_delete(f"data/{folder}/" + name)
                st.session_state[key] = load_folder(folder)
                st.rerun()
        else:
            st.info("Chưa có " + label.lower() + ".")
        st.divider()


# ============ PROMPT BÁM MẪU GIÁO ÁN ============
def build_exec_prompt(plan):
    mem = st.session_state["memory"]
    tpl_text = ""
    for f in plan.get("template_files", []):
        if f in st.session_state["templates"]:
            tpl_text += "\n=== FILE MẪU (BẮT BUỘC BÁM SÁT) ===\n" + st.session_state["templates"][f][:20000]
    src_text = ""
    for f in plan.get("source_files", []):
        if f in st.session_state["sources"]:
            src_text += "\n=== NGUỒN NỘI DUNG (SGK) ===\n" + st.session_state["sources"][f][:25000]

    prompt = "Bạn là trợ lý chuyên môn cao cấp cho giáo viên THCS Việt Nam.\n\n"
    prompt += "NHIỆM VỤ: Soạn KẾ HOẠCH BÀI DẠY theo ĐÚNG mẫu giáo án.\n\n"
    prompt += "THÔNG TIN:\n"
    prompt += "- Giáo viên: " + mem.get("teacher_name", "") + "\n"
    prompt += "- Trường: " + mem.get("school", "") + "\n"
    prompt += "- Môn: " + plan.get("subject", "") + " | Khối: " + plan.get("grade", "") + "\n"
    prompt += "- Tên bài: " + plan.get("topic", "") + "\n"
    prompt += "- Số tiết: " + (plan.get("period") or "1") + "\n\n"

    if tpl_text:
        prompt += tpl_text + "\n"
    if src_text:
        prompt += src_text + "\n"

    prompt += "=== CẤU TRÚC BẮT BUỘC CỦA BÀI SOẠN ===\n\n"

    prompt += "**I. MỤC TIÊU**\n"
    prompt += "1. Kiến thức: [nêu kiến thức trọng tâm từ SGK]\n"
    prompt += "2. Năng lực:\n"
    prompt += "   - Năng lực riêng: [2-3 năng lực đặc thù]\n"
    prompt += "   - Năng lực chung: Năng lực tư duy và lập luận toán học; năng lực giao tiếp toán học tự học; năng lực giải quyết vấn đề toán học, năng lực tư duy sáng tạo, năng lực hợp tác.\n"
    prompt += "3. Phẩm chất: Bồi dưỡng hứng thú học tập, ý thức làm việc nhóm, ý thức tìm tòi, khám phá và sáng tạo cho HS => độc lập, tự tin và tự chủ.\n\n"

    prompt += "**II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU**\n"
    prompt += "1. GV: Bài giảng, khbd.\n"
    prompt += "   (KHÔNG dùng từ 'máy chiếu'. Nếu cần thiết bị hiển thị, dùng 'tivi' hoặc 'máy vi tính'.)\n"
    prompt += "2. HS: SGK; Đồ dùng học tập; Ôn tập khái niệm về [kiến thức liên quan đã học].\n\n"

    prompt += "**III. TIẾN TRÌNH DẠY HỌC**\n\n"

    prompt += "**A. HOẠT ĐỘNG KHỞI ĐỘNG (MỞ ĐẦU)**\n"
    prompt += "a) Mục tiêu: [2 gạch đầu dòng: gây hứng thú + gợi mở nội dung]\n"
    prompt += "b) Nội dung: HS chú ý lắng nghe và thực hiện yêu cầu.\n"
    prompt += "c) Tổ chức thực hiện và sản phẩm:\n"
    prompt += "- Bước 1: Chuyển giao nhiệm vụ: [GV đặt vấn đề bằng bài toán mở đầu SGK]\n"
    prompt += "- Bước 2: Thực hiện nhiệm vụ: HS đọc, suy nghĩ, thảo luận nhóm và suy đoán, giải thích.\n"
    prompt += "- Bước 3: Báo cáo, thảo luận: GV gọi một số HS trả lời, HS khác nhận xét.\n"
    prompt += "- Bước 4: Kết luận, nhận định: GV đánh giá kết quả của HS, dẫn dắt vào bài mới.\n\n"

    prompt += "**B. HÌNH THÀNH KIẾN THỨC MỚI**\n\n"
    prompt += "Với MỖI Hoạt động (Hoạt động 1, 2, 3...), trình bày ĐẦY ĐỦ:\n\n"
    prompt += "**Hoạt động N: [Tên hoạt động]**\n"
    prompt += "a) Mục tiêu: [3-4 gạch đầu dòng cụ thể]\n"
    prompt += "b) Nội dung: HS quan sát SGK để tìm hiểu nội dung kiến thức theo yêu cầu của GV.\n"
    prompt += "c) Tổ chức thực hiện và sản phẩm:\n\n"

    prompt += "**BẮT BUỘC dùng BẢNG 2 CỘT:**\n"
    prompt += "| HOẠT ĐỘNG CỦA GV VÀ HS | SẢN PHẨM |\n"
    prompt += "|---|---|\n"
    prompt += "| - Bước 1: Chuyển giao nhiệm vụ:<br>+ GV hướng dẫn...<br>+ GV phân tích... | **N. [Tên kiến thức]**<br>* [Nội dung cụ thể]<br>+ Ví dụ 1: [chi tiết]<br>+ Ư(24) = {1; 2; 3; 4; 6; 8; 12; 24}<br>+ Ư(28) = {1; 2; 4; 7; 14; 28}<br>+ ƯC(24; 28) = {1; 2; 4} |\n"
    prompt += "| - Bước 2: Thực hiện nhiệm vụ:<br>+ HS chú ý lắng nghe...<br>+ GV quan sát và trợ giúp HS | [Sản phẩm tiếp] |\n"
    prompt += "| - Bước 3: Báo cáo, thảo luận:<br>+ HS: Chú ý, thảo luận và phát biểu... | [Sản phẩm tiếp] |\n"
    prompt += "| - Bước 4: Kết luận, nhận định: GV chính xác hóa và gọi 1 HS nhắc lại nội dung chính: ... | [Tổng kết] |\n\n"

    prompt += "**C. HOẠT ĐỘNG LUYỆN TẬP**\n"
    prompt += "a) Mục tiêu: Học sinh củng cố lại kiến thức thông qua một số bài tập.\n"
    prompt += "b) Nội dung: HS dựa vào kiến thức đã học vận dụng làm BT.\n"
    prompt += "c) Tổ chức thực hiện và sản phẩm:\n"
    prompt += "BẢNG 2 CỘT:\n"
    prompt += "| HOẠT ĐỘNG CỦA GV VÀ HS | SẢN PHẨM |\n"
    prompt += "|---|---|\n"
    prompt += "| - GV yêu cầu HS hoàn thành các bài tập: [số bài] (tr[XX] - SGK).<br>- HS tiếp nhận nhiệm vụ... | **Bài [số]:**<br>[Lời giải chi tiết] |\n\n"

    prompt += "**D. HOẠT ĐỘNG VẬN DỤNG**\n"
    prompt += "a) Mục tiêu: Học sinh thực hiện làm bài tập vận dụng để củng cố kiến thức và áp dụng kiến thức vào thực tế đời sống.\n"
    prompt += "b) Nội dung: HS sử dụng SGK và vận dụng kiến thức đã học để hoàn thành bài tập.\n"
    prompt += "c) Tổ chức thực hiện và sản phẩm: BẢNG 2 CỘT:\n"
    prompt += "| HOẠT ĐỘNG CỦA GV VÀ HS | SẢN PHẨM |\n"
    prompt += "|---|---|\n"
    prompt += "| - GV cho HS đọc...<br>- GV yêu cầu HS hoàn thành... | **Bài [số]:** [Lời giải chi tiết] |\n\n"

    prompt += "**\\* HƯỚNG DẪN VỀ NHÀ**\n"
    prompt += "- Đọc và ghi nhớ nội dung chính của bài.\n"
    prompt += "- Vận dụng kiến thức làm bài tập [số] (SGK- tr[XX]).\n"
    prompt += "- Chuẩn bị bài mới \"[tên bài tiếp theo]\".\n\n"

    prompt += "**KẾT BÀI (ghi ở cuối cùng, căn phải):**\n"
    prompt += "Sông Đốc, ngày ... tháng ... năm ...\n"
    prompt += "NGƯỜI SOẠN\n"
    prompt += "(Ký tên)\n"
    prompt += mem.get("teacher_name", "") + "\n\n"

    prompt += "=== QUY TẮC BẮT BUỘC ===\n"
    prompt += "1. NỘI DUNG PHẢI LẤY TỪ NGUỒN SGK. KHÔNG bịa.\n"
    prompt += "2. KHÔNG dùng từ 'máy chiếu' → thay bằng 'tivi' hoặc 'máy vi tính'.\n"
    prompt += "3. Ký hiệu toán Unicode: x², x³, √, ≤, ≥, ∈, π, Ư (KHÔNG LaTeX).\n"
    prompt += "4. Phân số viết dạng a/b.\n"
    prompt += "5. Tiếng Việt có dấu đầy đủ.\n"
    prompt += "6. Trong ô bảng: dùng <br> để xuống dòng, **bold** để in đậm.\n"
    prompt += "7. Độ dài tối thiểu: 1500 từ.\n"
    return prompt


# ============ WORD HELPERS ============
def fix_math(text):
    text = re.sub(r"\^2\b", "²", text)
    text = re.sub(r"\^3\b", "³", text)
    text = re.sub(r"_2\b", "₂", text)
    text = text.replace("sqrt(", "√(").replace("\\sqrt{", "√")
    text = text.replace("\\times", "×").replace("\\div", "÷")
    text = text.replace("\\leq", "≤").replace("\\geq", "≥")
    text = text.replace("\\pi", "π").replace("\\alpha", "α")
    text = re.sub(r"\$([^$]+)\$", r"\1", text)
    return text


def add_para(doc, text, style=None):
    text = text.replace("<br>", "\n").replace("<br/>", "\n")
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    for part in re.split(r"(\*\*[^*]+\*\*)", text):
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            p.add_run(part[2:-2]).bold = True
        else:
            p.add_run(part)
    return p


def add_table(doc, rows):
    if not rows:
        return
    n = max(len(r) for r in rows)
    t = doc.add_table(rows=len(rows), cols=n)
    t.style = "Table Grid"
    for i, row in enumerate(rows):
        for j in range(n):
            c = t.cell(i, j)
            c.text = ""
            p = c.paragraphs[0]
            val = (row[j] if j < len(row) else "").replace("<br>", "\n")
            for part in re.split(r"(\*\*[^*]+\*\*)", val):
                if part.startswith("**") and part.endswith("**") and len(part) > 4:
                    p.add_run(part[2:-2]).bold = True
                else:
                    p.add_run(part)


def write_content_to_doc(d, content):
    content = fix_math(content)
    lines = content.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if line.startswith("|") and line.endswith("|"):
            tl = []
            j = i
            while j < len(lines) and lines[j].strip().startswith("|") and lines[j].strip().endswith("|"):
                tl.append(lines[j].strip())
                j += 1
            rows = []
            for t in tl:
                if re.match(r"^\|[\s\-:]+\|", t):
                    continue
                rows.append([c.strip() for c in t.strip("|").split("|")])
            if rows:
                add_table(d, rows)
            i = j
            continue
        if line.startswith("### "):
            d.add_heading(line[4:], 3)
        elif line.startswith("## "):
            d.add_heading(line[3:], 2)
        elif line.startswith("# "):
            d.add_heading(line[2:], 1)
        elif line.startswith("- "):
            add_para(d, line[2:], style="List Bullet")
        elif re.match(r"^\d+[\.\)] ", line):
            add_para(d, re.sub(r"^\d+[\.\)] ", "", line), style="List Number")
        elif line == "":
            pass
        else:
            add_para(d, line)
        i += 1


def make_doc(title, body):
    d = Document()
    s = d.sections[0]
    s.top_margin = Inches(0.7)
    s.bottom_margin = Inches(0.7)
    s.left_margin = Inches(0.8)
    s.right_margin = Inches(0.8)
    d.styles["Normal"].font.name = "Times New Roman"
    d.styles["Normal"].font.size = Pt(12)
    d.add_heading(title, 0)
    mem = st.session_state.get("memory", {})
    if mem.get("school"):
        d.add_paragraph("Đơn vị: " + mem["school"])
    if mem.get("teacher_name"):
        d.add_paragraph("Người thực hiện: " + mem["teacher_name"])
    d.add_paragraph("Ngày tạo: " + datetime.now().strftime("%d/%m/%Y %H:%M"))
    write_content_to_doc(d, body)
    b = io.BytesIO()
    d.save(b)
    return b.getvalue()


# ============ PHIẾU BÁO GIẢNG ============
def add_phieu_bao_giang(doc, lich, tuan_start):
    mem = st.session_state.get("memory", {})

    # Dòng ngày tháng (căn phải)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    now = datetime.now()
    p.add_run("Sông Đốc, ngày " + now.strftime("%d") + " tháng " + now.strftime("%m") + " năm " + now.strftime("%Y"))

    # Tiêu đề
    h = doc.add_heading("PHIẾU BÁO GIẢNG", 0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Tuần: ").bold = True
    end_date = tuan_start + timedelta(days=6)
    p.add_run("Từ " + tuan_start.strftime("%d/%m/%Y")
              + " đến " + end_date.strftime("%d/%m/%Y"))

    doc.add_paragraph()

    # Bảng phiếu báo giảng — 7 cột
    table = doc.add_table(rows=1, cols=7)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    headers = ["Thứ ngày tháng", "Tiết TKB", "Môn", "Lớp",
               "Tiết PPCT", "Tên bài dạy", "Ghi chú"]
    for i, h_text in enumerate(headers):
        hdr[i].text = ""
        run = hdr[i].paragraphs[0].add_run(h_text)
        run.bold = True

    prev_ngay = ""
    for r in lich:
        row = table.add_row().cells
        key_ngay = r["thu"] + " " + r["ngay"]
        if key_ngay != prev_ngay:
            row[0].text = key_ngay
            prev_ngay = key_ngay
        else:
            row[0].text = ""
        row[1].text = r["tiet"].replace("Tiết ", "")
        row[2].text = r["mon"]
        row[3].text = r["lop"]
        row[4].text = str(r.get("stt_ppct", ""))
        row[5].text = r["ten_bai"]
        row[6].text = ""

    doc.add_paragraph()

    # Ký duyệt
    p = doc.add_paragraph()
    p.add_run("KÝ DUYỆT").bold = True
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.add_run("NGƯỜI SOẠN").bold = True
    doc.add_paragraph()
    doc.add_paragraph()
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p3.add_run(mem.get("teacher_name", ""))


# ============ SOẠN 1 BÀI ============
st.divider()
st.subheader("💬 Soạn 1 bài theo lệnh")
cmd = st.text_area("Lệnh", height=100,
                    placeholder="Soạn giáo án Toán 6 bài Ước chung lớn nhất...",
                    label_visibility="collapsed", key="single_cmd")
c1, c2 = st.columns([3, 1])
with c1:
    run = st.button("▶️ Thực hiện", type="primary", use_container_width=True)
with c2:
    if st.button("🗑️ Xóa", use_container_width=True, key="clear_single"):
        st.session_state["draft"] = ""
        st.rerun()

if run and cmd.strip():
    if not available_providers:
        st.error("Chưa có API key.")
    else:
        with st.spinner("🤖 Đang sinh nội dung..."):
            mem = st.session_state["memory"]
            plan = {
                "task": "Soạn kế hoạch bài dạy",
                "subject": mem.get("default_subject", "Toan"),
                "grade": mem.get("default_grade", "6"),
                "topic": cmd,
                "period": "",
                "template_files": list(st.session_state["templates"].keys())[:1],
                "source_files": list(st.session_state["sources"].keys())[:1],
            }
            result = call_ai(build_exec_prompt(plan))
        if result:
            st.session_state["draft"] = result
            st.session_state["last_topic"] = cmd
            save_history("Soạn bài", cmd, result)

if st.session_state.get("draft"):
    st.divider()
    st.subheader("📄 Kết quả")
    content = st.session_state["draft"]
    safe_name = re.sub(r"[^A-Za-z0-9_\-]", "_", st.session_state.get("last_topic", "tai_lieu"))[:45] or "tai_lieu"
    d1, d2 = st.columns(2)
    with d1:
        st.download_button("⬇️ Tải Word", make_doc("Tài liệu", content),
                           file_name=safe_name + ".docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           type="primary", use_container_width=True)
    with d2:
        st.download_button("⬇️ Tải TXT", content.encode("utf-8"),
                           file_name=safe_name + ".txt", mime="text/plain",
                           use_container_width=True)
    st.session_state["draft"] = st.text_area("Nội dung (có thể sửa)", content,
                                              height=500, key="edit_draft")


# ============ TẠO LỊCH BÁO GIẢNG ============
st.divider()
st.subheader("📅 1. Tạo lịch báo giảng từ TKB + PPCT")

st.markdown("**Chọn từ kho đã lưu:**")
col_ppct, col_tkb = st.columns(2)
with col_ppct:
    ppct_options = ["(Không dùng)"] + list(st.session_state["ppct"].keys())
    selected_ppct = st.selectbox("📗 PPCT đã lưu:", ppct_options, key="sel_ppct")
with col_tkb:
    tkb_options = ["(Không dùng)"] + list(st.session_state["tkb"].keys())
    selected_tkb = st.selectbox("📘 TKB đã lưu:", tkb_options, key="sel_tkb")

st.markdown("**Hoặc upload / dán:**")
tkb_file = st.file_uploader("📘 Upload file TKB", type=["pdf", "docx", "xlsx", "xls", "csv", "txt"], key="tkb_file")
tkb_manual = st.text_area("Hoặc dán TKB", height=100, key="tkb_manual")
ppct_file = st.file_uploader("📗 Upload file PPCT", type=["pdf", "docx", "xlsx", "xls", "csv", "txt"], key="ppct_file")
ppct_manual = st.text_area("Hoặc dán PPCT", height=100, key="ppct_manual")
start_date = st.date_input("📆 Ngày bắt đầu tuần (thứ 2)",
                            value=st.session_state["start_date_val"],
                            key="start_date_input")
st.session_state["start_date_val"] = start_date


def ai_extract_tkb(raw):
    system = "Ban la tro ly doc bang. Tra ve CHI text theo format."
    prompt = "Doc TKB sau, chuyen thanh moi dong: Thu | Tiet | Lop | Mon\n\nNOI DUNG:\n"
    prompt += raw[:12000] + "\n\nCHI tra ve cac dong."
    try:
        result = call_ai(prompt, system)
        return "\n".join([l.strip() for l in result.splitlines() if "|" in l and len(l.split("|")) >= 4])
    except Exception:
        return ""


def ai_extract_ppct(raw):
    system = "Bạn là trợ lý đọc bảng. Chỉ trả về text theo format, không giải thích."
    prompt = """Đọc bảng PHÂN PHỐI CHƯƠNG TRÌNH sau. Bảng có 5 cột:
STT | Bài học | Số tiết | Thời điểm | Ghi chú

QUAN TRỌNG:
1. Cột "Thời điểm" thường bị GỘP Ô — nhiều bài liên tiếp dùng chung 1 tuần.
   Hãy ĐIỀN tuần cho TẤT CẢ các bài, kể cả khi ô trống.
2. Các dòng có cùng tên bài và ghi "(Tiết 1)", "(Tiết 2)" → GỘP THÀNH 1 BÀI
   với số tiết = tổng các tiết.

Trả về MỖI BÀI 1 DÒNG format:
STT | Tên bài (KHÔNG có "(Tiết N)") | Số tiết (đã gộp) | Tuần | Ngày bắt đầu | Ngày kết thúc

Ví dụ:
1 | Bài 1. Tập hợp | 1 | 1 | 07/09/2026 | 13/09/2026
5 | Bài 5. Phép nhân và phép chia số tự nhiên | 2 | 2 | 14/09/2026 | 20/09/2026

CHỈ trả về các dòng theo format. KHÔNG giải thích.

NỘI DUNG PPCT:
""" + raw[:15000]

    try:
        result = call_ai(prompt, system)
        lines = []
        for line in result.splitlines():
            line = line.strip()
            if "|" not in line:
                continue
            low = line.lower()
            if "stt" in low and "bài" in low:
                continue
            if "---" in line:
                continue
            lines.append(line)
        return "\n".join(lines)
    except Exception as e:
        return ""


def auto_fix_tkb(text):
    fixed = []
    for line in text.strip().splitlines():
        line = line.strip()
        if not line:
            continue
        if "|" in line:
            fixed.append(line)
            continue
        m = re.match(r"^(\d)(\d)(6A\d+)(.+)$", line)
        if m:
            fixed.append("Thứ " + m.group(1) + " | Tiết " + m.group(2)
                         + " | " + m.group(3) + " | " + m.group(4).strip())
        else:
            fixed.append(line)
    return "\n".join(fixed)


def parse_tkb(text):
    result = {}
    for line in text.strip().splitlines():
        line = line.strip()
        if not line or "|" not in line:
            continue
        parts = [p.strip() for p in line.split("|")]
        if len(parts) < 4:
            continue
        thu = parts[0]
        if not thu.lower().startswith("thứ"):
            thu = "Thứ " + thu
        tiet = parts[1]
        if not tiet.lower().startswith("tiết"):
            num = re.search(r"\d+", tiet)
            tiet = "Tiết " + (num.group(0) if num else tiet)
        if thu not in result:
            result[thu] = []
        result[thu].append({"tiet": tiet, "lop": parts[2], "mon": parts[3]})
    for thu in result:
        result[thu].sort(key=lambda x: int(re.search(r"\d+", x["tiet"]).group(0)) if re.search(r"\d+", x["tiet"]) else 0)
    return result


def parse_ppct(text):
    """Parse PPCT — GỘP bài nhiều tiết."""
    raw_list = []
    for line in text.strip().splitlines():
        line = line.strip()
        if not line or "|" not in line:
            continue
        parts = [p.strip() for p in line.split("|")]
        if len(parts) < 3:
            continue
        if parts[0].lower() == "stt":
            continue
        m_stt = re.search(r"(\d+)", parts[0])
        if not m_stt:
            continue
        stt = int(m_stt.group(1))
        ten_bai = parts[1]
        so_tiet_str = parts[2] if len(parts) > 2 else "1"
        m = re.search(r"(\d+)", so_tiet_str)
        n_tiet = int(m.group(1)) if m else 1
        tuan = parts[3].strip() if len(parts) > 3 else ""
        tuan_start = None
        tuan_end = None
        if len(parts) >= 5:
            d1 = re.search(r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})", parts[4])
            if d1:
                try:
                    tuan_start = datetime(int(d1.group(3)), int(d1.group(2)), int(d1.group(1)))
                except Exception:
                    pass
        if len(parts) >= 6:
            d2 = re.search(r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})", parts[5])
            if d2:
                try:
                    tuan_end = datetime(int(d2.group(3)), int(d2.group(2)), int(d2.group(1)))
                except Exception:
                    pass
        raw_list.append({
            "stt": stt,
            "tiet_bd": stt,
            "ten_bai": ten_bai,
            "so_tiet": n_tiet,
            "tuan": tuan,
            "tuan_start": tuan_start,
            "tuan_end": tuan_end,
        })

    # GỘP bài có cùng tên (bài nhiều tiết)
    merged = []
    i = 0
    while i < len(raw_list):
        bai = raw_list[i]
        ten_goc = bai["ten_bai"]
        ten_chuan = re.sub(r"\s*\(Tiết\s*\d+\)", "", ten_goc).strip()
        so_tiet_gop = 0
        j = i
        while j < len(raw_list):
            ten_j = re.sub(r"\s*\(Tiết\s*\d+\)", "", raw_list[j]["ten_bai"]).strip()
            if ten_j == ten_chuan and raw_list[j]["tuan"] == bai["tuan"]:
                so_tiet_gop += raw_list[j]["so_tiet"]
                j += 1
            else:
                break
        merged.append({
            "stt": bai["stt"],
            "tiet_bd": bai["tiet_bd"],
            "ten_bai": ten_chuan,
            "so_tiet": so_tiet_gop,
            "tuan": bai["tuan"],
            "tuan_start": bai["tuan_start"],
            "tuan_end": bai["tuan_end"],
        })
        i = j if so_tiet_gop > 0 else i + 1

    return merged


def find_week_ppct(ppct, target_date):
    for bai in ppct:
        if bai.get("tuan_start") and bai.get("tuan_end"):
            if bai["tuan_start"] <= target_date <= bai["tuan_end"]:
                tuan_target = bai["tuan"]
                cung_tuan = [b for b in ppct if str(b.get("tuan")) == str(tuan_target)]
                return cung_tuan, tuan_target
    return None, None


if st.button("📅 Tạo lịch báo giảng", type="primary", use_container_width=True):
    if not available_providers:
        st.error("Cần API key.")
    else:
        tkb_raw = tkb_manual.strip()
        if tkb_file:
            _, text, _ = read_file(tkb_file, use_ocr=False)
            tkb_raw = text
        elif not tkb_raw and selected_tkb != "(Không dùng)":
            tkb_raw = st.session_state["tkb"][selected_tkb]
            st.info("📘 Dùng TKB: " + selected_tkb)
        elif not tkb_raw and st.session_state["tkb"]:
            first = list(st.session_state["tkb"].keys())[0]
            tkb_raw = st.session_state["tkb"][first]
            st.info("📘 Dùng TKB đã lưu: " + first)

        ppct_raw = ppct_manual.strip()
        if ppct_file:
            _, text, _ = read_file(ppct_file, use_ocr=False)
            ppct_raw = text
        elif not ppct_raw and selected_ppct != "(Không dùng)":
            ppct_raw = st.session_state["ppct"][selected_ppct]
            st.info("📗 Dùng PPCT: " + selected_ppct)
        elif not ppct_raw and st.session_state["ppct"]:
            first = list(st.session_state["ppct"].keys())[0]
            ppct_raw = st.session_state["ppct"][first]
            st.info("📗 Dùng PPCT đã lưu: " + first)

        if not tkb_raw or not ppct_raw:
            st.error("Cần TKB và PPCT.")
        else:
            tkb_auto = auto_fix_tkb(tkb_raw)
            if tkb_auto and "|" in tkb_auto and len(tkb_auto.splitlines()[0].split("|")) >= 4:
                tkb_clean = tkb_auto
            else:
                with st.spinner("🧠 AI đang đọc TKB..."):
                    tkb_clean = ai_extract_tkb(tkb_raw)

            with st.spinner("🧠 AI đang đọc PPCT (tách tuần, gộp bài)..."):
                ppct_clean = ai_extract_ppct(ppct_raw)

            with st.expander("📋 Xem kết quả AI trích xuất", expanded=False):
                st.markdown("**TKB đã chuẩn hóa:**")
                st.code(tkb_clean)
                st.markdown("**PPCT đã chuẩn hóa:**")
                st.code(ppct_clean)

            tkb = parse_tkb(tkb_clean)
            ppct = parse_ppct(ppct_clean)

            if not tkb or not ppct:
                st.error("Không parse được.")
            else:
                target = datetime.combine(start_date, datetime.min.time())
                week_bais, week_number = find_week_ppct(ppct, target)

                if week_bais:
                    st.success("✅ Tìm thấy **Tuần " + str(week_number)
                               + "** — " + str(len(week_bais)) + " bài")
                    ppct_week = week_bais
                else:
                    st.warning("⚠️ Không tìm thấy tuần khớp.")
                    ppct_week = ppct

                # Tạo ppct_tiets — mỗi bài chiếm n tiết
                ppct_tiets = []
                for bai in ppct_week:
                    for k in range(bai["so_tiet"]):
                        ppct_tiets.append({
                            "ten_bai": bai["ten_bai"],
                            "tiet_trong_bai": k + 1,
                            "tong_tiet": bai["so_tiet"],
                            "tuan": bai.get("tuan", ""),
                            "stt_ppct": bai.get("stt", ""),
                        })

                days_map = {
                    "thứ 2": 0, "thu 2": 0, "thứ hai": 0, "2": 0,
                    "thứ 3": 1, "thu 3": 1, "thứ ba": 1, "3": 1,
                    "thứ 4": 2, "thu 4": 2, "thứ tư": 2, "4": 2,
                    "thứ 5": 3, "thu 5": 3, "thứ năm": 3, "5": 3,
                    "thứ 6": 4, "thu 6": 4, "thứ sáu": 4, "6": 4,
                    "thứ 7": 5, "thu 7": 5, "thứ bảy": 5, "7": 5,
                }
                all_slots = []
                for thu in ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7"]:
                    for k in tkb:
                        if k.lower().startswith(thu.lower()[:5]):
                            for item in tkb[k]:
                                all_slots.append({
                                    "thu": k, "tiet": item["tiet"],
                                    "lop": item["lop"], "mon": item["mon"],
                                })

                # Ghép slot với tiết — GIỮ ĐỦ tất cả slot
                lich = []
                for idx, slot in enumerate(all_slots):
                    thu_lower = slot["thu"].lower()
                    offset = 0
                    for key in days_map:
                        if key in thu_lower:
                            offset = days_map[key]
                            break
                    ngay = start_date + timedelta(days=offset)

                    ten_bai = ""
                    tiet_trong_bai = 1
                    tong_tiet = 1
                    stt_ppct = ""
                    if idx < len(ppct_tiets):
                        tiet_info = ppct_tiets[idx]
                        ten_bai = tiet_info["ten_bai"]
                        tiet_trong_bai = tiet_info["tiet_trong_bai"]
                        tong_tiet = tiet_info["tong_tiet"]
                        stt_ppct = tiet_info.get("stt_ppct", "")

                    lich.append({
                        "thu": slot["thu"],
                        "ngay": ngay.strftime("%d/%m/%Y"),
                        "tiet": slot["tiet"],
                        "lop": slot["lop"],
                        "mon": slot["mon"],
                        "ten_bai": ten_bai,
                        "tiet_trong_bai": tiet_trong_bai,
                        "tong_tiet": tong_tiet,
                        "tuan": "",
                        "stt_ppct": stt_ppct,
                    })

                st.session_state["lich_bao_giang"] = lich
                st.success("✅ Đã tạo lịch: " + str(len(lich)) + " tiết")


# ============ HIỂN THỊ LỊCH ============
def hien_thi_lich_theo_ngay(lich):
    days_order = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7"]
    for thu in days_order:
        bai_trong_ngay = [r for r in lich if r["thu"].lower().startswith(thu.lower()[:5])]
        if not bai_trong_ngay:
            continue
        ngay = bai_trong_ngay[0]["ngay"]
        st.markdown("### 📅 **" + thu + " — " + ngay + "**")
        st.markdown("| Tiết | Lớp | Môn | Tên bài | Tiết |")
        st.markdown("|---|---|---|---|---|")
        for r in bai_trong_ngay:
            st.markdown(
                "| " + r["tiet"] + " | " + r["lop"] + " | " + r["mon"]
                + " | " + r["ten_bai"]
                + " | " + str(r["tiet_trong_bai"]) + "/" + str(r["tong_tiet"]) + " |"
            )
        st.markdown("")


if st.session_state.get("lich_bao_giang"):
    lich = st.session_state["lich_bao_giang"]
    st.divider()
    st.subheader("📋 Lịch báo giảng — Tuần (" + str(len(lich)) + " tiết)")

    st.markdown("#### 📊 Bảng tổng hợp")
    st.markdown("| Thứ | Ngày | Tiết | Lớp | Môn | Tiết PPCT | Tên bài | Tiết |")
    st.markdown("|---|---|---|---|---|---|---|---|")
    for r in lich:
        st.markdown("| " + r["thu"] + " | " + r["ngay"] + " | " + r["tiet"]
                    + " | " + r["lop"] + " | " + r["mon"]
                    + " | " + str(r.get("stt_ppct", ""))
                    + " | " + r["ten_bai"]
                    + " | " + str(r["tiet_trong_bai"]) + "/" + str(r["tong_tiet"]) + " |")

    st.markdown("")
    st.markdown("#### 📅 Chi tiết theo ngày")
    hien_thi_lich_theo_ngay(lich)

    if st.button("🚀 2. Soạn cả tuần (KHÔNG soạn trùng bài)",
                 type="primary", use_container_width=True):
        if not available_providers:
            st.error("Chưa có API key.")
        else:
            # Gom bài KHÔNG trùng để soạn
            bai_can_soan = {}
            for r in lich:
                if not r["ten_bai"]:
                    continue
                key = r["ten_bai"].strip()
                if key not in bai_can_soan:
                    bai_can_soan[key] = {
                        "ten_bai": key,
                        "cac_lop": [r["lop"]],
                        "tong_tiet": r["tong_tiet"],
                    }
                else:
                    if r["lop"] not in bai_can_soan[key]["cac_lop"]:
                        bai_can_soan[key]["cac_lop"].append(r["lop"])

            st.info("📚 Có **" + str(len(bai_can_soan)) + " bài** không trùng cần soạn "
                    + "(trong " + str(len(lich)) + " tiết dạy).")

            progress = st.progress(0)
            status_text = st.empty()
            results = []
            bai_list = list(bai_can_soan.values())
            for idx, bai in enumerate(bai_list):
                status_text.write("⏳ Soạn bài " + str(idx + 1) + "/" + str(len(bai_list))
                                   + ": " + bai["ten_bai"])
                plan = {
                    "task": "Soạn kế hoạch bài dạy",
                    "subject": st.session_state.get("memory", {}).get("default_subject", "Toan"),
                    "grade": st.session_state.get("memory", {}).get("default_grade", "6"),
                    "topic": bai["ten_bai"],
                    "period": str(bai["tong_tiet"]) + " tiết",
                    "template_files": list(st.session_state["templates"].keys())[:1],
                    "source_files": list(st.session_state["sources"].keys())[:1],
                }
                try:
                    content = call_ai(build_exec_prompt(plan))
                    results.append({
                        "ten_bai": bai["ten_bai"],
                        "cac_lop": bai["cac_lop"],
                        "tong_tiet": bai["tong_tiet"],
                        "content": content,
                    })
                except Exception as e:
                    results.append({
                        "ten_bai": bai["ten_bai"],
                        "cac_lop": bai["cac_lop"],
                        "tong_tiet": bai["tong_tiet"],
                        "content": "[Lỗi: " + str(e) + "]",
                    })
                progress.progress((idx + 1) / len(bai_list))
            status_text.write("✅ Đã soạn " + str(len(results)) + " bài (không trùng)!")
            st.session_state["weekly_results"] = results


# ============ XUẤT WORD TUẦN ============
if st.session_state.get("weekly_results"):
    results = st.session_state["weekly_results"]
    lich = st.session_state.get("lich_bao_giang", [])
    st.divider()
    st.subheader("📚 Kết quả tuần — " + str(len(results)) + " bài không trùng")
    for r in results:
        lop_str = ", ".join(r.get("cac_lop", []))
        st.write("- **" + r.get("ten_bai", "") + "** — Lớp: " + lop_str
                 + " (" + str(r.get("tong_tiet", 1)) + " tiết)")

    if st.button("📄 3. Xuất file Word tuần (Phiếu báo giảng + bài soạn)",
                 type="primary", use_container_width=True):
        d = Document()
        s = d.sections[0]
        s.top_margin = Inches(0.7)
        s.bottom_margin = Inches(0.7)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)
        d.styles["Normal"].font.name = "Times New Roman"
        d.styles["Normal"].font.size = Pt(12)

        if lich:
            add_phieu_bao_giang(d, lich, st.session_state["start_date_val"])

        for idx, r in enumerate(results):
            d.add_page_break()
            lop_str = ", ".join(r.get("cac_lop", []))
            d.add_heading("BÀI " + str(idx + 1) + ": " + r.get("ten_bai", ""), 1)
            p_info = d.add_paragraph()
            p_info.add_run("Lớp: ").bold = True
            p_info.add_run(lop_str)
            p_info.add_run(" | Số tiết: ").bold = True
            p_info.add_run(str(r.get("tong_tiet", 1)))
            d.add_paragraph("---")
            write_content_to_doc(d, r.get("content", ""))

        bio = io.BytesIO()
        d.save(bio)
        st.success("✅ File Word: Phiếu báo giảng + " + str(len(results))
                   + " bài soạn (không trùng)")
        st.download_button(
            "⬇️ Tải file Word tuần",
            bio.getvalue(),
            file_name="tuan_" + st.session_state["start_date_val"].strftime("%Y%m%d") + ".docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary", use_container_width=True,
        )
