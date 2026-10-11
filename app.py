import subprocess, sys, os, json, base64, urllib.request
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
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
import io, re, csv

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
st.caption("Bộ não tự động • Lịch báo giảng • Mẫu CV 5512 • Xuất Word")


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
    res = gh_api("DELETE", f"/repos/{GITHUB_REPO}/contents/{path}",
                 {"message": "Delete " + path, "sha": sha, "branch": GITHUB_BRANCH})
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
    if isinstance(res, dict) and "login" in res:
        login = res["login"]
        res2 = gh_api("GET", f"/repos/{GITHUB_REPO}")
        if isinstance(res2, dict) and "permissions" in res2:
            perms = res2["permissions"]
            if perms.get("push") or perms.get("admin"):
                return (True, "OK — " + login + " | Có quyền ghi")
            return (False, "Token thiếu quyền 'repo'")
        return (True, "OK — " + login)
    if isinstance(res, dict) and "error" in res:
        return (False, "Lỗi: " + res["error"])
    return (False, "Token không hợp lệ")


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


def load_templates(): return load_folder("templates")
def load_sources(): return load_folder("sources")
def load_ppct(): return load_folder("ppct")
def load_tkb(): return load_folder("tkb")
def load_schedule(): return load_folder("schedule")


def save_history(task, topic, content):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe = re.sub(r"[^a-zA-Z0-9]", "_", topic)[:30]
    return gh_save(f"data/history/{ts}_{safe}.md",
                   f"# {task}\n\n**Chủ đề:** {topic}\n\n---\n\n{content}",
                   f"History {topic}")


def ocr_image_with_gemini(data, fname):
    if not GEMINI_KEY: return "[" + fname + ": cần GEMINI_API_KEY]"
    if not HAS_GEMINI_VISION: return "[" + fname + ": chưa cài google-generativeai]"
    if not HAS_PIL: return "[" + fname + ": chưa cài pillow]"
    try:
        genai_vision.configure(api_key=GEMINI_KEY)
        img = Image.open(io.BytesIO(data))
        for mn in ["gemini-2.0-flash", "gemini-2.5-flash", "gemini-flash-latest"]:
            try:
                model = genai_vision.GenerativeModel(mn)
                p = ("Đây là ảnh trang sách giáo khoa. Hãy đọc TOÀN BỘ chữ trong ảnh, "
                     "giữ nguyên cấu trúc, dùng Unicode cho toán: x², √x, ≤, π. "
                     "Chỉ trả về text, không mô tả.")
                resp = model.generate_content([p, img])
                t = resp.text.strip()
                if len(t) >= 20:
                    return t
            except Exception:
                continue
        return "[" + fname + ": AI không đọc được]"
    except Exception as e:
        return "[" + fname + ": Lỗi OCR — " + str(e) + "]"


def _read_docx_safe(data):
    try:
        d = Document(io.BytesIO(data))
        out = [p.text for p in d.paragraphs if p.text.strip()]
        for t in d.tables:
            out.append("")
            for row in t.rows:
                out.append("| " + " | ".join(c.text.strip() for c in row.cells) + " |")
            out.append("")
        return "\n".join(out)
    except Exception as e:
        return "[Lỗi docx: " + str(e) + "]"


def _read_pdf_safe(data):
    if not HAS_PYPDF: return "[Chưa cài pypdf]"
    try:
        reader = PdfReader(io.BytesIO(data))
        texts = []
        for page in reader.pages:
            try:
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
            return "[PDF scan, không có text]" + full
        return full
    except Exception as e:
        return "[Lỗi pdf: " + str(e) + "]"


def _read_excel_safe(data):
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        out = []
        for sn in wb.sheetnames:
            ws = wb[sn]
            rows = [[str(c) if c is not None else "" for c in row] for row in ws.iter_rows(values_only=True)]
            rows = [r for r in rows if any(c.strip() for c in r)]
            if not rows: continue
            mc = max(len(r) for r in rows)
            rows = [r + [""] * (mc - len(r)) for r in rows]
            out.append("### Sheet: " + sn)
            out.append("| " + " | ".join(rows[0]) + " |")
            out.append("|" + "---|" * mc)
            for r in rows[1:]:
                out.append("| " + " | ".join(r) + " |")
            out.append("")
        return "\n".join(out)
    except Exception as e:
        return "[Lỗi Excel: " + str(e) + "]"


def _read_csv_safe(data):
    try:
        text = None
        for enc in ["utf-8", "utf-8-sig", "latin-1", "cp1258"]:
            try:
                text = data.decode(enc); break
            except Exception:
                continue
        if text is None:
            text = data.decode("utf-8", errors="ignore")
        sample = text[:2000]
        delim = "," if sample.count(",") >= sample.count("|") else "|"
        rows = list(csv.reader(io.StringIO(text), delimiter=delim))
        mc = max(len(r) for r in rows)
        rows = [r + [""] * (mc - len(r)) for r in rows]
        out = ["| " + " | ".join(rows[0]) + " |", "|" + "---|" * mc]
        for r in rows[1:]:
            out.append("| " + " | ".join(r) + " |")
        return "\n".join(out)
    except Exception as e:
        return "[Lỗi csv: " + str(e) + "]"


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

    if nl.endswith((".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp")):
        return (name, ocr_image_with_gemini(data, name) if use_ocr else "[" + name + ": ảnh]", 0)
    if magic[:4] == b"%PDF":
        text = _read_pdf_safe(data)
    elif magic[:4] == b"PK\x03\x04":
        text = _read_excel_safe(data) if nl.endswith((".xlsx", ".xlsm")) else _read_docx_safe(data)
    elif magic[:4] == b"\xd0\xcf\x11\xe0":
        text = "[File .doc cũ. Save As .docx trước]"
    elif nl.endswith(".docx"):
        text = _read_docx_safe(data)
    elif nl.endswith(".pdf"):
        text = _read_pdf_safe(data)
    elif nl.endswith((".xlsx", ".xls")):
        text = _read_excel_safe(data)
    elif nl.endswith(".csv"):
        text = _read_csv_safe(data)
    else:
        text = _read_txt_safe(data)
    return (name, text, len(text))


for k, v in {
    "memory": load_memory(),
    "templates": load_templates(),
    "sources": load_sources(),
    "ppct": load_ppct(),
    "tkb": load_tkb(),
    "schedule": load_schedule(),
    "draft": "",
    "last_topic": "",
    "weekly_results": [],
    "lich_bao_giang": [],
    "start_date_val": datetime.now().date(),
    "buoi_day": "BUỔI CHIỀU",
    "tuan_number": "",
}.items():
    if k not in st.session_state:
        st.session_state[k] = v


with st.sidebar:
    st.header("🧠 Bộ não")
    mem = st.session_state["memory"]

    if GITHUB_TOKEN:
        st.success("💾 Đã có GITHUB_TOKEN")
        if st.button("🧪 Test GitHub", use_container_width=True):
            ok, msg = gh_test()
            st.success("✅ " + msg) if ok else st.error("❌ " + msg)
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
        st.success("✅ " + msg) if ok else st.error("❌ " + msg)

    st.divider()
    for label, key in [("Mẫu", "templates"), ("Nguồn", "sources"), ("PPCT", "ppct"),
                       ("TKB", "tkb"), ("Lịch BG", "schedule")]:
        st.write(label + ": " + str(len(st.session_state[key])) + " file")

    if st.button("🔄 Tải lại", use_container_width=True):
        for key, fn in [("templates", load_templates), ("sources", load_sources),
                        ("ppct", load_ppct), ("tkb", load_tkb), ("schedule", load_schedule)]:
            st.session_state[key] = fn()
        st.rerun()
    st.write("AI: " + str(len(available_providers)) + " nhà cung cấp")


def call_ai(prompt, system=""):
    import freelm
    llm = freelm.FreeLLM.from_env()
    msgs = []
    if system: msgs.append({"role": "system", "content": system})
    msgs.append({"role": "user", "content": prompt})
    last_err = None
    for pk in available_providers:
        for mn in PREFERRED_MODELS.get(pk, []):
            try:
                resp = llm.chat(msgs, model=mn)
                t = getattr(resp, "text", None) or getattr(resp, "content", "")
                if t and len(t.strip()) > 50:
                    return t
            except Exception as e:
                last_err = e
                continue
    try:
        resp = llm.chat(msgs, model="auto")
        return getattr(resp, "text", None) or getattr(resp, "content", "")
    except Exception as e:
        if last_err: raise last_err
        raise e
