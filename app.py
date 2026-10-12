# ============================================================
# PHẦN 1/4: IMPORTS + CONFIG + GITHUB + ĐỌC FILE + SESSION
# ============================================================
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
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL
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


# ---------- GitHub API ----------
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


# ---------- Load / Save data ----------
def load_memory():
    content, _ = gh_get("data/memory.json")
    if content:
        try:
            return json.loads(content)
        except Exception:
            pass
    return {"teacher_name": "", "school": "", "default_subject": "Toán", "default_grade": "6"}


def save_memory(memory):
    memory["last_updated"] = datetime.now().strftime("%Y-%m-%d")
    return gh_save("data/memory.json",
                   json.dumps(memory, ensure_ascii=False, indent=2),
                   "Update memory")


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


# ---------- OCR ảnh bằng Gemini ----------
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


# ---------- Đọc file an toàn ----------
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
    if not HAS_PYPDF:
        return "[Chưa cài pypdf]"
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
            rows = [[str(c) if c is not None else "" for c in row]
                    for row in ws.iter_rows(values_only=True)]
            rows = [r for r in rows if any(c.strip() for c in r)]
            if not rows:
                continue
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
                text = data.decode(enc)
                break
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


# ---------- Khởi tạo session state ----------
if "app_started_v19" not in st.session_state:
    for k in list(st.session_state.keys()):
        if k.startswith("tuan_") or k.startswith("buoi_") or k.startswith("ngay_bd_"):
            del st.session_state[k]
    st.session_state["app_started_v19"] = True


defaults = {
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
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ============================================================
# PHẦN 2/4: SIDEBAR + UPLOAD/KHO + CALL_AI
# ============================================================

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
    mem["default_subject"] = st.text_input("Môn mặc định",
                                            value=mem.get("default_subject", "Toán"))
    mem["default_grade"] = st.text_input("Khối mặc định",
                                          value=mem.get("default_grade", "6"))

    if st.button("💾 Lưu vào bộ nhớ", use_container_width=True):
        ok, msg = save_memory(mem)
        st.success("✅ " + msg) if ok else st.error("❌ " + msg)

    st.divider()
    st.subheader("📊 Thống kê")
    for label, key in [("Mẫu", "templates"), ("Nguồn", "sources"),
                       ("PPCT", "ppct"), ("TKB", "tkb"), ("Lịch BG", "schedule")]:
        st.write(label + ": " + str(len(st.session_state[key])) + " file")

    if st.button("🔄 Tải lại", use_container_width=True):
        st.session_state["templates"] = load_templates()
        st.session_state["sources"] = load_sources()
        st.session_state["ppct"] = load_ppct()
        st.session_state["tkb"] = load_tkb()
        st.session_state["schedule"] = load_schedule()
        st.rerun()

    st.write("AI: " + str(len(available_providers)) + " nhà cung cấp")

    st.divider()
    st.subheader("🧹 Dọn dẹp")
    with st.expander("🗑️ Dọn rác và cache", expanded=False):
        st.caption("Chọn mục cần dọn:")

        if st.button("🧼 Xóa phiên làm việc tạm", use_container_width=True,
                     key="clean_session"):
            for k in ["draft", "last_topic", "weekly_results", "lich_bao_giang"]:
                if k in st.session_state:
                    del st.session_state[k]
            st.success("✅ Đã xóa phiên làm việc tạm")
            st.rerun()

        if st.button("♻️ Xóa cache Streamlit", use_container_width=True,
                     key="clean_cache"):
            st.cache_data.clear()
            st.cache_resource.clear()
            st.success("✅ Đã xóa cache Streamlit")
            st.rerun()

        if GITHUB_TOKEN:
            if st.button("📜 Xóa lịch sử soạn bài", use_container_width=True,
                         key="clean_history"):
                files = gh_list("data/history")
                deleted = 0
                for f in files:
                    ok, _ = gh_delete(f"data/history/{f}")
                    if ok:
                        deleted += 1
                st.success("✅ Đã xóa " + str(deleted) + " file lịch sử")

            if st.button("🗑️ Xóa file rác (file_*.md)", use_container_width=True,
                         key="clean_junk"):
                deleted = 0
                for folder in ["templates", "sources", "ppct", "tkb", "schedule"]:
                    files = gh_list(f"data/{folder}")
                    for f in files:
                        if f.startswith("file_"):
                            ok, _ = gh_delete(f"data/{folder}/{f}")
                            if ok:
                                deleted += 1
                st.success("✅ Đã xóa " + str(deleted) + " file rác")
                st.session_state["templates"] = load_templates()
                st.session_state["sources"] = load_sources()
                st.session_state["ppct"] = load_ppct()
                st.session_state["tkb"] = load_tkb()
                st.session_state["schedule"] = load_schedule()
                st.rerun()

            if st.button("📅 Xóa tất cả lịch đã lưu", use_container_width=True,
                         key="clean_sched"):
                files = gh_list("data/schedule")
                deleted = 0
                for f in files:
                    ok, _ = gh_delete(f"data/schedule/{f}")
                    if ok:
                        deleted += 1
                st.success("✅ Đã xóa " + str(deleted) + " file lịch")
                st.session_state["schedule"] = load_schedule()

        st.divider()
        st.caption("⚠️ **Cảnh báo**: các nút dưới đây xóa dữ liệu VĨNH VIỄN")

        if st.button("🔴 Reset toàn bộ app", use_container_width=True,
                     key="reset_all"):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.success("✅ Đã reset. Đang tải lại...")
            st.rerun()

        if GITHUB_TOKEN:
            if st.button("🔴🔴 XÓA TẤT CẢ dữ liệu GitHub",
                         use_container_width=True, key="del_all_gh"):
                st.session_state["confirm_delete_all"] = True

            if st.session_state.get("confirm_delete_all"):
                st.warning("⚠️ Bạn có chắc muốn xóa TẤT CẢ?")
                c_yes, c_no = st.columns(2)
                with c_yes:
                    if st.button("✅ XÓA", key="confirm_yes_x"):
                        deleted = 0
                        for folder in ["templates", "sources", "ppct", "tkb",
                                       "schedule", "history"]:
                            files = gh_list(f"data/{folder}")
                            for f in files:
                                ok, _ = gh_delete(f"data/{folder}/{f}")
                                if ok:
                                    deleted += 1
                        st.session_state["templates"] = {}
                        st.session_state["sources"] = {}
                        st.session_state["ppct"] = {}
                        st.session_state["tkb"] = {}
                        st.session_state["schedule"] = {}
                        del st.session_state["confirm_delete_all"]
                        st.success("✅ Đã xóa " + str(deleted) + " file")
                        st.rerun()
                with c_no:
                    if st.button("❌ Hủy", key="confirm_no_x"):
                        del st.session_state["confirm_delete_all"]
                        st.rerun()


# ============ HÀM GỌI AI ============
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
        if last_err:
            raise last_err
        raise e


# ============ TAB UPLOAD / KHO ============
tab_up, tab_kho = st.tabs(["📤 Upload", "📂 Kho tri thức"])

with tab_up:
    st.info("💡 MẪU = định dạng. NGUỒN = nội dung SGK. PPCT/TKB/Lịch BG = lịch dạy.")
    up_type = st.radio("Loại", ["Mẫu", "Nguồn", "PPCT", "TKB", "Lịch BG"],
                       horizontal=True)
    up_name = st.text_input("Tên file", placeholder="mau_cv5512")
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
                st.write("🖼️ " + f.name + " — sẽ OCR")
            else:
                st.write("📎 " + f.name)

    if st.button("💾 Lưu lên GitHub", type="primary", use_container_width=True):
        ok0, msg0 = gh_test()
        if not ok0:
            st.error("❌ " + msg0)
        else:
            parts = []
            if up_paste.strip():
                parts.append(up_paste.strip())
            for f in (up_files or []):
                nl = f.name.lower()
                if nl.endswith((".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif")):
                    with st.spinner("🔍 OCR " + f.name + "..."):
                        data = f.read()
                        txt = ocr_image_with_gemini(data, f.name)
                        parts.append("--- " + f.name + " ---\n" + txt)
                else:
                    _, txt, _ = read_file(f, use_ocr=False)
                    parts.append("--- " + f.name + " ---\n" + txt)

            if not parts:
                st.warning("Không có nội dung.")
            else:
                full = "\n\n".join(parts)
                safe = re.sub(r"[^a-zA-Z0-9_\-]", "_", up_name.strip())[:40] or \
                       "file_" + datetime.now().strftime("%Y%m%d_%H%M%S")
                folder = {"Mẫu": "templates", "Nguồn": "sources", "PPCT": "ppct",
                          "TKB": "tkb", "Lịch BG": "schedule"}.get(up_type, "sources")
                path = f"data/{folder}/{safe}.md"
                with st.spinner("💾 Đang lưu..."):
                    ok, msg = gh_save(path, full, f"Upload {safe}")
                    if ok:
                        st.success("✅ " + msg)
                        st.session_state["templates"] = load_templates()
                        st.session_state["sources"] = load_sources()
                        st.session_state["ppct"] = load_ppct()
                        st.session_state["tkb"] = load_tkb()
                        st.session_state["schedule"] = load_schedule()
                        st.rerun()
                    else:
                        st.error("❌ " + msg)


with tab_kho:
    for label, key, folder in [
        ("📐 Mẫu", "templates", "templates"),
        ("📚 Nguồn", "sources", "sources"),
        ("📗 PPCT", "ppct", "ppct"),
        ("📘 TKB", "tkb", "tkb"),
        ("📅 Lịch BG", "schedule", "schedule"),
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

# ============================================================
# PHẦN 3/4: BỘ QUY TẮC 11 MÔN + PROMPT + WORD HELPERS + SOẠN 1 BÀI
# ============================================================

# ============ BỘ QUY TẮC 11 MÔN HỌC ============
SUBJECT_RULES = {
    "Toán": {
        "ten_goi_khac": ["toan", "math", "toán học"],
        "ky_hieu": "📐 π ∑ √",
        "hinh_minh_hoa": "hình học, công thức, biểu đồ, trục số, mô hình toán học",
        "nang_luc_trong_tam": [
            "Năng lực tư duy và lập luận toán học",
            "Năng lực giải quyết vấn đề toán học",
            "Năng lực giao tiếp toán học",
            "Năng lực mô hình hóa toán học",
            "Năng lực sử dụng công cụ, phương tiện học toán",
        ],
        "phuong_phap_uu_tien": [
            "Khám phá quy luật, gợi mở",
            "Giải quyết vấn đề",
            "Mô hình hóa toán học",
            "Luyện tập thực hành phân hóa",
        ],
        "san_pham_hoc_tap": [
            "Lời giải và bài trình bày",
            "Phép tính, biểu thức, phương trình",
            "Hình vẽ, bảng số liệu, biểu đồ",
            "Lập luận hoặc chứng minh",
            "Mô hình toán học của tình huống thực tế",
        ],
        "danh_gia": [
            "Tính chính xác",
            "Lập luận logic",
            "Phương pháp giải",
            "Khả năng vận dụng",
        ],
        "loai_bai": [
            "Hình thành khái niệm",
            "Hình thành quy tắc",
            "Luyện tập",
            "Ôn tập",
            "Thực hành",
            "Vận dụng thực tế",
        ],
        "dac_thu_hoat_dong": {
            "mo_dau": "Đưa ra câu hỏi, tình huống thực tế, bài toán ngắn hoặc vấn đề cần giải quyết.",
            "trong_tam": "Hướng dẫn học sinh quan sát, dự đoán, phân tích, phát hiện quy luật và hình thành kiến thức mới.",
            "luyen_tap": "Tổ chức bài tập từ cơ bản đến nâng cao, có phân hóa.",
            "van_dung": "Yêu cầu học sinh giải quyết tình huống thực tế hoặc bài toán mới.",
        },
        "quy_tac_rieng": [
            "Các bước giải rõ ràng.",
            "Câu hỏi dẫn dắt tư duy.",
            "Dự kiến câu trả lời của học sinh.",
            "Những lỗi sai phổ biến và cách hỗ trợ.",
            "Bài tập có mức độ từ dễ đến khó.",
        ],
        "khong_nen": "Chỉ trình bày lý thuyết rồi cho một loạt bài tập mà không thể hiện quá trình học sinh hình thành kiến thức.",
    },

    "Ngữ văn": {
        "ten_goi_khac": ["ngu van", "văn", "van", "literature", "ngữ văn"],
        "ky_hieu": "📖 ✍️ 🎭",
        "hinh_minh_hoa": "sách, bút, trang văn bản, nhân vật văn học, sơ đồ phân tích",
        "nang_luc_trong_tam": [
            "Năng lực ngôn ngữ",
            "Năng lực văn học",
            "Năng lực đọc hiểu văn bản",
            "Năng lực viết",
            "Năng lực nói và nghe",
        ],
        "phuong_phap_uu_tien": [
            "Đọc hiểu",
            "Phân tích, thảo luận",
            "Thực hành viết",
            "Đóng vai, trình bày",
        ],
        "san_pham_hoc_tap": [
            "Phiếu học tập đọc hiểu",
            "Sơ đồ phân tích nhân vật hoặc nội dung văn bản",
            "Đoạn văn, bài văn",
            "Bài thuyết trình, tranh luận",
            "Bảng tổng hợp đặc điểm ngôn ngữ, thể loại, biện pháp tu từ",
        ],
        "danh_gia": [
            "Khả năng đọc hiểu",
            "Khả năng phân tích chi tiết, hình ảnh, nhân vật",
            "Khả năng viết đoạn văn, bài văn",
            "Khả năng trình bày, thảo luận, lắng nghe",
        ],
        "loai_bai": [
            "Văn bản văn học",
            "Văn bản thông tin",
            "Văn bản nghị luận",
            "Tiếng Việt",
            "Viết",
            "Nói và nghe",
        ],
        "dac_thu_hoat_dong": {
            "mo_dau": "Gợi vấn đề bằng hình ảnh, câu hỏi, tình huống, trải nghiệm hoặc liên hệ chủ đề văn bản.",
            "trong_tam": "Đọc, tìm hiểu từ ngữ, hình ảnh, chi tiết, nhân vật, bố cục, chủ đề hoặc đặc điểm thể loại.",
            "luyen_tap": "Trả lời câu hỏi đọc hiểu, phân tích ngữ liệu, luyện viết hoặc thực hành tiếng Việt.",
            "van_dung": "Viết, nói, trình bày quan điểm hoặc liên hệ vấn đề trong văn bản với thực tiễn.",
        },
        "quy_tac_rieng": [
            "Nhận diện đúng loại bài: văn bản văn học / thông tin / nghị luận / Tiếng Việt / Viết / Nói-nghe.",
            "Điều chỉnh quy trình theo dạng bài.",
            "Không đưa ra bài văn mẫu duy nhất để học sinh học thuộc.",
            "Câu hỏi đọc hiểu phải theo 4 mức độ: nhận biết, thông hiểu, vận dụng, vận dụng cao.",
        ],
        "khong_nen": "Ép tất cả bài Ngữ văn vào một quy trình phân tích tác phẩm văn học hoặc đưa ra một bài văn mẫu duy nhất để học sinh học thuộc.",
    },

    "Khoa học tự nhiên": {
        "ten_goi_khac": ["khtn", "khoa hoc tu nhien", "khoa học tự nhiên", "science"],
        "ky_hieu": "🔬 🧪 🌡️ ⚗️",
        "hinh_minh_hoa": "ống nghiệm, kính hiển vi, mô hình phân tử, biểu đồ thí nghiệm",
        "nang_luc_trong_tam": [
            "Năng lực nhận thức khoa học tự nhiên",
            "Năng lực tìm hiểu tự nhiên",
            "Năng lực vận dụng kiến thức, kỹ năng đã học",
        ],
        "phuong_phap_uu_tien": [
            "Quan sát",
            "Thí nghiệm, thực hành",
            "Điều tra, khảo sát",
            "Phân tích dữ liệu",
            "Giải thích hiện tượng bằng bằng chứng",
        ],
        "san_pham_hoc_tap": [
            "Báo cáo thí nghiệm",
            "Bảng quan sát",
            "Biểu đồ, sơ đồ hoặc mô hình",
            "Bản giải thích hiện tượng",
            "Phiếu điều tra",
            "Kết luận dựa trên bằng chứng",
        ],
        "danh_gia": [
            "Hiểu khái niệm và quy luật khoa học",
            "Biết quan sát và đặt câu hỏi",
            "Đưa ra dự đoán, giả thuyết",
            "Thực hiện thí nghiệm/khảo sát",
            "Phân tích dữ liệu và rút ra kết luận",
        ],
        "loai_bai": [
            "Hình thành khái niệm",
            "Tìm hiểu hiện tượng",
            "Thí nghiệm hoặc thực hành",
            "Giải thích và vận dụng",
            "Ôn tập, hệ thống hóa",
        ],
        "dac_thu_hoat_dong": {
            "mo_dau": "Đưa ra hiện tượng tự nhiên, câu hỏi khoa học hoặc tình huống cần giải thích.",
            "trong_tam": "Quan sát, dự đoán, thí nghiệm, đọc tài liệu, xử lý dữ liệu và xây dựng lời giải thích.",
            "luyen_tap": "Phân tích hiện tượng, đọc bảng số liệu, giải thích kết quả hoặc thực hành bài tập.",
            "van_dung": "Giải thích hiện tượng thực tế, đề xuất giải pháp hoặc thực hiện nhiệm vụ điều tra nhỏ.",
        },
        "quy_tac_rieng": [
            "Phân biệt dạng bài: hình thành khái niệm / hiện tượng / thí nghiệm / giải thích / ôn tập.",
            "Với bài thí nghiệm: nêu thiết bị, vật liệu, quy trình, yêu cầu an toàn, cách xử lý kết quả.",
            "Không tạo thí nghiệm không có trong điều kiện thiết bị thực tế.",
            "Không đưa kết luận khoa học không phù hợp với dữ liệu.",
        ],
        "khong_nen": "Tạo thí nghiệm không có trong điều kiện thiết bị thực tế hoặc đưa ra kết luận khoa học mà không phù hợp với dữ liệu và nội dung bài học.",
    },

    "Lịch sử": {
        "ten_goi_khac": ["lich su", "lịch sử", "history", "ls"],
        "ky_hieu": "🏛️ 📜 ⏳",
        "hinh_minh_hoa": "đền cổ, cuộn giấy, dòng thời gian, bản đồ lịch sử",
        "nang_luc_trong_tam": [
            "Năng lực tìm hiểu lịch sử",
            "Năng lực nhận thức và tư duy lịch sử",
            "Năng lực vận dụng kiến thức, kỹ năng đã học",
        ],
        "phuong_phap_uu_tien": [
            "Khai thác tư liệu",
            "Dòng thời gian",
            "So sánh sự kiện",
            "Phân tích nguyên nhân - diễn biến - kết quả",
        ],
        "san_pham_hoc_tap": [
            "Dòng thời gian",
            "Sơ đồ diễn biến sự kiện",
            "Bảng so sánh",
            "Phiếu khai thác tư liệu",
            "Bài thuyết trình lịch sử",
            "Bài viết giải thích hoặc đánh giá một sự kiện",
        ],
        "danh_gia": [
            "Xác định thời gian và không gian lịch sử",
            "Nhận biết sự kiện, nhân vật, quá trình",
            "Phân tích nguyên nhân, diễn biến, kết quả",
            "So sánh các sự kiện, giai đoạn",
            "Khai thác tư liệu lịch sử",
            "Liên hệ lịch sử với di sản và đời sống hiện nay",
        ],
        "loai_bai": [
            "Hình thành kiến thức mới",
            "Luyện tập",
            "Ôn tập, hệ thống hóa",
            "Thực hành khai thác tư liệu",
        ],
        "dac_thu_hoat_dong": {
            "mo_dau": "Sử dụng tranh ảnh, hiện vật, bản đồ, câu chuyện hoặc câu hỏi lịch sử.",
            "trong_tam": "Khai thác tư liệu, xác định bối cảnh, lập dòng thời gian, phân tích nguyên nhân và diễn biến.",
            "luyen_tap": "Sắp xếp sự kiện, lập bảng so sánh, giải thích nguyên nhân hoặc phân tích tư liệu.",
            "van_dung": "Tìm hiểu di tích, nhân vật, di sản hoặc vấn đề lịch sử gắn với địa phương.",
        },
        "quy_tac_rieng": [
            "Xác định rõ các yếu tố: thời gian, không gian, bối cảnh, nguyên nhân, diễn biến, kết quả, ý nghĩa, tư liệu.",
            "Không bắt học sinh học thuộc toàn bộ diễn biến.",
            "Phải tổ chức phân tích tư liệu, giải thích hoặc so sánh.",
        ],
        "khong_nen": "Bắt học sinh học thuộc toàn bộ diễn biến mà không tổ chức phân tích tư liệu, giải thích hoặc so sánh.",
    },

    "Địa lí": {
        "ten_goi_khac": ["dia li", "địa lí", "địa lý", "geography", "dl"],
        "ky_hieu": "🗺️ 🌍 🧭",
        "hinh_minh_hoa": "bản đồ, quả địa cầu, la bàn, biểu đồ khí hậu",
        "nang_luc_trong_tam": [
            "Năng lực nhận thức khoa học địa lí",
            "Năng lực tìm hiểu địa lí",
            "Năng lực vận dụng kiến thức, kỹ năng đã học",
        ],
        "phuong_phap_uu_tien": [
            "Đọc bản đồ, lược đồ",
            "Phân tích biểu đồ, số liệu",
            "So sánh vùng, quốc gia",
            "Khai thác thông tin địa lí",
        ],
        "san_pham_hoc_tap": [
            "Bản đồ được chú giải",
            "Biểu đồ",
            "Bảng thống kê",
            "Bài nhận xét số liệu",
            "Sơ đồ phân bổ",
            "Báo cáo khảo sát địa lí địa phương",
        ],
        "danh_gia": [
            "Xác định vị trí địa lí",
            "Nhận biết đặc điểm tự nhiên, dân cư, kinh tế - xã hội",
            "Đọc và sử dụng bản đồ",
            "Phân tích biểu đồ và bảng số liệu",
            "Giải thích sự phân bố của các đối tượng địa lí",
        ],
        "loai_bai": [
            "Hình thành kiến thức mới",
            "Luyện tập",
            "Thực hành đọc bản đồ, biểu đồ",
            "Ôn tập, hệ thống hóa",
        ],
        "dac_thu_hoat_dong": {
            "mo_dau": "Khai thác bản đồ, hình ảnh, số liệu hoặc tình huống địa lí.",
            "trong_tam": "Đọc bản đồ, phân tích số liệu, xác định vị trí và giải thích đặc điểm địa lí.",
            "luyen_tap": "Nhận xét biểu đồ, tính toán số liệu, so sánh vùng hoặc xác định vị trí.",
            "van_dung": "Phân tích vấn đề môi trường, thiên tai, dân cư, sản xuất hoặc phát triển địa phương.",
        },
        "quy_tac_rieng": [
            "Với bài có số liệu: xác định dữ liệu đầu vào, yêu cầu xử lý, phương pháp phù hợp.",
            "Tính toán chính xác nếu cần.",
            "Hướng dẫn nhận xét, giải thích xu hướng và mối quan hệ địa lí.",
            "Kiểm tra tính hợp lí của kết luận.",
        ],
        "khong_nen": "Chỉ yêu cầu học sinh đọc lại số liệu hoặc mô tả biểu đồ mà không giải thích xu hướng và mối quan hệ địa lí.",
    },

    "Giáo dục công dân": {
        "ten_goi_khac": ["gdcd", "giao duc cong dan", "giáo dục công dân", "civic"],
        "ky_hieu": "⚖️ 🤝 📋",
        "hinh_minh_hoa": "cán cân công lý, bắt tay, phiếu tình huống",
        "nang_luc_trong_tam": [
            "Năng lực điều chỉnh hành vi",
            "Năng lực phát triển bản thân",
            "Năng lực tìm hiểu và tham gia các hoạt động kinh tế - xã hội",
        ],
        "phuong_phap_uu_tien": [
            "Phân tích tình huống",
            "Xử lý tình huống",
            "Đóng vai",
            "Thảo luận nhóm",
        ],
        "san_pham_hoc_tap": [
            "Phiếu xử lý tình huống",
            "Bảng phân tích hành vi",
            "Bài trình bày quan điểm",
            "Kịch bản đóng vai",
            "Kế hoạch rèn luyện cá nhân",
            "Cam kết hoặc sản phẩm tuyên truyền",
        ],
        "danh_gia": [
            "Hiểu khái niệm và chuẩn mực đạo đức",
            "Nhận biết quyền, nghĩa vụ, trách nhiệm",
            "Phân tích hành vi đúng và sai",
            "Biết lựa chọn cách ứng xử phù hợp",
        ],
        "loai_bai": [
            "Hình thành kiến thức mới",
            "Luyện tập",
            "Thực hành xử lý tình huống",
            "Vận dụng vào đời sống",
        ],
        "dac_thu_hoat_dong": {
            "mo_dau": "Đưa ra câu chuyện, tình huống ứng xử hoặc vấn đề gần gũi.",
            "trong_tam": "Phân tích tình huống, xác định chuẩn mực, quyền, nghĩa vụ hoặc trách nhiệm.",
            "luyen_tap": "Nhận xét hành vi, xử lý tình huống, lựa chọn cách ứng xử và giải thích lí do.",
            "van_dung": "Xây dựng kế hoạch rèn luyện, thực hiện hành vi tích cực hoặc giải quyết tình huống thực tế.",
        },
        "quy_tac_rieng": [
            "Tình huống phải có: bối cảnh rõ ràng, nhân vật và hành vi cụ thể, vấn đề cần giải quyết.",
            "Có câu hỏi phân tích, các phương án ứng xử.",
            "Yêu cầu giải thích và liên hệ thực tế.",
            "Không đánh giá phẩm chất học sinh chỉ dựa vào một câu trả lời.",
        ],
        "khong_nen": "Chỉ đưa ra định nghĩa và yêu cầu học sinh học thuộc; cũng không nên đánh giá phẩm chất học sinh chỉ dựa vào một câu trả lời đơn lẻ.",
    },

    "Mỹ thuật": {
        "ten_goi_khac": ["my thuat", "mỹ thuật", "art"],
        "ky_hieu": "🎨 🖌️ 🖼️",
        "hinh_minh_hoa": "bảng màu, cọ vẽ, tranh mẫu, sản phẩm tạo hình",
        "nang_luc_trong_tam": [
            "Năng lực quan sát và nhận thức thẩm mỹ",
            "Năng lực sáng tạo và ứng dụng thẩm mỹ",
            "Năng lực phân tích và đánh giá thẩm mỹ",
        ],
        "phuong_phap_uu_tien": [
            "Quan sát",
            "Thực hành tạo hình",
            "Thiết kế",
            "Trưng bày, nhận xét sản phẩm",
        ],
        "san_pham_hoc_tap": [
            "Tranh vẽ",
            "Sản phẩm tạo hình",
            "Bản phác thảo",
            "Sản phẩm trang trí hoặc thiết kế",
            "Bộ sưu tập hình ảnh",
            "Bài giới thiệu và nhận xét tác phẩm",
        ],
        "danh_gia": [
            "Nhận biết yếu tố và nguyên lí tạo hình",
            "Quan sát, phân tích và cảm nhận tác phẩm",
            "Lựa chọn chất liệu, công cụ, phương pháp thể hiện",
            "Tạo ra sản phẩm theo yêu cầu",
            "Trình bày ý tưởng và nhận xét sản phẩm",
        ],
        "loai_bai": [
            "Vẽ theo mẫu",
            "Vẽ trang trí",
            "Vẽ tranh",
            "Tạo hình",
            "Thường thức mỹ thuật",
            "Trưng bày sản phẩm",
        ],
        "dac_thu_hoat_dong": {
            "mo_dau": "Quan sát tranh, vật thật, hình ảnh hoặc sản phẩm nghệ thuật.",
            "trong_tam": "Tìm hiểu bố cục, đường nét, hình khối, màu sắc, chất liệu và cách thể hiện.",
            "luyen_tap": "Thực hành phác thảo, tạo hình, vẽ hoặc thiết kế.",
            "van_dung": "Hoàn thiện sản phẩm, trưng bày, giới thiệu hoặc ứng dụng vào đời sống.",
        },
        "quy_tac_rieng": [
            "Xác định sản phẩm cần tạo ra, chất liệu, dụng cụ.",
            "Nêu rõ các bước thực hiện, thời gian thực hành.",
            "Có tiêu chí đánh giá cụ thể.",
            "Có cách trưng bày và nhận xét.",
            "Không đánh giá chỉ dựa trên đẹp/xấu.",
            "Không yêu cầu mọi học sinh tạo sản phẩm giống hệt nhau.",
        ],
        "khong_nen": "Đánh giá sản phẩm chỉ dựa trên việc đẹp hay xấu hoặc yêu cầu mọi học sinh tạo ra những sản phẩm giống hệt nhau.",
    },

    "Âm nhạc": {
        "ten_goi_khac": ["am nhac", "âm nhạc", "music"],
        "ky_hieu": "🎵 🎼 🎤",
        "hinh_minh_hoa": "nốt nhạc, khuông nhạc, nhạc cụ, ca sĩ biểu diễn",
        "nang_luc_trong_tam": [
            "Năng lực thể hiện âm nhạc",
            "Năng lực cảm thụ và hiểu biết âm nhạc",
            "Năng lực ứng dụng và sáng tạo âm nhạc",
        ],
        "phuong_phap_uu_tien": [
            "Nghe mẫu",
            "Luyện hát, đọc nhạc",
            "Thực hành nhạc cụ",
            "Trò chơi tiết tấu",
            "Biểu diễn",
        ],
        "san_pham_hoc_tap": [
            "Phần trình diễn bài hát",
            "Bài đọc nhạc",
            "Phần biểu diễn nhạc cụ",
            "Tiết tấu hoặc đoạn nhạc do học sinh sáng tạo",
            "Bài cảm nhận âm nhạc",
            "Video hoặc bản ghi âm biểu diễn",
        ],
        "danh_gia": [
            "Nhận biết đặc điểm tác phẩm âm nhạc",
            "Hát đúng giai điệu, lời ca, nhịp điệu",
            "Nghe và cảm nhận âm nhạc",
            "Đọc nhạc hoặc sử dụng nhạc cụ",
            "Hợp tác trong hoạt động biểu diễn",
        ],
        "loai_bai": [
            "Học hát",
            "Nghe nhạc",
            "Đọc nhạc",
            "Nhạc cụ",
            "Lí thuyết âm nhạc",
            "Thường thức âm nhạc",
        ],
        "dac_thu_hoat_dong": {
            "mo_dau": "Nghe một đoạn nhạc, nhận biết giai điệu hoặc thực hiện trò chơi tiết tấu.",
            "trong_tam": "Nghe mẫu, luyện hát, đọc nhạc, tìm hiểu tác phẩm hoặc kỹ thuật sử dụng nhạc cụ.",
            "luyen_tap": "Luyện theo câu, nhóm, tổ hoặc cá nhân; kết hợp tiết tấu và nhạc cụ khi phù hợp.",
            "van_dung": "Biểu diễn, phối hợp nhóm, sáng tạo tiết tấu hoặc chia sẻ cảm nhận.",
        },
        "quy_tac_rieng": [
            "Nhận diện dạng bài: học hát, nghe nhạc, đọc nhạc, nhạc cụ, lí thuyết, thường thức.",
            "Mỗi dạng bài có hoạt động và tiêu chí đánh giá riêng.",
            "Không chỉ trình bày kiến thức về tác phẩm mà bỏ qua hoạt động nghe, thực hành.",
        ],
        "khong_nen": "Chỉ trình bày kiến thức về tác phẩm mà bỏ qua hoạt động nghe, thực hành hoặc thể hiện âm nhạc khi bài học yêu cầu.",
    },

    "Giáo dục thể chất": {
        "ten_goi_khac": ["gdtc", "giao duc the chat", "giáo dục thể chất", "pe", "thể dục", "the duc"],
        "ky_hieu": "⚽ 🏃 🏀",
        "hinh_minh_hoa": "sân bãi, động tác thể dục, bóng, dụng cụ tập",
        "nang_luc_trong_tam": [
            "Năng lực vận động cơ bản",
            "Năng lực thể thao",
            "Năng lực thể chất và sức khỏe",
        ],
        "phuong_phap_uu_tien": [
            "Làm mẫu",
            "Luyện tập",
            "Sửa lỗi kỹ thuật",
            "Trò chơi vận động",
        ],
        "san_pham_hoc_tap": [
            "Động tác được thực hiện",
            "Chuỗi vận động",
            "Kết quả thực hành",
            "Bảng theo dõi tiến bộ",
            "Kết quả trò chơi vận động",
            "Khả năng phối hợp trong nhóm",
        ],
        "danh_gia": [
            "Hiểu yêu cầu và tác dụng của động tác",
            "Thực hiện đúng kỹ thuật vận động",
            "Phát triển tố chất thể lực",
            "Biết tự tập luyện và điều chỉnh vận động",
            "Ý thức an toàn và kỷ luật",
        ],
        "loai_bai": [
            "Bài lý thuyết",
            "Bài thực hành kỹ thuật",
            "Bài ôn luyện",
            "Trò chơi vận động",
            "Kiểm tra đánh giá",
        ],
        "dac_thu_hoat_dong": {
            "mo_dau": "Nhận lớp, kiểm tra tình trạng sẵn sàng, khởi động và chuẩn bị vận động.",
            "trong_tam": "Giáo viên làm mẫu, phân tích kỹ thuật, hướng dẫn học sinh tập theo tiến trình phù hợp.",
            "luyen_tap": "Tập cá nhân, theo cặp, nhóm hoặc đội hình; quan sát và sửa lỗi.",
            "van_dung": "Trò chơi vận động, bài tập tổng hợp hoặc vận dụng kỹ năng trong tình huống phù hợp.",
        },
        "quy_tac_rieng": [
            "Xác định địa điểm và dụng cụ.",
            "Nêu nội dung khởi động.",
            "Xác định kỹ thuật trọng tâm.",
            "Nêu cách tổ chức đội hình.",
            "Nêu khối lượng và thời gian luyện tập.",
            "Chỉ ra lỗi kỹ thuật thường gặp và cách sửa.",
            "Nêu biện pháp đảm bảo an toàn.",
        ],
        "khong_nen": "Soạn một giáo án chỉ gồm lời giảng, câu hỏi và bài tập viết cho một nội dung cần thực hành vận động.",
    },

    "Hoạt động trải nghiệm, hướng nghiệp": {
        "ten_goi_khac": ["hdtn", "hoat dong trai nghiem", "hướng nghiệp", "trai nghiem", "experiential"],
        "ky_hieu": "🧭 🌱 🎯",
        "hinh_minh_hoa": "la bàn, mầm cây, kế hoạch, hoạt động nhóm",
        "nang_luc_trong_tam": [
            "Năng lực thích ứng với cuộc sống",
            "Năng lực thiết kế và tổ chức hoạt động",
            "Năng lực định hướng nghề nghiệp",
        ],
        "phuong_phap_uu_tien": [
            "Trải nghiệm thực tế",
            "Dự án",
            "Khảo sát",
            "Thảo luận nhóm",
            "Phản hồi và tự đánh giá",
        ],
        "san_pham_hoc_tap": [
            "Nhật ký trải nghiệm",
            "Kế hoạch cá nhân",
            "Báo cáo khảo sát",
            "Poster hoặc video",
            "Sản phẩm dự án",
            "Bản phân công nhiệm vụ",
            "Bài trình bày kết quả hoạt động",
        ],
        "danh_gia": [
            "Nhận thức bản thân",
            "Kỹ năng giao tiếp và hợp tác",
            "Giải quyết vấn đề trong đời sống",
            "Thực hiện trách nhiệm với gia đình, nhà trường, cộng đồng",
            "Tìm hiểu nghề nghiệp",
            "Xây dựng kế hoạch phát triển bản thân",
        ],
        "loai_bai": [
            "Sinh hoạt dưới cờ",
            "Sinh hoạt lớp",
            "Hoạt động giáo dục theo chủ đề",
            "Hoạt động câu lạc bộ",
            "Hoạt động hướng nghiệp",
        ],
        "dac_thu_hoat_dong": {
            "mo_dau": "Nêu chủ đề, tình huống hoặc vấn đề cần trải nghiệm.",
            "trong_tam": "Học sinh thực hiện nhiệm vụ, khảo sát, thảo luận, trải nghiệm hoặc giải quyết tình huống.",
            "luyen_tap": "Báo cáo kết quả, phản hồi, tự đánh giá và đánh giá nhóm.",
            "van_dung": "Xây dựng kế hoạch hành động và áp dụng vào đời sống.",
        },
        "quy_tac_rieng": [
            "Xác định chủ đề trải nghiệm, nhiệm vụ thực tế.",
            "Nêu địa điểm, hình thức tổ chức, thời gian.",
            "Nêu nguồn lực và người tham gia.",
            "Xác định sản phẩm hoặc minh chứng.",
            "Nêu tiêu chí đánh giá.",
            "Có hoạt động phản tư và vận dụng.",
            "Không biến mọi tiết trải nghiệm thành tiết giáo viên thuyết trình.",
        ],
        "khong_nen": "Biến mọi tiết trải nghiệm thành một tiết giáo viên thuyết trình về kỹ năng sống.",
    },

    "Tin học": {
        "ten_goi_khac": ["tin hoc", "tin học", "informatics", "it", "cntt"],
        "ky_hieu": "💻 🖥️ ⌨️",
        "hinh_minh_hoa": "máy tính, sơ đồ thuật toán, code, màn hình",
        "nang_luc_trong_tam": [
            "Năng lực sử dụng và quản lý các phương tiện công nghệ thông tin",
            "Năng lực ứng dụng công nghệ thông tin và truyền thông trong học và tự học",
            "Năng lực giải quyết vấn đề với sự hỗ trợ của công nghệ thông tin",
            "Năng lực ứng dụng công nghệ thông tin và truyền thông trong giao tiếp và hợp tác",
            "Năng lực hiểu biết về an toàn, an ninh thông tin",
        ],
        "phuong_phap_uu_tien": [
            "Thực hành",
            "Giải quyết vấn đề",
            "Thử nghiệm",
            "Kiểm thử và sửa lỗi",
        ],
        "san_pham_hoc_tap": [
            "Tệp văn bản hoặc bảng tính",
            "Bài trình chiếu",
            "Sơ đồ thuật toán",
            "Chương trình máy tính",
            "Kết quả chạy thử",
            "Báo cáo thực hành",
            "Sản phẩm số phục vụ học tập",
        ],
        "danh_gia": [
            "Hiểu khái niệm tin học",
            "Sử dụng công cụ số phù hợp",
            "Thực hiện các thao tác trên máy tính",
            "Phân tích vấn đề và xây dựng thuật toán",
            "Viết, kiểm thử hoặc sửa chương trình",
            "Sử dụng công nghệ an toàn và có trách nhiệm",
        ],
        "loai_bai": [
            "Bài lí thuyết",
            "Bài sử dụng phần mềm",
            "Bài thực hành",
            "Bài thuật toán",
            "Bài lập trình",
            "Bài mạng và an toàn số",
        ],
        "dac_thu_hoat_dong": {
            "mo_dau": "Nêu vấn đề cần giải quyết hoặc giới thiệu sản phẩm số cần tạo.",
            "trong_tam": "Phân tích ví dụ, quan sát thao tác, tìm hiểu thuật toán hoặc công cụ.",
            "luyen_tap": "Thực hành theo nhiệm vụ, chạy thử, kiểm tra kết quả và sửa lỗi.",
            "van_dung": "Hoàn thiện sản phẩm số hoặc giải quyết một vấn đề mới.",
        },
        "quy_tac_rieng": [
            "Phân biệt: lí thuyết / sử dụng phần mềm / thực hành / thuật toán / lập trình / mạng-an toàn số.",
            "Với bài lập trình, giáo án nên có: mô tả bài toán, dữ liệu đầu vào, kết quả đầu ra, ý tưởng giải quyết, thuật toán, mã chương trình, bộ dữ liệu kiểm thử, kết quả mong đợi, lỗi thường gặp và cách sửa.",
        ],
        "khong_nen": "Chỉ cho học sinh chép mã chương trình mà không giải thích bài toán, thuật toán và cách kiểm tra kết quả.",
    },
}


def _chuan_hoa_ten_mon(ten_mon):
    """Chuẩn hóa tên môn (bỏ dấu, lowercase) để so khớp."""
    if not ten_mon:
        return ""
    ten_mon = ten_mon.strip().lower()
    import unicodedata
    ten_mon = unicodedata.normalize("NFD", ten_mon)
    ten_mon = "".join(c for c in ten_mon if unicodedata.category(c) != "Mn")
    ten_mon = ten_mon.replace("đ", "d").replace("Đ", "D")
    return ten_mon


def get_subject_rules(ten_mon):
    """Tra bộ quy tắc theo tên môn (linh hoạt với viết hoa, dấu, viết tắt)."""
    if not ten_mon:
        return None
    ten_chuan = _chuan_hoa_ten_mon(ten_mon)

    for key, rules in SUBJECT_RULES.items():
        if _chuan_hoa_ten_mon(key) == ten_chuan:
            return key, rules

    for key, rules in SUBJECT_RULES.items():
        for alias in rules.get("ten_goi_khac", []):
            if _chuan_hoa_ten_mon(alias) == ten_chuan:
                return key, rules

    for key, rules in SUBJECT_RULES.items():
        key_chuan = _chuan_hoa_ten_mon(key)
        if key_chuan in ten_chuan or ten_chuan in key_chuan:
            return key, rules
        for alias in rules.get("ten_goi_khac", []):
            alias_chuan = _chuan_hoa_ten_mon(alias)
            if len(alias_chuan) >= 3 and (
                alias_chuan in ten_chuan or ten_chuan in alias_chuan
            ):
                return key, rules

    return None


def detect_chapter_from_topic(topic):
    """Đoán chương từ tên bài để nhắc AI đúng mạch kiến thức.
    Trả về (ten_chuong, canh_bao) hoặc (None, None).
    """
    if not topic:
        return None, None
    t = topic.lower()

    toan6_rules = [
        (r"(tập hợp|ghi số tự nhiên|thứ tự trong tập hợp|cộng.*trừ số tự nhiên|"
         r"nhân.*chia số tự nhiên|lũy thừa|thứ tự thực hiện|bài tập cuối chương i)",
         "Chương I - Tập hợp các số tự nhiên",
         "CHỈ dùng số tự nhiên 0,1,2,3... KHÔNG dùng số nguyên âm, số thập phân."),

        (r"(quan hệ chia hết|dấu hiệu chia hết|số nguyên tố|hợp số|"
         r"ước chung|bội chung|ƯCLN|BCNN|bài tập cuối chương ii)",
         "Chương II - Tính chia hết trong tập hợp các số tự nhiên",
         "CHỈ dùng số tự nhiên. KHÔNG nhắc số nguyên âm, số nguyên tố âm, "
         "KHÔNG dùng 'số nguyên' khi chưa học Chương III."),

        (r"(tập hợp các số nguyên|số nguyên âm|cộng.*trừ số nguyên|"
         r"quy tắc dấu ngoặc|nhân.*chia số nguyên|ước.*bội.*số nguyên|"
         r"bài tập ôn chương iii)",
         "Chương III - Số nguyên",
         "Đã học số nguyên. Có thể dùng số âm, quy tắc dấu."),

        (r"(điểm|đường thẳng|đoạn thẳng|tia|trung điểm|góc|số đo góc)",
         "Chương VIII - Những hình học cơ bản",
         "CHỈ dùng kiến thức hình học cơ bản. KHÔNG dùng định lý, chứng minh phức tạp."),

        (r"(phân số|so sánh phân số|hỗn số|cộng.*trừ phân số|"
         r"nhân.*chia phân số|hai bài toán về phân số)",
         "Chương VI - Phân số",
         "Đã học phân số. KHÔNG dùng số thập phân (học Chương VII)."),

        (r"(số thập phân|tính toán với số thập phân|làm tròn|ước lượng|"
         r"tỉ số|tỉ số phần trăm)",
         "Chương VII - Số thập phân",
         "Đã học số thập phân. Có thể dùng dấu phẩy thập phân."),
    ]

    for pattern, chuong, canh_bao in toan6_rules:
        if re.search(pattern, t, re.IGNORECASE):
            return chuong, canh_bao

    return None, None


# ============ PROMPT CV 5512 THÔNG MINH THEO MÔN ============
def build_exec_prompt(plan):
    mem = st.session_state["memory"]

    ten_mon_input = plan.get("subject", "") or mem.get("default_subject", "")
    matched = get_subject_rules(ten_mon_input)

    if matched:
        ten_mon_chuan, rules = matched
    else:
        ten_mon_chuan = ten_mon_input or "Chung"
        rules = {
            "ky_hieu": "📚",
            "hinh_minh_hoa": "tài liệu, bảng, sơ đồ",
            "nang_luc_trong_tam": ["Năng lực chung", "Năng lực chuyên môn"],
            "phuong_phap_uu_tien": ["Thuyết trình", "Thảo luận", "Thực hành"],
            "san_pham_hoc_tap": ["Bài trình bày", "Sản phẩm thực hành"],
            "danh_gia": ["Kiến thức", "Kỹ năng", "Thái độ"],
            "loai_bai": ["Hình thành kiến thức mới", "Luyện tập", "Vận dụng"],
            "dac_thu_hoat_dong": {},
            "quy_tac_rieng": [],
            "khong_nen": "",
        }

    tpl_text = ""
    for f in plan.get("template_files", []):
        if f in st.session_state["templates"]:
            tpl_text += "\n=== FILE MẪU ===\n" + st.session_state["templates"][f][:20000]

    src_text = ""
    for f in plan.get("source_files", []):
        if f in st.session_state["sources"]:
            src_text += "\n=== NGUỒN SGK ===\n" + st.session_state["sources"][f][:25000]

    prompt = "Bạn là trợ lý chuyên môn cao cấp cho giáo viên THCS Việt Nam.\n\n"
    prompt += ("NHIỆM VỤ: Soạn KẾ HOẠCH BÀI DẠY môn **" + ten_mon_chuan
               + "** theo ĐÚNG mẫu Công văn 5512.\n\n")

    prompt += "=== THÔNG TIN BÀI DẠY ===\n"
    prompt += "- GV: " + mem.get("teacher_name", "") + " | Trường: " + mem.get("school", "") + "\n"
    prompt += "- Môn: " + ten_mon_chuan + " | Khối: " + (plan.get("grade", "") or "6") + "\n"
    prompt += "- Tên bài: " + plan.get("topic", "") + "\n"
    prompt += "- Số tiết: " + (plan.get("period") or "1") + "\n"
    _trang = plan.get("trang_sgk", "").strip()
    if _trang:
        prompt += "- 📖 PHẠM VI SGK: CHỈ dùng nội dung trang " + _trang + " SGK.\n"
        prompt += ("   KHÔNG dùng kiến thức ở các trang khác, "
                   "KHÔNG dùng kiến thức của bài khác.\n")
    prompt += "\n"

    _chuong_detect, _canh_bao_chuong = detect_chapter_from_topic(plan.get("topic", ""))
    if _chuong_detect:
        prompt += "=== 📖 CHƯƠNG DỰ KIẾN ===\n"
        prompt += "Chương: " + _chuong_detect + "\n"
        if _canh_bao_chuong:
            prompt += "⚠️ LƯU Ý MẠCH KIẾN THỨC: " + _canh_bao_chuong + "\n"
        prompt += "\n"

    prompt += "=== ⚠️ NGUỒN CHUẨN BẮT BUỘC ===\n"
    prompt += "1. CHỈ được dùng kiến thức có trong:\n"
    prompt += "   - NỘI DUNG SGK của ĐÚNG bài này (phần NGUỒN SGK bên dưới).\n"
    prompt += "   - PPCT của ĐÚNG tuần/bài này.\n"
    prompt += "   - File MẪU (định dạng 5512) — KHÔNG lấy nội dung từ mẫu.\n"
    prompt += "2. TUYỆT ĐỐI KHÔNG dùng kiến thức của:\n"
    prompt += "   - Bài khác, chương khác chưa học.\n"
    prompt += "   - Lớp khác, cấp khác.\n"
    prompt += "   - Kiến thức ngoài chương trình THCS.\n"
    prompt += "3. Nếu bài thuộc Chương I, II (Toán 6 HK1): KHÔNG nhắc số nguyên âm, "
    prompt += "số hữu tỉ, số vô tỉ, căn bậc hai, hàm số, phương trình... (CHƯA HỌC).\n"
    prompt += "4. Nếu không chắc bài thuộc chương nào, dựa vào TÊN BÀI + "
    prompt += "SỐ TIẾT PPCT để xác định đúng mạch kiến thức.\n"
    prompt += "5. KHÔNG tự suy diễn, KHÔNG mở rộng kiến thức ngoài bài. "
    prompt += "Nếu thiếu dữ liệu, ghi [Cần bổ sung SGK] thay vì bịa.\n\n"

    prompt += "=== ĐẶC TRƯNG MÔN " + ten_mon_chuan.upper() + " ===\n\n"
    if rules.get("ky_hieu"):
        prompt += "🎨 KÝ HIỆU ĐẶC TRƯNG: " + rules["ky_hieu"] + "\n"
    if rules.get("hinh_minh_hoa"):
        prompt += ("🖼️ HÌNH MINH HỌA ĐẶC TRƯNG: " + rules["hinh_minh_hoa"]
                   + " (chỉ mô tả bằng lời khi cần, KHÔNG chèn ảnh nhị phân)\n")
    prompt += "\n"

    if rules.get("nang_luc_trong_tam"):
        prompt += "🎯 NĂNG LỰC TRỌNG TÂM:\n"
        for nl in rules["nang_luc_trong_tam"]:
            prompt += "  - " + nl + "\n"
        prompt += "\n"

    if rules.get("phuong_phap_uu_tien"):
        prompt += "📚 PHƯƠNG PHÁP ƯU TIÊN:\n"
        for pp in rules["phuong_phap_uu_tien"]:
            prompt += "  - " + pp + "\n"
        prompt += "\n"

    if rules.get("san_pham_hoc_tap"):
        prompt += "📝 SẢN PHẨM HỌC TẬP đặc trưng:\n"
        for sp in rules["san_pham_hoc_tap"]:
            prompt += "  - " + sp + "\n"
        prompt += "\n"

    if rules.get("danh_gia"):
        prompt += "✅ TIÊU CHÍ ĐÁNH GIÁ:\n"
        for dg in rules["danh_gia"]:
            prompt += "  - " + dg + "\n"
        prompt += "\n"

    if rules.get("loai_bai"):
        prompt += "📖 CÁC DẠNG BÀI thường gặp:\n"
        for lb in rules["loai_bai"]:
            prompt += "  - " + lb + "\n"
        prompt += "\n"

    if rules.get("dac_thu_hoat_dong"):
        prompt += "🔄 ĐẶC THÙ 4 HOẠT ĐỘNG:\n"
        dthd = rules["dac_thu_hoat_dong"]
        if dthd.get("mo_dau"):
            prompt += "  • Mở đầu: " + dthd["mo_dau"] + "\n"
        if dthd.get("trong_tam"):
            prompt += "  • Hình thành kiến thức: " + dthd["trong_tam"] + "\n"
        if dthd.get("luyen_tap"):
            prompt += "  • Luyện tập: " + dthd["luyen_tap"] + "\n"
        if dthd.get("van_dung"):
            prompt += "  • Vận dụng: " + dthd["van_dung"] + "\n"
        prompt += "\n"

    if rules.get("quy_tac_rieng"):
        prompt += "⚙️ QUY TẮC RIÊNG CỦA MÔN:\n"
        for qt in rules["quy_tac_rieng"]:
            prompt += "  - " + qt + "\n"
        prompt += "\n"

    if rules.get("khong_nen"):
        prompt += "🚫 KHÔNG NÊN: " + rules["khong_nen"] + "\n\n"

    if tpl_text:
        prompt += tpl_text + "\n"
    if src_text:
        prompt += src_text + "\n"

    prompt += "=== 🔍 QUY TRÌNH TỰ KIỂM TRA TRƯỚC KHI TRẢ LỜI ===\n"
    prompt += "Bước 1: Xác định CHƯƠNG và MẠCH KIẾN THỨC của bài.\n"
    prompt += "Bước 2: Liệt kê các khái niệm CHỈ thuộc bài này (không lan sang bài khác).\n"
    prompt += "Bước 3: Kiểm tra từng khái niệm có nằm trong SGK/PPCT đã cung cấp không.\n"
    prompt += "Bước 4: LOẠI BỎ mọi kiến thức chưa học hoặc ngoài phạm vi bài.\n"
    prompt += "Bước 5: Chỉ sau khi kiểm tra xong mới soạn nội dung.\n\n"

    prompt += "=== CẤU TRÚC BẮT BUỘC (Công văn 5512) ===\n\n"
    prompt += "**TÊN BÀI DẠY: " + plan.get("topic", "") + "**\n"
    prompt += "**Môn học: " + ten_mon_chuan + "; lớp: " + (plan.get("grade", "") or "6") + "**\n"
    prompt += "**Thời gian thực hiện: (" + (plan.get("period") or "1") + " tiết)**\n\n"

    prompt += "**I. Mục tiêu**\n"
    prompt += "1. Kiến thức:\n"
    prompt += "[2-4 đoạn văn mô tả kiến thức, KHÔNG dùng dấu gạch đầu dòng]\n"
    prompt += "2. Năng lực:\n"
    prompt += ("Năng lực riêng: [chọn ĐÚNG các năng lực trọng tâm của môn "
               + ten_mon_chuan + " đã liệt kê ở trên]\n")
    prompt += ("Năng lực chung: Năng lực tự chủ và tự học; năng lực giao tiếp "
               "và hợp tác; năng lực giải quyết vấn đề và sáng tạo.\n")
    prompt += "3. Phẩm chất:\n"
    prompt += ("Bồi dưỡng hứng thú học tập, ý thức làm việc nhóm, ý thức tìm tòi, "
               "khám phá và sáng tạo cho HS => độc lập, tự tin và tự chủ.\n\n")

    prompt += "**II. Thiết bị dạy học và học liệu**\n"
    prompt += ("**1. Đối với giáo viên:** Kế hoạch bài dạy, SGK, SGV, "
               "phiếu học tập, đồ dùng đặc thù môn " + ten_mon_chuan + ", "
               "tivi (hoặc máy vi tính) nếu cần trình chiếu.\n")
    prompt += ("**2. Đối với học sinh:** SGK, vở ghi, đồ dùng học tập, "
               "ôn tập kiến thức đã học.\n\n")

    prompt += "**III. Tiến trình dạy học**\n\n"

    # Hướng dẫn chung cho cột "Sản phẩm"
    prompt += ("⚠️ LƯU Ý BẮT BUỘC: Cột **Sản phẩm** trong MỌI bảng dưới đây "
               "phải có NỘI DUNG CỤ THỂ, bao gồm:\n")
    prompt += ("  - Kiến thức SGK (định nghĩa, quy tắc, công thức...) — trích rõ tr.XX.\n")
    prompt += ("  - Các bước giải mẫu (nếu là bài toán).\n")
    prompt += ("  - Ký hiệu toán Unicode (x², √, ≤, ƯCLN, BCNN...).\n")
    prompt += ("  - Mô tả hình ảnh minh họa bằng lời (VD: 'Hình 4.5 SGK tr.45 — sơ đồ tập hợp ƯC(12,30)').\n")
    prompt += ("  - Dự kiến câu trả lời cụ thể của HS.\n")
    prompt += ("KHÔNG được để ô Sản phẩm nào trống hoặc ghi chung chung.\n\n")

    prompt += "**1. Hoạt động 1: Mở đầu**\n"
    prompt += "**a) Mục tiêu:** [2 câu mô tả ngắn gọn, bám đặc thù môn]\n"
    prompt += "**b) Nội dung:** [1-2 câu mô tả ngắn gọn]\n"
    prompt += "**c) Tổ chức thực hiện và sản phẩm:**\n\n"
    prompt += "| Hoạt động của GV và HS | Sản phẩm |\n"
    prompt += "|---|---|\n"
    prompt += "| Bước 1: Chuyển giao nhiệm vụ:<br>GV đặt vấn đề...<br>GV đặt câu hỏi: ... | [VD: Dự đoán của HS về...]<br>[Ký hiệu/công thức liên quan: ...]<br>[Hình ảnh minh họa: ...] |\n"
    prompt += "| Bước 2: Thực hiện nhiệm vụ:<br>HS suy nghĩ, thảo luận... | [VD: HS ghi nhận xét vào vở: ...]<br>[Câu trả lời dự kiến: ...] |\n"
    prompt += "| Bước 3: Báo cáo, thảo luận:<br>GV gọi đại diện HS trả lời... | [VD: 2-3 HS phát biểu: ...]<br>[Nhận xét bổ sung: ...] |\n"
    prompt += "| Bước 4: Kết luận, nhận định:<br>GV đánh giá, dẫn dắt vào bài mới | [Chốt: ...]<br>[Kiến thức nền cần nhớ: ...] |\n\n"

    prompt += "**2. Hoạt động 2: Hình thành kiến thức mới**\n"
    prompt += "**a) Mục tiêu:** [mô tả]\n"
    prompt += "**b) Nội dung:** [mô tả]\n"
    prompt += "**c) Tổ chức thực hiện và sản phẩm:**\n\n"
    prompt += "| Hoạt động của GV và HS | Sản phẩm |\n"
    prompt += "|---|---|\n"
    prompt += ("| Bước 1: Chuyển giao nhiệm vụ:<br>GV hướng dẫn và yêu cầu HS thực hiện... "
               "| 🔹 **Kiến thức SGK (tr.XX):**<br>- Định nghĩa/quy tắc: ...<br>- Ký hiệu: ...<br>"
               "🔹 **Ví dụ mẫu 1:**<br>- Đề: ...<br>- Các bước giải:<br>  + Bước 1: ...<br>  + Bước 2: ...<br>"
               "  + Kết luận: ...<br>🔹 **Ví dụ mẫu 2:**<br>- Đề: ...<br>- Giải: ...<br>"
               "🔹 **Hình ảnh minh họa:** Hình X.Y SGK tr.XX — mô tả... |\n")
    prompt += ("| Bước 2: Thực hiện nhiệm vụ:<br>HS chú ý lắng nghe, thực hiện yêu cầu của GV. "
               "| - HS ghi định nghĩa/quy tắc vào vở.<br>- HS chép ví dụ mẫu.<br>"
               "- Dự kiến: HS có thể nhầm lẫn ở bước ... |\n")
    prompt += ("| Bước 3: Báo cáo, thảo luận:<br>HS thảo luận và phát biểu, nhận xét và bổ sung cho nhau "
               "| - 2-3 HS phát biểu lại định nghĩa.<br>- HS nhận xét bài làm của bạn.<br>"
               "- Câu hỏi gợi mở của GV: ... |\n")
    prompt += ("| Bước 4: Kết luận, nhận định:<br>GV chính xác hóa và gọi 1 HS nhắc lại nội dung chính "
               "| - Chốt kiến thức: [công thức/định lý dạng Unicode].<br>"
               "- Lưu ý sai sót thường gặp: ...<br>- Ký hiệu chuẩn: ... |\n\n")

    prompt += "**3. Hoạt động 3: Luyện tập**\n"
    prompt += "**a) Mục tiêu:** Học sinh củng cố lại kiến thức thông qua một số bài tập.\n"
    prompt += "**b) Nội dung:** HS dựa vào kiến thức đã học vận dụng làm BT.\n"
    prompt += "**c) Tổ chức thực hiện và sản phẩm:**\n\n"
    prompt += "| Hoạt động của GV và HS | Sản phẩm |\n"
    prompt += "|---|---|\n"
    prompt += ("| GV yêu cầu HS hoàn thành các bài tập: [số bài] (tr[XX] - SGK). "
               "| **Bài [số] (tr.XX SGK):**<br>- Đề: ...<br>- Lời giải chi tiết:<br>"
               "  + Bước 1: ...<br>  + Bước 2: ...<br>  + Kết quả: ...<br>"
               "**Bài [số+1]:** ...<br>🔹 **Hình vẽ minh họa:** ...<br>"
               "🔹 **Lỗi sai thường gặp:** ... |\n\n")

    prompt += "**4. Hoạt động 4: Vận dụng**\n"
    prompt += "**a) Mục tiêu:** Học sinh thực hiện làm bài tập vận dụng để củng cố kiến thức và áp dụng kiến thức vào thực tế đời sống.\n"
    prompt += "**b) Nội dung:** HS sử dụng SGK và vận dụng kiến thức đã học để hoàn thành bài tập.\n"
    prompt += "**c) Tổ chức thực hiện và sản phẩm:**\n\n"
    prompt += "| Hoạt động của GV và HS | Sản phẩm |\n"
    prompt += "|---|---|\n"
    prompt += ("| GV cho HS đọc và tìm hiểu...<br>GV yêu cầu HS hoàn thành... "
               "| **Bài [số] (tr.XX SGK):**<br>- Đề: ...<br>- Lời giải chi tiết:<br>"
               "  + Bước 1: ...<br>  + Bước 2: ...<br>  + Kết quả: ...<br>"
               "🔹 **Liên hệ thực tế:** ...<br>🔹 **Ứng dụng/ký hiệu:** ... |\n\n")

    prompt += "**Kết thúc bài — thêm mục * HƯỚNG DẪN VỀ NHÀ với các gạch đầu dòng:**\n"
    prompt += "- Đọc và ghi nhớ nội dung chính của bài.\n"
    prompt += "- Vận dụng kiến thức làm bài tập [số] (SGK-tr[XX]).\n"
    prompt += "- Chuẩn bị bài mới \"[tên bài tiếp theo]\".\n\n"

    prompt += "=== QUY TẮC BẮT BUỘC ===\n"
    prompt += "1. NỘI DUNG PHẢI LẤY TỪ NGUỒN SGK. KHÔNG bịa.\n"
    prompt += "2. KHÔNG dùng từ 'máy chiếu' → thay bằng 'tivi' hoặc 'máy vi tính'.\n"
    prompt += "3. Ký hiệu toán Unicode: x², x³, √, ≤, ≥, ∈, π, Ư (KHÔNG LaTeX).\n"
    prompt += "4. Phân số viết dạng a/b.\n"
    prompt += "5. Trong ô bảng: dùng <br> để xuống dòng. Mỗi bước phải xuống dòng riêng.\n"
    prompt += "6. Độ dài tối thiểu 1500 từ.\n"
    prompt += "7. KHÔNG viết '---', '===', '****', '####' trong bài.\n"
    prompt += "8. KHÔNG dùng HTML tag <div>, <span>, <p>.\n"
    prompt += "9. KHÔNG viết dòng 'BẢNG 2 CỘT:' hay 'Bảng 2 cột:'.\n"
    prompt += "10. KHÔNG viết phần 'Người soạn', 'Ký tên', 'Sông Đốc' ở cuối.\n"
    prompt += "11. KHÔNG dùng dấu gạch đầu dòng (-, +, •) ở đầu các đoạn văn.\n"
    prompt += ("12. **HOẠT ĐỘNG PHẢI ĐÚNG ĐẶC THÙ MÔN "
               + ten_mon_chuan + "** — không copy máy móc hoạt động của môn khác.\n")
    prompt += "13. Sản phẩm học tập cuối bài phải thuộc danh sách sản phẩm đặc trưng của môn.\n"
    prompt += ("14. Nếu cột 'lớp' có nhiều lớp (VD: 6A10, 6A9, 6A8) thì giữ nguyên "
               "danh sách đó, KHÔNG rút gọn thành 1 lớp.\n")
    prompt += ("15. Trong phần 'Thiết bị dạy học' và các hoạt động, có thể gợi ý "
               "hình minh họa đặc trưng của môn bằng mô tả ngắn gọn. "
               "KHÔNG chèn ảnh nhị phân, chỉ mô tả bằng lời.\n")
    prompt += ("16. 🚫 **BÁM SÁT SGK**: Mọi kiến thức phải có trong NGUỒN SGK "
               "đã cung cấp. KHÔNG bịa, KHÔNG suy diễn, KHÔNG lấy từ bài khác.\n")
    prompt += ("17. 🚫 **ĐÚNG MẠCH KIẾN THỨC**: Xác định đúng chương/bài. "
               "Ví dụ: Chương II Toán 6 (Tính chia hết) KHÔNG có số nguyên. "
               "Chương III mới có số nguyên.\n")
    prompt += ("18. 🚫 **KHÔNG VƯỢT CẤP**: Chỉ dùng ngôn ngữ, khái niệm, ký hiệu "
               "phù hợp với lớp đang học. Không dùng kiến thức lớp trên.\n")
    prompt += ("19. 📖 **BÁM TRANG SGK**: Nếu có thông tin 'PHẠM VI SGK', "
               "chỉ được dùng kiến thức trong trang đó.\n")
    prompt += ("20. 📋 **CỘT SẢN PHẨM PHẢI ĐẦY ĐỦ**: Mọi ô trong cột 'Sản phẩm' "
               "phải chứa ít nhất 2 trong các mục sau: kiến thức SGK (tr.XX), "
               "các bước giải cụ thể, ký hiệu Unicode, mô tả hình ảnh minh họa, "
               "câu trả lời dự kiến của HS. KHÔNG để trống.\n")
    prompt += ("21. 🔢 **CÁC BƯỚC GIẢI**: Với bài có ví dụ/bài tập, phải trình bày "
               "đầy đủ các bước giải trong cột Sản phẩm (Bước 1: ..., Bước 2: ..., "
               "Kết luận: ...). KHÔNG chỉ ghi đáp án.\n")
    prompt += ("22. 🖼️ **MÔ TẢ HÌNH ẢNH**: Khi có hình vẽ, biểu đồ, sơ đồ SGK, "
               "phải mô tả bằng lời trong cột Sản phẩm (VD: 'Hình 4.5 tr.45 — "
               "sơ đồ Venn tập ƯC(12,30)'). KHÔNG chèn ảnh nhị phân.\n")
    prompt += ("23. 📐 **KÝ HIỆU ĐẶC TRƯNG MÔN**: Trong cột Sản phẩm, dùng ký hiệu "
               "đặc trưng của môn. VD môn Toán: ƯCLN, BCNN, ∈, ⊂, ⇒, ⇔, x², √, |x|.\n")

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
    text = re.sub(r"\$([^\$]+)\$", r"\1", text)
    return text


def _set_run_font(run, size=14, bold=False):
    run.bold = bold
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), "Times New Roman")
    rfonts.set(qn("w:hAnsi"), "Times New Roman")
    rfonts.set(qn("w:eastAsia"), "Times New Roman")
    rfonts.set(qn("w:cs"), "Times New Roman")


def add_para(doc, text, style=None):
    text = text.replace("<br>", "\n").replace("<br/>", "\n")
    text = re.sub(r"^[\-\+\*\•\#]+\s*", "", text)
    text = re.sub(r"^\d+[\.\)]\s+", "", text)
    text = re.sub(r"^[a-z]\)\s+", "", text)
    text = text.replace("**", "").replace("__", "")
    text = re.sub(r"(?<!\*)\*(?!\*)", "", text)
    text = text.strip()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Inches(0.39)
    p.paragraph_format.left_indent = Inches(0)
    p.paragraph_format.right_indent = Inches(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(text)
    _set_run_font(run, 14, bold=False)
    return p


def _tach_dong_trong_o(val):
    """Tách nội dung 1 ô bảng thành nhiều dòng hợp lý để justify không bị dãn xấu."""
    val = val.replace("<br>", "\n").replace("<br/>", "\n").replace("<br />", "\n")
    val = re.sub(r"(?<!^)(?<!\n)(?=Bước\s*\d+\s*:)", "\n", val)

    chu_hoa = (r"[A-ZĐÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊ"
               r"ÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴ]")
    val = re.sub(r"([\.\!\?\:])\s+(?=" + chu_hoa + r")", r"\1\n", val)

    keywords = ["Em hãy", "Vậy ", "Do đó", "Tuy nhiên", "Vì vậy",
                "GV yêu cầu", "GV gọi", "GV đánh giá", "GV chính xác",
                "GV hướng dẫn", "GV đặt", "GV chiếu",
                "HS suy nghĩ", "HS thảo luận", "HS thực hiện",
                "HS báo cáo", "HS nhận xét", "HS trả lời",
                "HS chú ý", "HS hoàn thành"]
    for kw in keywords:
        val = re.sub(r"(?<=[\.\!\?])\s+" + re.escape(kw), "\n" + kw, val)

    raw_lines = [ln.strip() for ln in val.split("\n")]
    raw_lines = [ln for ln in raw_lines if ln]

    merged = []
    for ln in raw_lines:
        if merged and len(ln) < 12 and not ln.startswith("Bước"):
            merged[-1] = merged[-1] + " " + ln
        else:
            merged.append(ln)
    return "\n".join(merged)


def add_table(doc, rows):
    if not rows:
        return
    cleaned_rows = []
    for row in rows:
        while len(row) > 2 and len(row[-1].strip()) <= 2 and row[-1].strip() in (
            "", "=", ".", ";", ",", "a", "b", "c", "5", "0", "1", "2",
            "3", "4", "6", "7", "8", "9", "|",
        ):
            row = row[:-1]
        cleaned_rows.append(row)
    max_n = max(len(r) for r in cleaned_rows)
    if max_n > 2:
        for i, r in enumerate(cleaned_rows):
            if len(r) > 2:
                extra = " ".join(c.strip() for c in r[2:] if c.strip())
                cleaned_rows[i] = [r[0], r[1] + (" " + extra if extra else "")]
        max_n = 2
    rows = cleaned_rows
    n = max_n
    if n < 1:
        return

    t = doc.add_table(rows=len(rows), cols=n)
    t.style = "Table Grid"

    for i, row in enumerate(rows):
        for j in range(n):
            c = t.cell(i, j)
            c.text = ""
            p = c.paragraphs[0]
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.first_line_indent = Inches(0)
            p.paragraph_format.left_indent = Inches(0)
            p.paragraph_format.right_indent = Inches(0)

            is_header = (i == 0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if is_header \
                else WD_ALIGN_PARAGRAPH.JUSTIFY

            val = row[j] if j < len(row) else ""

            if is_header:
                val = re.sub(r"^[\-\+\*\•\#]+\s*", "", val)
                val = re.sub(r"^\d+[\.\)]\s+", "", val)
                val = re.sub(r"^[a-z]\)\s+", "", val)
                val = val.replace("**", "").strip()
                run = p.add_run(val)
                _set_run_font(run, 13, bold=True)
            else:
                val = _tach_dong_trong_o(val)
                lines_clean = []
                for ln in val.split("\n"):
                    ln = re.sub(r"^[\-\+\*\•\#]+\s*", "", ln)
                    ln = re.sub(r"^\d+[\.\)]\s+", "", ln)
                    ln = re.sub(r"^[a-z]\)\s+", "", ln)
                    ln = re.sub(r"\s{2,}", " ", ln).strip()
                    if ln:
                        lines_clean.append(ln)

                first = True
                for line_content in lines_clean:
                    if first:
                        target_p = p
                        first = False
                    else:
                        target_p = c.add_paragraph()
                        target_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        target_p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                        target_p.paragraph_format.space_before = Pt(0)
                        target_p.paragraph_format.space_after = Pt(2)
                        target_p.paragraph_format.first_line_indent = Inches(0)
                        target_p.paragraph_format.left_indent = Inches(0)
                        target_p.paragraph_format.right_indent = Inches(0)

                    for part in re.split(r"(\*\*[^\*]+\*\*)", line_content):
                        if part.startswith("**") and part.endswith("**") and len(part) > 4:
                            run = target_p.add_run(part[2:-2])
                            _set_run_font(run, 13, bold=True)
                        else:
                            run = target_p.add_run(part.replace("**", ""))
                            _set_run_font(run, 13, bold=False)


def _add_heading(doc, text):
    text = re.sub(r"^[\*#\-\+\\]+\s*", "", text)
    text = text.rstrip("*").strip()

    p = doc.add_paragraph()
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)

    is_title_ten_bai = bool(re.match(r"^TÊN\s*BÀI\s*DẠY\s*:", text, re.IGNORECASE))
    is_title_mon_hoc = bool(re.match(r"^Môn\s*học\s*:", text, re.IGNORECASE))
    is_title_thoi_gian = bool(re.match(r"^Thời\s*gian\s*thực\s*hiện\s*:", text, re.IGNORECASE))

    if is_title_ten_bai or is_title_mon_hoc or is_title_thoi_gian:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Inches(0)
        p.paragraph_format.left_indent = Inches(0)
        p.paragraph_format.right_indent = Inches(0)

        # Với dòng "Môn học: Toán; lớp: 6A10, 6A9"
        if is_title_mon_hoc:
            idx = text.index(":")
            label = text[:idx + 1].strip()
            rest = text[idx + 1:].strip()

            run_label = p.add_run(label + " ")
            _set_run_font(run_label, 14, bold=True)

            # Tách "Toán; lớp: 6A10, 6A9, 6A8"
            m_mon = re.match(r"^([^;]+);\s*(lớp\s*:\s*)(.+)$", rest, re.IGNORECASE)
            if m_mon:
                _mon_text = m_mon.group(1).strip()
                _lop_label = m_mon.group(2).strip()
                _lop_value = m_mon.group(3).strip()

                run_mon = p.add_run(_mon_text + "; ")
                _set_run_font(run_mon, 14, bold=False)

                run_lop_label = p.add_run(_lop_label + " ")
                _set_run_font(run_lop_label, 14, bold=False)

                run_lop_value = p.add_run(_lop_value)
                _set_run_font(run_lop_value, 14, bold=True)
            else:
                run_rest = p.add_run(rest)
                _set_run_font(run_rest, 14, bold=False)
            return p

        # Với "TÊN BÀI DẠY:" và "Thời gian thực hiện:"
        if ":" in text:
            idx = text.index(":")
            label = text[:idx + 1].strip()
            content = text[idx + 1:].strip()
        else:
            label = text
            content = ""

        run_label = p.add_run(label + " ")
        _set_run_font(run_label, 14, bold=True)
        if content:
            run_content = p.add_run(content)
            _set_run_font(run_content, 14, bold=False)
        return p

    if re.match(r"^[a-z]\)\s", text):
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Inches(0.39)
        p.paragraph_format.left_indent = Inches(0)
        p.paragraph_format.right_indent = Inches(0)
        if ":" in text:
            idx_colon = text.index(":")
            heading_part = text[:idx_colon + 1]
            content_part = text[idx_colon + 1:]
            heading_part = re.sub(r"^\*\*", "", heading_part).strip()
            heading_part = re.sub(r"\*\*$", "", heading_part).strip()
            content_part = re.sub(r"^\*\*\s*", "", content_part)
            content_part = re.sub(r"\*\*$", "", content_part).strip()
            run1 = p.add_run(heading_part)
            _set_run_font(run1, 14, bold=True)
            run1.italic = True
            if content_part:
                if not content_part.startswith(" "):
                    content_part = " " + content_part
                run2 = p.add_run(content_part)
                _set_run_font(run2, 14, bold=False)
                run2.italic = True
        else:
            clean = re.sub(r"\*\*", "", text).strip()
            run = p.add_run(clean)
            _set_run_font(run, 14, bold=True)
            run.italic = True
        return p

    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Inches(0.39)
    p.paragraph_format.left_indent = Inches(0)
    p.paragraph_format.right_indent = Inches(0)
    run = p.add_run(text)
    _set_run_font(run, 14, bold=True)
    return p


def write_content_to_doc(d, content):
    content = fix_math(content)
    content = re.sub(r"^\s*BẢNG\s*2\s*CỘT\s*:?\s*$", "", content, flags=re.MULTILINE)
    content = re.sub(r"^\s*Bảng\s*2\s*cột\s*:?\s*$", "", content, flags=re.MULTILINE)
    for tag in ["div", "strong", "span", "p", "em"]:
        content = re.sub(r"<" + tag + r"[^>]*>", "", content)
        content = re.sub(r"</" + tag + r">", "", content)
    content = re.sub(r"^\s*[-=]{3,}\s*$", "", content, flags=re.MULTILINE)

    lines = content.splitlines()
    start_idx = 0
    for idx, line in enumerate(lines):
        s = line.strip()
        if re.match(r"^\*\*\s*TÊN\s*BÀI\s*DẠY", s, re.IGNORECASE):
            start_idx = idx
            break
        if re.match(r"^\*\*\s*I\.\s*(Mục|MỤC)\s*tiêu", s, re.IGNORECASE):
            start_idx = idx
            break
    lines = lines[start_idx:]

    cleaned = []
    for line in lines:
        s = line.strip()
        if re.match(r"^Sông\s+Đốc", s, re.IGNORECASE):
            break
        if re.search(r"NGƯỜI\s*SOẠN", s, re.IGNORECASE):
            break
        if re.search(r"KÝ\s*DUYỆT", s, re.IGNORECASE):
            break
        if re.match(r"^\(?Ký\s*tên\)?\s*$", s, re.IGNORECASE):
            break
        cleaned.append(line)
    content = "\n".join(cleaned)

    lines = content.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        stripped = line.strip()

        if stripped == "":
            if i > 0 and lines[i - 1].strip() != "":
                p = d.add_paragraph("")
                p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
            i += 1
            continue

        if stripped.startswith("|") and stripped.endswith("|"):
            tl = []
            j = i
            while j < len(lines) and lines[j].strip().startswith("|") \
                    and lines[j].strip().endswith("|"):
                tl.append(lines[j].strip())
                j += 1
            is_valid_table = False
            if len(tl) >= 2:
                if re.match(r"^\|[\s\-:\|]+\|$", tl[1]):
                    is_valid_table = True
            if is_valid_table:
                rows = []
                for t in tl:
                    if re.match(r"^\|[\s\-:\|]+\|$", t):
                        continue
                    cells = [c.strip() for c in t.strip("|").split("|")]
                    if 2 <= len(cells) <= 4:
                        rows.append(cells)
                if rows:
                    add_table(d, rows)
                else:
                    for t in tl:
                        text_line = t.strip("|").strip()
                        if text_line:
                            add_para(d, text_line)
            else:
                for t in tl:
                    text_line = t.strip("|").strip()
                    if text_line:
                        add_para(d, text_line)
            i = j
            continue

        if re.match(r"^\*\*\s*TÊN\s*BÀI\s*DẠY", stripped, re.IGNORECASE):
            _add_heading(d, stripped.replace("**", "").replace("*", "").strip())
            i += 1
            continue

        if re.match(r"^\*\*\s*Môn\s*học", stripped, re.IGNORECASE):
            _add_heading(d, stripped.replace("**", "").replace("*", "").strip())
            i += 1
            continue

        if re.match(r"^\*\*\s*Thời\s*gian\s*thực\s*hiện", stripped, re.IGNORECASE):
            _add_heading(d, stripped.replace("**", "").replace("*", "").strip())
            i += 1
            continue

        if re.match(r"^\*\*\s*(I|II|III|IV|V)\.\s+\S", stripped):
            _add_heading(d, stripped.replace("**", "").replace("*", "").strip())
            i += 1
            continue

        if re.match(r"^\*\*\s*\d+\.\s+[A-Za-zÀ-ỹ]", stripped) and len(stripped) < 120:
            _add_heading(d, stripped.replace("**", "").replace("*", "").strip())
            i += 1
            continue

        if re.match(r"^\*\*\s*[a-z]\)\s+\S", stripped):
            _add_heading(d, stripped.replace("**", "").replace("*", "").strip())
            i += 1
            continue

        if stripped.startswith("#### "):
            _add_heading(d, re.sub(r"[\*#\\]", "", stripped[5:].strip()))
        elif stripped.startswith("### "):
            _add_heading(d, re.sub(r"[\*#\\]", "", stripped[4:].strip()))
        elif stripped.startswith("## "):
            _add_heading(d, re.sub(r"[\*#\\]", "", stripped[3:].strip()))
        elif stripped.startswith("# "):
            _add_heading(d, re.sub(r"[\*#\\]", "", stripped[2:].strip()))
        elif re.match(r"^[-\*+\•]\s+", stripped):
            add_para(d, stripped)
        elif re.match(r"^\d+[\.\)]\s+", stripped):
            add_para(d, stripped)
        else:
            text = re.sub(r"^[\*#\-\+\\]+\s*", "", stripped)
            text = re.sub(r"^\*\*\\?\*?\s*", "", text)
            text = re.sub(r"\s*\*\*\\?\*?$", "", text)
            if text:
                add_para(d, text)
        i += 1


def make_doc(title, body, subject_name=""):
    """Tạo file Word. ĐÃ BỎ dòng ký hiệu môn ở đầu."""
    d = Document()
    s = d.sections[0]
    s.top_margin = Inches(0.7)
    s.bottom_margin = Inches(0.7)
    s.left_margin = Inches(0.8)
    s.right_margin = Inches(0.8)
    style = d.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(14)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), "Times New Roman")
    rfonts.set(qn("w:hAnsi"), "Times New Roman")
    rfonts.set(qn("w:eastAsia"), "Times New Roman")
    rfonts.set(qn("w:cs"), "Times New Roman")
    pf = style.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)

    d.add_heading(title, 0)
    # ❌ KHÔNG thêm dòng ký hiệu môn ở đây nữa

    mem = st.session_state.get("memory", {})
    if mem.get("school"):
        d.add_paragraph("Đơn vị: " + mem["school"])
    if mem.get("teacher_name"):
        d.add_paragraph("Người thực hiện: " + mem["teacher_name"])
    d.add_paragraph("Ngày tạo: " + datetime.now().strftime("%d/%m/%Y %H:%M"))
    d.add_paragraph("")
    write_content_to_doc(d, body)
    b = io.BytesIO()
    d.save(b)
    return b.getvalue()


def add_phieu_bao_giang(doc, lich, tuan_start, buoi, tuan_num):
    mem = st.session_state.get("memory", {})
    now = datetime.now()
    end_date = tuan_start + timedelta(days=6)

    table_head = doc.add_table(rows=1, cols=2)
    table_head.autofit = True
    cell = table_head.cell(0, 0)
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run("UBND XÃ SÔNG ĐỐC")
    _set_run_font(run, 13, bold=False)
    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(0)
    run = p2.add_run("TRƯỜNG THCS SÔNG ĐỐC")
    _set_run_font(run, 13, bold=True)

    cell_r = table_head.cell(0, 1)
    cell_r.text = ""
    p = cell_r.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    _set_run_font(run, 13, bold=True)
    p2 = cell_r.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(0)
    run = p2.add_run("Độc lập -- Tự do -- Hạnh phúc")
    _set_run_font(run, 13, bold=True)

    p_date = doc.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p_date.add_run("Sông Đốc, ngày " + now.strftime("%d")
                          + " tháng " + now.strftime("%m")
                          + " năm " + now.strftime("%Y"))
    _set_run_font(run, 13, bold=False)
    run.italic = True

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_title.add_run("PHIẾU BÁO GIẢNG ")
    _set_run_font(run, 14, bold=True)
    run = p_title.add_run(buoi)
    _set_run_font(run, 14, bold=True)
    run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)

    p_tuan = doc.add_paragraph()
    p_tuan.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_tuan.add_run("Tuần " + str(tuan_num)
                          + ", năm học 2026-2027 "
                          + "(Từ " + tuan_start.strftime("%d/%m/%Y")
                          + " đến " + end_date.strftime("%d/%m/%Y") + ")")
    _set_run_font(run, 13, bold=False)
    run.italic = True

    doc.add_paragraph("")

    table = doc.add_table(rows=1, cols=7)
    table.style = "Table Grid"
    table.autofit = False
    widths = [Cm(2.5), Cm(1.2), Cm(1.5), Cm(1.5), Cm(1.3), Cm(9.0), Cm(1.5)]

    hdr = table.rows[0].cells
    headers_lines = [
        ["Thứ", "ngày, tháng"],
        ["Tiết", "TKB"],
        ["Môn"],
        ["Lớp"],
        ["Tiết", "PP CT"],
        ["Tên bài dạy"],
        ["Ghi", "chú"],
    ]
    for i, lines in enumerate(headers_lines):
        hdr[i].text = ""
        first = True
        for line in lines:
            if first:
                p = hdr[i].paragraphs[0]
                first = False
            else:
                p = hdr[i].add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(line)
            _set_run_font(run, 13, bold=True)
        hdr[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    for p in hdr[6].paragraphs:
        for r in p.runs:
            r.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
    for i, w in enumerate(widths):
        table.rows[0].cells[i].width = w

    nhom_ngay = []
    for r in lich:
        thu_clean = r["thu"].strip()
        ngay_clean = r.get("ngay", "").strip()
        key = (thu_clean + "|" + ngay_clean).lower().replace(" ", "")
        if not nhom_ngay or nhom_ngay[-1]["key"] != key:
            nhom_ngay.append({
                "key": key,
                "thu": thu_clean,
                "ngay": ngay_clean,
                "items": [r],
            })
        else:
            nhom_ngay[-1]["items"].append(r)

    for nhom in nhom_ngay:
        for idx, r in enumerate(nhom["items"]):
            row = table.add_row().cells
            if idx == 0:
                row[0].text = ""
                p = row[0].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                run = p.add_run(nhom["thu"])
                _set_run_font(run, 13, bold=False)
                p2 = row[0].add_paragraph()
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p2.paragraph_format.space_before = Pt(0)
                p2.paragraph_format.space_after = Pt(0)
                run = p2.add_run(nhom["ngay"])
                _set_run_font(run, 13, bold=False)
            else:
                row[0].text = ""

            for c_idx, key, align in [
                (1, "tiet", WD_ALIGN_PARAGRAPH.CENTER),
                (2, "mon", WD_ALIGN_PARAGRAPH.CENTER),
                (3, "lop", WD_ALIGN_PARAGRAPH.CENTER),
                (4, "stt_ppct", WD_ALIGN_PARAGRAPH.CENTER),
                (5, "ten_bai", WD_ALIGN_PARAGRAPH.LEFT),
            ]:
                row[c_idx].text = ""
                p = row[c_idx].paragraphs[0]
                p.alignment = align
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.first_line_indent = Inches(0)
                p.paragraph_format.left_indent = Inches(0)
                val = r.get(key, "")
                if key == "tiet":
                    val = str(val).replace("Tiết ", "")
                if key == "stt_ppct":
                    m_num = re.search(r"\d+", str(val))
                    val = m_num.group(0) if m_num else str(val).strip()
                if key == "ten_bai":
                    _trang = str(r.get("trang_sgk", "")).strip()
                    if _trang:
                        val = str(val) + " (tr. " + _trang + ")"
                run = p.add_run(str(val))
                _set_run_font(run, 13, bold=False)
            row[6].text = ""

    start_row_idx = 1
    for nhom in nhom_ngay:
        so_dong = len(nhom["items"])
        if so_dong > 1:
            try:
                cell_start = table.cell(start_row_idx, 0)
                cell_end = table.cell(start_row_idx + so_dong - 1, 0)
                merged = cell_start.merge(cell_end)
                merged.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for p in merged.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            except Exception:
                pass
        start_row_idx += so_dong

    for row in table.rows:
        for i, w in enumerate(widths):
            if i < len(row.cells):
                row.cells[i].width = w

    doc.add_paragraph("")


# ============ SOẠN 1 BÀI ============
st.divider()
st.subheader("💬 Soạn 1 bài theo lệnh")

_lich_hien_tai = st.session_state.get("lich_bao_giang", [])
_lop_tu_lich = []
for _r in _lich_hien_tai:
    _l = _r.get("lop", "").strip()
    if _l and _l not in _lop_tu_lich:
        _lop_tu_lich.append(_l)
_lop_goi_y = ", ".join(_lop_tu_lich) if _lop_tu_lich else ""

_col_mon, _col_lop, _col_tiet, _col_trang = st.columns([2, 1.5, 0.8, 1])
with _col_mon:
    _danh_sach_mon = list(SUBJECT_RULES.keys())
    _options_hien_thi = [
        SUBJECT_RULES[m].get("ky_hieu", "📚") + " " + m
        for m in _danh_sach_mon
    ]
    _default_mon = st.session_state["memory"].get("default_subject", "Toán")
    _idx_default = 0
    for _i, _m in enumerate(_danh_sach_mon):
        if _chuan_hoa_ten_mon(_m) == _chuan_hoa_ten_mon(_default_mon):
            _idx_default = _i
            break

    _mon_hien_thi = st.selectbox("📚 Môn học", _options_hien_thi,
                                   index=_idx_default, key="chon_mon_soan")
    _mon_chon = _danh_sach_mon[_options_hien_thi.index(_mon_hien_thi)]
with _col_lop:
    _lop_mac_dinh = _lop_goi_y or st.session_state["memory"].get("default_grade", "6")
    _lop_chon = st.text_input("🎓 Lớp (trống = lấy từ lịch BG)",
                               value=_lop_mac_dinh,
                               key="chon_lop_soan",
                               help="VD: 6A10, 6A9, 6A8.")
with _col_tiet:
    _tiet_chon = st.text_input("⏱️ Số tiết", value="", key="chon_tiet_soan")
with _col_trang:
    _trang_chon = st.text_input("📖 Trang SGK", value="",
                                 key="chon_trang_soan",
                                 help="VD: 44-46.")

_rules_hien_tai = SUBJECT_RULES.get(_mon_chon, {})
if _rules_hien_tai:
    _ky_hieu = _rules_hien_tai.get("ky_hieu", "")
    _hinh_mh = _rules_hien_tai.get("hinh_minh_hoa", "")
    if _ky_hieu or _hinh_mh:
        st.markdown(
            f"<div style='background: linear-gradient(90deg,#e0f7ff,#f0f9ff); "
            f"padding: 10px 14px; border-radius: 8px; border-left: 4px solid #0ea5e9; "
            f"margin-bottom: 10px;'>"
            f"<b style='font-size: 18px;'>{_ky_hieu}</b> "
            f"<b>Môn {_mon_chon}</b> — "
            f"<i>Hình minh họa đặc trưng: {_hinh_mh}</i>"
            f"</div>",
            unsafe_allow_html=True,
        )

cmd = st.text_area("Lệnh", height=100,
                    placeholder="VD: Soạn bài Ước chung lớn nhất...",
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
        _lop_dung = _lop_chon.strip()
        if not _lop_dung:
            _cac_lop_set = []
            for _r in _lich_hien_tai:
                _l = _r.get("lop", "").strip()
                if _l and _l not in _cac_lop_set:
                    _cac_lop_set.append(_l)
            _lop_dung = ", ".join(_cac_lop_set) if _cac_lop_set else \
                        st.session_state.get("memory", {}).get("default_grade", "6")

        _trang_tu_lich = ""
        for _r_lich in _lich_hien_tai:
            _tb_lich = _r_lich.get("ten_bai", "").strip().lower()
            if _tb_lich and (_tb_lich in cmd.lower() or cmd.lower() in _tb_lich):
                _trang_tu_lich = str(_r_lich.get("trang_sgk", "")).strip()
                if _trang_tu_lich:
                    break

        _trang_dung = _trang_chon.strip() or _trang_tu_lich

        with st.spinner("🤖 Đang sinh nội dung môn " + _mon_chon + "..."):
            mem = st.session_state["memory"]
            plan = {
                "task": "Soạn kế hoạch bài dạy",
                "subject": _mon_chon,
                "grade": _lop_dung,
                "topic": cmd,
                "period": _tiet_chon,
                "trang_sgk": _trang_dung,
                "template_files": list(st.session_state["templates"].keys())[:1],
                "source_files": list(st.session_state["sources"].keys())[:1],
            }
            result = call_ai(build_exec_prompt(plan))
            if result:
                st.session_state["draft"] = result
                st.session_state["last_topic"] = _mon_chon + " - " + cmd
                save_history("Soạn bài " + _mon_chon, cmd, result)

if st.session_state.get("draft"):
    st.divider()
    st.subheader("📄 Kết quả")
    content = st.session_state["draft"]
    safe_name = re.sub(r"[^A-Za-z0-9_\-]", "_",
                        st.session_state.get("last_topic", "tai_lieu"))[:45] or "tai_lieu"
    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            "⬇️ Tải Word",
            make_doc("Tài liệu", content, subject_name=_mon_chon),
            file_name=safe_name + ".docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary", use_container_width=True,
        )
    with d2:
        st.download_button(
            "⬇️ Tải TXT",
            content.encode("utf-8"),
            file_name=safe_name + ".txt",
            mime="text/plain",
            use_container_width=True,
        )
    st.session_state["draft"] = st.text_area("Nội dung (có thể sửa)", content,
                                              height=500, key="edit_draft")

# ============================================================
# PHẦN 4/4: TẠO LỊCH BÁO GIẢNG + NẠP + HIỂN THỊ + SOẠN TUẦN
# ============================================================

st.divider()
st.subheader("📅 Lịch báo giảng")
st.caption("Chọn 1 trong 3 cách: tự sinh từ TKB+PPCT, nạp phiếu có sẵn, "
           "hoặc nạp từ kho đã lưu")

sub_t1, sub_t2, sub_t3 = st.tabs([
    "🔧 Tự sinh từ TKB + PPCT",
    "✏️ Nạp phiếu báo giảng có sẵn",
    "📂 Nạp lịch đã lưu",
])


# ==================== TAB 1: TỰ SINH ====================
with sub_t1:
    st.markdown("### 🔧 Tự sinh lịch từ Thời khóa biểu + Phân phối chương trình")
    st.info(
        "App đọc TKB (5 cột: **Thứ | Tiết | Lớp | Môn | Tiết PPCT**) "
        "và PPCT (STT = Tiết PPCT, có thể có cột **Trang SGK**). "
        "Tra tên bài + trang SGK theo cột Tiết PPCT của TKB."
    )

    st.markdown("**Chọn từ kho đã lưu:**")
    col_ppct, col_tkb = st.columns(2)
    with col_ppct:
        ppct_options = ["(Không dùng)"] + list(st.session_state["ppct"].keys())
        selected_ppct = st.selectbox("📗 PPCT đã lưu:", ppct_options, key="sel_ppct")
    with col_tkb:
        tkb_options = ["(Không dùng)"] + list(st.session_state["tkb"].keys())
        selected_tkb = st.selectbox("📘 TKB đã lưu:", tkb_options, key="sel_tkb")

    st.markdown("**Hoặc upload / dán:**")
    tkb_file = st.file_uploader(
        "📘 Upload TKB",
        type=["pdf", "docx", "xlsx", "xls", "csv", "txt"],
        key="tkb_file",
    )
    tkb_manual = st.text_area("Hoặc dán TKB", height=100, key="tkb_manual")
    ppct_file = st.file_uploader(
        "📗 Upload PPCT",
        type=["pdf", "docx", "xlsx", "xls", "csv", "txt"],
        key="ppct_file",
    )
    ppct_manual = st.text_area("Hoặc dán PPCT", height=100, key="ppct_manual")

    col1_d, col2_d, col3_d = st.columns(3)
    with col1_d:
        start_date = st.date_input(
            "📆 Ngày bắt đầu tuần (thứ 2)",
            value=datetime.now().date(),
            key="ngay_bd_w",
        )
    with col2_d:
        buoi_day = st.selectbox(
            "🌅 Buổi dạy", ["BUỔI CHIỀU", "BUỔI SÁNG"], key="buoi_w"
        )
    with col3_d:
        tuan_number = st.number_input(
            "📅 Số tuần", min_value=1, max_value=52, value=1, step=1, key="tuan_w"
        )

    def ai_extract_tkb(raw):
        system = "Bạn là trợ lý đọc bảng thời khóa biểu. CHỈ trả về text."
        prompt = """Đọc bảng TKB sau. Mỗi dòng dữ liệu có 5 cột:
Thứ | Tiết | Lớp | Môn | Tiết PPCT

QUY TẮC:
- Cột 1 "Thứ": "Thứ 2", "Thứ 3", ...
- Cột 2 "Tiết": tiết trong ngày (1, 2, 3, 4, 5)
- Cột 3 "Lớp": ví dụ 6A10, 6A9, 6A8
- Cột 4 "Môn": Toán, Văn, Anh...
- Cột 5 "Tiết PPCT": SỐ NGUYÊN (25, 26, 27, 28...)

QUAN TRỌNG: Nhiều dòng gốc bị dính liền nhau, ví dụ:
"216A10Toán25" → "Thứ 2 | Tiết 1 | 6A10 | Toán | 25"
"226A10Toán26" → "Thứ 2 | Tiết 2 | 6A10 | Toán | 26"
"546A10Toán27" → "Thứ 5 | Tiết 4 | 6A10 | Toán | 27"

CHỈ trả về 5 cột, cách nhau bằng dấu |.
KHÔNG bịa thêm. Bỏ dòng tiêu đề.

NỘI DUNG TKB:
""" + raw[:15000]
        try:
            result = call_ai(prompt, system)
            lines = []
            for l in result.splitlines():
                l = l.strip()
                if not l or "|" not in l:
                    continue
                parts = [p.strip() for p in l.split("|")]
                if len(parts) < 5:
                    if len(parts) == 4:
                        parts.append("")
                    else:
                        continue
                if parts[0].lower().startswith("thứ") and \
                   "tiết" in parts[1].lower() and "lớp" in parts[2].lower():
                    continue
                lines.append(" | ".join(parts[:5]))
            return "\n".join(lines)
        except Exception:
            return ""

    def ai_extract_ppct(raw):
        system = "Bạn là trợ lý đọc bảng phân phối chương trình."
        prompt = """Đọc PPCT sau. Các cột có thể có:
STT | Tên bài | Số tiết | Thời điểm (Tuần) | Trang SGK | Ghi chú

QUY TẮC QUAN TRỌNG:
1. Cột "STT" = Tiết PPCT (số thứ tự tiết toàn cấp). GIỮ NGUYÊN SỐ.
2. Cột "Thời điểm" chứa text dạng: "Tuần: 6 (12/10/2026 - 18/10/2026)"
   - Trích số tuần: 6
   - Trích ngày bắt đầu: 12/10/2026
   - Trích ngày kết thúc: 18/10/2026
3. Cột "Trang SGK" có thể có dạng: "44", "44-46", "tr. 44-46".
   - Trích đúng số trang, bỏ chữ "tr.", "trang".
   - Nếu PPCT không có cột này, ghi rỗng.
4. Nếu cột Thời điểm bị GỘP Ô, PHẢI ĐIỀN tuần cho TẤT CẢ các bài thuộc tuần đó.
5. Các dòng "Bài X (Tiết 1)", "Bài X (Tiết 2)" → GỘP thành 1 bài, cộng số tiết.
6. Trả về MỖI BÀI 1 dòng, 7 cột:
   STT | Tên bài | Số tiết | Tuần | Ngày BĐ | Ngày KT | Trang SGK
   (Ngày định dạng dd/mm/yyyy, Trang SGK giữ dạng "44-46")

VÍ DỤ 1 (có cột Trang SGK):
Input:
"21 | Bài 11. Ước chung. Ước chung lớn nhất (Tiết 1) | 1 | Tuần: 6 (12/10/2026 - 18/10/2026) | 44-46 |"
"22 | Ước chung. Ước chung lớn nhất (Tiết 2) | 1 | (gộp) | |"
Output:
"21 | Bài 11. Ước chung. Ước chung lớn nhất | 2 | 6 | 12/10/2026 | 18/10/2026 | 44-46"

VÍ DỤ 2 (không có cột Trang SGK):
Input:
"21 | Bài 11. Ước chung. Ước chung lớn nhất (Tiết 1) | 1 | Tuần: 6 (12/10/2026 - 18/10/2026) |"
"22 | Ước chung. Ước chung lớn nhất (Tiết 2) | 1 | (gộp) |"
Output:
"21 | Bài 11. Ước chung. Ước chung lớn nhất | 2 | 6 | 12/10/2026 | 18/10/2026 | "

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
        except Exception:
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
            m = re.match(r"^(\d)(\d)(6A\d+)([^\d]+?)(\d+)$", line)
            if m:
                fixed.append(
                    "Thứ " + m.group(1) +
                    " | Tiết " + m.group(2) +
                    " | " + m.group(3) +
                    " | " + m.group(4).strip() +
                    " | " + m.group(5)
                )
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
            if parts[0].lower() in ("thứ", "thu") and "tiết" in parts[1].lower():
                continue
            thu = parts[0]
            if not thu.lower().startswith("thứ"):
                thu = "Thứ " + thu
            tiet = parts[1]
            if not tiet.lower().startswith("tiết"):
                num = re.search(r"\d+", tiet)
                tiet = "Tiết " + (num.group(0) if num else tiet)
            tiet_ppct = ""
            if len(parts) >= 5:
                m = re.search(r"\d+", parts[4])
                if m:
                    tiet_ppct = m.group(0)
            if thu not in result:
                result[thu] = []
            result[thu].append({
                "tiet": tiet,
                "lop": parts[2],
                "mon": parts[3],
                "tiet_ppct": tiet_ppct,
            })
        for thu in result:
            result[thu].sort(
                key=lambda x: int(re.search(r"\d+", x["tiet"]).group(0))
                if re.search(r"\d+", x["tiet"]) else 0
            )
        return result

    def parse_ppct(text):
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
            tuan_num = ""
            m_t = re.search(r"(?:Tuần|tuần|Tuan|tuan)\s*[:\-]?\s*(\d+)", tuan)
            if m_t:
                tuan_num = m_t.group(1)
            elif tuan.isdigit():
                tuan_num = tuan
            ts = parts[4].strip() if len(parts) >= 5 else ""
            te = parts[5].strip() if len(parts) >= 6 else ""
            trang = parts[6].strip() if len(parts) >= 7 else ""
            trang = re.sub(r"(?i)^(tr\.?|trang)\s*", "", trang).strip()

            raw_list.append({
                "stt": stt,
                "ten_bai": ten_bai,
                "so_tiet": n_tiet,
                "tuan": tuan_num,
                "ngay_bd": ts,
                "ngay_kt": te,
                "trang_sgk": trang,
            })
        merged = []
        i = 0
        while i < len(raw_list):
            bai = raw_list[i]
            ten_chuan = re.sub(r"\s*\(Tiết\s*\d+\)", "", bai["ten_bai"]).strip()
            so_tiet_gop = 0
            j = i
            while j < len(raw_list):
                ten_j = re.sub(r"\s*\(Tiết\s*\d+\)", "",
                               raw_list[j]["ten_bai"]).strip()
                if ten_j == ten_chuan and \
                   str(raw_list[j]["tuan"]) == str(bai["tuan"]):
                    so_tiet_gop += raw_list[j]["so_tiet"]
                    j += 1
                else:
                    break
            merged.append({
                "stt": bai["stt"],
                "ten_bai": ten_chuan,
                "so_tiet": so_tiet_gop if so_tiet_gop > 0 else bai["so_tiet"],
                "tuan": bai["tuan"],
                "ngay_bd": bai["ngay_bd"],
                "ngay_kt": bai["ngay_kt"],
                "trang_sgk": bai.get("trang_sgk", ""),
            })
            i = j if so_tiet_gop > 0 else i + 1
        return merged

    def find_ppct_by_tiet(ppct, tiet_ppct):
        if tiet_ppct is None or str(tiet_ppct).strip() == "":
            return None
        try:
            target = int(str(tiet_ppct).strip())
        except Exception:
            return None
        for bai in ppct:
            stt = bai.get("stt", 0)
            so_tiet = bai.get("so_tiet", 1)
            if stt <= target < stt + so_tiet:
                return {
                    "ten_bai": bai["ten_bai"],
                    "tiet_trong_bai": target - stt + 1,
                    "tong_tiet": so_tiet,
                    "stt_ppct": target,
                    "trang_sgk": bai.get("trang_sgk", ""),
                }
        return None

    if st.button("📅 Tạo lịch báo giảng", type="primary",
                 use_container_width=True):
        if not available_providers:
            st.error("Cần API key.")
        else:
            tkb_raw = tkb_manual.strip()
            if tkb_file:
                _, text, _ = read_file(tkb_file, use_ocr=False)
                tkb_raw = text
            elif not tkb_raw and selected_tkb != "(Không dùng)":
                tkb_raw = st.session_state["tkb"][selected_tkb]

            ppct_raw = ppct_manual.strip()
            if ppct_file:
                _, text, _ = read_file(ppct_file, use_ocr=False)
                ppct_raw = text
            elif not ppct_raw and selected_ppct != "(Không dùng)":
                ppct_raw = st.session_state["ppct"][selected_ppct]

            if not tkb_raw or not ppct_raw:
                st.error("Cần TKB và PPCT.")
            else:
                tkb_auto = auto_fix_tkb(tkb_raw)
                first_line = tkb_auto.splitlines()[0] if tkb_auto else ""
                if tkb_auto and "|" in first_line and \
                   len(first_line.split("|")) >= 5:
                    tkb_clean = tkb_auto
                    st.success("✅ TKB parse trực tiếp (không cần AI)")
                else:
                    with st.spinner("🧠 AI đọc TKB..."):
                        tkb_clean = ai_extract_tkb(tkb_raw)

                with st.spinner("🧠 AI đọc PPCT..."):
                    ppct_clean = ai_extract_ppct(ppct_raw)

                with st.expander("📋 Xem kết quả AI trích xuất", expanded=False):
                    st.markdown("**TKB:**")
                    st.code(tkb_clean)
                    st.markdown("**PPCT:**")
                    st.code(ppct_clean)

                tkb = parse_tkb(tkb_clean)
                ppct = parse_ppct(ppct_clean)

                if not tkb or not ppct:
                    st.error("Không parse được TKB hoặc PPCT.")
                else:
                    n_slots = sum(len(v) for v in tkb.values())
                    n_co_trang = sum(1 for b in ppct if b.get("trang_sgk"))
                    st.success(
                        "✅ TKB: " + str(n_slots) + " tiết | "
                        + "PPCT: " + str(len(ppct)) + " bài | "
                        + "Có trang SGK: " + str(n_co_trang) + "/" + str(len(ppct))
                    )

                    days_map = {
                        "thứ 2": 0, "thu 2": 0, "2": 0,
                        "thứ 3": 1, "thu 3": 1, "3": 1,
                        "thứ 4": 2, "thu 4": 2, "4": 2,
                        "thứ 5": 3, "thu 5": 3, "5": 3,
                        "thứ 6": 4, "thu 6": 4, "6": 4,
                        "thứ 7": 5, "thu 7": 5, "7": 5,
                    }
                    thu_order = {
                        "thứ 2": 2, "thứ 3": 3, "thứ 4": 4,
                        "thứ 5": 5, "thứ 6": 6, "thứ 7": 7,
                    }

                    lich = []
                    canh_bao = []

                    for thu, items in tkb.items():
                        thu_lower = thu.lower()
                        offset = 0
                        for key in days_map:
                            if key in thu_lower:
                                offset = days_map[key]
                                break
                        ngay = start_date + timedelta(days=offset)
                        ngay_str = ngay.strftime("%d/%m/%Y")

                        for item in items:
                            tkb_ppct = str(item.get("tiet_ppct", "")).strip()
                            ten_bai = ""
                            tiet_trong_bai = 1
                            tong_tiet = 1
                            stt_ppct = tkb_ppct
                            trang_sgk = ""

                            if tkb_ppct and tkb_ppct.isdigit():
                                found = find_ppct_by_tiet(ppct, tkb_ppct)
                                if found:
                                    ten_bai = found["ten_bai"]
                                    tiet_trong_bai = found["tiet_trong_bai"]
                                    tong_tiet = found["tong_tiet"]
                                    stt_ppct = str(found["stt_ppct"])
                                    trang_sgk = found.get("trang_sgk", "")
                                else:
                                    ten_bai = ("[Tiết " + tkb_ppct +
                                               " chưa có trong PPCT]")
                                    stt_ppct = str(int(tkb_ppct))
                                    canh_bao.append(
                                        "Tiết PPCT " + tkb_ppct +
                                        " lớp " + item["lop"] +
                                        " không tìm thấy trong PPCT"
                                    )
                            else:
                                ten_bai = "[TKB thiếu mã Tiết PPCT]"
                                canh_bao.append(
                                    thu + " lớp " + item["lop"] +
                                    " (" + item["tiet"] + "): thiếu mã Tiết PPCT"
                                )

                            lich.append({
                                "thu": thu,
                                "ngay": ngay_str,
                                "tiet": item["tiet"],
                                "lop": item["lop"],
                                "mon": item["mon"],
                                "ten_bai": ten_bai,
                                "tiet_trong_bai": tiet_trong_bai,
                                "tong_tiet": tong_tiet,
                                "stt_ppct": stt_ppct,
                                "trang_sgk": trang_sgk,
                            })

                    def key_lich(r):
                        t = r["thu"].lower()
                        t_num = 99
                        for k, v in thu_order.items():
                            if k in t:
                                t_num = v
                                break
                        m = re.search(r"\d+", r["tiet"])
                        tiet_num = int(m.group(0)) if m else 99
                        return (t_num, tiet_num)
                    lich.sort(key=key_lich)

                    if canh_bao:
                        with st.expander("⚠️ Cảnh báo (" +
                                         str(len(canh_bao)) + ")"):
                            for msg in sorted(set(canh_bao)):
                                st.warning(msg)

                    st.session_state["lich_bao_giang"] = lich
                    st.success("✅ Đã tạo lịch: " + str(len(lich)) + " tiết")

                    if GITHUB_TOKEN:
                        lich_json = json.dumps(lich, ensure_ascii=False)
                        schedule_name = "lich_tuan_" + \
                            start_date.strftime("%Y%m%d")
                        with st.spinner("💾 Đang lưu..."):
                            ok_s, msg_s = gh_save(
                                f"data/schedule/{schedule_name}.json",
                                lich_json,
                                "Save schedule",
                            )
                            if ok_s:
                                st.success("✅ Đã lưu: " + schedule_name)
                                st.session_state["schedule"] = load_schedule()


# ==================== TAB 2: NẠP PHIẾU CÓ SẴN ====================
with sub_t2:
    st.markdown("### ✏️ Nạp phiếu báo giảng có sẵn")
    st.info("💡 Upload file phiếu báo giảng (Word/Excel/PDF/CSV). "
            "App sẽ đọc trực tiếp, KHÔNG tra PPCT, KHÔNG ghép TKB.")

    manual_file = st.file_uploader(
        "Chọn file phiếu báo giảng",
        type=["docx", "xlsx", "xls", "csv", "pdf", "txt"],
        key="manual_sched_file_v2",
    )
    st.markdown("**Hoặc dán trực tiếp** (mỗi dòng 1 tiết, dấu `|` phân cách)")
    st.caption("Mẫu: Thứ | Ngày | Tiết | Lớp | Môn | Tiết PPCT | Tên bài")
    manual_paste = st.text_area(
        "Dán lịch",
        height=150,
        placeholder=(
            "Thứ Hai | 19/10/2026 | 1 | 6A10 | Toán | 25 | Bài 19. Hình chữ nhật...\n"
            "Thứ Hai | 19/10/2026 | 2 | 6A10 | Toán | 26 | Bài 20. Chu vi và diện tích..."
        ),
        key="manual_sched_paste_v2",
    )
    manual_name = st.text_input(
        "Tên lịch (nếu muốn lưu)",
        placeholder="lich_tuan_06",
        key="manual_sched_name_v2",
    )

    if st.button("📥 Nạp phiếu báo giảng", type="primary",
                 use_container_width=True):
        lich = []
        loi = []

        if manual_file:
            fname = manual_file.name.lower()
            raw_bytes = manual_file.read()

            if fname.endswith(".docx"):
                try:
                    from docx import Document as DocxDoc
                    d = DocxDoc(io.BytesIO(raw_bytes))
                    for tbl in d.tables:
                        for row in tbl.rows:
                            cells = [c.text.strip() for c in row.cells]
                            if not any(cells):
                                continue
                            low_first = cells[0].lower() if cells else ""
                            if low_first.startswith("thứ") and "ngày" in low_first:
                                continue
                            if len(cells) >= 6:
                                thu = cells[0]
                                tiet = cells[1].replace("Tiết ", "").strip()
                                mon = cells[2]
                                lop = cells[3]
                                stt = cells[4].strip()
                                ten = cells[5].strip()
                                if not thu and lich:
                                    thu = lich[-1]["thu"]
                                if not thu:
                                    continue
                                lich.append({
                                    "thu": thu, "ngay": "", "tiet": tiet,
                                    "mon": mon, "lop": lop, "stt_ppct": stt,
                                    "ten_bai": ten, "tiet_trong_bai": 1,
                                    "tong_tiet": 1, "trang_sgk": "",
                                })
                except Exception as e:
                    loi.append("Lỗi đọc Word: " + str(e))

            elif fname.endswith((".xlsx", ".xls")):
                try:
                    import openpyxl
                    wb = openpyxl.load_workbook(io.BytesIO(raw_bytes),
                                                 data_only=True)
                    for sn in wb.sheetnames:
                        ws = wb[sn]
                        for row in ws.iter_rows(values_only=True):
                            cells = [str(c).strip() if c is not None else ""
                                     for c in row]
                            if not any(cells):
                                continue
                            low_first = cells[0].lower()
                            if low_first.startswith("thứ") and "ngày" in low_first:
                                continue
                            while cells and not cells[0]:
                                cells.pop(0)
                            if len(cells) < 5:
                                continue
                            thu = cells[0]
                            tiet = cells[1].replace("Tiết ", "").strip() \
                                if len(cells) > 1 else ""
                            mon = cells[2] if len(cells) > 2 else ""
                            lop = cells[3] if len(cells) > 3 else ""
                            stt = cells[4] if len(cells) > 4 else ""
                            ten = cells[5] if len(cells) > 5 else ""
                            if not thu and lich:
                                thu = lich[-1]["thu"]
                            if not thu:
                                continue
                            lich.append({
                                "thu": thu, "ngay": "", "tiet": tiet,
                                "mon": mon, "lop": lop, "stt_ppct": stt,
                                "ten_bai": ten, "tiet_trong_bai": 1,
                                "tong_tiet": 1, "trang_sgk": "",
                            })
                except Exception as e:
                    loi.append("Lỗi đọc Excel: " + str(e))

            else:
                try:
                    _, raw_text, _ = read_file(manual_file, use_ocr=False)
                    for line in raw_text.splitlines():
                        line = line.strip()
                        if not line or "|" not in line:
                            continue
                        parts = [p.strip() for p in line.split("|")]
                        if len(parts) < 5:
                            continue
                        if parts[0].lower() in ("thứ", "thu", "stt"):
                            continue
                        if len(parts) >= 7:
                            lich.append({
                                "thu": parts[0], "ngay": parts[1],
                                "tiet": parts[2], "lop": parts[3],
                                "mon": parts[4], "stt_ppct": parts[5],
                                "ten_bai": parts[6], "tiet_trong_bai": 1,
                                "tong_tiet": 1, "trang_sgk": "",
                            })
                        elif len(parts) == 6:
                            lich.append({
                                "thu": parts[0], "ngay": parts[1],
                                "tiet": parts[2], "lop": parts[3],
                                "mon": parts[4], "stt_ppct": "",
                                "ten_bai": parts[5], "tiet_trong_bai": 1,
                                "tong_tiet": 1, "trang_sgk": "",
                            })
                except Exception as e:
                    loi.append("Lỗi đọc file: " + str(e))

        elif manual_paste.strip():
            for line in manual_paste.strip().splitlines():
                line = line.strip()
                if not line or "|" not in line:
                    continue
                parts = [p.strip() for p in line.split("|")]
                if len(parts) < 5:
                    continue
                if parts[0].lower() in ("thứ", "thu", "stt"):
                    continue
                if len(parts) >= 7:
                    lich.append({
                        "thu": parts[0], "ngay": parts[1], "tiet": parts[2],
                        "lop": parts[3], "mon": parts[4], "stt_ppct": parts[5],
                        "ten_bai": parts[6], "tiet_trong_bai": 1, "tong_tiet": 1,
                        "trang_sgk": "",
                    })
                elif len(parts) == 6:
                    lich.append({
                        "thu": parts[0], "ngay": parts[1], "tiet": parts[2],
                        "lop": parts[3], "mon": parts[4], "stt_ppct": "",
                        "ten_bai": parts[5], "tiet_trong_bai": 1, "tong_tiet": 1,
                        "trang_sgk": "",
                    })
                else:
                    loi.append("Dòng thiếu cột: " + line[:80])

        for msg in loi:
            st.warning(msg)

        if not lich:
            st.error("❌ Không đọc được phiếu báo giảng.")
        else:
            st.session_state["lich_bao_giang"] = lich
            st.success("✅ Đã nạp phiếu báo giảng: " + str(len(lich)) + " tiết")

            with st.expander("👁️ Xem phiếu đã nạp", expanded=True):
                st.markdown("| Thứ | Ngày | Tiết | Lớp | Môn | Tiết PPCT | Tên bài |")
                st.markdown("|---|---|---|---|---|---|---|")
                for r in lich[:30]:
                    st.markdown(
                        "| " + r["thu"] + " | " + r.get("ngay", "") + " | " + r["tiet"]
                        + " | " + r["lop"] + " | " + r["mon"]
                        + " | " + str(r["stt_ppct"])
                        + " | " + r["ten_bai"] + " |"
                    )

            if GITHUB_TOKEN and manual_name.strip():
                safe_name = re.sub(r"[^a-zA-Z0-9_\-]", "_",
                                    manual_name.strip())[:40]
                with st.spinner("💾 Đang lưu..."):
                    ok_s, msg_s = gh_save(
                        f"data/schedule/{safe_name}.json",
                        json.dumps(lich, ensure_ascii=False),
                        "Manual schedule",
                    )
                    if ok_s:
                        st.success("✅ Đã lưu: " + safe_name)
                        st.session_state["schedule"] = load_schedule()


# ==================== TAB 3: NẠP TỪ KHO ====================
with sub_t3:
    st.markdown("### 📂 Nạp lịch đã lưu từ GitHub")
    if st.session_state["schedule"]:
        sched_options = [f for f in st.session_state["schedule"].keys()
                         if f.endswith(".json")]
        if sched_options:
            selected_sched = st.selectbox("Chọn lịch:", sched_options,
                                           key="sel_sched")
            if st.button("📂 Nạp lịch này", type="primary"):
                try:
                    lich = json.loads(
                        st.session_state["schedule"][selected_sched])
                    st.session_state["lich_bao_giang"] = lich
                    st.success("✅ Đã nạp: " + selected_sched +
                               " (" + str(len(lich)) + " tiết)")
                except Exception as e:
                    st.error("Lỗi: " + str(e))
        else:
            st.info("Chưa có file JSON.")
    else:
        st.info("Chưa có lịch.")


# ============ HIỂN THỊ LỊCH ============
def hien_thi_lich_theo_ngay(lich):
    for thu in ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7"]:
        bai_trong_ngay = [r for r in lich
                          if r["thu"].lower().startswith(thu.lower()[:5])]
        if not bai_trong_ngay:
            continue
        ngay = bai_trong_ngay[0].get("ngay", "")
        st.markdown("### 📅 **" + thu +
                    (" — " + ngay if ngay else "") + "**")
        st.markdown("| Tiết | Lớp | Môn | Tên bài | Tiết |")
        st.markdown("|---|---|---|---|---|")
        for r in bai_trong_ngay:
            _ten_hien_thi = r["ten_bai"]
            _tr = str(r.get("trang_sgk", "")).strip()
            if _tr:
                _ten_hien_thi += " (tr. " + _tr + ")"
            st.markdown(
                "| " + r["tiet"] + " | " + r["lop"] + " | " + r["mon"]
                + " | " + _ten_hien_thi + " | "
                + str(r.get("tiet_trong_bai", 1)) + "/"
                + str(r.get("tong_tiet", 1)) + " |"
            )
        st.markdown("")


if st.session_state.get("lich_bao_giang"):
    lich = st.session_state["lich_bao_giang"]
    st.divider()
    st.subheader("📋 Lịch báo giảng — " + str(len(lich)) + " tiết")

    st.markdown("#### 📊 Bảng tổng hợp")
    st.markdown("| Thứ | Ngày | Tiết | Lớp | Môn | Tiết PPCT | Trang SGK | Tên bài |")
    st.markdown("|---|---|---|---|---|---|---|---|")
    for r in lich:
        st.markdown(
            "| " + r["thu"] + " | " + str(r.get("ngay", "")) + " | " + r["tiet"]
            + " | " + r["lop"] + " | " + r["mon"]
            + " | " + str(r.get("stt_ppct", ""))
            + " | " + str(r.get("trang_sgk", ""))
            + " | " + r["ten_bai"] + " |"
        )

    st.markdown("#### 📅 Chi tiết theo ngày")
    hien_thi_lich_theo_ngay(lich)

    if st.button("🚀 Soạn cả tuần", type="primary", use_container_width=True):
        if not available_providers:
            st.error("Chưa có API key.")
        else:
            bai_can_soan = {}
            for r in lich:
                ten = r.get("ten_bai", "").strip()
                if not ten:
                    continue
                key = ten
                if key not in bai_can_soan:
                    bai_can_soan[key] = {
                        "ten_bai": key,
                        "cac_lop": [r.get("lop", "")],
                        "tong_tiet": r.get("tong_tiet", 1),
                        "trang_sgk": r.get("trang_sgk", ""),
                    }
                else:
                    lop = r.get("lop", "")
                    if lop and lop not in bai_can_soan[key]["cac_lop"]:
                        bai_can_soan[key]["cac_lop"].append(lop)

            if not bai_can_soan:
                st.warning("⚠️ Không có bài nào để soạn. "
                           "Kiểm tra cột 'Tên bài' trong lịch.")
            else:
                st.info("📚 Có **" + str(len(bai_can_soan)) +
                        " bài** không trùng cần soạn.")
                progress = st.progress(0)
                status_text = st.empty()
                results = []
                bai_list = list(bai_can_soan.values())

                for idx, bai in enumerate(bai_list):
                    status_text.write("⏳ Soạn " + str(idx + 1) + "/"
                                       + str(len(bai_list)) + ": "
                                       + bai["ten_bai"])

                    _cac_lop = bai.get("cac_lop", [])
                    _lop_str = ", ".join([l for l in _cac_lop if l]) or \
                               st.session_state.get("memory", {}) \
                                   .get("default_grade", "6")

                    _mon_tuan = st.session_state.get("memory", {}) \
                        .get("default_subject", "Toán")

                    _trang_tuan = bai.get("trang_sgk", "")

                    plan = {
                        "task": "Soạn kế hoạch bài dạy",
                        "subject": _mon_tuan,
                        "grade": _lop_str,
                        "topic": bai["ten_bai"],
                        "period": str(bai["tong_tiet"]) + " tiết",
                        "trang_sgk": _trang_tuan,
                        "template_files": list(st.session_state["templates"].keys())[:1],
                        "source_files": list(st.session_state["sources"].keys())[:1],
                    }
                    try:
                        content = call_ai(build_exec_prompt(plan))
                        results.append({
                            "ten_bai": bai["ten_bai"],
                            "cac_lop": bai["cac_lop"],
                            "tong_tiet": bai["tong_tiet"],
                            "trang_sgk": _trang_tuan,
                            "content": content,
                        })
                    except Exception as e:
                        results.append({
                            "ten_bai": bai["ten_bai"],
                            "cac_lop": bai["cac_lop"],
                            "tong_tiet": bai["tong_tiet"],
                            "trang_sgk": _trang_tuan,
                            "content": "[Lỗi: " + str(e) + "]",
                        })
                    progress.progress((idx + 1) / len(bai_list))

                status_text.write("✅ Đã soạn " + str(len(results)) + " bài!")
                st.session_state["weekly_results"] = results


# ============ XUẤT WORD TUẦN ============
if st.session_state.get("weekly_results"):
    results = st.session_state["weekly_results"]
    lich = st.session_state.get("lich_bao_giang", [])

    st.divider()
    st.subheader("📚 Kết quả tuần — " + str(len(results)) + " bài")

    for r in results:
        lop_str = ", ".join(r.get("cac_lop", []))
        _trang_hien_thi = r.get("trang_sgk", "")
        _suffix = " (tr. " + _trang_hien_thi + ")" if _trang_hien_thi else ""
        st.write("- **" + r.get("ten_bai", "") + _suffix
                 + "** — Lớp: " + lop_str
                 + " (" + str(r.get("tong_tiet", 1)) + " tiết)")

    if st.button("📄 Xuất file Word tuần", type="primary",
                 use_container_width=True):
        d = Document()
        s = d.sections[0]
        s.top_margin = Inches(0.7)
        s.bottom_margin = Inches(0.7)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)
        style = d.styles["Normal"]
        style.font.name = "Times New Roman"
        style.font.size = Pt(14)
        rpr = style.element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = rpr.makeelement(qn("w:rFonts"), {})
            rpr.append(rfonts)
        rfonts.set(qn("w:ascii"), "Times New Roman")
        rfonts.set(qn("w:hAnsi"), "Times New Roman")
        rfonts.set(qn("w:eastAsia"), "Times New Roman")
        rfonts.set(qn("w:cs"), "Times New Roman")
        pf = style.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)

        if lich:
            buoi = st.session_state.get("buoi_w", "BUỔI CHIỀU")
            tuan_num = str(st.session_state.get("tuan_w", 1))
            ngay_bd = st.session_state.get("ngay_bd_w", datetime.now().date())
            add_phieu_bao_giang(d, lich, ngay_bd, buoi, tuan_num)

        for idx, r in enumerate(results):
            d.add_page_break()
            # ❌ ĐÃ BỎ ký hiệu môn ở đầu mỗi bài
            write_content_to_doc(d, r.get("content", ""))

        bio = io.BytesIO()
        d.save(bio)
        st.success("✅ File Word: Phiếu báo giảng + " + str(len(results))
                   + " bài soạn")
        st.download_button(
            "⬇️ Tải file Word tuần",
            bio.getvalue(),
            file_name="tuan_" + st.session_state.get(
                "ngay_bd_w", datetime.now().date()).strftime("%Y%m%d") + ".docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary", use_container_width=True,
        )
