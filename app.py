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

# ============ UPLOAD TAB ============
tab_up, tab_kho = st.tabs(["📤 Upload", "📂 Kho tri thức"])

with tab_up:
    st.info("💡 MẪU = định dạng. NGUỒN = nội dung SGK. PPCT/TKB/Lịch BG = lịch dạy.")
    up_type = st.radio("Loại", ["Mẫu", "Nguồn", "PPCT", "TKB", "Lịch BG"], horizontal=True)
    up_name = st.text_input("Tên file", placeholder="mau_cv5512")
    up_files = st.file_uploader("Chọn file",
        type=["docx", "pdf", "txt", "md", "csv", "xlsx", "xls", "doc",
              "png", "jpg", "jpeg", "webp", "bmp"],
        accept_multiple_files=True, key="up_files")
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
            if up_paste.strip(): parts.append(up_paste.strip())
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
                    for key, fn in [("templates", load_templates), ("sources", load_sources),
                                    ("ppct", load_ppct), ("tkb", load_tkb), ("schedule", load_schedule)]:
                        st.session_state[key] = fn()
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


def build_exec_prompt(plan):
    mem = st.session_state["memory"]
    tpl_text = ""
    for f in plan.get("template_files", []):
        if f in st.session_state["templates"]:
            tpl_text += "\n=== FILE MẪU ===\n" + st.session_state["templates"][f][:20000]
    src_text = ""
    for f in plan.get("source_files", []):
        if f in st.session_state["sources"]:
            src_text += "\n=== NGUỒN SGK ===\n" + st.session_state["sources"][f][:25000]

    prompt = "Bạn là trợ lý chuyên môn cao cấp cho giáo viên THCS Việt Nam.\n\n"
    prompt += "NHIỆM VỤ: Soạn KẾ HOẠCH BÀI DẠY theo ĐÚNG mẫu CV 5512.\n\n"
    prompt += "THÔNG TIN:\n"
    prompt += "- GV: " + mem.get("teacher_name", "") + " | Trường: " + mem.get("school", "") + "\n"
    prompt += "- Môn: " + plan.get("subject", "") + " | Khối: " + plan.get("grade", "") + "\n"
    prompt += "- Tên bài: " + plan.get("topic", "") + "\n"
    prompt += "- Số tiết: " + (plan.get("period") or "1") + "\n\n"

    if tpl_text: prompt += tpl_text + "\n"
    if src_text: prompt += src_text + "\n"

    prompt += "=== CẤU TRÚC BẮT BUỘC ===\n\n"
    prompt += "**I. Mục tiêu**\n"
    prompt += "1. Về kiến thức:\n+ [2-4 gạch đầu dòng]\n"
    prompt += "2. Về năng lực:\n+ Năng lực chung: ...\n+ Năng lực đặc thù: ...\n"
    prompt += "3. Về phẩm chất:\n+ Chăm chỉ, trung thực, trách nhiệm.\n\n"

    prompt += "**II. Thiết bị dạy học và học liệu**\n"
    prompt += "1. Đối với giáo viên: Kế hoạch bài dạy, SGK, SGV, phiếu học tập, thước kẻ, phấn màu, tivi (hoặc máy vi tính) nếu cần trình chiếu.\n"
    prompt += "2. Đối với học sinh: SGK, vở ghi, đồ dùng học tập, ôn tập kiến thức đã học.\n\n"

    prompt += "**III. Tiến trình dạy học**\n\n"

    prompt += "**1. Hoạt động 1: Mở đầu (Khởi động)**\n"
    prompt += "a) Mục tiêu: [2 gạch đầu dòng]\n"
    prompt += "b) Nội dung: [Mô tả ngắn gọn]\n"
    prompt += "c) Tổ chức thực hiện và sản phẩm:\n\n"
    prompt += "| Hoạt động của GV và HS | Sản phẩm |\n"
    prompt += "|---|---|\n"
    prompt += "| - Bước 1: Chuyển giao nhiệm vụ:<br>+ GV đặt vấn đề... | [Sản phẩm dự kiến] |\n"
    prompt += "| - Bước 2: Thực hiện nhiệm vụ:<br>+ HS suy nghĩ, thảo luận... | |\n"
    prompt += "| - Bước 3: Báo cáo, thảo luận:<br>+ HS trả lời... | |\n"
    prompt += "| - Bước 4: Kết luận, nhận định:<br>+ GV đánh giá, dẫn dắt vào bài mới | |\n\n"

    prompt += "**2. Hoạt động 2: Hình thành kiến thức mới**\n"
    prompt += "a) Mục tiêu: [mục tiêu]\n"
    prompt += "b) Nội dung: [nội dung kiến thức]\n"
    prompt += "c) Tổ chức thực hiện và sản phẩm:\n\n"
    prompt += "| Hoạt động của GV và HS | Sản phẩm |\n"
    prompt += "|---|---|\n"
    prompt += "| - Bước 1: Chuyển giao nhiệm vụ:<br>+ GV hướng dẫn và yêu cầu HS thực hiện... | **1. [Tên kiến thức]**<br>[Nội dung chi tiết từ SGK]<br>+ Ví dụ 1: ...<br>+ Ví dụ 2: ... |\n"
    prompt += "| - Bước 2: Thực hiện nhiệm vụ:<br>+ HS chú ý lắng nghe, thực hiện yêu cầu của GV.<br>+ GV quan sát và trợ giúp HS | |\n"
    prompt += "| - Bước 3: Báo cáo, thảo luận:<br>+ HS thảo luận và phát biểu, nhận xét và bổ sung cho nhau | |\n"
    prompt += "| - Bước 4: Kết luận, nhận định:<br>+ GV chính xác hóa và gọi 1 HS nhắc lại nội dung chính | |\n\n"

    prompt += "**3. Hoạt động 3: Luyện tập**\n"
    prompt += "a) Mục tiêu: Học sinh củng cố lại kiến thức thông qua một số bài tập.\n"
    prompt += "b) Nội dung: HS dựa vào kiến thức đã học vận dụng làm BT.\n"
    prompt += "c) Tổ chức thực hiện và sản phẩm:\n\n"
    prompt += "| Hoạt động của GV và HS | Sản phẩm |\n"
    prompt += "|---|---|\n"
    prompt += "| - GV yêu cầu HS hoàn thành các bài tập: [số bài] (tr[XX] - SGK).<br>- HS tiếp nhận nhiệm vụ... | **Bài [số]:**<br>[Lời giải chi tiết] |\n\n"

    prompt += "**4. Hoạt động 4: Vận dụng**\n"
    prompt += "a) Mục tiêu: Học sinh thực hiện làm bài tập vận dụng để củng cố kiến thức và áp dụng kiến thức vào thực tế đời sống.\n"
    prompt += "b) Nội dung: HS sử dụng SGK và vận dụng kiến thức đã học để hoàn thành bài tập.\n"
    prompt += "c) Tổ chức thực hiện và sản phẩm:\n\n"
    prompt += "| Hoạt động của GV và HS | Sản phẩm |\n"
    prompt += "|---|---|\n"
    prompt += "| - GV cho HS đọc và tìm hiểu...<br>- GV yêu cầu HS hoàn thành... | **Bài [số]:** [Lời giải chi tiết] |\n\n"

    prompt += "=== QUY TẮC BẮT BUỘC ===\n"
    prompt += "1. NỘI DUNG PHẢI LẤY TỪ NGUỒN SGK. KHÔNG bịa.\n"
    prompt += "2. KHÔNG dùng từ 'máy chiếu' → thay bằng 'tivi' hoặc 'máy vi tính'.\n"
    prompt += "3. Ký hiệu toán Unicode: x², x³, √, ≤, ≥, ∈, π, Ư (KHÔNG LaTeX).\n"
    prompt += "4. Phân số viết dạng a/b.\n"
    prompt += "5. Trong ô bảng: dùng <br> để xuống dòng.\n"
    prompt += "6. Độ dài tối thiểu 1500 từ.\n"
    prompt += "7. KHÔNG viết '---', '===', '***', '####' trong bài.\n"
    prompt += "8. KHÔNG dùng HTML tag <div>, <span>, <p>.\n"
    prompt += "9. KHÔNG viết dòng 'BẢNG 2 CỘT:' hay 'Bảng 2 cột:'.\n"
    prompt += "10. KHÔNG viết phần 'Người soạn', 'Ký tên', 'Sông Đốc' ở cuối.\n"
    prompt += "11. Kết thúc bài bằng '4. Hoạt động 4'.\n"
    return prompt


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
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if style is None:
        p.paragraph_format.first_line_indent = Inches(0.39)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    for part in re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", text):
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            run = p.add_run(part[2:-2])
            _set_run_font(run, 14, bold=True)
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            run = p.add_run(part[1:-1])
            run.italic = True
            _set_run_font(run, 14, bold=False)
        else:
            run = p.add_run(part)
            _set_run_font(run, 14, bold=False)
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
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            val = (row[j] if j < len(row) else "").replace("<br>", "\n")
            for part in re.split(r"(\*\*[^*]+\*\*)", val):
                if part.startswith("**") and part.endswith("**") and len(part) > 4:
                    run = p.add_run(part[2:-2])
                    _set_run_font(run, 14, bold=True)
                else:
                    run = p.add_run(part)
                    _set_run_font(run, 14, bold=False)


def _add_heading(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.first_line_indent = Inches(0)
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

    lines = content.splitlines()
    start_idx = 0
    for idx, line in enumerate(lines):
        s = line.strip()
        if re.match(r"^\**\s*I\.\s*(Mục|MỤC)\s*tiêu", s, re.IGNORECASE):
            start_idx = idx
            break
        if re.match(r"^\**\s*1\.\s*Về\s*kiến\s*thức", s, re.IGNORECASE):
            start_idx = idx
            break
    lines = lines[start_idx:]

    cleaned = []
    for line in lines:
        s = line.strip()
        if re.match(r"^Sông\s+Đốc", s, re.IGNORECASE): break
        if re.search(r"NGƯỜI\s*SOẠN", s, re.IGNORECASE): break
        if re.search(r"KÝ\s*DUYỆT", s, re.IGNORECASE): break
        if re.match(r"^\(?Ký\s*tên\)?\s*$", s, re.IGNORECASE): break
        if re.match(r"^-{3,}\s*$", s): continue
        cleaned.append(line)
    content = "\n".join(cleaned)

    lines = content.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        stripped = line.strip()

        if stripped == "":
            if i > 0 and lines[i-1].strip() != "":
                p = d.add_paragraph("")
                p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
            i += 1
            continue

        if stripped.startswith("|") and stripped.endswith("|"):
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

        if re.match(r"^\**\s*(I|II|III|IV|V)\.\s+\S", stripped):
            text = re.sub(r"^\**\s*", "", stripped).rstrip("*")
            _add_heading(d, text)
            i += 1
            continue
        if re.match(r"^\**\s*\d+\.\s+[A-Za-zÀ-ỹ]", stripped) and len(stripped) < 100:
            text = re.sub(r"^\**\s*", "", stripped).rstrip("*")
            _add_heading(d, text)
            i += 1
            continue
        if re.match(r"^\**\s*[a-z]\)\s+\S", stripped):
            text = re.sub(r"^\**\s*", "", stripped).rstrip("*")
            _add_heading(d, text)
            i += 1
            continue

        if stripped.startswith("#### "):
            _add_heading(d, re.sub(r"[*\\]", "", stripped[5:].strip()))
        elif stripped.startswith("### "):
            _add_heading(d, re.sub(r"[*\\]", "", stripped[4:].strip()))
        elif stripped.startswith("## "):
            _add_heading(d, re.sub(r"[*\\]", "", stripped[3:].strip()))
        elif stripped.startswith("# "):
            _add_heading(d, re.sub(r"[*\\]", "", stripped[2:].strip()))
        elif re.match(r"^[-*+]\s+", stripped):
            text = re.sub(r"^[-*+]\s+", "", stripped)
            text = re.sub(r"^[*\\]+\s*", "", text)
            add_para(d, text, style="List Bullet")
        elif re.match(r"^\d+[\.\)]\s+", stripped):
            text = re.sub(r"^\d+[\.\)]\s+", "", stripped)
            add_para(d, text, style="List Number")
        else:
            text = stripped
            text = re.sub(r"^\*\*\\?\*?\s*", "", text)
            text = re.sub(r"\s*\*\*\\?\*?$", "", text)
            if text:
                add_para(d, text)
        i += 1


def make_doc(title, body):
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


def add_phieu_bao_giang(doc, lich, tuan_start):
    mem = st.session_state.get("memory", {})
    now = datetime.now()
    end_date = tuan_start + timedelta(days=6)
    buoi = st.session_state.get("buoi_day", "BUỔI CHIỀU")
    tuan_num = st.session_state.get("tuan_number", "").strip() or "........"

    table_head = doc.add_table(rows=1, cols=2)
    table_head.autofit = True

    cell = table_head.cell(0, 0)
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("UBND XÃ SÔNG ĐỐC")
    _set_run_font(run, 13, bold=False)
    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p2.add_run("TRƯỜNG THCS SÔNG ĐỐC")
    _set_run_font(run, 13, bold=True)

    cell_r = table_head.cell(0, 1)
    cell_r.text = ""
    p = cell_r.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    _set_run_font(run, 13, bold=True)
    p2 = cell_r.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p2.add_run("Độc lập – Tự do – Hạnh phúc")
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
            run = p.add_run(line)
            _set_run_font(run, 13, bold=True)

    for p in hdr[6].paragraphs:
        for r in p.runs:
            r.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)

    prev_ngay_key = ""
    for r in lich:
        row = table.add_row().cells
        ngay_key = r["thu"] + "|" + r["ngay"]
        if ngay_key != prev_ngay_key:
            row[0].text = ""
            p = row[0].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(r["thu"])
            _set_run_font(run, 13, bold=False)
            p2 = row[0].add_paragraph()
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p2.add_run(r["ngay"])
            _set_run_font(run, 13, bold=False)
            prev_ngay_key = ngay_key
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
            val = r.get(key, "")
            if key == "tiet":
                val = val.replace("Tiết ", "")
            run = p.add_run(str(val))
            _set_run_font(run, 13, bold=False)

        row[6].text = ""

    doc.add_paragraph("")


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


st.divider()
st.subheader("📅 Tạo / Nạp lịch báo giảng")

sub_t1, sub_t2, sub_t3 = st.tabs([
    "🔧 Tạo từ TKB + PPCT",
    "📂 Nạp từ kho đã lưu",
    "✏️ Nạp thủ công",
])

with sub_t1:
    st.markdown("**Chọn từ kho đã lưu:**")
    col_ppct, col_tkb = st.columns(2)
    with col_ppct:
        ppct_options = ["(Không dùng)"] + list(st.session_state["ppct"].keys())
        selected_ppct = st.selectbox("📗 PPCT đã lưu:", ppct_options, key="sel_ppct")
    with col_tkb:
        tkb_options = ["(Không dùng)"] + list(st.session_state["tkb"].keys())
        selected_tkb = st.selectbox("📘 TKB đã lưu:", tkb_options, key="sel_tkb")

    st.markdown("**Hoặc upload / dán:**")
    tkb_file = st.file_uploader("📘 Upload TKB", type=["pdf", "docx", "xlsx", "xls", "csv", "txt"], key="tkb_file")
    tkb_manual = st.text_area("Hoặc dán TKB", height=100, key="tkb_manual")
    ppct_file = st.file_uploader("📗 Upload PPCT", type=["pdf", "docx", "xlsx", "xls", "csv", "txt"], key="ppct_file")
    ppct_manual = st.text_area("Hoặc dán PPCT", height=100, key="ppct_manual")

    col1_d, col2_d, col3_d = st.columns(3)
    with col1_d:
        start_date = st.date_input(
            "📆 Ngày bắt đầu tuần (thứ 2)",
            value=st.session_state.get("start_date_val", datetime.now().date()),
            key="date_widget_v2",
        )
        st.session_state["start_date_val"] = start_date
    with col2_d:
        buoi_day = st.selectbox(
            "🌅 Buổi dạy",
            ["BUỔI CHIỀU", "BUỔI SÁNG"],
            index=0 if st.session_state.get("buoi_day", "BUỔI CHIỀU") == "BUỔI CHIỀU" else 1,
            key="buoi_widget_v2",
        )
        st.session_state["buoi_day"] = buoi_day
    with col3_d:
        tuan_number = st.text_input(
            "📅 Số tuần (VD: 6)",
            key="tuan_widget_v2",
            placeholder="6",
        )
        st.session_state["tuan_number"] = tuan_number

    def ai_extract_tkb(raw):
        system = "Ban la tro ly doc bang. Tra ve CHI text theo format."
        prompt = ("Doc TKB sau, chuyen thanh moi dong: Thu | Tiet TKB | Lop | Mon | Tiet PPCT\n\n"
                  "Vi du:\nThứ 2 | 1 | 6A10 | Toán | 19\nThứ 2 | 2 | 6A10 | Toán | 20\n\n"
                  "Neu khong co cot Tiet PPCT thi de trong.\n\n"
                  "NOI DUNG:\n" + raw[:12000] + "\n\nCHI tra ve cac dong.")
        try:
            result = call_ai(prompt, system)
            return "\n".join([l.strip() for l in result.splitlines()
                              if "|" in l and len(l.split("|")) >= 4])
        except Exception:
            return ""

    def ai_extract_ppct(raw):
        system = "Bạn là trợ lý đọc bảng. Chỉ trả về text theo format."
        prompt = """Đọc PPCT sau. Bảng 5 cột: STT | Bài học | Số tiết | Thời điểm | Ghi chú
Cột "Thời điểm" bị GỘP Ô — điền tuần cho TẤT CẢ các bài.
Các dòng cùng tên và "(Tiết 1)", "(Tiết 2)" → GỘP thành 1 bài.

Trả về MỖI BÀI 1 DÒNG: STT | Tên bài (không có "(Tiết N)") | Số tiết (đã gộp) | Tuần | Ngày BĐ | Ngày KT

NỘI DUNG PPCT:
""" + raw[:15000]
        try:
            result = call_ai(prompt, system)
            lines = []
            for line in result.splitlines():
                line = line.strip()
                if "|" not in line: continue
                low = line.lower()
                if "stt" in low and "bài" in low: continue
                if "---" in line: continue
                lines.append(line)
            return "\n".join(lines)
        except Exception:
            return ""

    def auto_fix_tkb(text):
        fixed = []
        for line in text.strip().splitlines():
            line = line.strip()
            if not line: continue
            if "|" in line:
                fixed.append(line)
                continue
            m = re.match(r"^(\d)(\d)(6A\d+)(.+)$", line)
            if m:
                fixed.append("Thứ " + m.group(1) + " | Tiết " + m.group(2)
                             + " | " + m.group(3) + " | " + m.group(4).strip() + " | ")
            else:
                fixed.append(line)
        return "\n".join(fixed)

    def parse_tkb(text):
        result = {}
        for line in text.strip().splitlines():
            line = line.strip()
            if not line or "|" not in line: continue
            parts = [p.strip() for p in line.split("|")]
            if len(parts) < 4: continue
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
                if m: tiet_ppct = m.group(0)
            if thu not in result:
                result[thu] = []
            result[thu].append({
                "tiet": tiet, "lop": parts[2], "mon": parts[3],
                "tiet_ppct": tiet_ppct,
            })
        for thu in result:
            result[thu].sort(key=lambda x: int(re.search(r"\d+", x["tiet"]).group(0)) if re.search(r"\d+", x["tiet"]) else 0)
        return result

    def parse_ppct(text):
        raw_list = []
        for line in text.strip().splitlines():
            line = line.strip()
            if not line or "|" not in line: continue
            parts = [p.strip() for p in line.split("|")]
            if len(parts) < 3: continue
            if parts[0].lower() == "stt": continue
            m_stt = re.search(r"(\d+)", parts[0])
            if not m_stt: continue
            stt = int(m_stt.group(1))
            ten_bai = parts[1]
            so_tiet_str = parts[2] if len(parts) > 2 else "1"
            m = re.search(r"(\d+)", so_tiet_str)
            n_tiet = int(m.group(1)) if m else 1
            tuan = parts[3].strip() if len(parts) > 3 else ""
            ts, te = None, None
            if len(parts) >= 5:
                d1 = re.search(r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})", parts[4])
                if d1:
                    try: ts = datetime(int(d1.group(3)), int(d1.group(2)), int(d1.group(1)))
                    except Exception: pass
            if len(parts) >= 6:
                d2 = re.search(r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})", parts[5])
                if d2:
                    try: te = datetime(int(d2.group(3)), int(d2.group(2)), int(d2.group(1)))
                    except Exception: pass
            raw_list.append({
                "stt": stt, "ten_bai": ten_bai, "so_tiet": n_tiet,
                "tuan": tuan, "tuan_start": ts, "tuan_end": te,
            })

        merged = []
        i = 0
        while i < len(raw_list):
            bai = raw_list[i]
            ten_chuan = re.sub(r"\s*\(Tiết\s*\d+\)", "", bai["ten_bai"]).strip()
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
                "stt": bai["stt"], "ten_bai": ten_chuan,
                "so_tiet": so_tiet_gop, "tuan": bai["tuan"],
                "tuan_start": bai["tuan_start"], "tuan_end": bai["tuan_end"],
            })
            i = j if so_tiet_gop > 0 else i + 1
        return merged

    def find_ppct_by_tiet(ppct, tiet_str):
        if not tiet_str: return None
        try:
            target = int(tiet_str)
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
                    "stt_ppct": stt,
                }
        return None

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
                if tkb_auto and "|" in tkb_auto and len(tkb_auto.splitlines()[0].split("|")) >= 4:
                    tkb_clean = tkb_auto
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
                    st.error("Không parse được.")
                else:
                    target = datetime.combine(start_date, datetime.min.time())
                    week_bais, week_num = find_week_ppct(ppct, target)
                    if week_bais:
                        st.success("✅ Tuần " + str(week_num)
                                   + " — " + str(len(week_bais)) + " bài")
                        ppct_week = week_bais
                    else:
                        st.warning("⚠️ Không tìm thấy tuần khớp. Dùng toàn bộ PPCT.")
                        ppct_week = ppct

                    ppct_tiets = []
                    for bai in ppct_week:
                        for k in range(bai["so_tiet"]):
                            ppct_tiets.append({
                                "ten_bai": bai["ten_bai"],
                                "tiet_trong_bai": k + 1,
                                "tong_tiet": bai["so_tiet"],
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
                                        "tiet_ppct": item.get("tiet_ppct", ""),
                                    })

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

                        tkb_ppct = slot.get("tiet_ppct", "")
                        if tkb_ppct:
                            found = find_ppct_by_tiet(ppct, tkb_ppct)
                            if found:
                                ten_bai = found["ten_bai"]
                                tiet_trong_bai = found["tiet_trong_bai"]
                                tong_tiet = found["tong_tiet"]
                                stt_ppct = str(found["stt_ppct"])
                            else:
                                stt_ppct = tkb_ppct
                        else:
                            if idx < len(ppct_tiets):
                                ti = ppct_tiets[idx]
                                ten_bai = ti["ten_bai"]
                                tiet_trong_bai = ti["tiet_trong_bai"]
                                tong_tiet = ti["tong_tiet"]
                                stt_ppct = str(ti.get("stt_ppct", ""))

                        lich.append({
                            "thu": slot["thu"],
                            "ngay": ngay.strftime("%d/%m/%Y"),
                            "tiet": slot["tiet"],
                            "lop": slot["lop"],
                            "mon": slot["mon"],
                            "ten_bai": ten_bai,
                            "tiet_trong_bai": tiet_trong_bai,
                            "tong_tiet": tong_tiet,
                            "stt_ppct": stt_ppct,
                        })

                    st.session_state["lich_bao_giang"] = lich
                    st.success("✅ Đã tạo lịch: " + str(len(lich)) + " tiết")

                    if GITHUB_TOKEN:
                        lich_json = json.dumps(lich, ensure_ascii=False)
                        schedule_name = "lich_tuan_" + start_date.strftime("%Y%m%d")
                        with st.spinner("💾 Đang lưu lịch..."):
                            ok_s, msg_s = gh_save(f"data/schedule/{schedule_name}.json",
                                                  lich_json, "Save schedule")
                        if ok_s:
                            st.success("✅ Đã lưu: " + schedule_name)
                            st.session_state["schedule"] = load_schedule()


with sub_t2:
    st.markdown("**Nạp lịch đã lưu từ GitHub**")
    if st.session_state["schedule"]:
        sched_options = [f for f in st.session_state["schedule"].keys() if f.endswith(".json")]
        if sched_options:
            selected_sched = st.selectbox("Chọn lịch:", sched_options, key="sel_sched")
            if st.button("📂 Nạp lịch này", type="primary"):
                try:
                    lich = json.loads(st.session_state["schedule"][selected_sched])
                    st.session_state["lich_bao_giang"] = lich
                    st.success("✅ Đã nạp: " + selected_sched + " (" + str(len(lich)) + " tiết)")
                except Exception as e:
                    st.error("Lỗi: " + str(e))
        else:
            st.info("Chưa có file lịch JSON.")
    else:
        st.info("Chưa có lịch nào được lưu.")


with sub_t3:
    st.markdown("**Nạp lịch thủ công**")
    st.caption("Dán lịch theo mẫu: Thứ | Ngày | Tiết | Lớp | Môn | Tiết PPCT | Tên bài")

    manual_file = st.file_uploader("Upload file Excel/Word/CSV/PDF",
        type=["xlsx", "xls", "docx", "csv", "pdf", "txt"], key="manual_sched_file")
    manual_paste = st.text_area("Hoặc dán lịch", height=150,
        placeholder="Thứ Hai | 12/10/2026 | 1 | 6A10 | Toán | 19 | Luyện tập chung",
        key="manual_sched_paste")
    manual_name = st.text_input("Tên lịch (nếu muốn lưu)",
                                 placeholder="lich_tuan_06", key="manual_sched_name")

    if st.button("📥 Nạp lịch thủ công", type="primary", use_container_width=True):
        raw_text = manual_paste.strip()
        if manual_file:
            _, raw_text, _ = read_file(manual_file, use_ocr=False)

        if not raw_text:
            st.warning("Chưa có nội dung.")
        else:
            lich = []
            for line in raw_text.splitlines():
                line = line.strip()
                if not line or "|" not in line: continue
                parts = [p.strip() for p in line.split("|")]
                if len(parts) < 5: continue
                if parts[0].lower() in ("thứ", "thu", "stt"): continue
                lich.append({
                    "thu": parts[0],
                    "ngay": parts[1] if len(parts) > 1 else "",
                    "tiet": parts[2] if len(parts) > 2 else "",
                    "lop": parts[3] if len(parts) > 3 else "",
                    "mon": parts[4] if len(parts) > 4 else "",
                    "stt_ppct": parts[5] if len(parts) > 5 else "",
                    "ten_bai": parts[6] if len(parts) > 6 else "",
                    "tiet_trong_bai": 1,
                    "tong_tiet": 1,
                })

            if not lich:
                st.error("Không đọc được. Kiểm tra định dạng.")
            else:
                st.session_state["lich_bao_giang"] = lich
                st.success("✅ Đã nạp: " + str(len(lich)) + " tiết")

                if GITHUB_TOKEN and manual_name.strip():
                    safe_name = re.sub(r"[^a-zA-Z0-9_\-]", "_", manual_name.strip())[:40]
                    with st.spinner("💾 Đang lưu..."):
                        ok_s, msg_s = gh_save(f"data/schedule/{safe_name}.json",
                                              json.dumps(lich, ensure_ascii=False),
                                              "Manual schedule")
                    if ok_s:
                        st.success("✅ Đã lưu: " + safe_name)
                        st.session_state["schedule"] = load_schedule()


def hien_thi_lich_theo_ngay(lich):
    for thu in ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7"]:
        bai_trong_ngay = [r for r in lich if r["thu"].lower().startswith(thu.lower()[:5])]
        if not bai_trong_ngay: continue
        ngay = bai_trong_ngay[0]["ngay"]
        st.markdown("### 📅 **" + thu + " — " + ngay + "**")
        st.markdown("| Tiết | Lớp | Môn | Tên bài | Tiết |")
        st.markdown("|---|---|---|---|---|")
        for r in bai_trong_ngay:
            st.markdown("| " + r["tiet"] + " | " + r["lop"] + " | " + r["mon"]
                        + " | " + r["ten_bai"]
                        + " | " + str(r["tiet_trong_bai"]) + "/" + str(r["tong_tiet"]) + " |")
        st.markdown("")


if st.session_state.get("lich_bao_giang"):
    lich = st.session_state["lich_bao_giang"]
    st.divider()
    st.subheader("📋 Lịch báo giảng — " + str(len(lich)) + " tiết")

    st.markdown("#### 📊 Bảng tổng hợp")
    st.markdown("| Thứ | Ngày | Tiết | Lớp | Môn | Tiết PPCT | Tên bài |")
    st.markdown("|---|---|---|---|---|---|---|")
    for r in lich:
        st.markdown("| " + r["thu"] + " | " + r["ngay"] + " | " + r["tiet"]
                    + " | " + r["lop"] + " | " + r["mon"]
                    + " | " + str(r.get("stt_ppct", ""))
                    + " | " + r["ten_bai"] + " |")

    st.markdown("#### 📅 Chi tiết theo ngày")
    hien_thi_lich_theo_ngay(lich)

    if st.button("🚀 Soạn cả tuần (không soạn trùng bài)", type="primary", use_container_width=True):
        if not available_providers:
            st.error("Chưa có API key.")
        else:
            bai_can_soan = {}
            for r in lich:
                if not r["ten_bai"]: continue
                key = r["ten_bai"].strip()
                if key not in bai_can_soan:
                    bai_can_soan[key] = {
                        "ten_bai": key, "cac_lop": [r["lop"]],
                        "tong_tiet": r["tong_tiet"],
                    }
                else:
                    if r["lop"] not in bai_can_soan[key]["cac_lop"]:
                        bai_can_soan[key]["cac_lop"].append(r["lop"])

            st.info("📚 Có **" + str(len(bai_can_soan)) + " bài** không trùng cần soạn "
                    + "(trong " + str(len(lich)) + " tiết).")

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
            status_text.write("✅ Đã soạn " + str(len(results)) + " bài!")
            st.session_state["weekly_results"] = results


if st.session_state.get("weekly_results"):
    results = st.session_state["weekly_results"]
    lich = st.session_state.get("lich_bao_giang", [])
    st.divider()
    st.subheader("📚 Kết quả tuần — " + str(len(results)) + " bài")
    for r in results:
        lop_str = ", ".join(r.get("cac_lop", []))
        st.write("- **" + r.get("ten_bai", "") + "** — Lớp: " + lop_str
                 + " (" + str(r.get("tong_tiet", 1)) + " tiết)")

    if st.button("📄 Xuất file Word tuần (Phiếu báo giảng + bài soạn)",
                 type="primary", use_container_width=True):
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
            add_phieu_bao_giang(d, lich, st.session_state["start_date_val"])

        for idx, r in enumerate(results):
            d.add_page_break()
            lop_str = ", ".join(r.get("cac_lop", []))
            h = d.add_heading(r.get("ten_bai", ""), 1)
            h.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_info = d.add_paragraph()
            p_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p_info.add_run("Lớp: " + lop_str + " | Số tiết: " + str(r.get("tong_tiet", 1)))
            _set_run_font(run, 14, bold=False)
            write_content_to_doc(d, r.get("content", ""))

        bio = io.BytesIO()
        d.save(bio)
        st.success("✅ File Word: Phiếu báo giảng + " + str(len(results)) + " bài soạn")
        st.download_button(
            "⬇️ Tải file Word tuần",
            bio.getvalue(),
            file_name="tuan_" + st.session_state["start_date_val"].strftime("%Y%m%d") + ".docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary", use_container_width=True,
        )
