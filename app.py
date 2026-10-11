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

st.set_page_config(page_title="Trợ lý Giáo viên THCS", page_icon="🧠", layout="wide")
st.title("🧠 Trợ lý Giáo viên THCS")
st.caption("Bộ não tự động • Lịch báo giảng • Soạn tuần • Xuất Word")


# ============ CẤU HÌNH ============
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
    if not GITHUB_REPO:
        return (False, "Chưa có GITHUB_REPO")

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
        return (False, "Không gọi được GitHub API")
    if isinstance(res, dict) and "error" in res:
        return (False, "Lỗi: " + str(res["error"]))
    if isinstance(res, dict) and "content" in res:
        return (True, "Đã lưu: " + path)
    return (False, "Phản hồi không hợp lệ")


def gh_delete(path):
    """Xóa file trên GitHub. Trả về (success, message)."""
    if not GITHUB_TOKEN:
        return (False, "Chưa có GITHUB_TOKEN")
    old, sha = gh_get(path)
    if not sha:
        return (False, "File không tồn tại: " + path)
    res = gh_api(
        "DELETE",
        f"/repos/{GITHUB_REPO}/contents/{path}",
        {"message": "Delete " + path, "sha": sha, "branch": GITHUB_BRANCH},
    )
    if isinstance(res, dict) and "error" in res:
        return (False, "Lỗi: " + str(res["error"]))
    if isinstance(res, dict) and "content" in res:
        return (True, "Đã xóa: " + path)
    if isinstance(res, dict) and "commit" in res:
        return (True, "Đã xóa: " + path)
    return (False, "Phản hồi không hợp lệ")


def gh_list(folder):
    res = gh_api("GET", f"/repos/{GITHUB_REPO}/contents/{folder}?ref={GITHUB_BRANCH}")
    if not res or "error" in res or not isinstance(res, list):
        return []
    return [item["name"] for item in res if item["type"] == "file"]


def gh_test():
    if not GITHUB_TOKEN:
        return (False, "Thiếu GITHUB_TOKEN trong Secrets")
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
                return (True, "OK — " + login + " | Có quyền ghi vào " + GITHUB_REPO)
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


def load_templates():
    files = gh_list("data/templates")
    out = {}
    for f in files:
        c, _ = gh_get(f"data/templates/{f}")
        if c:
            out[f] = c
    return out


def load_sources():
    files = gh_list("data/sources")
    out = {}
    for f in files:
        c, _ = gh_get(f"data/sources/{f}")
        if c:
            out[f] = c
    return out


def save_history(task, topic, content):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe = re.sub(r"[^a-zA-Z0-9]", "_", topic)[:30]
    fname = f"data/history/{ts}_{safe}.md"
    md = f"# {task}\n\n**Chủ đề:** {topic}\n\n---\n\n{content}"
    return gh_save(fname, md, f"History {topic}")


# ============ ĐỌC FILE AN TOÀN ============
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
                t = page.extract_text() or ""
                if t.strip():
                    texts.append(t)
            except Exception:
                continue
        full = "\n".join(texts)
        if len(full.strip()) < 50:
            return "[PDF scan, không có text. Dùng Google Drive OCR hoặc dán trực tiếp.]\n" + full
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


def _read_image_info(data, fname):
    try:
        img = Image.open(io.BytesIO(data))
        return "[" + fname + ": ảnh " + str(img.size[0]) + "x" + str(img.size[1]) + " px. Cần OCR.]"
    except Exception as e:
        return "[Lỗi ảnh " + fname + ": " + str(e) + "]"


def read_file(file):
    if file is None:
        return ("", "", 0)
    try:
        data = file.read()
    except Exception as e:
        return (file.name, "[Lỗi đọc file: " + str(e) + "]", 0)
    if not data:
        return (file.name, "[File rỗng]", 0)

    name = file.name
    nl = name.lower()
    magic = data[:8]
    text = ""

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
        elif nl.endswith((".png", ".jpg", ".jpeg", ".gif", ".webp")):
            text = _read_image_info(data, name)
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
        st.caption("Branch: " + GITHUB_BRANCH)
        if st.button("🧪 Test GitHub", use_container_width=True):
            ok, msg = gh_test()
            if ok:
                st.success("✅ " + msg)
            else:
                st.error("❌ " + msg)
    else:
        st.error("❌ Chưa có GITHUB_TOKEN")
        st.caption("Thêm vào Secrets: GITHUB_TOKEN = \"ghp_...\"")

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
    if st.button("🔄 Tải lại", use_container_width=True):
        st.session_state["templates"] = load_templates()
        st.session_state["sources"] = load_sources()
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
tab_up, tab_kho = st.tabs(["📤 Upload mẫu/nguồn", "📂 Kho tri thức"])

with tab_up:
    st.info("💡 **MẪU** = định dạng bài soạn (bảng 2 cột). **NGUỒN** = nội dung SGK.")
    up_type = st.radio("Loại", ["Mẫu (định dạng)", "Nguồn (nội dung)"], horizontal=True)
    up_name = st.text_input("Tên file", placeholder="mau_bai_soan")
    up_files = st.file_uploader(
        "Chọn file (docx / pdf / xlsx / csv / txt / ảnh)",
        type=["docx", "pdf", "txt", "md", "csv", "xlsx", "xls", "doc", "png", "jpg", "jpeg"],
        accept_multiple_files=True, key="up_files",
    )
    up_paste = st.text_area("Hoặc dán nội dung", height=120, key="up_paste")

    if up_files:
        st.markdown("**📋 Trạng thái file:**")
        for f in up_files:
            nm, txt, sz = read_file(f)
            if sz == 0 or txt.startswith("[Lỗi") or txt.startswith("[File .doc"):
                st.write("⚠️ " + nm + ": " + txt[:80])
            elif sz < 200:
                st.write("⚠️ " + nm + ": " + str(sz) + " ký tự (có thể lỗi)")
            else:
                st.write("✅ " + nm + ": " + str(sz) + " ký tự")

    if st.button("💾 Lưu lên GitHub", type="primary"):
        ok0, msg0 = gh_test()
        if not ok0:
            st.error("❌ Lỗi kết nối: " + msg0)
            st.info("Kiểm tra GITHUB_TOKEN, GITHUB_REPO trong Secrets")
        else:
            st.success("✅ " + msg0)
            parts = []
            if up_paste.strip():
                parts.append(up_paste.strip())
            for f in (up_files or []):
                _, txt, _ = read_file(f)
                parts.append("--- " + f.name + " ---\n" + txt)

            if not parts:
                st.warning("Không có nội dung. Hãy chọn file hoặc dán text.")
            else:
                full = "\n\n".join(parts)
                safe = re.sub(r"[^a-zA-Z0-9_\-]", "_", up_name.strip())[:40]
                if not safe:
                    safe = "file_" + datetime.now().strftime("%Y%m%d_%H%M%S")
                folder = "templates" if "Mẫu" in up_type else "sources"
                path = f"data/{folder}/{safe}.md"

                with st.spinner("💾 Đang lưu..."):
                    ok, msg = gh_save(path, full, f"Upload {safe}")

                if ok:
                    st.success("✅ " + msg)
                    st.info("Đã lưu " + str(len(full)) + " ký tự")
                    st.session_state["templates"] = load_templates()
                    st.session_state["sources"] = load_sources()
                    st.rerun()
                else:
                    st.error("❌ Lỗi: " + msg)

with tab_kho:
    st.subheader("📐 Mẫu đã lưu")
    if st.session_state["templates"]:
        for name in list(st.session_state["templates"].keys()):
            col1, col2 = st.columns([5, 1])
            with col1:
                with st.expander("📄 " + name):
                    content = st.session_state["templates"][name]
                    st.caption("Kích thước: " + str(len(content)) + " ký tự")
                    st.text(content[:2000])
            with col2:
                if st.button("🗑️", key="del_tpl_" + name, help="Xóa mẫu này"):
                    with st.spinner("Đang xóa..."):
                        ok, msg = gh_delete("data/templates/" + name)
                    if ok:
                        st.success(msg)
                        st.session_state["templates"] = load_templates()
                        st.rerun()
                    else:
                        st.error(msg)
    else:
        st.info("Chưa có mẫu. Upload ở tab bên trái.")

    if st.session_state["templates"]:
        if st.button("🗑️ Xóa TẤT CẢ mẫu", key="del_all_tpl"):
            for name in list(st.session_state["templates"].keys()):
                gh_delete("data/templates/" + name)
            st.session_state["templates"] = load_templates()
            st.success("Đã xóa tất cả mẫu.")
            st.rerun()

    st.divider()
    st.subheader("📚 Nguồn đã lưu")
    if st.session_state["sources"]:
        for name in list(st.session_state["sources"].keys()):
            col1, col2 = st.columns([5, 1])
            with col1:
                with st.expander("📄 " + name):
                    content = st.session_state["sources"][name]
                    st.caption("Kích thước: " + str(len(content)) + " ký tự")
                    st.text(content[:2000])
            with col2:
                if st.button("🗑️", key="del_src_" + name, help="Xóa nguồn này"):
                    with st.spinner("Đang xóa..."):
                        ok, msg = gh_delete("data/sources/" + name)
                    if ok:
                        st.success(msg)
                        st.session_state["sources"] = load_sources()
                        st.rerun()
                    else:
                        st.error(msg)
    else:
        st.info("Chưa có nguồn. Upload ở tab bên trái.")

    if st.session_state["sources"]:
        if st.button("🗑️ Xóa TẤT CẢ nguồn", key="del_all_src"):
            for name in list(st.session_state["sources"].keys()):
                gh_delete("data/sources/" + name)
            st.session_state["sources"] = load_sources()
            st.success("Đã xóa tất cả nguồn.")
            st.rerun()


# ============ PROMPT ============
def build_exec_prompt(plan):
    mem = st.session_state["memory"]
    tpl_text = ""
    for f in plan.get("template_files", []):
        if f in st.session_state["templates"]:
            tpl_text += "\n=== FILE MẪU (ĐỊNH DẠNG) ===\n" + st.session_state["templates"][f][:15000]
    src_text = ""
    for f in plan.get("source_files", []):
        if f in st.session_state["sources"]:
            src_text += "\n=== NGUỒN NỘI DUNG (SGK) ===\n" + st.session_state["sources"][f][:25000]

    prompt = "Bạn là trợ lý chuyên môn cao cấp cho giáo viên THCS Việt Nam.\n\n"
    prompt += "NGUYÊN TẮC BẮT BUỘC:\n"
    prompt += "- NỘI DUNG bài soạn: LẤY TỪ NGUỒN SÁCH GIÁO KHOA (nếu có). KHÔNG bịa.\n"
    prompt += "- ĐỊNH DẠNG bài soạn: BÁM THEO FILE MẪU (nếu có).\n\n"
    prompt += "THÔNG TIN:\n"
    prompt += "GV: " + mem.get("teacher_name", "") + " | Trường: " + mem.get("school", "") + "\n"
    prompt += "Môn: " + plan.get("subject", "") + " | Khối: " + plan.get("grade", "") + "\n"
    prompt += "Tên bài: " + plan.get("topic", "") + "\n"
    prompt += "Thời lượng: " + (plan.get("period") or "Chưa xác định") + "\n\n"

    if tpl_text:
        prompt += tpl_text + "\n"
    if src_text:
        prompt += src_text + "\n"

    prompt += "YÊU CẦU NỘI DUNG:\n"
    prompt += "1. Kiến thức, số liệu, ví dụ: PHẢI LẤY TỪ NGUỒN SGK.\n"
    prompt += "2. Nếu nguồn thiếu: ghi [Cần bổ sung: ...]. KHÔNG bịa.\n"
    prompt += "3. Chi tiết, cụ thể, dùng được ngay. Tối thiểu 1500 từ.\n"
    prompt += "4. Mỗi hoạt động dạy học phải có: mục tiêu, cách tổ chức, sản phẩm.\n\n"

    prompt += "YÊU CẦU ĐỊNH DẠNG (BÁM THEO MẪU):\n"
    prompt += "5. Nếu mẫu có mục nào → bài soạn phải có mục đó, đúng thứ tự.\n"
    prompt += "6. Nếu mẫu dùng bảng → dùng bảng markdown tương ứng.\n"
    prompt += "7. Bảng hoạt động dạy học: ĐÚNG 2 CỘT.\n"
    prompt += "   | Hoạt động của GV và HS | Sản phẩm |\n"
    prompt += "   |---|---|\n"
    prompt += "8. Trong ô: dùng <br> xuống dòng, **bold** in đậm.\n\n"

    prompt += "KÝ HIỆU TOÁN (dùng Unicode):\n"
    prompt += "9. x², x³, √x, H₂O, ≤, ≥, ≠, ∈, π, °, α, β\n"
    prompt += "10. KHÔNG dùng LaTeX ($...$, \\frac, \\sqrt).\n\n"

    prompt += "ĐỊNH DẠNG MARKDOWN:\n"
    prompt += "11. '# ' cấp 1, '## ' cấp 2, '### ' cấp 3\n"
    prompt += "12. '- ' bullet, '1. ' số\n"
    prompt += "13. Tiếng Việt có dấu đầy đủ.\n"
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
    h = doc.add_heading("PHIẾU BÁO GIẢNG", 0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Trường: ").bold = True
    p.add_run(mem.get("school", "..."))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Giáo viên: ").bold = True
    p.add_run(mem.get("teacher_name", "..."))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Tuần: ").bold = True
    end_date = tuan_start + timedelta(days=6)
    p.add_run(
        "Từ ngày " + tuan_start.strftime("%d/%m/%Y")
        + " đến ngày " + end_date.strftime("%d/%m/%Y")
    )

    doc.add_paragraph()

    table = doc.add_table(rows=1, cols=6)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    for i, h_text in enumerate(["Thứ", "Ngày", "Tiết", "Lớp", "Tên bài dạy", "Ghi chú"]):
        hdr[i].text = ""
        run = hdr[i].paragraphs[0].add_run(h_text)
        run.bold = True

    for r in lich:
        row = table.add_row().cells
        row[0].text = r["thu"]
        row[1].text = r["ngay"]
        row[2].text = r["tiet"]
        row[3].text = r["lop"]
        row[4].text = r["ten_bai"]
        row[5].text = ""

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    now = datetime.now()
    p.add_run("Ngày " + now.strftime("%d") + " tháng " + now.strftime("%m") + " năm " + now.strftime("%Y"))

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.add_run("Giáo viên").bold = True

    doc.add_paragraph()
    doc.add_paragraph()

    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p3.add_run(mem.get("teacher_name", "..."))


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
    system = "Ban la tro ly doc bang. Tra ve CHI text theo format."
    prompt = "Doc PPCT sau, chuyen thanh moi dong: Tiet | Ten bai | So tiet\n\nNOI DUNG:\n"
    prompt += raw[:12000] + "\n\nCHI tra ve cac dong."
    try:
        result = call_ai(prompt, system)
        return "\n".join([l.strip() for l in result.splitlines() if "|" in l and len(l.split("|")) >= 2])
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
    result = []
    for line in text.strip().splitlines():
        line = line.strip()
        if not line or "|" not in line:
            continue
        parts = [p.strip() for p in line.split("|")]
        if len(parts) < 2:
            continue
        m2 = re.search(r"(\d+)", parts[0])
        tiet_start = int(m2.group(1)) if m2 else 0
        so_tiet_str = parts[2] if len(parts) > 2 else "1 tiết"
        m = re.search(r"(\d+)", so_tiet_str)
        n_tiet = int(m.group(1)) if m else 1
        result.append({"tiet_bd": tiet_start, "ten_bai": parts[1], "so_tiet": n_tiet})
    result.sort(key=lambda x: x["tiet_bd"])
    return result


if st.button("📅 Tạo lịch báo giảng", type="primary", use_container_width=True):
    if not available_providers:
        st.error("Cần API key.")
    else:
        tkb_raw = tkb_manual.strip()
        if tkb_file:
            _, text, _ = read_file(tkb_file)
            tkb_raw = text
        ppct_raw = ppct_manual.strip()
        if ppct_file:
            _, text, _ = read_file(ppct_file)
            ppct_raw = text

        if not tkb_raw or not ppct_raw:
            st.error("Cần TKB và PPCT.")
        else:
            tkb_auto = auto_fix_tkb(tkb_raw)
            if tkb_auto and "|" in tkb_auto and len(tkb_auto.splitlines()[0].split("|")) >= 4:
                tkb_clean = tkb_auto
            else:
                with st.spinner("🧠 AI đang đọc TKB..."):
                    tkb_clean = ai_extract_tkb(tkb_raw)

            with st.spinner("🧠 AI đang đọc PPCT..."):
                if ppct_raw and "|" in ppct_raw:
                    ppct_clean = ppct_raw
                else:
                    ppct_clean = ai_extract_ppct(ppct_raw)

            with st.expander("📋 Xem kết quả AI trích xuất"):
                st.markdown("**TKB:**")
                st.code(tkb_clean)
                st.markdown("**PPCT:**")
                st.code(ppct_clean)

            tkb = parse_tkb(tkb_clean)
            ppct = parse_ppct(ppct_clean)

            if not tkb or not ppct:
                st.error("Không parse được.")
            else:
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

                ppct_tiets = []
                for bai in ppct:
                    for k in range(bai["so_tiet"]):
                        ppct_tiets.append({
                            "ten_bai": bai["ten_bai"],
                            "tiet_trong_bai": k + 1,
                            "tong_tiet": bai["so_tiet"],
                        })

                lich = []
                for idx, slot in enumerate(all_slots):
                    if idx >= len(ppct_tiets):
                        break
                    tiet_info = ppct_tiets[idx]
                    thu_lower = slot["thu"].lower()
                    offset = 0
                    for key in days_map:
                        if key in thu_lower:
                            offset = days_map[key]
                            break
                    ngay = start_date + timedelta(days=offset)
                    lich.append({
                        "thu": slot["thu"],
                        "ngay": ngay.strftime("%d/%m/%Y"),
                        "tiet": slot["tiet"],
                        "lop": slot["lop"],
                        "mon": slot["mon"],
                        "ten_bai": tiet_info["ten_bai"],
                        "tiet_trong_bai": tiet_info["tiet_trong_bai"],
                        "tong_tiet": tiet_info["tong_tiet"],
                    })

                st.session_state["lich_bao_giang"] = lich
                st.success("✅ Đã tạo lịch: " + str(len(lich)) + " tiết")


# ============ HIỂN THỊ LỊCH ============
if st.session_state.get("lich_bao_giang"):
    lich = st.session_state["lich_bao_giang"]
    st.divider()
    st.subheader("📋 Lịch báo giảng — " + str(len(lich)) + " tiết")
    st.markdown("| Thứ | Ngày | Tiết | Lớp | Môn | Tên bài | Tiết |")
    st.markdown("|---|---|---|---|---|---|---|")
    for r in lich:
        st.markdown("| " + r["thu"] + " | " + r["ngay"] + " | " + r["tiet"]
                    + " | " + r["lop"] + " | " + r["mon"]
                    + " | " + r["ten_bai"]
                    + " | " + str(r["tiet_trong_bai"]) + "/" + str(r["tong_tiet"]) + " |")

    if st.button("🚀 2. Soạn cả tuần (có Phiếu báo giảng ở trang đầu)",
                 type="primary", use_container_width=True):
        if not available_providers:
            st.error("Chưa có API key.")
        else:
            progress = st.progress(0)
            status_text = st.empty()
            results = []
            for idx, r in enumerate(lich):
                status_text.write("⏳ Soạn bài " + str(idx + 1) + "/" + str(len(lich)) + ": " + r["ten_bai"])
                plan = {
                    "task": "Soạn kế hoạch bài dạy",
                    "subject": r["mon"],
                    "grade": st.session_state.get("memory", {}).get("default_grade", "6"),
                    "topic": r["ten_bai"],
                    "period": str(r["tiet_trong_bai"]) + "/" + str(r["tong_tiet"]),
                    "template_files": list(st.session_state["templates"].keys())[:1],
                    "source_files": list(st.session_state["sources"].keys())[:1],
                }
                try:
                    content = call_ai(build_exec_prompt(plan))
                    results.append({
                        "thu": r["thu"] + " " + r["ngay"],
                        "ten_bai": r["ten_bai"],
                        "mon_lop": r["mon"] + " " + r["lop"],
                        "so_tiet": str(r["tiet_trong_bai"]) + "/" + str(r["tong_tiet"]),
                        "content": content,
                    })
                except Exception as e:
                    results.append({
                        "thu": r["thu"] + " " + r["ngay"],
                        "ten_bai": r["ten_bai"],
                        "mon_lop": r["mon"] + " " + r["lop"],
                        "so_tiet": "",
                        "content": "[Lỗi: " + str(e) + "]",
                    })
                progress.progress((idx + 1) / len(lich))
            status_text.write("✅ Đã soạn " + str(len(results)) + " bài!")
            st.session_state["weekly_results"] = results


# ============ XUẤT WORD TUẦN ============
if st.session_state.get("weekly_results"):
    results = st.session_state["weekly_results"]
    lich = st.session_state.get("lich_bao_giang", [])
    st.divider()
    st.subheader("📚 Kết quả tuần — " + str(len(results)) + " bài")
    for r in results:
        st.write("- " + r.get("thu", "") + " | " + r.get("ten_bai", ""))

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
        else:
            d.add_heading("PHIẾU BÁO GIẢNG", 0)
            mem = st.session_state.get("memory", {})
            d.add_paragraph("Trường: " + mem.get("school", ""))
            d.add_paragraph("Giáo viên: " + mem.get("teacher_name", ""))

        for idx, r in enumerate(results):
            d.add_page_break()
            d.add_heading("BÀI " + str(idx + 1) + ": " + r.get("ten_bai", ""), 1)
            d.add_paragraph("Thứ: " + r.get("thu", "") + " | "
                            + r.get("mon_lop", "") + " | " + r.get("so_tiet", ""))
            d.add_paragraph("---")
            write_content_to_doc(d, r.get("content", ""))

        bio = io.BytesIO()
        d.save(bio)
        st.success("✅ Đã tạo file Word: Phiếu báo giảng + " + str(len(results)) + " bài soạn")
        st.download_button(
            "⬇️ Tải file Word tuần",
            bio.getvalue(),
            file_name="tuan_" + st.session_state["start_date_val"].strftime("%Y%m%d") + ".docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary", use_container_width=True,
        )
