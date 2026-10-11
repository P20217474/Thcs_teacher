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

st.set_page_config(
    page_title="Trợ lý Giáo viên THCS",
    page_icon="🧠",
    layout="wide",
)

st.title("🧠 Trợ lý Giáo viên THCS")
st.caption("Bộ não tự động • Lịch báo giảng • Soạn hàng loạt • Xuất Word")


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


def gh_api(method, path, data=None):
    if not GITHUB_TOKEN:
        return None
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
    old, sha = gh_get(path)
    data = {
        "message": message,
        "content": base64.b64encode(content.encode("utf-8")).decode("ascii"),
        "branch": GITHUB_BRANCH,
    }
    if sha:
        data["sha"] = sha
    res = gh_api("PUT", f"/repos/{GITHUB_REPO}/contents/{path}", data)
    return res and "error" not in res


def gh_list(folder):
    res = gh_api("GET", f"/repos/{GITHUB_REPO}/contents/{folder}?ref={GITHUB_BRANCH}")
    if not res or "error" in res or not isinstance(res, list):
        return []
    return [item["name"] for item in res if item["type"] == "file"]


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


def load_skills():
    c, _ = gh_get("data/skills/README.md")
    return c or ""


def read_docx(data):
    try:
        d = Document(io.BytesIO(data))
        out = [p.text for p in d.paragraphs if p.text.strip()]
        for t in d.tables:
            out.append("")
            for row in t.rows:
                out.append("| " + " | ".join(c.text.strip() for c in row.cells) + " |")
        return "\n".join(out)
    except Exception as e:
        return "[Loi docx: " + str(e) + "]"


def read_pdf(data):
    if not HAS_PYPDF:
        return "[Chua cai pypdf]"
    try:
        reader = PdfReader(io.BytesIO(data))
        texts = [(p.extract_text() or "") for p in reader.pages]
        full = "\n".join(t for t in texts if t.strip())
        if len(full) < 100:
            return "[PDF scan, khong co text]" + full
        return full
    except Exception as e:
        return "[Loi pdf: " + str(e) + "]"


def read_csv_text(data):
    try:
        text = data.decode("utf-8", errors="ignore")
        rows = list(csv.reader(io.StringIO(text)))
        if not rows:
            return "[CSV rong]"
        out = ["| " + " | ".join(rows[0]) + " |", "|" + "---|" * len(rows[0])]
        for row in rows[1:]:
            row = (row + [""] * len(rows[0]))[:len(rows[0])]
            out.append("| " + " | ".join(row) + " |")
        return "\n".join(out)
    except Exception as e:
        return "[Loi csv: " + str(e) + "]"


def read_excel(data):
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        out = []
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            rows = []
            for row in ws.iter_rows(values_only=True):
                rows.append([str(c) if c is not None else "" for c in row])
            if not rows:
                continue
            out.append("| " + " | ".join(rows[0]) + " |")
            out.append("|" + "---|" * len(rows[0]))
            for r in rows[1:]:
                out.append("| " + " | ".join(r) + " |")
        return "\n".join(out)
    except Exception as e:
        return "[Loi Excel: " + str(e) + "]"


def read_file(file):
    if file is None:
        return ("", "", 0)
    data = file.read()
    if not data:
        return (file.name, "[File rong]", 0)
    n = file.name.lower()
    if n.endswith(".docx"):
        text = read_docx(data)
    elif n.endswith(".pdf"):
        text = read_pdf(data)
    elif n.endswith(".csv"):
        text = read_csv_text(data)
    elif n.endswith((".xlsx", ".xls")):
        text = read_excel(data)
    elif n.endswith((".txt", ".md", ".log")):
        text = data.decode("utf-8", errors="ignore")
    elif n.endswith((".png", ".jpg", ".jpeg", ".gif", ".webp")):
        text = "[" + file.name + ": anh, can OCR]"
    else:
        text = "[Dinh dang chua ho tro]"
    return (file.name, text, len(text))


if "memory" not in st.session_state:
    st.session_state["memory"] = load_memory() if GITHUB_TOKEN else {}
if "templates" not in st.session_state:
    st.session_state["templates"] = load_templates() if GITHUB_TOKEN else {}
if "sources" not in st.session_state:
    st.session_state["sources"] = load_sources() if GITHUB_TOKEN else {}
if "skills" not in st.session_state:
    st.session_state["skills"] = load_skills() if GITHUB_TOKEN else ""
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


with st.sidebar:
    st.header("🧠 Bộ não")
    mem = st.session_state["memory"]

    if GITHUB_TOKEN:
        st.success("💾 Bộ nhớ vĩnh viễn đang bật")
    else:
        st.warning("⚠️ Chưa có GITHUB_TOKEN")

    mem["teacher_name"] = st.text_input("Giáo viên", value=mem.get("teacher_name", ""))
    mem["school"] = st.text_input("Trường", value=mem.get("school", ""))
    mem["default_subject"] = st.text_input("Môn mặc định", value=mem.get("default_subject", "Toan"))
    mem["default_grade"] = st.text_input("Khối mặc định", value=mem.get("default_grade", "6"))

    if st.button("💾 Lưu vào bộ nhớ", use_container_width=True):
        if GITHUB_TOKEN and save_memory(mem):
            st.success("✅ Đã lưu")
        elif not GITHUB_TOKEN:
            st.warning("Chưa có GITHUB_TOKEN")

    st.divider()
    st.subheader("📚 Kho tri thức")
    st.write("Mẫu: " + str(len(st.session_state["templates"])) + " file")
    st.write("Nguồn: " + str(len(st.session_state["sources"])) + " file")

    if st.button("🔄 Tải lại", use_container_width=True):
        st.session_state["templates"] = load_templates()
        st.session_state["sources"] = load_sources()
        st.rerun()

    st.divider()
    st.write("AI: " + str(len(available_providers)) + " nhà cung cấp")


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


tab_up, tab_kho = st.tabs(["📤 Upload mẫu/nguồn", "📂 Kho tri thức"])

with tab_up:
    st.markdown("**MẪU** (định dạng) và **NGUỒN** (nội dung) — đều tùy chọn.")
    up_type = st.radio("Loại", ["Mẫu (định dạng)", "Nguồn (nội dung)"], horizontal=True)
    up_name = st.text_input("Tên file (không cần đuôi)", placeholder="giao_an_toan6")
    up_files = st.file_uploader(
        "Chọn file",
        type=["docx", "pdf", "txt", "md", "csv", "xlsx"],
        accept_multiple_files=True,
        key="up_files",
    )
    up_paste = st.text_area("Hoặc dán nội dung", height=120, key="up_paste")

    if st.button("💾 Lưu lên GitHub", type="primary"):
        if not GITHUB_TOKEN:
            st.error("Cần GITHUB_TOKEN để lưu.")
        else:
            parts = []
            if up_paste.strip():
                parts.append(up_paste.strip())
            for f in (up_files or []):
                _, txt, _ = read_file(f)
                parts.append("--- " + f.name + " ---\n" + txt)
            if not parts:
                st.warning("Không có nội dung.")
            else:
                full = "\n\n".join(parts)
                safe = re.sub(r"[^a-zA-Z0-9_\-]", "_", up_name.strip())[:40] or "file"
                folder = "templates" if "Mẫu" in up_type else "sources"
                path = f"data/{folder}/{safe}.md"
                if gh_save(path, full, f"Upload {safe}"):
                    st.success("✅ Đã lưu: " + path)
                    st.session_state["templates"] = load_templates()
                    st.session_state["sources"] = load_sources()
                else:
                    st.error("❌ Lỗi lưu GitHub.")

with tab_kho:
    st.subheader("📐 Mẫu")
    if st.session_state["templates"]:
        for name in st.session_state["templates"]:
            st.write("- " + name)
    else:
        st.info("Chưa có mẫu.")
    st.divider()
    st.subheader("📚 Nguồn")
    if st.session_state["sources"]:
        for name in st.session_state["sources"]:
            st.write("- " + name)
    else:
        st.info("Chưa có nguồn.")


def analyze_command(cmd):
    mem = st.session_state["memory"]
    tpl_list = list(st.session_state["templates"].keys())
    src_list = list(st.session_state["sources"].keys())
    system = "Ban la bo nao phan tich lenh. Tra ve CHI JSON."
    prompt = 'Phan tich lenh sau, tra ve JSON:\n{"task":"","subject":"","grade":"","topic":"","period":"","template_files":[],"source_files":[]}\n\n'
    prompt += "LENH: " + cmd + "\n"
    prompt += "GV: " + mem.get("teacher_name", "") + "\n"
    prompt += "Mon mac dinh: " + mem.get("default_subject", "Toan") + "\n"
    prompt += "Khoi mac dinh: " + mem.get("default_grade", "6") + "\n"
    prompt += "Mau co: " + str(tpl_list) + "\n"
    prompt += "Nguon co: " + str(src_list) + "\n"
    try:
        result = call_ai(prompt, system)
        m = re.search(r"\{[\s\S]*\}", result)
        if m:
            return json.loads(m.group(0))
    except Exception:
        pass
    low = cmd.lower()
    return {
        "task": "Soan ke hoach bai day" if "giao an" in low else "Soan noi dung theo yeu cau rieng",
        "subject": mem.get("default_subject", "Toan"),
        "grade": mem.get("default_grade", "6"),
        "topic": cmd,
        "period": "",
        "template_files": tpl_list[:1],
        "source_files": src_list[:1],
    }


def build_exec_prompt(plan):
    mem = st.session_state["memory"]
    tpl_text = ""
    for f in plan.get("template_files", []):
        if f in st.session_state["templates"]:
            tpl_text += "\n=== MAU: " + f + " ===\n" + st.session_state["templates"][f][:15000]
    src_text = ""
    for f in plan.get("source_files", []):
        if f in st.session_state["sources"]:
            src_text += "\n=== NGUON: " + f + " ===\n" + st.session_state["sources"][f][:25000]
    prompt = "Ban la tro ly chuyen mon cho giao vien THCS Viet Nam.\n\n"
    prompt += "GV: " + mem.get("teacher_name", "") + " | Truong: " + mem.get("school", "") + "\n"
    prompt += "NHIEM VU: " + plan.get("task", "") + "\n"
    prompt += "MON: " + plan.get("subject", "") + " | KHOI: " + plan.get("grade", "") + "\n"
    prompt += "CHU DE: " + plan.get("topic", "") + "\n"
    prompt += "THOI LUONG: " + (plan.get("period") or "Chua xac dinh") + "\n\n"
    if tpl_text:
        prompt += tpl_text + "\n"
    if src_text:
        prompt += src_text + "\n"
    prompt += "QUY TAC:\n"
    prompt += "1. Bam sat mau (neu co).\n"
    prompt += "2. Lay noi dung tu nguon (neu co), khong bia.\n"
    prompt += "3. Chi tiet, cu the, toi thieu 1500 tu.\n"
    prompt += "4. Bang markdown: | Cot 1 | Cot 2 |, dung <br> va **bold** trong o.\n"
    prompt += "5. Ky hieu toan: x², √x, H₂O, ≤, ∈, π, ° (khong dung LaTeX).\n"
    prompt += "6. Tieng Viet co dau.\n"
    prompt += "7. Markdown: '# ', '## ', '### ', '- ', '1. '.\n"
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
        d.add_paragraph("Don vi: " + mem["school"])
    if mem.get("teacher_name"):
        d.add_paragraph("Nguoi thuc hien: " + mem["teacher_name"])
    d.add_paragraph("Ngay tao: " + datetime.now().strftime("%d/%m/%Y %H:%M"))
    write_content_to_doc(d, body)
    b = io.BytesIO()
    d.save(b)
    return b.getvalue()


st.divider()
st.subheader("💬 Soạn 1 bài theo lệnh")
st.caption("Ví dụ: Soạn giáo án Toán 6 bài Ước chung lớn nhất, 2 tiết")

cmd = st.text_area(
    "Lệnh",
    height=100,
    placeholder="Soạn giáo án Toán 6 bài Ước chung lớn nhất...",
    label_visibility="collapsed",
    key="single_cmd",
)

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
        with st.spinner("🧠 Phân tích lệnh..."):
            plan = analyze_command(cmd)
        with st.expander("📋 Kế hoạch", expanded=False):
            st.json(plan)
        with st.spinner("🤖 Đang sinh nội dung..."):
            prompt = build_exec_prompt(plan)
            result = call_ai(prompt)
        if result:
            st.session_state["draft"] = result
            st.session_state["last_topic"] = plan.get("topic", "")
            if GITHUB_TOKEN:
                save_history(plan.get("task", ""), plan.get("topic", ""), result)


if st.session_state.get("draft"):
    st.divider()
    st.subheader("📄 Kết quả")
    content = st.session_state["draft"]
    safe_name = re.sub(r"[^A-Za-z0-9_\-]", "_", st.session_state.get("last_topic", "tai_lieu"))[:45] or "tai_lieu"
    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            "⬇️ Tải Word",
            make_doc("Tài liệu", content),
            file_name=safe_name + ".docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )
    with d2:
        st.download_button(
            "⬇️ Tải TXT",
            content.encode("utf-8"),
            file_name=safe_name + ".txt",
            mime="text/plain",
            use_container_width=True,
        )
    st.session_state["draft"] = st.text_area(
        "Nội dung (có thể sửa)", content, height=500, key="edit_draft"
    )


st.divider()
st.subheader("📅 Tạo lịch báo giảng từ TKB + PPCT")
st.caption("Upload TKB và PPCT (PDF/DOCX/Excel) → AI tự đọc → tạo lịch báo giảng")

with st.expander("📖 Hướng dẫn sử dụng", expanded=False):
    st.markdown("""
**Bước 1:** Upload **Thời khóa biểu** (TKB) — file PDF, DOCX hoặc Excel do trường cung cấp.

**Bước 2:** Upload **Phân phối chương trình** (PPCT) — file PDF, DOCX hoặc Excel.

**Bước 3:** Nhập **ngày bắt đầu tuần** (thứ 2 của tuần đó).

**Bước 4:** Bấm **Tạo lịch báo giảng** — AI tự đọc file, trích xuất TKB và PPCT, ghép lại thành lịch.

**Lưu ý:**
- File PDF scan (ảnh) → cần OCR trước hoặc dán text thủ công
- File Excel có nhiều sheet → AI sẽ đọc sheet đầu tiên
- Nếu file phức tạp, có thể dán text thủ công vào ô bên dưới
""")

st.markdown("### 📘 1. Thời khóa biểu")
tkb_file = st.file_uploader(
    "Upload file TKB (PDF/DOCX/Excel)",
    type=["pdf", "docx", "xlsx", "xls", "csv", "txt"],
    key="tkb_file",
)
tkb_manual = st.text_area(
    "Hoặc dán TKB trực tiếp (nếu file không đọc được)",
    height=120,
    placeholder="Thứ 2 | Tiết 1 | 6A | Toán\nThứ 2 | Tiết 3 | 6B | Toán\n...",
    key="tkb_manual",
)

st.markdown("### 📗 2. Phân phối chương trình")
ppct_file = st.file_uploader(
    "Upload file PPCT (PDF/DOCX/Excel)",
    type=["pdf", "docx", "xlsx", "xls", "csv", "txt"],
    key="ppct_file",
)
ppct_manual = st.text_area(
    "Hoặc dán PPCT trực tiếp",
    height=120,
    placeholder="Tiết 22 | Bài 11: Ước chung lớn nhất | 2 tiết\nTiết 24 | Bài 12: Bội chung nhỏ nhất | 2 tiết\n...",
    key="ppct_manual",
)

st.markdown("### 📆 3. Ngày bắt đầu tuần")
start_date = st.date_input(
    "Chọn thứ 2 của tuần cần tạo lịch",
    value=st.session_state["start_date_val"],
    key="start_date_input",
)
st.session_state["start_date_val"] = start_date


def ai_extract_tkb(raw_text):
    system = "Ban la tro ly doc bang. Tra ve CHI text theo format, khong giai thich."
    prompt = "Doc thong tin Thoi khoa bieu sau, chuyen thanh format moi dong 1 tiet:\n"
    prompt += "Thu | Tiet | Lop | Mon\n\n"
    prompt += "Vi du:\nThu 2 | Tiet 1 | 6A | Toan\nThu 2 | Tiet 3 | 6B | Toan\n\n"
    prompt += "NOI DUNG TKB:\n" + raw_text[:12000]
    prompt += "\n\nCHI tra ve cac dong theo format tren."
    try:
        result = call_ai(prompt, system)
        lines = []
        for line in result.splitlines():
            line = line.strip()
            if "|" in line and len(line.split("|")) >= 4:
                lines.append(line)
        return "\n".join(lines)
    except Exception:
        return ""


def ai_extract_ppct(raw_text):
    system = "Ban la tro ly doc bang. Tra ve CHI text theo format, khong giai thich."
    prompt = "Doc thong tin Phan phoi chuong trinh sau, chuyen thanh format moi dong 1 bai:\n"
    prompt += "Tiet bat dau | Ten bai | So tiet\n\n"
    prompt += "Vi du:\nTiet 22 | Bai 11: Uoc chung lon nhat | 2 tiet\nTiet 24 | Bai 12: Boi chung nho nhat | 2 tiet\n\n"
    prompt += "NOI DUNG PPCT:\n" + raw_text[:12000]
    prompt += "\n\nCHI tra ve cac dong theo format tren."
    try:
        result = call_ai(prompt, system)
        lines = []
        for line in result.splitlines():
            line = line.strip()
            if "|" in line and len(line.split("|")) >= 2:
                lines.append(line)
        return "\n".join(lines)
    except Exception:
        return ""


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
        if thu not in result:
            result[thu] = []
        result[thu].append({
            "tiet": parts[1],
            "lop": parts[2],
            "mon": parts[3],
        })
    for thu in result:
        result[thu].sort(key=lambda x: x["tiet"])
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
        result.append({
            "tiet_bd": tiet_start,
            "ten_bai": parts[1],
            "so_tiet": n_tiet,
            "so_tiet_str": so_tiet_str,
        })
    result.sort(key=lambda x: x["tiet_bd"])
    return result


if st.button("📅 Tạo lịch báo giảng", type="primary", use_container_width=True):
    if not available_providers:
        st.error("Cần có ít nhất 1 API key để AI đọc file.")
    else:
        tkb_raw = tkb_manual.strip()
        if tkb_file:
            _, text, _ = read_file(tkb_file)
            tkb_raw = text
            st.info("📘 Đã đọc TKB: " + str(len(tkb_raw)) + " ký tự")

        ppct_raw = ppct_manual.strip()
        if ppct_file:
            _, text, _ = read_file(ppct_file)
            ppct_raw = text
            st.info("📗 Đã đọc PPCT: " + str(len(ppct_raw)) + " ký tự")

        if not tkb_raw or not ppct_raw:
            st.error("Cần có cả TKB và PPCT.")
        else:
            with st.spinner("🧠 AI đang đọc TKB..."):
                if tkb_raw and "|" in tkb_raw and len(tkb_raw.splitlines()[0].split("|")) >= 4:
                    tkb_clean = tkb_raw
                else:
                    tkb_clean = ai_extract_tkb(tkb_raw)

            with st.spinner("🧠 AI đang đọc PPCT..."):
                if ppct_raw and "|" in ppct_raw and len(ppct_raw.splitlines()[0].split("|")) >= 2:
                    ppct_clean = ppct_raw
                else:
                    ppct_clean = ai_extract_ppct(ppct_raw)

            with st.expander("📋 Xem kết quả AI trích xuất"):
                st.markdown("**TKB đã chuẩn hóa:**")
                st.code(tkb_clean)
                st.markdown("**PPCT đã chuẩn hóa:**")
                st.code(ppct_clean)

            tkb = parse_tkb(tkb_clean)
            ppct = parse_ppct(ppct_clean)

            if not tkb or not ppct:
                st.error("Không parse được TKB hoặc PPCT. Kiểm tra lại nội dung.")
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
                                    "thu": k,
                                    "tiet": item["tiet"],
                                    "lop": item["lop"],
                                    "mon": item["mon"],
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
                    if idx < len(ppct_tiets):
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
                st.success("✅ Đã tạo lịch báo giảng: " + str(len(lich)) + " tiết")


if st.session_state.get("lich_bao_giang"):
    lich = st.session_state["lich_bao_giang"]
    st.divider()
    st.subheader("📋 Lịch báo giảng — " + str(len(lich)) + " tiết")

    st.markdown("| Thứ | Ngày | Tiết | Lớp | Môn | Tên bài | Tiết |")
    st.markdown("|---|---|---|---|---|---|---|")
    for r in lich:
        st.markdown(
            "| " + r["thu"] + " | " + r["ngay"] + " | " + r["tiet"]
            + " | " + r["lop"] + " | " + r["mon"]
            + " | " + r["ten_bai"]
            + " | " + str(r["tiet_trong_bai"]) + "/" + str(r["tong_tiet"]) + " |"
        )

    if st.button("📄 Xuất lịch báo giảng ra Word", type="primary", use_container_width=True):
        d = Document()
        s = d.sections[0]
        s.top_margin = Inches(0.7)
        s.bottom_margin = Inches(0.7)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)
        d.styles["Normal"].font.name = "Times New Roman"
        d.styles["Normal"].font.size = Pt(12)

        mem = st.session_state.get("memory", {})
        d.add_heading("LỊCH BÁO GIẢNG TUẦN", 0)
        if mem.get("school"):
            d.add_paragraph("Đơn vị: " + mem["school"])
        if mem.get("teacher_name"):
            d.add_paragraph("Giáo viên: " + mem["teacher_name"])
        d.add_paragraph("Tuần bắt đầu: " + st.session_state["start_date_val"].strftime("%d/%m/%Y"))

        table = d.add_table(rows=1, cols=7)
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        for i, h in enumerate(["Thứ", "Ngày", "Tiết", "Lớp", "Môn", "Tên bài", "Tiết/Bài"]):
            hdr[i].text = h
        for r in lich:
            row = table.add_row().cells
            row[0].text = r["thu"]
            row[1].text = r["ngay"]
            row[2].text = r["tiet"]
            row[3].text = r["lop"]
            row[4].text = r["mon"]
            row[5].text = r["ten_bai"]
            row[6].text = str(r["tiet_trong_bai"]) + "/" + str(r["tong_tiet"])

        bio = io.BytesIO()
        d.save(bio)
        st.download_button(
            "⬇️ Tải Word lịch báo giảng",
            bio.getvalue(),
            file_name="lich_bao_giang_" + st.session_state["start_date_val"].strftime("%Y%m%d") + ".docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )

    if st.button("🚀 Soạn hàng loạt các bài trong lịch này", type="primary", use_container_width=True):
        if not available_providers:
            st.error("Chưa có API key.")
        else:
            progress = st.progress(0)
            status_text = st.empty()
            results = []
            for idx, r in enumerate(lich):
                status_text.write(
                    "⏳ Soạn bài " + str(idx + 1) + "/" + str(len(lich))
                    + ": " + r["ten_bai"]
                )
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
                    prompt = build_exec_prompt(plan)
                    content = call_ai(prompt)
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

            status_text.write("✅ Đã soạn xong " + str(len(results)) + " tiết!")
            st.session_state["weekly_results"] = results


if st.session_state.get("weekly_results"):
    results = st.session_state["weekly_results"]
    st.divider()
    st.subheader("📚 Kết quả tuần — " + str(len(results)) + " bài")
    for r in results:
        st.write("- " + r.get("thu", "") + " | " + r.get("ten_bai", ""))

    if st.button("📄 Xuất 1 file Word gộp tất cả bài", type="primary", use_container_width=True):
        d = Document()
        s = d.sections[0]
        s.top_margin = Inches(0.7)
        s.bottom_margin = Inches(0.7)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)
        d.styles["Normal"].font.name = "Times New Roman"
        d.styles["Normal"].font.size = Pt(12)

        mem = st.session_state.get("memory", {})
        d.add_heading("KẾ HOẠCH BÀI DẠY TUẦN", 0)
        if mem.get("school"):
            d.add_paragraph("Đơn vị: " + mem["school"])
        if mem.get("teacher_name"):
            d.add_paragraph("Giáo viên: " + mem["teacher_name"])
        d.add_paragraph("Ngày tạo: " + datetime.now().strftime("%d/%m/%Y %H:%M"))
        d.add_paragraph("Tổng số bài: " + str(len(results)))

        for idx, r in enumerate(results):
            d.add_page_break()
            d.add_heading("BÀI " + str(idx + 1) + ": " + r.get("ten_bai", ""), 1)
            d.add_paragraph(
                "Thứ: " + r.get("thu", "") + " | " + r.get("mon_lop", "")
                + " | " + r.get("so_tiet", "")
            )
            d.add_paragraph("---")
            write_content_to_doc(d, r.get("content", ""))

        bio = io.BytesIO()
        d.save(bio)
        st.success("✅ Đã tạo file Word gộp " + str(len(results)) + " bài.")
        st.download_button(
            "⬇️ Tải file Word gộp",
            bio.getvalue(),
            file_name="ke_hoach_tuan_" + datetime.now().strftime("%Y%m%d") + ".docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )
