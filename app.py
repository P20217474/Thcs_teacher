import subprocess
import sys
import os
import json
import base64
import urllib.request
from datetime import datetime


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
st.caption("Bộ não tự động • Nhớ mọi thứ • Tự ra quyết định")


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


def append_skill(s):
    old = load_skills()
    new = old + "\n- [" + datetime.now().strftime("%Y-%m-%d") + "] " + s
    return gh_save("data/skills/README.md", new, "Learn")


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


tab_up, tab_kho = st.tabs(["📤 Upload mẫu/nguồn (tùy chọn)", "📂 Kho tri thức"])

with tab_up:
    st.markdown("**MẪU** (định dạng) và **NGUỒN** (nội dung) — đều tùy chọn.")
    up_type = st.radio("Loại", ["Mẫu (định dạng)", "Nguồn (nội dung)"], horizontal=True)
    up_name = st.text_input("Tên file (không cần đuôi)", placeholder="giao_an_toan6")
    up_files = st.file_uploader(
        "Chọn file",
        type=["docx", "pdf", "txt", "md", "csv"],
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


st.divider()
st.subheader("💬 Ra lệnh cho trợ lý")
st.caption("Ví dụ: Soạn giáo án Toán 6 bài Ước chung lớn nhất, 2 tiết")

cmd = st.text_area(
    "Lệnh",
    height=100,
    placeholder="Soạn giáo án Toán 6 bài Ước chung lớn nhất...",
    label_visibility="collapsed",
)

c1, c2 = st.columns([3, 1])
with c1:
    run = st.button("▶️ Thực hiện", type="primary", use_container_width=True)
with c2:
    if st.button("🗑️ Xóa", use_container_width=True):
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


def make_doc(title, body):
    body = fix_math(body)
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

    lines = body.splitlines()
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

    b = io.BytesIO()
    d.save(b)
    return b.getvalue()


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
        "Nội dung (có thể sửa)", content, height=500
    )
