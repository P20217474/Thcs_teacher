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


# ============ ĐỌC CẤU HÌNH ============
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
    "OPENROUTER_API_KEY": ["google/gemini-2.0-flash-exp:free", "meta-llama/llama-3.3-70b-instruct:free"],
    "NVIDIA_API_KEY": ["meta/llama-3.3-70b-instruct"],
}

GITHUB_TOKEN = cfg("GITHUB_TOKEN", "")
GITHUB_REPO = cfg("GITHUB_REPO", "P20217474/Thcs_teacher")
GITHUB_BRANCH = cfg("GITHUB_BRANCH", "main")


# ============ GITHUB API ============
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


def gh_delete(path):
    _, sha = gh_get(path)
    if not sha:
        return True
    res = gh_api(
        "DELETE",
        f"/repos/{GITHUB_REPO}/contents/{path}",
        {"message": "Delete", "sha": sha, "branch": GITHUB_BRANCH},
    )
    return res and "error" not in res


def gh_list(folder):
    res = gh_api("GET", f"/repos/{GITHUB_REPO}/contents/{folder}?ref={GITHUB_BRANCH}")
    if not res or "error" in res or not isinstance(res, list):
        return []
    return [item["name"] for item in res if item["type"] == "file"]

# ============ BỘ NHỚ DÀI HẠN ============
def load_memory():
    content, _ = gh_get("data/memory.json")
    if content:
        try:
            return json.loads(content)
        except Exception:
            pass
    return {
        "teacher_name": "",
        "school": "",
        "default_subject": "Toan",
        "default_grade": "6",
        "preferences": {},
        "created_at": datetime.now().strftime("%Y-%m-%d"),
    }


def save_memory(memory):
    memory["last_updated"] = datetime.now().strftime("%Y-%m-%d")
    return gh_save("data/memory.json", json.dumps(memory, ensure_ascii=False, indent=2), "Update memory")


def load_templates():
    """Đọc tất cả mẫu trong data/templates/."""
    files = gh_list("data/templates")
    templates = {}
    for fname in files:
        content, _ = gh_get(f"data/templates/{fname}")
        if content:
            templates[fname] = content
    return templates


def load_sources():
    """Đọc tất cả nguồn trong data/sources/."""
    files = gh_list("data/sources")
    sources = {}
    for fname in files:
        content, _ = gh_get(f"data/sources/{fname}")
        if content:
            sources[fname] = content
    return sources


def save_history(task, topic, content):
    """Lưu lịch sử soạn thảo."""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    fname = f"data/history/{ts}_{re.sub(r'[^a-zA-Z0-9]', '_', topic)[:30]}.md"
    md = f"# {task}\n\n**Chủ đề:** {topic}\n**Thời gian:** {datetime.now().strftime('%d/%m/%Y %H:%M')}\n\n---\n\n{content}"
    return gh_save(fname, md, f"History: {topic}")


def load_skills():
    """Đọc các kỹ năng đã học."""
    content, _ = gh_get("data/skills/README.md")
    return content or ""


def append_skill(new_skill):
    """Thêm kỹ năng mới vào file skills."""
    old = load_skills()
    updated = old + "\n- [" + datetime.now().strftime("%Y-%m-%d") + "] " + new_skill
    return gh_save("data/skills/README.md", updated, "Learn new skill")


# ============ ĐỌC FILE ============
def read_docx(data):
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


def read_pdf(data):
    if not HAS_PYPDF:
        return "[Chưa cài pypdf]"
    try:
        reader = PdfReader(io.BytesIO(data))
        texts = [(p.extract_text() or "") for p in reader.pages]
        full = "\n".join(t for t in texts if t.strip())
        if len(full) < 100:
            return "[PDF scan, không có text]" + full
        return full
    except Exception as e:
        return "[Lỗi pdf: " + str(e) + "]"


def read_csv_text(data):
    try:
        text = data.decode("utf-8", errors="ignore")
        rows = list(csv.reader(io.StringIO(text)))
        if not rows:
            return "[CSV rỗng]"
        out = ["| " + " | ".join(rows[0]) + " |", "|" + "---|" * len(rows[0])]
        for row in rows[1:]:
            row = (row + [""] * len(rows[0]))[:len(rows[0])]
            out.append("| " + " | ".join(row) + " |")
        return "\n".join(out)
    except Exception as e:
        return "[Lỗi csv: " + str(e) + "]"


def read_file(file):
    if file is None:
        return ("", "", 0)
    data = file.read()
    if not data:
        return (file.name, "[File rỗng]", 0)
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
        text = "[" + file.name + ": ảnh, cần OCR]"
    else:
        text = "[Định dạng chưa hỗ trợ]"
    return (file.name, text, len(text))


# ============ KHỞI TẠO BỘ NHỚ ============
if "memory" not in st.session_state:
    st.session_state["memory"] = load_memory() if GITHUB_TOKEN else {}
if "templates" not in st.session_state:
    st.session_state["templates"] = load_templates() if GITHUB_TOKEN else {}
if "sources" not in st.session_state:
    st.session_state["sources"] = load_sources() if GITHUB_TOKEN else {}
if "skills" not in st.session_state:
    st.session_state["skills"] = load_skills() if GITHUB_TOKEN else ""
if "chat" not in st.session_state:
    st.session_state["chat"] = []
if "draft" not in st.session_state:
    st.session_state["draft"] = ""


# ============ SIDEBAR ============
with st.sidebar:
    st.header("🧠 Bộ não")
    mem = st.session_state["memory"]
    
    if GITHUB_TOKEN:
        st.success("💾 Bộ nhớ vĩnh viễn đang bật")
    else:
        st.warning("⚠️ Chưa có GITHUB_TOKEN. Bộ nhớ tạm thời.")
    
    st.divider()
    st.subheader("📋 Thông tin cơ bản")
    mem["teacher_name"] = st.text_input("Giáo viên", value=mem.get("teacher_name", ""))
    mem["school"] = st.text_input("Trường", value=mem.get("school", ""))
    mem["default_subject"] = st.selectbox(
        "Môn mặc định",
        ["Toan", "Ngu van", "Tieng Anh", "Vat li", "Hoa hoc", "Sinh hoc", "Tin hoc"],
        index=["Toan", "Ngu van", "Tieng Anh", "Vat li", "Hoa hoc", "Sinh hoc", "Tin hoc"].index(mem.get("default_subject", "Toan")),
    )
    mem["default_grade"] = st.selectbox(
        "Khối mặc định", ["6", "7", "8", "9"],
        index=["6", "7", "8", "9"].index(mem.get("default_grade", "6")),
    )
    
    if st.button("💾 Lưu vào bộ nhớ", use_container_width=True):
        if GITHUB_TOKEN:
            if save_memory(mem):
                st.success("✅ Đã lưu vào bộ nhớ vĩnh viễn")
        else:
            st.warning("Chưa có GITHUB_TOKEN")
    
    st.divider()
    st.subheader("📚 Kho tri thức")
    st.write("**Mẫu:** " + str(len(st.session_state["templates"])) + " file")
    st.write("**Nguồn:** " + str(len(st.session_state["sources"])) + " file")
    
    if st.button("🔄 Tải lại từ GitHub", use_container_width=True):
        st.session_state["templates"] = load_templates()
        st.session_state["sources"] = load_sources()
        st.session_state["memory"] = load_memory()
        st.rerun()
    
    st.divider()
    st.write("**AI:** " + str(len(available_providers)) + " nhà cung cấp")
    for k in available_providers:
        st.caption("✅ " + k.replace("_API_KEY", ""))

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


# ============ BỘ NÃO PLANNER ============
def analyze_command(cmd):
    """Bộ não phân tích lệnh → kế hoạch hành động."""
    mem = st.session_state["memory"]
    templates_list = list(st.session_state["templates"].keys())
    sources_list = list(st.session_state["sources"].keys())
    
    system = "Bạn là bộ não phân tích lệnh cho trợ lý giáo viên. Trả về JSON."
    prompt = f"""Phân tích lệnh sau và trả về JSON với cấu trúc:
{{
  "task": "loại công việc",
  "subject": "môn học",
  "grade": "khối lớp",
  "topic": "chủ đề",
  "period": "thời lượng",
  "template_files": ["tên file mẫu cần dùng"],
  "source_files": ["tên file nguồn cần dùng"],
  "needs_clarification": false,
  "clarification_question": ""
}}

LỆNH: {cmd}

THÔNG TIN CÓ SẴN:
- Giáo viên: {mem.get('teacher_name', 'chưa rõ')}
- Trường: {mem.get('school', 'chưa rõ')}
- Môn mặc định: {mem.get('default_subject', 'Toan')}
- Khối mặc định: {mem.get('default_grade', '6')}
- Mẫu có sẵn: {templates_list}
- Nguồn có sẵn: {sources_list}

Trả về CHỈ JSON, không giải thích."""
    
    try:
        result = call_ai(prompt, system)
        # Tìm JSON trong kết quả
        m = re.search(r"\{[\s\S]*\}", result)
        if m:
            return json.loads(m.group(0))
    except Exception as e:
        st.warning("Bộ não không phân tích được lệnh: " + str(e))
    
    # Fallback: phân tích thủ công
    low = cmd.lower()
    return {
        "task": "Soạn kế hoạch bài dạy" if "giáo án" in low else "Soạn nội dung theo yêu cầu riêng",
        "subject": mem.get("default_subject", "Toan"),
        "grade": mem.get("default_grade", "6"),
        "topic": cmd,
        "period": "",
        "template_files": templates_list[:1],
        "source_files": sources_list[:1],
        "needs_clarification": False,
        "clarification_question": "",
    }


def build_execution_prompt(plan):
    """Bộ não xây dựng prompt thực thi từ kế hoạch."""
    mem = st.session_state["memory"]
    
    # Lấy nội dung mẫu
    tpl_text = ""
    for fname in plan.get("template_files", []):
        if fname in st.session_state["templates"]:
            tpl_text += "\n=== MẪU: " + fname + " ===\n"
            tpl_text += st.session_state["templates"][fname][:15000]
    
    # Lấy nội dung nguồn
    src_text = ""
    for fname in plan.get("source_files", []):
        if fname in st.session_state["sources"]:
            src_text += "\n=== NGUỒN: " + fname + " ===\n"
            src_text += st.session_state["sources"][fname][:25000]
    
    skills = st.session_state.get("skills", "")
    
    prompt = "Bạn là trợ lý chuyên môn cao cấp cho giáo viên THCS Việt Nam.\n\n"
    prompt += "GIÁO VIÊN: " + mem.get("teacher_name", "chưa rõ") + "\n"
    prompt += "TRƯỜNG: " + mem.get("school", "chưa rõ") + "\n\n"
    prompt += "NHIỆM VỤ: " + plan.get("task", "") + "\n"
    prompt += "MÔN: " + plan.get("subject", "") + " | KHỐI: " + plan.get("grade", "") + "\n"
    prompt += "CHỦ ĐỀ: " + plan.get("topic", "") + "\n"
    prompt += "THỜI LƯỢNG: " + (plan.get("period") or "Chưa xác định") + "\n\n"
    
    if tpl_text:
        prompt += "=== FILE MẪU (BẮT BUỘC BÁM THEO) ===" + tpl_text + "\n=== HẾT MẪU ===\n\n"
    if src_text:
        prompt += "=== NGUỒN TÀI LIỆU ===" + src_text + "\n=== HẾT NGUỒN ===\n\n"
    if skills:
        prompt += "=== KỸ NĂNG ĐÃ HỌC ===\n" + skills[:2000] + "\n\n"
    
    prompt += "QUY TẮC:\n"
    prompt += "1. NẾU CÓ MẪU: bám sát cấu trúc, thứ tự mục, cách đánh số.\n"
    prompt += "2. NẾU CÓ NGUỒN: lấy nội dung từ nguồn, KHÔNG bịa.\n"
    prompt += "3. Viết CHI TIẾT, cụ thể, tối thiểu 1500 từ cho kế hoạch bài dạy.\n"
    prompt += "4. Với mỗi hoạt động: mục tiêu, cách tổ chức, sản phẩm.\n"
    prompt += "5. Bảng markdown: | Cột 1 | Cột 2 |, dùng <br> và **bold** trong ô.\n"
    prompt += "6. Ký hiệu toán: x², √x, H₂O, ≤, ∈, π, ° (KHÔNG dùng LaTeX).\n"
    prompt += "7. Tiếng Việt có dấu, văn phong sư phạm.\n"
    prompt += "8. Markdown: '# ', '## ', '### ', '- ', '1. '.\n"
    prompt += "9. Chỉ xuất nội dung tài liệu, không lời chào.\n"
    
    return prompt


def reflect_on_result(plan, result):
    """Bộ não tự đánh giá kết quả → học kỹ năng mới."""
    if not GITHUB_TOKEN:
        return
    prompt = f"""Đánh giá tài liệu sau có tốt không. Trả lời ngắn gọn 1 câu, có "TỐT" hoặc "CẦN CẢI THIỆN".

Loại: {plan.get('task')}
Chủ đề: {plan.get('topic')}

Nội dung (500 ký tự đầu):
{result[:500]}

Trả lời:"""
    try:
        judge = call_ai(prompt)
        if "CẦN CẢI THIỆN" in judge.upper():
            # Ghi lại như một bài học
            append_skill(f"Cần chú ý khi làm {plan.get('task')} chủ đề {plan.get('topic')}")
    except Exception:
        pass


# ============ THỰC THI ============
def execute_command(cmd):
    """Bộ não thực thi lệnh đầy đủ: phân tích → lập kế hoạch → sinh → lưu."""
    if not available_providers:
        st.error("Chưa có API key.")
        return None, None
    
    with st.spinner("🧠 Bộ não đang phân tích lệnh..."):
        plan = analyze_command(cmd)
    
    with st.expander("📋 Kế hoạch của bộ não", expanded=False):
        st.json(plan)
    
    if plan.get("needs_clarification") and plan.get("clarification_question"):
        st.warning("⚠️ Bộ não cần làm rõ: " + plan["clarification_question"])
    
    with st.spinner("🤖 Đang sinh nội dung..."):
        prompt = build_execution_prompt(plan)
        result = call_ai(prompt)
    
    if not result:
        st.error("Không sinh được nội dung.")
        return None, plan
    
    # Lưu lịch sử
    if GITHUB_TOKEN:
        with st.spinner("💾 Đang lưu lịch sử..."):
            save_history(plan.get("task", ""), plan.get("topic", ""), result)
            reflect_on_result(plan, result)
    
    return result, plan

# ============ GIAO DIỆN CHÍNH ============
st.divider()
st.subheader("💬 Ra lệnh cho trợ lý")
st.caption("Nhập lệnh tự nhiên. Ví dụ: `Soạn giáo án Toán 6 bài Ước chung lớn nhất, 2 tiết`")

cmd = st.text_area(
    "Lệnh",
    height=100,
    placeholder="Soạn giáo án Toán 6 bài Ước chung lớn nhất, có ví dụ thực tiễn...",
    label_visibility="collapsed",
    key="main_cmd",
)

cbtn1, cbtn2 = st.columns([3, 1])
with cbtn1:
    run = st.button("▶️ Thực hiện", type="primary", use_container_width=True)
with cbtn2:
    if st.button("🗑️ Xóa", use_container_width=True):
        st.session_state["draft"] = ""
        st.rerun()

if run and cmd.strip():
    result, plan = execute_command(cmd)
    if result:
        st.session_state["draft"] = result
        st.session_state["draft_title"] = plan.get("task", "Tài liệu") if plan else "Tài liệu"
        st.session_state["last_topic"] = plan.get("topic", "") if plan else ""


# ============ BẢN NHÁP ============
if st.session_state.get("draft"):
    st.divider()
    st.subheader("📄 Kết quả")
    
    # Nút tải
    d = st.session_state["draft"]
    name = re.sub(r"[^A-Za-z0-9_\-]", "_", st.session_state.get("last_topic", "tai_lieu"))[:45] or "tai_lieu"
    
    dc1, dc2, dc3 = st.columns([1, 1, 2])
    with dc1:
        st.download_button(
            "⬇️ Word",
            make_doc(st.session_state.get("draft_title", "Tài liệu"), d),
            file_name=name + ".docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )
    with dc2:
        st.download_button(
            "⬇️ TXT",
            d.encode("utf-8"),
            file_name=name + ".txt",
            mime="text/plain",
            use_container_width=True,
        )
    
    # Cho phép sửa
    st.session_state["draft"] = st.text_area(
        "Nội dung (có thể chỉnh sửa)",
        st.session_state["draft"],
        height=500,
    )


# ============ LỊCH SỬ ============
with st.expander("🕘 Lịch sử soạn thảo"):
    if not GITHUB_TOKEN:
        st.info("Cần GITHUB_TOKEN để lưu lịch sử.")
    else:
        files = gh_list("data/history")
        if not files:
            st.info("Chưa có lịch sử.")
        else:
            for f in sorted(files, reverse=True)[:20]:
                if st.button("📄 " + f, key="hist_" + f):
                    content, _ = gh_get("data/history/" + f)
                    if content:
                        st.session_state["draft"] = content
                        st.rerun()


# ============ HÀM XUẤT WORD ============
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
                else
