import subprocess
import sys
import os


def _ensure(pkg, mod=None):
    mod = mod or pkg
    try:
        __import__(mod)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "--quiet"])


_ensure("streamlit")
_ensure("python-docx", "docx")
_ensure("pypdf", "pypdf")
_ensure("freelm")

import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from datetime import datetime
import io
import re

try:
    from pypdf import PdfReader
    HAS_PYPDF = True
except Exception:
    HAS_PYPDF = False

st.set_page_config(
    page_title="Agent Tro ly Giao vien THCS AI",
    page_icon="🎓",
    layout="wide",
)

st.title("🎓 Agent Tro ly Giao vien THCS AI")
st.caption("Ban 3.1 • freelm • 6 nha cung cap AI • Tu dong chuyen doi • Xuat Word")


# ============ ĐỌC CẤU HÌNH TỪ SECRETS ============
def cfg(key, default=""):
    try:
        val = st.secrets.get(key, default)
        return val if val else default
    except Exception:
        return default


# Đọc tất cả API key có thể có
GEMINI_KEY = cfg("GEMINI_API_KEY", "")
GROQ_KEY = cfg("GROQ_API_KEY", "")
OPENROUTER_KEY = cfg("OPENROUTER_API_KEY", "")
MISTRAL_KEY = cfg("MISTRAL_API_KEY", "")
CEREBRAS_KEY = cfg("CEREBRAS_API_KEY", "")
NVIDIA_KEY = cfg("NVIDIA_API_KEY", "")

# Copy tất cả key vào os.environ để freelm đọc
# (freelm đọc từ biến môi trường)
for k, v in {
    "GEMINI_API_KEY": GEMINI_KEY,
    "GROQ_API_KEY": GROQ_KEY,
    "OPENROUTER_API_KEY": OPENROUTER_KEY,
    "MISTRAL_API_KEY": MISTRAL_KEY,
    "CEREBRAS_API_KEY": CEREBRAS_KEY,
    "NVIDIA_API_KEY": NVIDIA_KEY,
}.items():
    if v:
        os.environ[k] = v

DEFAULT_SCHOOL = cfg("SCHOOL", "")
DEFAULT_TEACHER = cfg("TEACHER", "")
DEFAULT_SUBJECT = cfg("SUBJECT", "Toan")
DEFAULT_GRADE = cfg("GRADE", "6")


# ============ SESSION STATE ============
for k, v in {
    "draft": "",
    "draft_title": "",
    "template_text": "",
    "source_text": "",
    "template_names": [],
    "source_names": [],
}.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ============ ĐỌC FILE ============
def read_docx(data: bytes) -> str:
    try:
        d = Document(io.BytesIO(data))
        return "\n".join(p.text for p in d.paragraphs if p.text.strip())
    except Exception as e:
        return "[Loi docx: " + str(e) + "]"


def read_pdf(data: bytes) -> str:
    if not HAS_PYPDF:
        return "[Chua cai pypdf]"
    try:
        reader = PdfReader(io.BytesIO(data))
        return "\n".join((pg.extract_text() or "") for pg in reader.pages)
    except Exception as e:
        return "[Loi pdf: " + str(e) + "]"


def read_file(file):
    if file is None:
        return ("", "", 0)
    data = file.read()
    if not data:
        return (file.name, "[File rong]", 0)
    name = file.name.lower()
    if name.endswith(".docx"):
        text = read_docx(data)
    elif name.endswith(".pdf"):
        text = read_pdf(data)
    elif name.endswith((".txt", ".md")):
        text = data.decode("utf-8", errors="ignore")
    else:
        text = "[Khong ho tro]"
    return (file.name, text, len(text))


# ============ SIDEBAR ============
SUBJECTS = [
    "Toan", "Ngu van", "Tieng Anh", "Khoa hoc tu nhien",
    "Vat li", "Hoa hoc", "Sinh hoc", "Lich su va Dia li",
    "Tin hoc", "Cong nghe", "GDCD", "Nghe thuat",
    "Giao duc the chat", "Hoat dong trai nghiem", "Khac",
]
GRADES = ["6", "7", "8", "9", "Nhieu khoi"]

if DEFAULT_SUBJECT not in SUBJECTS:
    DEFAULT_SUBJECT = "Toan"
if DEFAULT_GRADE not in GRADES:
    DEFAULT_GRADE = "6"


with st.sidebar:
    st.header("⚙️ Cấu hình AI")

    # Đếm số nhà cung cấp có key
    available = []
    if GEMINI_KEY:
        available.append("Gemini")
    if GROQ_KEY:
        available.append("Groq")
    if OPENROUTER_KEY:
        available.append("OpenRouter")
    if MISTRAL_KEY:
        available.append("Mistral")
    if CEREBRAS_KEY:
        available.append("Cerebras")
    if NVIDIA_KEY:
        available.append("NVIDIA")

    if available:
        st.success("✅ " + ", ".join(available))
        use_ai = st.checkbox("Dùng AI để sinh nội dung", value=True)
        st.caption("freelm tự chuyển đổi khi 1 nhà cung cấp lỗi.")
    else:
        st.warning("Chưa có API key nào.")
        st.caption("Nhập ít nhất 1 key để dùng AI.")
        use_ai = st.checkbox("Dùng AI để sinh nội dung", value=False)

    # Cho phép nhập key tạm thời nếu chưa có
    if not available:
        with st.expander("🔑 Nhập API key tạm thời"):
            temp_gemini = st.text_input("Gemini API Key", type="password")
            temp_groq = st.text_input("Groq API Key", type="password")
            temp_openrouter = st.text_input("OpenRouter API Key", type="password")
            if temp_gemini.strip():
                os.environ["GEMINI_API_KEY"] = temp_gemini.strip()
                GEMINI_KEY = temp_gemini.strip()
            if temp_groq.strip():
                os.environ["GROQ_API_KEY"] = temp_groq.strip()
                GROQ_KEY = temp_groq.strip()
            if temp_openrouter.strip():
                os.environ["OPENROUTER_API_KEY"] = temp_openrouter.strip()
                OPENROUTER_KEY = temp_openrouter.strip()

    st.divider()
    st.header("📋 Thông tin")
    school = st.text_input("Trường/đơn vị", value=DEFAULT_SCHOOL)
    teacher = st.text_input("Giáo viên", value=DEFAULT_TEACHER)
    subject = st.selectbox(
        "Môn học", SUBJECTS, index=SUBJECTS.index(DEFAULT_SUBJECT)
    )
    grade = st.selectbox(
        "Khối lớp", GRADES, index=GRADES.index(DEFAULT_GRADE)
    )

    st.divider()
    st.subheader("📝 Công việc")
    task = st.selectbox("Chọn công việc", [
        "Soan ke hoach bai day",
        "Tao de kiem tra va khung dap an",
        "Tao phieu hoc tap",
        "Soan bao cao/ho so",
        "Lap kich ban video bai giang",
        "Soan noi dung theo yeu cau rieng",
    ])
    topic = st.text_input("Tên bài/chủ đề")
    details = st.text_area("Yêu cầu cụ thể", height=100)
    period = st.text_input("Thời lượng", placeholder="Vi du: 1 tiet / 2 tiet")


# ============ 3 TAB ============
tab_mau, tab_nguon, tab_xem = st.tabs([
    "📐 TÙY CHỌN MẪU (định dạng)",
    "📚 TÙY CHỌN NGUỒN (nội dung)",
    "👁️ XEM TRƯỚC ĐÃ NẠP",
])

with tab_mau:
    st.markdown(
        "**MẪU** quy định cấu trúc, thứ tự mục, cách trình bày. "
        "AI sẽ bám sát cấu trúc này."
    )
    tpl_files = st.file_uploader(
        "📎 Nạp 1 hoặc nhiều file MẪU (docx/pdf/txt)",
        type=["docx", "pdf", "txt"],
        accept_multiple_files=True,
        key="tpl_uploader",
    )
    tpl_paste = st.text_area(
        "📝 Hoặc dán mẫu trực tiếp", height=140, key="tpl_paste"
    )
    if st.button("🗑️ Xóa mẫu đã nạp", key="clear_tpl"):
        st.session_state["template_text"] = ""
        st.session_state["template_names"] = []
        st.rerun()

with tab_nguon:
    st.markdown(
        "**NGUỒN** chứa kiến thức, số liệu, dữ kiện. "
        "AI chỉ lấy nội dung từ nguồn."
    )
    src_files = st.file_uploader(
        "📎 Nạp 1 hoặc nhiều file NGUỒN (docx/pdf/txt)",
        type=["docx", "pdf", "txt"],
        accept_multiple_files=True,
        key="src_uploader",
    )
    src_paste = st.text_area(
        "📝 Hoặc dán nội dung nguồn trực tiếp", height=140, key="src_paste"
    )
    if st.button("🗑️ Xóa nguồn đã nạp", key="clear_src"):
        st.session_state["source_text"] = ""
        st.session_state["source_names"] = []
        st.rerun()


def collect_text(files, paste, store_key, names_key):
    parts = []
    names = []
    if paste and paste.strip():
        parts.append(paste.strip())
        names.append("(van ban dan)")
    if files:
        for f in files:
            nm, text, sz = read_file(f)
            if nm:
                parts.append("--- TAI LIEU: " + nm + " ---\n" + text)
                names.append(nm + " (" + str(sz) + " ky tu)")
    if parts:
        st.session_state[store_key] = "\n\n".join(parts)
        st.session_state[names_key] = names
    return st.session_state[store_key], st.session_state[names_key]


template_text, template_names = collect_text(
    tpl_files, tpl_paste, "template_text", "template_names"
)
source_text, source_names = collect_text(
    src_files, src_paste, "source_text", "source_names"
)

with tab_xem:
    st.subheader("📐 MẪU đã nạp")
    if template_names:
        for n in template_names:
            st.write("- " + n)
        st.caption("Tổng: " + str(len(template_text)) + " ký tự")
        with st.expander("Xem nội dung mẫu"):
            st.text(template_text[:3000])
    else:
        st.info("Chưa nạp mẫu.")
    st.divider()
    st.subheader("📚 NGUỒN đã nạp")
    if source_names:
        for n in source_names:
            st.write("- " + n)
        st.caption("Tổng: " + str(len(source_text)) + " ký tự")
        with st.expander("Xem nội dung nguồn"):
            st.text(source_text[:3000])
    else:
        st.info("Chưa nạp nguồn.")


# ============ GỌI AI QUA FREELM ============
def call_ai(prompt):
    """Gọi AI qua freelm với tự động chuyển đổi nhà cung cấp."""
    import freelm

    llm = freelm.FreeLLM.from_env()
    resp = llm.chat(
        [{"role": "user", "content": prompt}],
        model="auto",
    )
    return resp.text, resp.provider


def build_prompt():
    t = topic.strip() or "[Chua co ten]"

    if template_text:
        tpl_block = (
            "=== FILE MAU (BAT BUOC SAO CHEP NGUYEN CAU TRUC) ===\n"
            + template_text[:12000]
            + "\n=== HET MAU ==="
        )
    else:
        tpl_block = "(Khong co mau - dung cau truc chuan GDPT 2018.)"

    if source_text:
        src_block = (
            "=== NGUON TAI LIEU (chi lay noi dung tu day) ===\n"
            + source_text[:20000]
            + "\n=== HET NGUON ==="
        )
    else:
        src_block = "(Khong co nguon - dung kien thuc chuan mon hoc.)"

    prompt = "Ban la tro ly chuyen mon cho giao vien THCS Viet Nam.\n\n"
    prompt += "NHIEM VU: " + task + "\n"
    prompt += "MON: " + subject + " | KHOI: " + grade + "\n"
    prompt += "CHU DE: " + t + "\n"
    prompt += "THOI LUONG: " + (period or "Chua xac dinh") + "\n"
    prompt += "YEU CAU RIENG: " + (details or "Khong co") + "\n\n"
    prompt += tpl_block + "\n\n" + src_block + "\n\n"

    prompt += "QUY TAC BAT BUOC (phai tuan thu tuyet doi):\n"
    prompt += "1. NEU CO FILE MAU: phai SAO CHEP NGUYEN cau truc cua mau:\n"
    prompt += "   - Giu DUNG thu tu cac muc (I, II, III... hoac 1, 2, 3...)\n"
    prompt += "   - Giu DUNG ten muc, tieu de nhu trong mau\n"
    prompt += "   - Giu DUNG cach danh so, cach trinh bay, vi tri cac muc\n"
    prompt += "   - CHI THAY phan noi dung ben trong, KHONG doi cau truc\n"
    prompt += "   - Neu mau co muc nao thi phai co muc do, khong tu y them/bot\n"
    prompt += "2. NEU CO NGUON: kien thuc, so lieu phai lay tu nguon. KHONG bia.\n"
    prompt += "3. NEU NGUON THIEU: ghi ro [Can bo sung: ...].\n"
    prompt += "4. KY HIEU TOAN HOC - BAT BUOC dung Unicode dep:\n"
    prompt += "   - Luy thua: x², x³, aⁿ (KHONG viet x^2)\n"
    prompt += "   - Chi so duoi: H₂O, CO₂, a₁, a₂ (KHONG viet H_2O)\n"
    prompt += "   - Can bac 2: √x, √(a+b) (KHONG viet sqrt(x))\n"
    prompt += "   - Phan so: viet dang 1/2, 3/4 (KHONG viet \\frac{1}{2})\n"
    prompt += "   - Nhan chia: × ÷ (KHONG viet * /)\n"
    prompt += "   - So sanh: ≤ ≥ ≠ ≈ ≡ (KHONG viet <= >= !=)\n"
    prompt += "   - Tap hop: ∈ ∉ ⊂ ⊃ ∪ ∩ ∅\n"
    prompt += "   - Hy Lap: α β γ δ π θ Δ Σ Ω\n"
    prompt += "   - Mui ten: → ⇒ ↔ ⇔\n"
    prompt += "   - Do: ° (vi du 90°, 45°)\n"
    prompt += "   - TUYET DOI KHONG dung LaTeX: khong dung $...$, \\frac, \\sqrt, ^, _\n"
    prompt += "5. Viet tieng Viet CO DAU day du, van phong su pham, cu the, dung ngay.\n"
    prompt += "6. Markdown: '# ' cap 1, '## ' cap 2, '### ' cap 3, '- ' bullet, '1. ' so.\n"
    prompt += "7. Chi xuat noi dung tai lieu, KHONG them loi chao, KHONG giai thich.\n"
    return prompt


def generate_local():
    t = topic.strip() or "[Dien ten bai]"
    common = (
        "- Mon hoc: " + subject
        + "\n- Lop: " + grade
        + "\n- Chu de: " + t
        + "\n- Thoi gian: " + (period or "Chua xac dinh")
    )
    return (
        "# " + task.upper() + "\n" + common
        + "\n\n## I. Thong tin chung\n" + common
        + "\n\n## II. Noi dung chinh\n"
        + (details or "[Bo sung yeu cau cu the]")
        + "\n\n## III. Ghi chu\n- Ban mau cuc bo."
    )


def generate():
    if use_ai:
        if not (GEMINI_KEY or GROQ_KEY or OPENROUTER_KEY or MISTRAL_KEY or CEREBRAS_KEY or NVIDIA_KEY):
            st.error("Chưa có API key nào.")
            return ""
        try:
            with st.spinner("🤖 freelm đang dò tìm nhà cung cấp..."):
                text, provider = call_ai(build_prompt())
            st.success(f"✅ Đã sinh bằng: {provider}")
            return text
        except Exception as e:
            st.error(f"Lỗi freelm: {e}")
            st.info("Chuyển sang mẫu cục bộ.")
            return generate_local()
    return generate_local()


# ============ CHUYỂN KÝ HIỆU TOÁN ============
def fix_math_symbols(text: str) -> str:
    text = re.sub(r"\^2\b", "²", text)
    text = re.sub(r"\^3\b", "³", text)
    text = re.sub(r"\^4\b", "⁴", text)
    text = re.sub(r"\^5\b", "⁵", text)
    text = re.sub(r"\^0\b", "⁰", text)
    text = re.sub(r"\^1\b", "¹", text)
    text = re.sub(r"\^n\b", "ⁿ", text)
    text = re.sub(r"_2\b", "₂", text)
    text = re.sub(r"_3\b", "₃", text)
    text = re.sub(r"_1\b", "₁", text)
    text = re.sub(r"_0\b", "₀", text)
    text = text.replace("sqrt(", "√(")
    text = text.replace("\\sqrt{", "√")
    text = text.replace("\\times", "×").replace("\\div", "÷")
    text = text.replace("\\leq", "≤").replace("\\geq", "≥")
    text = text.replace("\\neq", "≠").replace("\\approx", "≈")
    text = text.replace("\\in", "∈").replace("\\notin", "∉")
    text = text.replace("\\subset", "⊂").replace("\\cup", "∪")
    text = text.replace("\\cap", "∩").replace("\\emptyset", "∅")
    text = text.replace("\\pi", "π").replace("\\alpha", "α")
    text = text.replace("\\beta", "β").replace("\\gamma", "γ")
    text = text.replace("\\Delta", "Δ").replace("\\theta", "θ")
    text = text.replace("\\rightarrow", "→").replace("\\to", "→")
    text = text.replace("\\Rightarrow", "⇒")
    text = text.replace("\\frac{1}{2}", "½").replace("\\frac{1}{3}", "⅓")
    text = text.replace("\\frac{1}{4}", "¼").replace("\\frac{3}{4}", "¾")
    text = re.sub(r"\\frac\{([^}]+)\}\{([^}]+)\}", r"(\1)/(\2)", text)
    text = re.sub(r"\$([^$]+)\$", r"\1", text)
    return text


# ============ TẠO FILE WORD ============
def make_doc(title, body):
    body = fix_math_symbols(body)
    d = Document()
    s = d.sections[0]
    s.top_margin = Inches(0.7)
    s.bottom_margin = Inches(0.7)
    s.left_margin = Inches(0.8)
    s.right_margin = Inches(0.8)
    d.styles["Normal"].font.name = "Times New Roman"
    d.styles["Normal"].font.size = Pt(12)
    d.add_heading(title, 0)
    if school.strip():
        d.add_paragraph("Don vi: " + school.strip())
    if teacher.strip():
        d.add_paragraph("Nguoi thuc hien: " + teacher.strip())
    d.add_paragraph("Ngay tao: " + datetime.now().strftime("%d/%m/%Y %H:%M"))
    for line in body.splitlines():
        if line.startswith("### "):
            d.add_heading(line[4:], 3)
        elif line.startswith("## "):
            d.add_heading(line[3:], 2)
        elif line.startswith("# "):
            d.add_heading(line[2:], 1)
        elif line.startswith("- "):
            d.add_paragraph(line[2:], style="List Bullet")
        elif re.match(r"^\d+[\.\)] ", line):
            d.add_paragraph(
                re.sub(r"^\d+[\.\)] ", "", line), style="List Number"
            )
        elif line.strip() == "":
            continue
        else:
            d.add_paragraph(line)
    b = io.BytesIO()
    d.save(b)
    return b.getvalue()


# ============ KHU VỰC CHÍNH ============
st.divider()
st.subheader("🚀 Tạo tài liệu")

col1, col2 = st.columns([1, 2])
with col1:
    if st.button("✨ Tạo bản nháp", type="primary", use_container_width=True):
        if not topic.strip():
            st.warning("⚠️ Nên nhập 'Tên bài/chủ đề'.")
        result = generate()
        if result:
            st.session_state["draft"] = result
            st.session_state["draft_title"] = task
with col2:
    status = []
    if template_text:
        status.append(str(len(template_names)) + " mẫu")
    if source_text:
        status.append(str(len(source_names)) + " nguồn")
    if use_ai and available:
        status.append("AI: " + ", ".join(available))
    if status:
        st.caption(" | ".join(status))

if st.session_state["draft"]:
    st.divider()
    st.subheader("Bản nháp — có thể chỉnh sửa")
    st.session_state["draft"] = st.text_area(
        "Nội dung", st.session_state["draft"], height=420
    )
    name = (
        re.sub(r"[^A-Za-z0-9_\-]", "_", topic.strip())[:45]
        or "tai_lieu_giao_vien"
    )
    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        st.download_button(
            "⬇️ Tải Word (.docx)",
            make_doc(
                st.session_state["draft_title"], st.session_state["draft"]
            ),
            file_name=name + ".docx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            type="primary",
        )
    with c2:
        st.download_button(
            "⬇️ Tải văn bản (.txt)",
            st.session_state["draft"].encode("utf-8"),
            file_name=name + ".txt",
            mime="text/plain",
        )
    with c3:
        if st.button("🗑️ Xóa bản nháp"):
            st.session_state["draft"] = ""
            st.session_state["draft_title"] = ""
            st.rerun()

with st.expander("📖 Hướng dẫn cài đặt vĩnh viễn (1 lần)"):
    st.write("**Bước 1:** Lấy API key miễn phí (chỉ cần 1, càng nhiều càng tốt):")
    st.write("- Gemini: aistudio.google.com/app/apikey")
    st.write("- Groq: console.groq.com/keys")
    st.write("- OpenRouter: openrouter.ai/keys")
    st.write("- Mistral: console.mistral.ai/api-keys")
    st.write("- Cerebras: cloud.cerebras.ai")
    st.write("- NVIDIA: build.nvidia.com")
    st.write("**Bước 2:** Mở app Streamlit → **Manage app** → **Settings** → tab **Secrets**")
    st.write("**Bước 3:** Dán nội dung sau (bỏ dòng nào không có key):")
    st.code(
        'GEMINI_API_KEY = "AIzaSy..."\n'
        'GROQ_API_KEY = "gsk_..."\n'
        'OPENROUTER_API_KEY = "sk-or-..."\n'
        'MISTRAL_API_KEY = "..."\n'
        'CEREBRAS_API_KEY = "..."\n'
        'NVIDIA_API_KEY = "nvapi-..."\n'
        'SCHOOL = "Truong THCS ..."\n'
        'TEACHER = "Nguyen Van A"\n'
        'SUBJECT = "Toan"\n'
        'GRADE = "6"',
        language="toml",
    )
    st.write("**Bước 4:** Bấm **Save** → **Reboot**")
    st.write("**Cách dùng:** freelm tự thử từng nhà cung cấp, không cần chọn thủ công.")
