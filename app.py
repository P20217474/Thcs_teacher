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
st.caption("Ban 3.2 • 4 nha cung cap AI • Tu dong chuyen doi • Xuat Word")


# ============ ĐỌC CẤU HÌNH TỪ SECRETS ============
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

# Đặt vào os.environ để freelm đọc
for _k, _v in PROVIDER_KEYS.items():
    if _v:
        os.environ[_k] = _v

available_providers = [_k for _k, _v in PROVIDER_KEYS.items() if _v]

NAME_MAP = {
    "GEMINI_API_KEY": "Gemini",
    "GROQ_API_KEY": "Groq",
    "OPENROUTER_API_KEY": "OpenRouter",
    "NVIDIA_API_KEY": "NVIDIA",
}

DEFAULT_SCHOOL = cfg("SCHOOL", "")
DEFAULT_TEACHER = cfg("TEACHER", "")
DEFAULT_SUBJECT = cfg("SUBJECT", "Toan")
DEFAULT_GRADE = cfg("GRADE", "6")


# ============ SESSION STATE ============
for _k, _v in {
    "draft": "",
    "draft_title": "",
    "template_text": "",
    "source_text": "",
    "template_names": [],
    "source_names": [],
}.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


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

    available_names = [NAME_MAP[k] for k in available_providers]

    if available_names:
        st.success("✅ " + ", ".join(available_names))
        use_ai = st.checkbox("Dùng AI để sinh nội dung", value=True)
        st.caption("freelm tự chuyển đổi khi 1 nhà cung cấp lỗi.")
    else:
        st.warning("Chưa có API key nào.")
        st.caption("Thêm key trong Secrets để bật AI.")
        use_ai = st.checkbox("Dùng AI để sinh nội dung", value=False)

    # Nhập key tạm thời nếu chưa có trong Secrets
    if not available_names:
        with st.expander("🔑 Nhập API key tạm thời"):
            temp_gem = st.text_input("Gemini API Key", type="password", key="t_gem")
            temp_groq = st.text_input("Groq API Key", type="password", key="t_groq")
            temp_or = st.text_input("OpenRouter API Key", type="password", key="t_or")
            temp_nv = st.text_input("NVIDIA API Key", type="password", key="t_nv")
            for env_key, val in [
                ("GEMINI_API_KEY", temp_gem),
                ("GROQ_API_KEY", temp_groq),
                ("OPENROUTER_API_KEY", temp_or),
                ("NVIDIA_API_KEY", temp_nv),
            ]:
                if val.strip():
                    os.environ[env_key] = val.strip()
                    available_providers.append(env_key)

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
        "Soạn kế hoạch bài dạy",
        "Tạo đề kiểm tra và khung đáp án",
        "Tạo phiếu học tập",
        "Soạn báo cáo/hồ sơ",
        "Lập kịch bản video bài giảng",
        "Soạn nội dung theo yêu cầu riêng",
    ])
    topic = st.text_input("Tên bài/chủ đề")
    details = st.text_area("Yêu cầu cụ thể", height=100)
    period = st.text_input("Thời lượng", placeholder="Ví dụ: 1 tiết / 2 tiết")


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
        names.append("(văn bản dán)")
    if files:
        for f in files:
            nm, text, sz = read_file(f)
            if nm:
                parts.append("--- TÀI LIỆU: " + nm + " ---\n" + text)
                names.append(nm + " (" + str(sz) + " ký tự)")
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
    text = getattr(resp, "text", None) or getattr(resp, "content", "")
    provider = getattr(resp, "provider", "AI")
    return text, provider


def build_prompt():
    t = topic.strip() or "[Chưa có tên]"

    if template_text:
        tpl_block = (
            "=== FILE MẪU (BẮT BUỘC SAO CHÉP NGUYÊN CẤU TRÚC) ===\n"
            + template_text[:12000]
            + "\n=== HẾT MẪU ==="
        )
    else:
        tpl_block = "(Không có mẫu - dùng cấu trúc chuẩn GDPT 2018.)"

    if source_text:
        src_block = (
            "=== NGUỒN TÀI LIỆU (chỉ lấy nội dung từ đây) ===\n"
            + source_text[:20000]
            + "\n=== HẾT NGUỒN ==="
        )
    else:
        src_block = "(Không có nguồn - dùng kiến thức chuẩn môn học.)"

    prompt = "Bạn là trợ lý chuyên môn cho giáo viên THCS Việt Nam.\n\n"
    prompt += "NHIỆM VỤ: " + task + "\n"
    prompt += "MÔN: " + subject + " | KHỐI: " + grade + "\n"
    prompt += "CHỦ ĐỀ: " + t + "\n"
    prompt += "THỜI LƯỢNG: " + (period or "Chưa xác định") + "\n"
    prompt += "YÊU CẦU RIÊNG: " + (details or "Không có") + "\n\n"
    prompt += tpl_block + "\n\n" + src_block + "\n\n"

    prompt += "QUY TẮC BẮT BUỘC (phải tuân thủ tuyệt đối):\n"
    prompt += "1. NẾU CÓ FILE MẪU: phải SAO CHÉP NGUYÊN cấu trúc của mẫu:\n"
    prompt += "   - Giữ ĐÚNG thứ tự các mục (I, II, III... hoặc 1, 2, 3...)\n"
    prompt += "   - Giữ ĐÚNG tên mục, tiêu đề như trong mẫu\n"
    prompt += "   - Giữ ĐÚNG cách đánh số, cách trình bày, vị trí các mục\n"
    prompt += "   - CHỈ THAY phần nội dung bên trong, KHÔNG đổi cấu trúc\n"
    prompt += "   - Nếu mẫu có mục nào thì phải có mục đó, không tự ý thêm/bớt\n"
    prompt += "2. NẾU CÓ NGUỒN: kiến thức, số liệu phải lấy từ nguồn. KHÔNG bịa.\n"
    prompt += "3. NẾU NGUỒN THIẾU: ghi rõ [Cần bổ sung: ...].\n"
    prompt += "4. KÝ HIỆU TOÁN HỌC - BẮT BUỘC dùng Unicode đẹp:\n"
    prompt += "   - Lũy thừa: x², x³, aⁿ (KHÔNG viết x^2)\n"
    prompt += "   - Chỉ số dưới: H₂O, CO₂, a₁, a₂ (KHÔNG viết H_2O)\n"
    prompt += "   - Căn bậc 2: √x, √(a+b) (KHÔNG viết sqrt(x))\n"
    prompt += "   - Phân số: viết dạng 1/2, 3/4 (KHÔNG viết \\frac{1}{2})\n"
    prompt += "   - Nhân chia: × ÷ (KHÔNG viết * /)\n"
    prompt += "   - So sánh: ≤ ≥ ≠ ≈ ≡ (KHÔNG viết <= >= !=)\n"
    prompt += "   - Tập hợp: ∈ ∉ ⊂ ⊃ ∪ ∩ ∅\n"
    prompt += "   - Hy Lạp: α β γ δ π θ Δ Σ Ω\n"
    prompt += "   - Mũi tên: → ⇒ ↔ ⇔\n"
    prompt += "   - Độ: ° (ví dụ 90°, 45°)\n"
    prompt += "   - TUYỆT ĐỐI KHÔNG dùng LaTeX: không dùng $...$, \\frac, \\sqrt, ^, _\n"
    prompt += "5. Viết tiếng Việt CÓ DẤU đầy đủ, văn phong sư phạm, cụ thể, dùng ngay.\n"
    prompt += "6. Markdown: '# ' cấp 1, '## ' cấp 2, '### ' cấp 3, '- ' bullet, '1. ' số.\n"
    prompt += "7. Chỉ xuất nội dung tài liệu, KHÔNG thêm lời chào, KHÔNG giải thích.\n"
    return prompt


def generate_local():
    t = topic.strip() or "[Điền tên bài]"
    common = (
        "- Môn học: " + subject
        + "\n- Lớp: " + grade
        + "\n- Chủ đề: " + t
        + "\n- Thời gian: " + (period or "Chưa xác định")
    )
    return (
        "# " + task.upper() + "\n" + common
        + "\n\n## I. Thông tin chung\n" + common
        + "\n\n## II. Nội dung chính\n"
        + (details or "[Bổ sung yêu cầu cụ thể]")
        + "\n\n## III. Ghi chú\n- Bản mẫu cục bộ."
    )


def generate():
    if use_ai:
        if not available_providers:
            st.error("Chưa có API key nào.")
            return ""
        try:
            with st.spinner("🤖 freelm đang dò tìm nhà cung cấp..."):
                text, provider = call_ai(build_prompt())
            st.success("✅ Đã sinh bằng: " + str(provider))
            return text
        except Exception as e:
            st.error("Lỗi freelm: " + str(e))
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
        d.add_paragraph("Đơn vị: " + school.strip())
    if teacher.strip():
        d.add_paragraph("Người thực hiện: " + teacher.strip())
    d.add_paragraph("Ngày tạo: " + datetime.now().strftime("%d/%m/%Y %H:%M"))
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
    if use_ai and available_names:
        status.append("AI: " + ", ".join(available_names))
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
    st.write("**Bước 1:** Lấy API key (đã có sẵn 4 key):")
    st.write("- Gemini: aistudio.google.com/app/apikey")
    st.write("- Groq: console.groq.com/keys")
    st.write("- OpenRouter: openrouter.ai/keys")
    st.write("- NVIDIA: build.nvidia.com")
    st.write("**Bước 2:** Mở app Streamlit → **Manage app** → **Settings** → tab **Secrets**")
    st.write("**Bước 3:** Dán nội dung sau (thay giá trị thật):")
    st.code(
        'GEMINI_API_KEY = "AIzaSy..."\n'
        'GROQ_API_KEY = "gsk_..."\n'
        'OPENROUTER_API_KEY = "sk-or-..."\n'
        'NVIDIA_API_KEY = "nvapi-..."\n'
        'SCHOOL = "Truong THCS ..."\n'
        'TEACHER = "Nguyen Van A"\n'
        'SUBJECT = "Toan"\n'
        'GRADE = "6"',
        language="toml",
    )
    st.write("**Bước 4:** Bấm **Save** → **Reboot**")
    st.write("**Cách dùng:** freelm tự thử từng nhà cung cấp, không cần chọn thủ công.")
