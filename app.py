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
    page_title="Agent Tro ly Giao vien THCS AI",
    page_icon="🎓",
    layout="wide",
)

st.title("🎓 Agent Trợ lý Giáo viên THCS AI")
st.caption("Bản 4.0 • Quét đa định dạng • Bảng Word • 4 nhà cung cấp AI")


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

PREFERRED_MODELS = {
    "GEMINI_API_KEY": [
        "gemini-2.5-flash",
        "gemini-2.0-flash",
        "gemini-flash-latest",
        "gemini-1.5-flash",
    ],
    "GROQ_API_KEY": [
        "llama-3.3-70b-versatile",
        "llama-3.1-70b-versatile",
        "mixtral-8x7b-32768",
        "llama-3.1-8b-instant",
    ],
    "OPENROUTER_API_KEY": [
        "google/gemini-2.0-flash-exp:free",
        "meta-llama/llama-3.3-70b-instruct:free",
        "deepseek/deepseek-chat:free",
        "mistralai/mistral-7b-instruct:free",
    ],
    "NVIDIA_API_KEY": [
        "meta/llama-3.3-70b-instruct",
        "meta/llama-3.1-70b-instruct",
        "mistralai/mixtral-8x7b-instruct-v0.1",
    ],
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


# ============ HÀM ĐỌC FILE ĐA ĐỊNH DẠNG ============
def read_docx(data: bytes) -> str:
    try:
        d = Document(io.BytesIO(data))
        out = []
        for p in d.paragraphs:
            if p.text.strip():
                out.append(p.text)
        # Đọc cả bảng trong docx
        for table in d.tables:
            out.append("")
            for row in table.rows:
                cells = [c.text.strip() for c in row.cells]
                out.append("| " + " | ".join(cells) + " |")
            out.append("")
        return "\n".join(out)
    except Exception as e:
        return "[Lỗi docx: " + str(e) + "]"


def read_pdf(data: bytes) -> str:
    if not HAS_PYPDF:
        return "[Chưa cài pypdf]"
    try:
        reader = PdfReader(io.BytesIO(data))
        texts = []
        for page in reader.pages:
            t = page.extract_text() or ""
            if t.strip():
                texts.append(t)
        full = "\n".join(texts)
        if len(full.strip()) < 100:
            return (
                "[PDF này có vẻ là ảnh scan, không có text. "
                "Hãy dùng Google Drive OCR hoặc dán trực tiếp nội dung.]\n"
                + full
            )
        return full
    except Exception as e:
        return "[Lỗi pdf: " + str(e) + "]"


def read_csv_text(data: bytes) -> str:
    try:
        text = data.decode("utf-8", errors="ignore")
        reader = csv.reader(io.StringIO(text))
        rows = list(reader)
        if not rows:
            return "[File CSV rỗng]"
        out = []
        header = rows[0]
        out.append("| " + " | ".join(header) + " |")
        out.append("|" + "---|" * len(header))
        for row in rows[1:]:
            row = (row + [""] * len(header))[:len(header)]
            out.append("| " + " | ".join(row) + " |")
        return "\n".join(out)
    except Exception as e:
        return "[Lỗi csv: " + str(e) + "]"


def read_gif_info(data: bytes) -> str:
    """Đọc metadata của GIF. Không trích xuất được text, chỉ báo info."""
    if not HAS_PIL:
        return "[GIF: chưa cài PIL để đọc. Hãy chuyển sang .txt/.docx]"
    try:
        img = Image.open(io.BytesIO(data))
        info = "Ảnh GIF: " + str(img.size[0]) + "x" + str(img.size[1]) + " px"
        try:
            n_frames = getattr(img, "n_frames", 1)
            info += ", " + str(n_frames) + " frame"
        except Exception:
            pass
        info += ". GIF thường là ảnh động, không có text. "
        info += "Nếu muốn AI đọc nội dung, hãy chuyển sang PDF có text hoặc dán text."
        return "[" + info + "]"
    except Exception as e:
        return "[Lỗi gif: " + str(e) + "]"


def read_image_info(data: bytes, fname: str) -> str:
    """Ảnh thường: báo info, khuyến nghị OCR."""
    if not HAS_PIL:
        return "[" + fname + ": ảnh, cần OCR để lấy text]"
    try:
        img = Image.open(io.BytesIO(data))
        return (
            "[Ảnh " + fname + ": " + str(img.size[0]) + "x" + str(img.size[1])
            + " px. Ảnh không chứa text trực tiếp. "
            + "Hãy dùng Google Drive OCR hoặc dán text thủ công.]"
        )
    except Exception as e:
        return "[Lỗi ảnh " + fname + ": " + str(e) + "]"


def read_file(file):
    if file is None:
        return ("", "", 0)
    data = file.read()
    if not data:
        return (file.name, "[File rỗng]", 0)
    name = file.name.lower()
    if name.endswith(".docx"):
        text = read_docx(data)
    elif name.endswith(".pdf"):
        text = read_pdf(data)
    elif name.endswith(".csv"):
        text = read_csv_text(data)
    elif name.endswith((".txt", ".md", ".log")):
        text = data.decode("utf-8", errors="ignore")
    elif name.endswith(".gif"):
        text = read_gif_info(data)
    elif name.endswith((".png", ".jpg", ".jpeg", ".webp", ".bmp")):
        text = read_image_info(data, file.name)
    else:
        text = "[Định dạng " + name + " chưa hỗ trợ. Dùng .docx/.pdf/.txt/.csv]"
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
        "📎 Nạp file MẪU (docx / pdf / txt / csv / ảnh / gif)",
        type=["docx", "pdf", "txt", "md", "csv", "png", "jpg", "jpeg", "gif", "webp"],
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
        "📎 Nạp file NGUỒN (docx / pdf / txt / csv / ảnh / gif)",
        type=["docx", "pdf", "txt", "md", "csv", "png", "jpg", "jpeg", "gif", "webp"],
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
        with st.expander("Xem nội dung mẫu (2000 ký tự đầu)"):
            st.text(template_text[:2000])
    else:
        st.info("Chưa nạp mẫu.")
    st.divider()
    st.subheader("📚 NGUỒN đã nạp")
    if source_names:
        for n in source_names:
            st.write("- " + n)
        st.caption("Tổng: " + str(len(source_text)) + " ký tự")
        with st.expander("Xem nội dung nguồn (2000 ký tự đầu)"):
            st.text(source_text[:2000])
    else:
        st.info("Chưa nạp nguồn.")


# ============ GỌI AI QUA FREELM ============
def call_ai(prompt):
    import freelm

    llm = freelm.FreeLLM.from_env()

    last_err = None
    for provider_key in available_providers:
        models = PREFERRED_MODELS.get(provider_key, [])
        for model_name in models:
            try:
                resp = llm.chat(
                    [{"role": "user", "content": prompt}],
                    model=model_name,
                )
                text = getattr(resp, "text", None) or getattr(resp, "content", "")
                if text and len(text.strip()) > 100:
                    return text, model_name
            except Exception as e:
                last_err = e
                continue

    try:
        resp = llm.chat(
            [{"role": "user", "content": prompt}],
            model="auto",
        )
        text = getattr(resp, "text", None) or getattr(resp, "content", "")
        return text, "auto"
    except Exception as e:
        if last_err:
            raise last_err
        raise e


def build_prompt():
    t = topic.strip() or "[Chưa có tên]"

    if template_text:
        tpl_block = (
            "=== FILE MẪU (BẮT BUỘC SAO CHÉP NGUYÊN CẤU TRÚC) ===\n"
            + template_text[:15000]
            + "\n=== HẾT MẪU ==="
        )
    else:
        tpl_block = "(Không có mẫu - dùng cấu trúc chuẩn GDPT 2018.)"

    if source_text:
        src_block = (
            "=== NGUỒN TÀI LIỆU (chỉ lấy nội dung từ đây) ===\n"
            + source_text[:25000]
            + "\n=== HẾT NGUỒN ==="
        )
    else:
        src_block = "(Không có nguồn - dùng kiến thức chuẩn môn học.)"

    prompt = "Bạn là trợ lý chuyên môn cao cấp cho giáo viên THCS Việt Nam.\n"
    prompt += "Bạn có nhiệm vụ soạn tài liệu CHI TIẾT, ĐẦY ĐỦ, DÙNG ĐƯỢC NGAY.\n\n"
    prompt += "NHIỆM VỤ: " + task + "\n"
    prompt += "MÔN: " + subject + " | KHỐI: " + grade + "\n"
    prompt += "CHỦ ĐỀ: " + t + "\n"
    prompt += "THỜI LƯỢNG: " + (period or "Chưa xác định") + "\n"
    prompt += "YÊU CẦU RIÊNG: " + (details or "Không có") + "\n\n"
    prompt += tpl_block + "\n\n" + src_block + "\n\n"

    prompt += "QUY TẮC BẮT BUỘC (tuân thủ tuyệt đối):\n"
    prompt += "\n[A] VỀ CẤU TRÚC:\n"
    prompt += "1. NẾU CÓ FILE MẪU: phải SAO CHÉP NGUYÊN cấu trúc của mẫu:\n"
    prompt += "   - Giữ ĐÚNG thứ tự các mục (I, II, III... hoặc 1, 2, 3...)\n"
    prompt += "   - Giữ ĐÚNG tên mục, tiêu đề, cách đánh số\n"
    prompt += "   - Giữ ĐÚNG vị trí các mục con, không tự thêm/bớt\n"
    prompt += "   - Nếu mẫu có bảng, PHẢI tạo bảng markdown tương tự\n"
    prompt += "   - CHỈ THAY nội dung bên trong, KHÔNG đổi cấu trúc\n"
    prompt += "2. NẾU KHÔNG CÓ MẪU: dùng cấu trúc chuẩn của loại tài liệu đó.\n"

    prompt += "\n[B] VỀ NỘI DUNG (RẤT QUAN TRỌNG):\n"
    prompt += "3. Viết CHI TIẾT, CỤ THỂ, ĐẦY ĐỦ. KHÔNG viết chung chung, sáo rỗng.\n"
    prompt += "4. Mỗi mục phải có nội dung thực chất:\n"
    prompt += "   - Mục 'Mục tiêu': nêu rõ kiến thức, kỹ năng, thái độ cụ thể\n"
    prompt += "   - Mục 'Hoạt động': mô tả từng bước: GV làm gì, HS làm gì, sản phẩm\n"
    prompt += "   - Mục 'Câu hỏi/bài tập': có câu hỏi cụ thể, đáp án mẫu\n"
    prompt += "5. Với mỗi hoạt động dạy học, ghi rõ:\n"
    prompt += "   - Mục tiêu hoạt động\n"
    prompt += "   - Nội dung/cách thức tổ chức\n"
    prompt += "   - Sản phẩm học tập dự kiến\n"
    prompt += "6. Độ dài tối thiểu: 1500 từ cho kế hoạch bài dạy, 1000 từ cho các loại khác.\n"
    prompt += "7. NẾU CÓ NGUỒN: kiến thức, số liệu phải lấy từ nguồn. KHÔNG bịa.\n"
    prompt += "8. NẾU NGUỒN THIẾU: ghi rõ [Cần bổ sung: ...].\n"

    prompt += "\n[C] VỀ BẢNG BIỂU (QUAN TRỌNG):\n"
    prompt += "9. Nếu cần trình bày dạng bảng, dùng markdown table:\n"
    prompt += "   | Cột 1 | Cột 2 |\n"
    prompt += "   |---|---|\n"
    prompt += "   | Ô 1 | Ô 2 |\n"
    prompt += "   KHÔNG dùng HTML table.\n"
    prompt += "10. Trong ô bảng, dùng <br> để xuống dòng.\n"
    prompt += "11. Dùng **bold** để làm đậm tiêu đề trong ô bảng.\n"

    prompt += "\n[D] VỀ KÝ HIỆU TOÁN HỌC:\n"
    prompt += "12. Dùng Unicode đẹp, KHÔNG dùng LaTeX:\n"
    prompt += "   - Lũy thừa: x², x³, aⁿ (KHÔNG viết x^2)\n"
    prompt += "   - Chỉ số dưới: H₂O, CO₂, a₁, a₂ (KHÔNG viết H_2O)\n"
    prompt += "   - Căn bậc 2: √x (KHÔNG viết sqrt(x))\n"
    prompt += "   - Phân số: 1/2, 3/4 (KHÔNG viết \\frac{1}{2})\n"
    prompt += "   - Nhân chia: × ÷ ; So sánh: ≤ ≥ ≠ ≈ ; Tập hợp: ∈ ∉ ⊂ ∪ ∩ ∅\n"
    prompt += "   - Hy Lạp: α β γ π θ Δ Σ Ω ; Mũi tên: → ⇒ ↔ ; Độ: °\n"

    prompt += "\n[E] VỀ ĐỊNH DẠNG:\n"
    prompt += "13. Tiếng Việt có dấu đầy đủ, văn phong sư phạm trang trọng.\n"
    prompt += "14. Markdown: '# ' cấp 1, '## ' cấp 2, '### ' cấp 3, '- ' bullet, '1. ' số.\n"
    prompt += "15. Chỉ xuất nội dung tài liệu. KHÔNG lời chào, KHÔNG giải thích.\n"

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
            with st.spinner("🤖 Đang thử model mạnh nhất..."):
                text, model_used = call_ai(build_prompt())
            st.success("✅ Đã sinh bằng model: " + str(model_used))
            return text
        except Exception as e:
            st.error("Lỗi tất cả nhà cung cấp: " + str(e))
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


# ============ TẠO FILE WORD (CÓ BẢNG) ============
def add_formatted_paragraph(doc, text, style=None):
    """Thêm đoạn văn có xử lý **bold** và <br>."""
    text = text.replace("<br>", "\n").replace("<br/>", "\n").replace("<br />", "\n")
    if style:
        p = doc.add_paragraph(style=style)
    else:
        p = doc.add_paragraph()
    parts = re.split(r"(\*\*[^*]+\*\*)", text)
    for part in parts:
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            run = p.add_run(part[2:-2])
            run.bold = True
        else:
            p.add_run(part)
    return p


def add_markdown_table(doc, rows):
    """Chuyển list các dòng markdown thành bảng Word."""
    if not rows:
        return
    n_cols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=n_cols)
    table.style = "Table Grid"
    for i, row in enumerate(rows):
        for j in range(n_cols):
            cell = table.cell(i, j)
            cell.text = ""
            p = cell.paragraphs[0]
            val = row[j] if j < len(row) else ""
            val = val.replace("<br>", "\n").replace("<br/>", "\n")
            parts = re.split(r"(\*\*[^*]+\*\*)", val)
            for part in parts:
                if part.startswith("**") and part.endswith("**") and len(part) > 4:
                    run = p.add_run(part[2:-2])
                    run.bold = True
                else:
                    p.add_run(part)
    return table


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

    lines = body.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # Bảng markdown
        if stripped.startswith("|") and stripped.endswith("|"):
            table_lines = []
            j = i
            while j < len(lines) and lines[j].strip().startswith("|") and lines[j].strip().endswith("|"):
                table_lines.append(lines[j].strip())
                j += 1
            rows = []
            for tl in table_lines:
                if re.match(r"^\|[\s\-:]+\|", tl):
                    continue
                cells = [c.strip() for c in tl.strip("|").split("|")]
                rows.append(cells)
            if rows:
                add_markdown_table(d, rows)
            i = j
            continue

        if stripped.startswith("### "):
            d.add_heading(stripped[4:], 3)
        elif stripped.startswith("## "):
            d.add_heading(stripped[3:], 2)
        elif stripped.startswith("# "):
            d.add_heading(stripped[2:], 1)
        elif stripped.startswith("- "):
            add_formatted_paragraph(d, stripped[2:], style="List Bullet")
        elif re.match(r"^\d+[\.\)] ", stripped):
            add_formatted_paragraph(
                d, re.sub(r"^\d+[\.\)] ", "", stripped), style="List Number"
            )
        elif stripped == "":
            pass
        else:
            add_formatted_paragraph(d, stripped)
        i += 1

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

with st.expander("📖 Hướng dẫn sử dụng"):
    st.markdown("""
### 📐 File MẪU
- Quy định **cấu trúc, thứ tự mục, cách trình bày**
- Hỗ trợ: `.docx`, `.pdf`, `.txt`, `.csv`, ảnh, `.gif`
- AI sẽ bám sát cấu trúc này

### 📚 File NGUỒN
- Chứa **kiến thức, số liệu, dữ kiện**
- Hỗ trợ: `.docx`, `.pdf`, `.txt`, `.csv`, ảnh, `.gif`
- AI chỉ lấy nội dung từ nguồn

### ⚠️ Lưu ý về PDF scan
- PDF scan (ảnh) chỉ đọc được vài ký tự metadata
- Cần PDF có text thật (xuất từ Word) hoặc dán trực tiếp
- Có thể dùng Google Drive OCR để chuyển scan thành text

### 🖼️ Ảnh và GIF
- Ảnh/GIF không chứa text trực tiếp
- App sẽ báo info (kích thước, số frame)
- Cần OCR hoặc dán text thủ công

### 📊 Bảng biểu
- AI dùng markdown table: `| Cột 1 | Cột 2 |`
- Word sẽ render thành bảng có viền thật
- Trong ô có thể dùng `<br>` để xuống dòng
""")
