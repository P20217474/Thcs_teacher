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
_ensure("google-generativeai", "google.generativeai")
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
st.caption("Ban 2.2 • MAU + NGUON • AI tu dong • Xuat Word")


# ============ ĐỌC CẤU HÌNH TỪ SECRETS ============
def cfg(key, default=""):
    """Đọc giá trị từ Secrets, nếu không có thì trả về default."""
    try:
        val = st.secrets.get(key, default)
        return val if val else default
    except Exception:
        return default


DEFAULT_API_KEY = cfg("GEMINI_API_KEY", "")
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
    if DEFAULT_API_KEY:
        st.success("✅ Đã có API Key từ Secrets")
        api_key = DEFAULT_API_KEY
        use_ai = st.checkbox("Dùng AI để sinh nội dung", value=True)
        with st.expander("🔄 Đổi API Key khác"):
            override = st.text_input(
                "Nhập key mới (bỏ trống để dùng key mặc định)",
                type="password",
                key="override_key",
            )
            if override.strip():
                api_key = override.strip()
    else:
        api_key = st.text_input(
            "Google Gemini API Key", type="password", key="manual_key"
        )
        use_ai = st.checkbox("Dùng AI để sinh nội dung", value=bool(api_key))
        st.caption("💡 Có thể cài key vĩnh viễn trong Settings > Secrets.")

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
    else:
        st.info("Chưa nạp mẫu.")
    st.divider()
    st.subheader("📚 NGUỒN đã nạp")
    if source_names:
        for n in source_names:
            st.write("- " + n)
        st.caption("Tổng: " + str(len(source_text)) + " ký tự")
    else:
        st.info("Chưa nạp nguồn.")


# ============ GỌI AI ============
def call_gemini(prompt, key):
    import google.generativeai as genai
    genai.configure(api_key=key)
    last_err = None
    for model_name in [
        "gemini-3.8-flash",
        "gemini-flash-latest",
        "gemini-2.5-flash",
        "gemini-2.0-flash",
        "gemini-1.5-flash-latest",
        "gemini-1.5-flash",
    ]:
        try:
            model = genai.GenerativeModel(model_name)
            resp = model.generate_content(prompt)
            return resp.text
        except Exception as e:
            last_err = e
            continue
    raise last_err


def build_prompt():
    t = topic.strip() or "[Chua co ten]"
    if template_text:
        tpl_block = (
            "=== FILE MAU (BAT BUOC bam theo) ===\n"
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
    prompt += "QUY TAC BAT BUOC:\n"
    prompt += "1. Neu co FILE MAU: giu dung cau truc, thu tu muc, cach danh so.\n"
    prompt += "2. Neu co NGUON: kien thuc, so lieu phai lay tu nguon. KHONG bia.\n"
    prompt += "3. Neu nguon thieu: ghi ro [Can bo sung: ...].\n"
    prompt += "4. Viet tieng Viet, van phong su pham, cu the, dung duoc ngay.\n"
    prompt += "5. Markdown: '# ' cap 1, '## ' cap 2, '### ' cap 3, '- ' bullet, '1. ' so.\n"
    prompt += "6. Chi xuat noi dung tai lieu, khong them loi chao.\n"
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
        if not api_key or not api_key.strip():
            st.error("Chua co API Key.")
            return ""
        try:
            with st.spinner("🤖 AI đang sinh nội dung..."):
                return call_gemini(build_prompt(), api_key.strip())
        except Exception as e:
            st.error("Loi AI: " + str(e))
            st.info("Chuyen sang mau cuc bo.")
            return generate_local()
    return generate_local()


# ============ TẠO FILE WORD ============
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
    if use_ai and api_key:
        status.append("AI bật")
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
    st.write("**Bước 1:** Lấy API Key miễn phí tại aistudio.google.com/app/apikey")
    st.write("**Bước 2:** Mở app Streamlit → **Manage app** → **Settings** → tab **Secrets**")
    st.write("**Bước 3:** Dán nội dung sau vào ô Secrets (thay giá trị thật):")
    st.code(
        'GEMINI_API_KEY = "AIzaSy...key_cua_ban..."\n'
        'SCHOOL = "Truong THCS ..."\n'
        'TEACHER = "Nguyen Van A"\n'
        'SUBJECT = "Toan"\n'
        'GRADE = "6"',
        language="toml",
    )
    st.write("**Bước 4:** Bấm **Save** → **Reboot**")
    st.write("Từ đó mở app là có sẵn tất cả thông tin, không cần nhập lại.")
