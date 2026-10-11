"""
AGENT TRỢ LÝ GIÁO VIÊN THCS AI — BẢN 2.0 (2 TÙY CHỌN RÕ RÀNG)
- Tùy chọn MẪU: định dạng trình bày
- Tùy chọn NGUỒN: nội dung tài liệu
- AI bám mẫu + lấy nội dung từ nguồn
"""
import subprocess
import sys

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
import io, re

try:
    from pypdf import PdfReader
    HAS_PYPDF = True
except Exception:
    HAS_PYPDF = False

st.set_page_config(
    page_title="Agent Trợ lý Giáo viên THCS AI",
    page_icon="🎓",
    layout="wide",
)

st.title("🎓 Agent Trợ lý Giáo viên THCS AI")
st.caption("Bản 2.0 • 2 tùy chọn: MẪU (định dạng) + NGUỒN (nội dung) • AI • Xuất Word")

# SESSION STATE
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
        return f"[Lỗi docx: {e}]"

def read_pdf(data: bytes) -> str:
    if not HAS_PYPDF:
        return "[Chưa cài pypdf]"
    try:
        reader = PdfReader(io.BytesIO(data))
        return "\n".join((pg.extract_text() or "") for pg in reader.pages)
    except Exception as e:
        return f"[Lỗi pdf: {e}]"

def read_file(file) -> tuple:
    """Trả về (tên, nội dung, kích thước)"""
    if file is None:
        return ("", "", 0)
    data = file.read()
    if not data:
        return (file.name, f"[File {file.name} rỗng — kiểm tra lại file gốc]", 0)
    name = file.name.lower()
    if name.endswith(".docx"):
        text = read_docx(data)
    elif name.endswith(".pdf"):
        text = read_pdf(data)
    elif name.endswith((".txt", ".md")):
        text = data.decode("utf-8", errors="ignore")
    else:
        text = f"[Không hỗ trợ định dạng {file.name}]"
    return (file.name, text, len(text))


# ============ SIDEBAR: CẤU HÌNH + THÔNG TIN ============
with st.sidebar:
    st.header("⚙️ Cấu hình AI")
    api_key = st.text_input(
        "Google Gemini API Key",
        type="password",
        help="Lấy miễn phí tại aistudio.google.com/app/apikey",
    )
    use_ai = st.checkbox("Dùng AI để sinh nội dung", value=bool(api_key))

    st.divider()
    st.header("📋 Thông tin")
    school = st.text_input("Trường/đơn vị")
    teacher = st.text_input("Giáo viên")
    subject = st.selectbox("Môn học", [
        "Toán", "Ngữ văn", "Tiếng Anh", "Khoa học tự nhiên",
        "Vật lí", "Hóa học", "Sinh học", "Lịch sử và Địa lí",
        "Tin học", "Công nghệ", "GDCD/GDKT&PL", "Nghệ thuật",
        "Giáo dục thể chất", "Hoạt động trải nghiệm", "Khác",
    ])
    grade = st.selectbox("Khối lớp", ["6", "7", "8", "9", "Nhiều khối"])

    st.divider()
    st.subheader("Công việc")
    task = st.selectbox("Chọn công việc", [
        "Soạn kế hoạch bài dạy",
        "Tạo đề kiểm tra và khung đáp án",
        "Tạo phiếu học tập",
        "Soạn báo cáo/hồ sơ",
        "Lập kịch bản video bài giảng",
        "Soạn nội dung theo yêu cầu riêng",
    ])
    topic = st.text_input("Tên bài/chủ đề/công việc")
    details = st.text_area("Yêu cầu cụ thể", height=100)
    period = st.text_input("Thời lượng", placeholder="Ví dụ: 1 tiết / 2 tiết")


# ============ 2 TAB RÕ RÀNG: MẪU + NGUỒN ============
tab_mau, tab_nguon, tab_xem = st.tabs([
    "📐 TÙY CHỌN MẪU (định dạng trình bày)",
    "📚 TÙY CHỌN NGUỒN (nội dung tài liệu)",
    "👁️ XEM TRƯỚC ĐÃ NẠP",
])

# ---------- TAB MẪU ----------
with tab_mau:
    st.markdown("""
**MẪU dùng để:** quy định **cấu trúc, thứ tự mục, cách trình bày** của tài liệu.
Ví dụ: giáo án mẫu của trường, đề kiểm tra mẫu, phiếu học tập mẫu...
AI sẽ **bám sát cấu trúc này** khi soạn.
""")
    tpl_files = st.file_uploader(
        "📎 Nạp 1 hoặc nhiều file MẪU (docx / pdf / txt)",
        type=["docx", "pdf", "txt"],
        accept_multiple_files=True,
        key="tpl_uploader",
    )
    tpl_paste = st.text_area(
        "📝 Hoặc dán mẫu trực tiếp vào đây",
        height=140,
        placeholder="Dán nội dung mẫu (giáo án, đề, phiếu...) vào đây...",
        key="tpl_paste",
    )
    if st.button("🗑️ Xóa mẫu đã nạp", key="clear_tpl"):
        st.session_state["template_text"] = ""
        st.session_state["template_names"] = []
        st.rerun()

# ---------- TAB NGUỒN ----------
with tab_nguon:
    st.markdown("""
**NGUỒN dùng để:** chứa **kiến thức, số liệu, dữ kiện** để AI lấy nội dung.
Ví dụ: bài học SGK, tài liệu tham khảo, số liệu thống kê, văn bản gốc...
AI **chỉ được lấy nội dung từ nguồn**, không bịa thêm.
""")
    src_files = st.file_uploader(
        "📎 Nạp 1 hoặc nhiều file NGUỒN (docx / pdf / txt)",
        type=["docx", "pdf", "txt"],
        accept_multiple_files=True,
        key="src_uploader",
    )
    src_paste = st.text_area(
        "📝 Hoặc dán nội dung nguồn trực tiếp vào đây",
        height=140,
        placeholder="Dán kiến thức, số liệu, dữ kiện vào đây...",
        key="src_paste",
    )
    if st.button("🗑️ Xóa nguồn đã nạp", key="clear_src"):
        st.session_state["source_text"] = ""
        st.session_state["source_names"] = []
        st.rerun()

# ---------- XỬ LÝ NỘI DUNG ĐÃ NẠP ----------
# Ưu tiên dùng nội dung mới nạp (file + paste), nếu trống thì dùng session_state (đã nạp trước)
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
                parts.append(f"--- TÀI LIỆU: {nm} ---\n{text}")
                names.append(f"{nm} ({sz} ký tự)")
    if parts:
        st.session_state[store_key] = "\n\n".join(parts)
        st.session_state[names_key] = names
    return st.session_state[store_key], st.session_state[names_key]

template_text, template_names = collect_text(tpl_files, tpl_paste, "template_text", "template_names")
source_text, source_names = collect_text(src_files, src_paste, "source_text", "source_names")

# ---------- TAB XEM TRƯỚC ----------
with tab_xem:
    st.subheader("📐 MẪU đã nạp")
    if template_names:
        for n in template_names:
            st.write(f"✅ {n}")
        st.caption(f"Tổng: {len(template_text):,} ký tự")
        with st.expander("Xem nội dung mẫu"):
            st.text(template_text[:3000] + ("..." if len(template_text) > 3000 else ""))
    else:
        st.info("Chưa nạp mẫu.")

    st.divider()
    st.subheader("📚 NGUỒN đã nạp")
    if source_names:
        for n in source_names:
            st.write(f"✅ {n}")
        st.caption(f"Tổng: {len(source_text):,} ký tự")
        with st.expander("Xem nội dung nguồn"):
            st.text(source_text[:3000] + ("..." if len(source_text) > 3000 else ""))
    else:
        st.info("Chưa nạp nguồn.")


# ============ GỌI AI ============
def call_gemini(prompt: str, key: str) -> str:
    import google.generativeai as genai
    genai.configure(api_key=key)
    last_err = None
    for model_name in [
        "gemini-3.8-flash",
        "gemini-flash-latest",
        "gemini-2.5-flash",
        "gemini-2.0-flash",
        "gemini-1.5-flash-latest",
    ]:
        try:
            model = genai.GenerativeModel(model_name)
            resp = model.generate_content(prompt)
            return resp.text
        except Exception as e:
            last_err = e
    raise last_err


def build_prompt() -> str:
    t = topic.strip() or "[Chưa có tên]"

    tpl_block = f"""
╔══════════════════════════════════════════╗
║  FILE MẪU (BẮT BUỘC bám theo cấu trúc)  ║
╚══════════════════════════════════════════╝
{template_text[:12000]}
═════════════ HẾT MẪU ═════════════
""" if template_text else """
(Không có file mẫu — dùng cấu trúc chuẩn GDPT 2018.)
"""

    src_block = f"""
╔══════════════════════════════════════════╗
║  NGUỒN TÀI LIỆU (chỉ lấy nội dung ở đây) ║
╚══════════════════════════════════════════╝
{source_text[:20000]}
═════════════ HẾT NGUỒN ═════════════
""" if source_text else """
(Không có nguồn — dùng kiến thức chuẩn môn học.)
"""

    return f"""Bạn là trợ lý chuyên môn cho giáo viên THCS Việt Nam.

NHIỆM VỤ: {task}
MÔN: {subject} | KHỐI: {grade}
CHỦ ĐỀ: {t}
THỜI LƯỢNG: {period or 'Chưa xác định'}
YÊU CẦU RIÊNG: {details or 'Không có'}

{tpl_block}

{src_block}

QUY TẮC BẮT BUỘC:
1. Nếu có FILE MẪU: giữ đúng cấu trúc, thứ tự mục, cách đánh số, cách trình bày của mẫu. Không tự thêm/bớt mục.
2. Nếu có NGUỒN TÀI LIỆU: kiến thức, số liệu phải lấy từ nguồn. KHÔNG bịa.
3. Nếu nguồn thiếu: ghi rõ [Cần bổ sung: ...] thay vì tự bịa.
4. Viết tiếng Việt, văn phong sư phạm, cụ thể, dùng được ngay.
5. Markdown: '# ' cấp 1, '## ' cấp 2, '### ' cấp 3, '- ' bullet, '1. ' số.
6. Chỉ xuất nội dung tài liệu, không thêm lời chào.
"""


# ============ MẪU CỤC BỘ ============
def generate_local() -> str:
    t = topic.strip() or "[Điền tên bài/chủ đề]"
    common = (f"- Môn học: {subject}\n- Lớp: {grade}\n- Chủ đề: {t}\n"
              f"- Thời gian: {period or 'Chưa xác định'}")
    return f"""# {task.upper()}
{common}

## I. Thông tin chung
{common}

## II. Nội dung chính
{details or '[Bổ sung yêu cầu cụ thể]'}

## III. Ghi chú
- Đây là bản mẫu cục bộ. Nhập API key để AI sinh nội dung đầy đủ."""


def generate() -> str:
    if use_ai:
        if not api_key.strip():
            st.error("Chưa nhập API Key.")
            return ""
        try:
            with st.spinner("🤖 AI đang sinh nội dung..."):
                return call_gemini(build_prompt(), api_key.strip())
        except Exception as e:
            st.error(f"Lỗi AI: {e}")
            st.info("Chuyển sang mẫu cục bộ.")
            return generate_local()
    return generate_local()


# ============ WORD ============
def make_doc(title: str, body: str) -> bytes:
    d = Document()
    s = d.sections[0]
    s.top_margin = Inches(0.7); s.bottom_margin = Inches(0.7)
    s.left_margin = Inches(0.8); s.right_margin = Inches(0.8)
    d.styles["Normal"].font.name = "Times New Roman"
    d.styles["Normal"].font.size = Pt(12)

    d.add_heading(title, 0)
    if school.strip(): d.add_paragraph("Đơn vị: " + school.strip())
    if teacher.strip(): d.add_paragraph("Người thực hiện: " + teacher.strip())
    d.add_paragraph("Ngày tạo: " + datetime.now().strftime("%d/%m/%Y %H:%M"))

    for line in body.splitlines():
        if line.startswith("### "): d.add_heading(line[4:], 3)
        elif line.startswith("## "): d.add_heading(line[3:], 2)
        elif line.startswith("# "): d.add_heading(line[2:], 1)
        elif line.startswith("- "): d.add_paragraph(line[2:], style="List Bullet")
        elif re.match(r"^\d+[\.\)] ", line):
            d.add_paragraph(re.sub(r"^\d+[\.\)] ", "", line), style="List Number")
        elif line.strip() == "": continue
        else: d.add_paragraph(line)

    b = io.BytesIO(); d.save(b); return b.getvalue()


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
    if template_text: status.append(f"📐 {len(template_names)} mẫu")
    if source_text: status.append(f"📚 {len(source_names)} nguồn")
    if use_ai and api_key: status.append("🤖 AI bật")
    if status:
        st.caption(" | ".join(status))
    else:
        st.caption("Chưa nạp mẫu/nguồn, chưa bật AI.")

if st.session_state["draft"]:
    st.divider()
    st.subheader("Bản nháp — có thể chỉnh sửa")
    st.session_state["draft"] = st.text_area(
        "Nội dung", st.session_state["draft"], height=420
    )
    name = re.sub(r"[^A-Za-z0-9_\-]", "_", topic.strip())[:45] or "tai_lieu_giao_vien"

    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        st.download_button(
            "⬇️ Tải Word (.docx)",
            make_doc(st.session_state["draft_title"], st.session_state["draft"]),
            file_name=name + ".docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
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

with st.expander("📖 Hướng dẫn sử dụng 2 tùy chọn"):
    st.markdown("""
### 📐 TÙY CHỌN MẪU (định dạng)
- Dùng khi bạn có **file mẫu** quy định cấu trúc (giáo án mẫu, đề mẫu...)
- AI sẽ **giữ đúng thứ tự mục, cách trình bày** của mẫu
- Nạp **nhiều file** cùng lúc để gộp nhiều mẫu

### 📚 TÙY CHỌN NGUỒN (nội dung)
- Dùng khi bạn có **tài liệu chứa nội dung** (SGK, bài học, số liệu...)
- AI **chỉ lấy nội dung từ nguồn**, không bịa
- Nếu nguồn thiếu, AI ghi `[Cần bổ sung: ...]`

### 🎯 Kết hợp cả hai
- MẪU → định dạng trình bày
- NGUỒN → nội dung kiến thức
- AI → sinh tài liệu bám mẫu + lấy nội dung từ nguồn

### 🔑 Lấy API Key Gemini miễn phí
1. Vào `aistudio.google.com/app/apikey`
2. Đăng nhập Google → Create API key → Copy
3. Dán vào ô sidebar
4. Tick **"Dùng AI để sinh nội dung"**
""")
