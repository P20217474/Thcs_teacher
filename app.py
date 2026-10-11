"""
AGENT TRỢ LÝ GIÁO VIÊN THCS AI — BẢN NÂNG CẤP
- Tùy chọn AI (Google Gemini miễn phí)
- Nạp mẫu (docx/pdf/txt hoặc dán text) để AI bám theo
- Vẫn có chế độ mẫu cục bộ khi chưa có API key
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

# Thư viện đọc PDF (nếu có)
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
st.caption("Bản nâng cấp • Nạp mẫu • AI sinh nội dung • Xuất Word")

# ============ SESSION STATE ============
for k, v in {
    "draft": "",
    "draft_title": "",
    "template_text": "",
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ============ SIDEBAR ============
with st.sidebar:
    st.header("⚙️ Cấu hình AI")
    api_key = st.text_input(
        "Google Gemini API Key (tùy chọn)",
        type="password",
        help="Lấy miễn phí tại aistudio.google.com/app/apikey. Để trống nếu chỉ dùng mẫu cục bộ.",
    )
    use_ai = st.checkbox("Dùng AI để sinh nội dung", value=bool(api_key))

    st.divider()
    st.header("📄 Mẫu tham khảo")
    template_file = st.file_uploader(
        "Nạp file mẫu (docx / pdf / txt)",
        type=["docx", "pdf", "txt"],
    )
    template_paste = st.text_area(
        "Hoặc dán mẫu trực tiếp vào đây", height=120,
        placeholder="Dán giáo án mẫu, đề mẫu, phiếu học tập mẫu...",
    )

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
    details = st.text_area("Yêu cầu cụ thể, mục tiêu, số liệu hoặc ghi chú", height=120)
    period = st.text_input("Thời lượng hoặc thời gian (nếu cần)",
                           placeholder="Ví dụ: 1 tiết / tháng 10 / video 2 phút")


# ============ ĐỌC FILE MẪU ============
def read_docx(file_bytes: bytes) -> str:
    d = Document(io.BytesIO(file_bytes))
    return "\n".join(p.text for p in d.paragraphs if p.text.strip())


def read_pdf(file_bytes: bytes) -> str:
    if not HAS_PYPDF:
        return ""
    reader = PdfReader(io.BytesIO(file_bytes))
    return "\n".join((page.extract_text() or "") for page in reader.pages)


def get_template_text() -> str:
    parts = []
    if template_paste.strip():
        parts.append(template_paste.strip())
    if template_file is not None:
        data = template_file.read()
        name = template_file.name.lower()
        try:
            if name.endswith(".docx"):
                parts.append(read_docx(data))
            elif name.endswith(".pdf"):
                parts.append(read_pdf(data))
            elif name.endswith(".txt"):
                parts.append(data.decode("utf-8", errors="ignore"))
        except Exception as e:
            st.warning(f"Không đọc được file mẫu: {e}")
    text = "\n\n---\n\n".join(p for p in parts if p)
    return text[:8000]  # giới hạn để không vượt context


# ============ GỌI AI ============
def call_gemini(prompt: str, api_key: str) -> str:
    import google.generativeai as genai
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-2.0-flash")
    resp = model.generate_content(prompt)
    return resp.text


def build_prompt() -> str:
    t = topic.strip() or "[Chưa có tên]"
    template = get_template_text()
    tpl_block = f"""
DƯỚI ĐÂY LÀ MẪU THAM KHẢO (giáo viên cung cấp). Hãy bám sát cấu trúc, cách trình bày và văn phong của mẫu này:

\"\"\"
{template}
\"\"\"
""" if template else "Không có mẫu cụ thể; hãy dùng cấu trúc chuẩn của giáo dục THCS Việt Nam."

    return f"""Bạn là trợ lý chuyên môn cho giáo viên THCS Việt Nam.

NHIỆM VỤ: {task}
MÔN: {subject} | KHỐI: {grade}
CHỦ ĐỀ: {t}
THỜI LƯỢNG: {period or 'Chưa xác định'}
YÊU CẦU RIÊNG: {details or 'Không có'}

{tpl_block}

YÊU CẦU ĐẦU RA:
- Viết bằng tiếng Việt, văn phong sư phạm, trang trọng.
- Nội dung cụ thể, chi tiết, có thể dùng được ngay (không viết chung chung).
- Dùng markdown: '# ' cho tiêu đề cấp 1, '## ' cho cấp 2, '### ' cho cấp 3, '- ' cho gạch đầu dòng, '1. ' cho danh sách số.
- Bám sát chương trình GDPT 2018 và đặc thù môn {subject} lớp {grade}.
- Nếu có mẫu ở trên, phải theo đúng cấu trúc của mẫu.
- Không thêm lời chào, không giải thích, chỉ xuất nội dung tài liệu.
"""


# ============ MẪU CỤC BỘ (khi không dùng AI) ============
def generate_local() -> str:
    t = topic.strip() or "[Điền tên bài/chủ đề]"
    common = (f"- Môn học: {subject}\n- Lớp: {grade}\n- Chủ đề: {t}\n"
              f"- Thời gian: {period or 'Chưa xác định'}")

    if task == "Soạn kế hoạch bài dạy":
        return f"""# KẾ HOẠCH BÀI DẠY

## I. Thông tin chung
{common}

## II. Mục tiêu / yêu cầu cần đạt
- Kiến thức: nêu được kiến thức trọng tâm của bài.
- Năng lực: vận dụng kiến thức để giải quyết nhiệm vụ học tập.
- Phẩm chất: tích cực, trung thực và hợp tác trong học tập.

## III. Thiết bị và học liệu
- Giáo viên: sách giáo khoa, bảng/phương tiện trình chiếu, phiếu học tập.
- Học sinh: sách giáo khoa, vở ghi và dụng cụ học tập.

## IV. Tiến trình dạy học
### 1. Khởi động
- Nêu tình huống/câu hỏi gợi mở liên quan đến {t}.
### 2. Hình thành kiến thức
- Giao nhiệm vụ cá nhân hoặc nhóm; học sinh thảo luận và trình bày.
- Giáo viên nhận xét, chuẩn hóa kiến thức và đưa ví dụ.
### 3. Luyện tập
- Tổ chức bài tập từ nhận biết đến vận dụng; phản hồi cách làm.
### 4. Vận dụng
- Giao nhiệm vụ gắn với thực tiễn hoặc tình huống gần gũi.
### 5. Đánh giá
- Đánh giá qua câu trả lời, sản phẩm học tập và mức độ tham gia.

## V. Điều chỉnh sau bài dạy
- Nội dung cần điều chỉnh: bổ sung sau khi dạy thực tế.

## VI. Yêu cầu riêng của giáo viên
{details or 'Đối chiếu yêu cầu cần đạt, sách giáo khoa và kế hoạch giáo dục của nhà trường trước khi sử dụng.'}"""

    if task == "Tạo đề kiểm tra và khung đáp án":
        return f"""# ĐỀ KIỂM TRA — KHUNG SOẠN THẢO
{common}

## I. Trắc nghiệm
1. [Soạn câu hỏi phù hợp với nội dung đã học; 4 lựa chọn A, B, C, D.]
2. [Soạn câu hỏi ở mức độ thông hiểu.]
3. [Soạn câu hỏi vận dụng.]

## II. Tự luận
1. [Ghi rõ yêu cầu và điểm thành phần.]
2. [Ghi rõ yêu cầu và điểm thành phần.]

## III. Đáp án và hướng dẫn chấm
- Trắc nghiệm: ghi đáp án đúng và giải thích ngắn.
- Tự luận: ghi lời giải, thang điểm và tiêu chí chấp nhận cách làm tương đương.

## IV. Ghi chú
{details or 'Bổ sung ma trận/đặc tả, tổng điểm và thời gian làm bài theo yêu cầu của nhà trường.'}"""

    if task == "Tạo phiếu học tập":
        return f"""# PHIẾU HỌC TẬP
{common}

## Mục tiêu
{details or 'Nắm kiến thức trọng tâm và vận dụng vào nhiệm vụ học tập.'}

## Nhiệm vụ 1 — Khám phá
1. Ghi lại 3 ý quan trọng của bài học.
2. Nêu một điều em còn thắc mắc.

## Nhiệm vụ 2 — Luyện tập
1. [Câu hỏi nhận biết về chủ đề.]
2. [Câu hỏi thông hiểu.]
3. [Bài tập vận dụng.]

## Nhiệm vụ 3 — Vận dụng
Nêu ví dụ thực tế liên quan đến chủ đề và giải thích.

## Tự đánh giá
- [ ] Em hiểu kiến thức chính.
- [ ] Em hoàn thành nhiệm vụ.
- [ ] Em cần được hỗ trợ thêm."""

    if task == "Soạn báo cáo/hồ sơ":
        return f"""# BÁO CÁO / HỒ SƠ
{common}

## I. Mục đích
Nêu mục đích, phạm vi và thời gian thực hiện.

## II. Kết quả thực hiện
{details or '[Bổ sung công việc đã thực hiện, số liệu và minh chứng thực tế.]'}

## III. Đánh giá
- Ưu điểm: [nêu theo minh chứng].
- Hạn chế và nguyên nhân: [nêu khách quan].

## IV. Phương hướng
- Nhiệm vụ tiếp theo: [việc cần làm, thời hạn].
- Đề xuất/kiến nghị: [nội dung cần hỗ trợ].

## V. Lưu ý
Đối chiếu số liệu với hồ sơ gốc trước khi nộp."""

    if task == "Lập kịch bản video bài giảng":
        return f"""# KỊCH BẢN VIDEO BÀI GIẢNG
{common}

## 1. Mở đầu (khoảng 10%)
- Hiển thị tên bài và tình huống gợi mở về {t}.

## 2. Giải thích kiến thức (khoảng 55%)
- Chia nội dung thành 2–4 ý nhỏ, mỗi ý có ví dụ/hình minh họa.
- Hiển thị từ khóa ngắn, tránh nhiều chữ trên màn hình.

## 3. Ví dụ và luyện tập (khoảng 25%)
- Đưa câu hỏi, dành thời gian suy nghĩ rồi trình bày lời giải.

## 4. Tổng kết (khoảng 10%)
- Nhắc lại 3 ý chính và giao nhiệm vụ vận dụng.

## 5. Yêu cầu bổ sung
{details or 'Tách cảnh, lời thoại, chữ trên màn hình, âm thanh và motion thành các phần có thể tái sử dụng.'}"""

    return f"""# PHIẾU YÊU CẦU CÔNG VIỆC
{common}

## Yêu cầu của giáo viên
{details or '[Mô tả cụ thể sản phẩm cần tạo.]'}

## Quy trình xử lý đề xuất
1. Xác định mục tiêu và đối tượng sử dụng.
2. Thu thập dữ liệu đầu vào và tài liệu cần đối chiếu.
3. Tạo bản nháp, rà soát nội dung.
4. Giáo viên duyệt trước khi sử dụng chính thức."""


def generate() -> str:
    if use_ai:
        if not api_key.strip():
            st.error("Bạn đã bật AI nhưng chưa nhập API Key.")
            return ""
        try:
            with st.spinner("🤖 AI đang sinh nội dung..."):
                return call_gemini(build_prompt(), api_key.strip())
        except Exception as e:
            st.error(f"Lỗi khi gọi AI: {e}")
            st.info("Đã chuyển sang dùng mẫu cục bộ.")
            return generate_local()
    return generate_local()


# ============ TẠO FILE WORD ============
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
if use_ai:
    st.success("🤖 Chế độ AI đang bật — nội dung sẽ được sinh tự động.")
else:
    st.info("📝 Chế độ mẫu cục bộ — chưa dùng AI. Nhập API key ở sidebar để bật AI.")

if get_template_text():
    st.success("📄 Đã nhận mẫu tham khảo — AI sẽ bám theo cấu trúc này.")

if st.button("✨ Tạo bản nháp", type="primary"):
    if not topic.strip():
        st.warning("⚠️ Bạn nên nhập 'Tên bài/chủ đề' để tài liệu rõ ràng hơn.")
    result = generate()
    if result:
        st.session_state["draft"] = result
        st.session_state["draft_title"] = task

if st.session_state["draft"]:
    st.divider()
    st.subheader("Bản nháp — có thể chỉnh sửa")
    st.session_state["draft"] = st.text_area(
        "Nội dung", st.session_state["draft"], height=420
    )
    name = re.sub(r"[^A-Za-z0-9_\-]", "_", topic.strip())[:45] or "tai_lieu_giao_vien"

    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        st.download_button(
            "⬇️ Tải Word (.docx)",
            make_doc(st.session_state["draft_title"], st.session_state["draft"]),
            file_name=name + ".docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
        )
    with col2:
        st.download_button(
            "⬇️ Tải văn bản (.txt)",
            st.session_state["draft"].encode("utf-8"),
            file_name=name + ".txt",
            mime="text/plain",
        )
    with col3:
        if st.button("🗑️ Xóa bản nháp"):
            st.session_state["draft"] = ""
            st.session_state["draft_title"] = ""
            st.rerun()

with st.expander("📖 Hướng dẫn lấy API Key Gemini miễn phí"):
    st.markdown("""
1. Mở trình duyệt, vào **aistudio.google.com/app/apikey**
2. Đăng nhập bằng tài khoản Google
3. Bấm **"Create API key"** → chọn project → **Copy** key
4. Dán vào ô **"Google Gemini API Key"** ở sidebar
5. Tick chọn **"Dùng AI để sinh nội dung"**
6. Nạp mẫu (file docx/pdf/txt hoặc dán text) nếu muốn AI bám theo mẫu
7. Bấm **"Tạo bản nháp"** — AI sẽ sinh nội dung hoàn chỉnh

**Miễn phí**: Gemini 1.5 Flash có hạn mức miễn phí 15 request/phút, 1500 request/ngày — quá đủ cho nhu cầu giáo viên.
""")
