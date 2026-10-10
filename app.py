import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from datetime import datetime
import io, re

st.set_page_config(page_title="Agent Trợ lý Giáo viên THCS AI", page_icon="🎓", layout="wide")
st.title("🎓 Agent Trợ lý Giáo viên THCS AI")
st.caption("Bản MVP chạy cục bộ • Hỗ trợ theo yêu cầu • Xuất Word")

with st.sidebar:
    st.header("Thông tin")
    school = st.text_input("Trường/đơn vị")
    teacher = st.text_input("Giáo viên")
    subject = st.selectbox("Môn học", ["Toán","Ngữ văn","Tiếng Anh","Khoa học tự nhiên","Vật lí","Hóa học","Sinh học","Lịch sử và Địa lí","Tin học","Công nghệ","GDCD/GDKT&PL","Nghệ thuật","Giáo dục thể chất","Hoạt động trải nghiệm","Khác"])
    grade = st.selectbox("Khối lớp", ["6","7","8","9","Nhiều khối"])
    st.caption("Bản này dùng mẫu cục bộ, không cần API AI.")

task = st.selectbox("Chọn công việc", ["Soạn kế hoạch bài dạy","Tạo đề kiểm tra và khung đáp án","Tạo phiếu học tập","Soạn báo cáo/hồ sơ","Lập kịch bản video bài giảng","Soạn nội dung theo yêu cầu riêng"])
topic = st.text_input("Tên bài/chủ đề/công việc")
details = st.text_area("Yêu cầu cụ thể, mục tiêu, số liệu hoặc ghi chú", height=120)
period = st.text_input("Thời lượng hoặc thời gian (nếu cần)", placeholder="Ví dụ: 1 tiết / tháng 10 / video 2 phút")

def make_doc(title, body):
    d=Document()
    s=d.sections[0]
    s.top_margin=Inches(.7); s.bottom_margin=Inches(.7); s.left_margin=Inches(.8); s.right_margin=Inches(.8)
    d.styles["Normal"].font.name="Times New Roman"; d.styles["Normal"].font.size=Pt(12)
    d.add_heading(title, 0)
    if school.strip(): d.add_paragraph("Đơn vị: "+school.strip())
    if teacher.strip(): d.add_paragraph("Người thực hiện: "+teacher.strip())
    d.add_paragraph("Ngày tạo: "+datetime.now().strftime("%d/%m/%Y %H:%M"))
    for line in body.splitlines():
        if line.startswith("# "): d.add_heading(line[2:],1)
        elif line.startswith("## "): d.add_heading(line[3:],2)
        elif line.startswith("### "): d.add_heading(line[4:],3)
        elif line.startswith("- "): d.add_paragraph(line[2:],style="List Bullet")
        elif re.match(r"^\d+[\.\)] ",line): d.add_paragraph(re.sub(r"^\d+[\.\)] ","",line),style="List Number")
        else: d.add_paragraph(line)
    b=io.BytesIO(); d.save(b); return b.getvalue()

def generate():
    t=topic.strip() or "[Điền tên bài/chủ đề]"
    common=f"- Môn học: {subject}\n- Lớp: {grade}\n- Chủ đề: {t}\n- Thời gian: {period or 'Chưa xác định'}"
    if task=="Soạn kế hoạch bài dạy":
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
    if task=="Tạo đề kiểm tra và khung đáp án":
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
{details or 'Bổ sung ma trận/đặc tả, tổng điểm và thời gian làm bài theo yêu cầu của nhà trường.'}
LƯU Ý: Đây là khung đề, chưa phải đề hoàn chỉnh; cần thay phần trong ngoặc vuông bằng câu hỏi và đáp án chính xác."""
    if task=="Tạo phiếu học tập":
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
    if task=="Soạn báo cáo/hồ sơ":
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
    if task=="Lập kịch bản video bài giảng":
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

if st.button("✨ Tạo bản nháp", type="primary"):
    st.session_state["draft"] = generate()
    st.session_state["draft_title"] = task

if "draft" in st.session_state:
    st.divider()
    st.subheader("Bản nháp — có thể chỉnh sửa")
    st.warning("Bản MVP tạo khung tài liệu theo mẫu, chưa dùng mô hình AI để tự sinh kiến thức hoàn chỉnh. Hãy kiểm tra nội dung, số liệu và quy định trước khi sử dụng.")
    st.session_state["draft"] = st.text_area("Nội dung", st.session_state["draft"], height=420)
    name=re.sub(r"[^A-Za-z0-9_-]+","_",topic.strip())[:45] or "tai_lieu_giao_vien"
    st.download_button("⬇️ Tải Word (.docx)", make_doc(st.session_state["draft_title"],st.session_state["draft"]), file_name=name+".docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary")
    st.download_button("⬇️ Tải văn bản (.txt)", st.session_state["draft"].encode("utf-8"), file_name=name+".txt", mime="text/plain")
    if st.button("Xóa bản nháp"):
        del st.session_state["draft"]; del st.session_state["draft_title"]; st.rerun()

with st.expander("Lộ trình nâng cấp"):
    st.markdown("- Bản hiện tại: giao diện, mẫu công việc và xuất Word.\n- Bước 2: tích hợp AI API để sinh nội dung theo yêu cầu.\n- Bước 3: lưu kho tài liệu và lịch sử bằng SQLite.\n- Bước 4: Excel, PowerPoint, đọc tài liệu đầu vào.\n- Bước 5: tạo hình ảnh, video và điều phối Blender/FFmpeg.")
