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
st.caption("Ban 2.1 • MAU + NGUON • AI tu dong • Xuat Word")


def get_default_api_key() -> str:
    try:
        key = st.secrets.get("GEMINI_API_KEY", "")
        if key:
            return key
    except Exception:
        pass
    return os.environ.get("GEMINI_API_KEY", "")


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


with st.sidebar:
    st.header("Cau hinh AI")
    default_key = get_default_api_key()
    if default_key:
        st.success("Da co API Key tu Secrets")
        api_key = default_key
        use_ai = st.checkbox("Dung AI de sinh noi dung", value=True)
        with st.expander("Doi API Key khac"):
            override = st.text_input("Nhap key moi", type="password", key="override_key")
            if override.strip():
                api_key = override.strip()
    else:
        api_key = st.text_input("Google Gemini API Key", type="password", key="manual_key")
        use_ai = st.checkbox("Dung AI de sinh noi dung", value=bool(api_key))
        st.caption("Co the cai dat key vinh vien trong Settings > Secrets.")

    st.divider()
    st.header("Thong tin")
    school = st.text_input("Truong/don vi")
    teacher = st.text_input("Giao vien")
    subject = st.selectbox("Mon hoc", [
        "Toan", "Ngu van", "Tieng Anh", "Khoa hoc tu nhien",
        "Vat li", "Hoa hoc", "Sinh hoc", "Lich su va Dia li",
        "Tin hoc", "Cong nghe", "GDCD", "Nghe thuat",
        "Giao duc the chat", "Hoat dong trai nghiem", "Khac",
    ])
    grade = st.selectbox("Khoi lop", ["6", "7", "8", "9", "Nhieu khoi"])

    st.divider()
    st.subheader("Cong viec")
    task = st.selectbox("Chon cong viec", [
        "Soan ke hoach bai day",
        "Tao de kiem tra va khung dap an",
        "Tao phieu hoc tap",
        "Soan bao cao/ho so",
        "Lap kich ban video bai giang",
        "Soan noi dung theo yeu cau rieng",
    ])
    topic = st.text_input("Ten bai/chu de")
    details = st.text_area("Yeu cau cu the", height=100)
    period = st.text_input("Thoi luong")


tab_mau, tab_nguon, tab_xem = st.tabs([
    "TUY CHON MAU (dinh dang)",
    "TUY CHON NGUON (noi dung)",
    "XEM TRUOC DA NAP",
])

with tab_mau:
    st.markdown("**MAU** quy dinh cau truc, thu tu muc, cach trinh bay. AI se bam sat cau truc nay.")
    tpl_files = st.file_uploader(
        "Nap 1 hoac nhieu file MAU (docx/pdf/txt)",
        type=["docx", "pdf", "txt"],
        accept_multiple_files=True,
        key="tpl_uploader",
    )
    tpl_paste = st.text_area("Hoac dan mau truc tiep", height=140, key="tpl_paste")
    if st.button("Xoa mau da nap", key="clear_tpl"):
        st.session_state["template_text"] = ""
        st.session_state["template_names"] = []
        st.rerun()

with tab_nguon:
    st.markdown("**NGUON** chua kien thuc, so lieu, du kien. AI chi lay noi dung tu nguon.")
    src_files = st.file_uploader(
        "Nap 1 hoac nhieu file NGUON (docx/pdf/txt)",
        type=["docx", "pdf", "txt"],
        accept_multiple_files=True,
        key="src_uploader",
    )
    src_paste = st.text_area("Hoac dan noi dung nguon truc tiep", height=140, key="src_paste")
    if st.button("Xoa nguon da nap", key="clear_src"):
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


template_text, template_names = collect_text(tpl_files, tpl_paste, "template_text", "template_names")
source_text, source_names = collect_text(src_files, src_paste, "source_text", "source_names")

with tab_xem:
    st.subheader("MAU da nap")
    if template_names:
        for n in template_names:
            st.write("- " + n)
        st.caption("Tong: " + str(len(template_text)) + " ky tu")
    else:
        st.info("Chua nap mau.")
    st.divider()
    st.subheader("NGUON da nap")
    if source_names:
        for n in source_names:
            st.write("- " + n)
        st.caption("Tong: " + str(len(source_text)) + " ky tu")
    else:
        st.info("Chua nap nguon.")


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
        tpl_block = "=== FILE MAU (BAT BUOC bam theo) ===\n" + template_text[:12000] + "\n=== HET MAU ==="
    else:
        tpl_block = "(Khong co mau - dung cau truc chuan GDPT 2018.)"
    if source_text:
        src_block = "=== NGUON TAI LIEU (chi lay noi dung tu day) ===\n" + source_text[:20000] + "\n=== HET NGUON ==="
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
    common = "- Mon hoc: " + subject + "\n- Lop: " + grade + "\n- Chu de: " + t + "\n- Thoi gian: " + (period or "Chua xac dinh")
    return "# " + task.upper() + "\n" + common + "\n\n## I. Thong tin chung\n" + common + "\n\n## II. Noi dung chinh\n" + (details or "[Bo sung yeu cau cu the]") + "\n\n## III. Ghi chu\n- Ban mau cuc bo."


def generate():
    if use_ai:
        if not api_key or not api_key.strip():
            st.error("Chua co API Key.")
            return ""
        try:
            with st.spinner("AI dang sinh noi dung..."):
                return call_gemini(build_prompt(), api_key.strip())
        except Exception as e:
            st.error("Loi AI: " + str(e))
            st.info("Chuyen sang mau cuc bo.")
            return generate_local()
    return generate_local()


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
            d.add_paragraph(re.sub(r"^\d+[\.\)] ", "", line), style="List Number")
        elif line.strip() == "":
            continue
        else:
            d.add_paragraph(line)
    b = io.BytesIO()
    d.save(b)
    return b.getvalue()


st.divider()
st.subheader("Tao tai lieu")
col1, col2 = st.columns([1, 2])
with col1:
    if st.button("Tao ban nhap", type="primary", use_container_width=True):
        if not topic.strip():
            st.warning("Nen nhap 'Ten bai/chu de'.")
        result = generate()
        if result:
            st.session_state["draft"] = result
            st.session_state["draft_title"] = task
with col2:
    status = []
    if template_text:
        status.append(str(len(template_names)) + " mau")
    if source_text:
        status.append(str(len(source_names)) + " nguon")
    if use_ai and api_key:
        status.append("AI bat")
    if status:
        st.caption(" | ".join(status))

if st.session_state["draft"]:
    st.divider()
    st.subheader("Ban nhap - co the chinh sua")
    st.session_state["draft"] = st.text_area("Noi dung", st.session_state["draft"], height=420)
    name = re.sub(r"[^A-Za-z0-9_\-]", "_", topic.strip())[:45] or "tai_lieu_giao_vien"
    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        st.download_button(
            "Tai Word (.docx)",
            make_doc(st.session_state["draft_title"], st.session_state["draft"]),
            file_name=name + ".docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
        )
    with c2:
        st.download_button(
            "Tai van ban (.txt)",
            st.session_state["draft"].encode("utf-8"),
            file_name=name + ".txt",
            mime="text/plain",
        )
    with c3:
        if st.button("Xoa ban nhap"):
            st.session_state["draft"] = ""
            st.session_state["draft_title"] = ""
            st.rerun()

with st.expander("Huong dan cai API Key vinh vien"):
    st.write("Buoc 1: Lay API Key mien phi tai aistudio.google.com/app/apikey")
    st.write("Buoc 2: Mo app Streamlit > bam ... goc duoi phai > Settings > tab Secrets")
    st.write("Buoc 3: Dan dong sau (thay bang key cua ban):")
    st.code('GEMINI_API_KEY = "AIzaSy...key_cua_ban..."', language="toml")
    st.write("Buoc 4: Bam Save > Reboot")
    st.write("Tu do AI se tu dong bat, khong can dan key nua.")
