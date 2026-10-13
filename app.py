# ============================================================
# PHẦN 1/6: HẠ TẦNG
# ============================================================
import os, json, base64, urllib.request
from datetime import datetime, timedelta

import streamlit as st
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
import io, re, csv

try:
    from pypdf import PdfReader
    HAS_PYPDF = True
except Exception:
    HAS_PYPDF = False

try:
    import fitz
    HAS_PYMUPDF = True
except Exception:
    HAS_PYMUPDF = False

try:
    from PIL import Image
    HAS_PIL = True
except Exception:
    HAS_PIL = False

try:
    import google.generativeai as genai_vision
    HAS_GEMINI_VISION = True
except Exception:
    HAS_GEMINI_VISION = False

st.set_page_config(
    page_title="Trợ lý Giáo viên THCS",
    page_icon="🧠",
    layout="wide",
)
st.title("🧠 Trợ lý Giáo viên THCS")
st.caption("Bộ não tự học • Mẫu CV 5512 • Sách Kết nối tri thức • Xuất Word")


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
GEMINI_KEY = PROVIDER_KEYS.get("GEMINI_API_KEY", "")

PREFERRED_MODELS = {
    "GEMINI_API_KEY": ["gemini-2.5-flash", "gemini-2.0-flash",
                       "gemini-flash-latest"],
    "GROQ_API_KEY": ["llama-3.3-70b-versatile",
                     "llama-3.1-70b-versatile"],
    "OPENROUTER_API_KEY": ["google/gemini-2.0-flash-exp:free"],
    "NVIDIA_API_KEY": ["meta/llama-3.3-70b-instruct"],
}

GITHUB_TOKEN = cfg("GITHUB_TOKEN", "")
GITHUB_REPO = cfg("GITHUB_REPO", "P20217474/Thcs_teacher")
GITHUB_BRANCH = cfg("GITHUB_BRANCH", "main")

CACHE_DIR = "cache"
SGK_BAI_DIR = os.path.join(CACHE_DIR, "sgk_bai")
SGK_PAGES_DIR = os.path.join(CACHE_DIR, "sgk_pages")
OCR_CACHE_DIR = os.path.join(CACHE_DIR, "ocr")

for _d in [CACHE_DIR, SGK_BAI_DIR, SGK_PAGES_DIR, OCR_CACHE_DIR]:
    os.makedirs(_d, exist_ok=True)


def gh_api(method, path, data=None):
    if not GITHUB_TOKEN:
        return {"error": "Chưa có GITHUB_TOKEN"}
    url = "https://api.github.com" + path
    headers = {
        "Authorization": "token " + GITHUB_TOKEN,
        "Accept": "application/vnd.github+json",
        "User-Agent": "agent-thcs",
    }
    body = json.dumps(data).encode("utf-8") if data else None
    req = urllib.request.Request(url, data=body, headers=headers,
                                 method=method)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        return {"error": str(e)}


def gh_get(path):
    res = gh_api("GET",
                 f"/repos/{GITHUB_REPO}/contents/{path}?ref={GITHUB_BRANCH}")
    if not res or "error" in res or "content" not in res:
        return (None, None)
    try:
        content = base64.b64decode(res["content"]).decode("utf-8")
        return (content, res["sha"])
    except Exception:
        return (None, None)


def gh_save(path, content, message="Update"):
    if not GITHUB_TOKEN:
        return (False, "Chưa có GITHUB_TOKEN")
    old, sha = gh_get(path)
    data = {
        "message": message,
        "content": base64.b64encode(content.encode("utf-8")).decode("ascii"),
        "branch": GITHUB_BRANCH,
    }
    if sha:
        data["sha"] = sha
    res = gh_api("PUT", f"/repos/{GITHUB_REPO}/contents/{path}", data)
    if isinstance(res, dict) and "error" in res:
        return (False, "Lỗi: " + str(res["error"]))
    if isinstance(res, dict) and "content" in res:
        return (True, "Đã lưu: " + path)
    return (False, "Phản hồi không hợp lệ")


def gh_delete(path):
    if not GITHUB_TOKEN:
        return (False, "Chưa có GITHUB_TOKEN")
    old, sha = gh_get(path)
    if not sha:
        return (False, "File không tồn tại")
    res = gh_api("DELETE", f"/repos/{GITHUB_REPO}/contents/{path}",
                 {"message": "Delete " + path, "sha": sha,
                  "branch": GITHUB_BRANCH})
    if isinstance(res, dict) and "error" in res:
        return (False, "Lỗi: " + str(res["error"]))
    if isinstance(res, dict) and ("content" in res or "commit" in res):
        return (True, "Đã xóa: " + path)
    return (False, "Phản hồi không hợp lệ")


def gh_list(folder):
    res = gh_api("GET",
                 f"/repos/{GITHUB_REPO}/contents/{folder}?ref={GITHUB_BRANCH}")
    if not res or "error" in res or not isinstance(res, list):
        return []
    return [item["name"] for item in res if item["type"] == "file"]


def gh_test():
    if not GITHUB_TOKEN:
        return (False, "Thiếu GITHUB_TOKEN")
    res = gh_api("GET", "/user")
    if isinstance(res, dict) and "login" in res:
        login = res["login"]
        res2 = gh_api("GET", f"/repos/{GITHUB_REPO}")
        if isinstance(res2, dict) and "permissions" in res2:
            perms = res2["permissions"]
            if perms.get("push") or perms.get("admin"):
                return (True, "OK --- " + login + " | Có quyền ghi")
            return (False, "Token thiếu quyền 'repo'")
        return (True, "OK --- " + login)
    if isinstance(res, dict) and "error" in res:
        return (False, "Lỗi: " + res["error"])
    return (False, "Token không hợp lệ")


def load_memory():
    content, _ = gh_get("data/memory.json")
    if content:
        try:
            return json.loads(content)
        except Exception:
            pass
    return {"teacher_name": "", "school": "",
            "default_subject": "Toán", "default_grade": "6"}


def save_memory(memory):
    memory["last_updated"] = datetime.now().strftime("%Y-%m-%d")
    return gh_save("data/memory.json",
                   json.dumps(memory, ensure_ascii=False, indent=2),
                   "Update memory")


def load_folder(folder):
    files = gh_list(f"data/{folder}")
    out = {}
    for f in files:
        c, _ = gh_get(f"data/{folder}/{f}")
        if c:
            out[f] = c
    return out


def load_templates(): return load_folder("templates")
def load_sources():   return load_folder("sources")
def load_schedule():  return load_folder("schedule")


def save_history(task, topic, content):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe = re.sub(r"[^a-zA-Z0-9]", "_", topic)[:30]
    return gh_save(
        f"data/history/{ts}_{safe}.md",
        f"# {task}\n\n**Chủ đề:** {topic}\n\n---\n\n{content}",
        f"History {topic}")


# Session state
if "app_started_v23" not in st.session_state:
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    st.session_state["app_started_v23"] = True

defaults = {
    "memory": load_memory(),
    "templates": load_templates(),
    "sources": load_sources(),
    "schedule": load_schedule(),
    "draft": "",
    "last_topic": "",
    "weekly_results": [],
    "lich_bao_giang": [],
    "mau_5512_dang_dung": None,
    "ocr_session_cache": {},
    "matran_hientai": None,
    "dacta_hientai": None,
    "de_kiem_tra_hientai": None,
    "ai_prompt_cache": {},
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v
# ============ KẾT THÚC PHẦN 1/6 ============
# ============================================================
# PHẦN 2/6: ĐỌC FILE + OCR + TÁCH SGK THEO BÀI + CACHE
# ============================================================
# Chức năng:
#   - Đọc file: PDF, Word, Excel, CSV, TXT, ảnh
#   - OCR ảnh bằng Gemini Vision (cache local + session)
#   - Tách SGK PDF theo BÀI (quét từng trang, không dùng mục lục)
#   - Gom nhóm trang thành file bài riêng
#   - Tạo metadata.json mapping bài → file
#   - Tìm bài theo TÊN (fuzzy match)
#   - Quản lý cache local
# ============================================================

import hashlib
import unicodedata
from pathlib import Path


# ============================================================
# PHẦN A: OCR ẢNH BẰNG GEMINI VISION
# ============================================================

def _ocr_cache_key(data):
    """Tạo hash cho ảnh để cache OCR."""
    return hashlib.md5(data).hexdigest()


def ocr_image_with_gemini(data, fname=""):
    """OCR ảnh bằng Gemini Vision.
    - Ưu tiên cache session
    - Sau đó cache local (file .txt)
    - Cuối cùng mới gọi API
    """
    if not GEMINI_KEY:
        return "[" + fname + ": cần GEMINI_API_KEY]"
    if not HAS_GEMINI_VISION:
        return "[" + fname + ": chưa cài google-generativeai]"
    if not HAS_PIL:
        return "[" + fname + ": chưa cài pillow]"

    # 1. Cache session
    cache_key = _ocr_cache_key(data)
    if cache_key in st.session_state.get("ocr_session_cache", {}):
        return st.session_state["ocr_session_cache"][cache_key]

    # 2. Cache local
    local_cache_file = os.path.join(OCR_CACHE_DIR, cache_key + ".txt")
    if os.path.exists(local_cache_file):
        try:
            with open(local_cache_file, "r", encoding="utf-8") as f:
                text = f.read()
            st.session_state["ocr_session_cache"][cache_key] = text
            return text
        except Exception:
            pass

    # 3. Gọi API Gemini Vision
    try:
        genai_vision.configure(api_key=GEMINI_KEY)
        img = Image.open(io.BytesIO(data))
        for mn in ["gemini-2.0-flash", "gemini-2.5-flash",
                   "gemini-flash-latest"]:
            try:
                model = genai_vision.GenerativeModel(mn)
                prompt = (
                    "Đây là ảnh trang sách giáo khoa. "
                    "Hãy đọc TOÀN BỘ chữ trong ảnh, giữ nguyên cấu trúc, "
                    "dùng Unicode cho toán: x², √x, ≤, π. "
                    "Chỉ trả về text, không mô tả."
                )
                resp = model.generate_content([prompt, img])
                text = resp.text.strip()
                if len(text) >= 20:
                    st.session_state["ocr_session_cache"][cache_key] = text
                    try:
                        with open(local_cache_file, "w",
                                  encoding="utf-8") as f:
                            f.write(text)
                    except Exception:
                        pass
                    return text
            except Exception:
                continue
        return "[" + fname + ": AI không đọc được]"
    except Exception as e:
        return "[" + fname + ": Lỗi OCR — " + str(e) + "]"


# ============================================================
# PHẦN B: ĐỌC FILE AN TOÀN
# ============================================================

def _read_docx_safe(data):
    """Đọc file Word an toàn."""
    try:
        d = Document(io.BytesIO(data))
        out = [p.text for p in d.paragraphs if p.text.strip()]
        for t in d.tables:
            out.append("")
            for row in t.rows:
                out.append(
                    "| " + " | ".join(c.text.strip() for c in row.cells)
                    + " |")
            out.append("")
        return "\n".join(out)
    except Exception as e:
        return "[Lỗi docx: " + str(e) + "]"


def _read_pdf_safe(data):
    """Đọc text từ PDF (chỉ hoạt động nếu PDF có text layer)."""
    if not HAS_PYPDF:
        return "[Chưa cài pypdf]"
    try:
        reader = PdfReader(io.BytesIO(data))
        texts = []
        for page in reader.pages:
            try:
                try:
                    t = page.extract_text(extraction_mode="layout") or ""
                except TypeError:
                    t = page.extract_text() or ""
                if t.strip():
                    texts.append(t)
            except Exception:
                continue
        full = "\n".join(texts)
        if len(full.strip()) < 50:
            return "[PDF scan, không có text layer]" + full
        return full
    except Exception as e:
        return "[Lỗi pdf: " + str(e) + "]"


def _read_excel_safe(data):
    """Đọc file Excel an toàn."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        out = []
        for sn in wb.sheetnames:
            ws = wb[sn]
            rows = [
                [str(c) if c is not None else "" for c in row]
                for row in ws.iter_rows(values_only=True)
            ]
            rows = [r for r in rows if any(c.strip() for c in r)]
            if not rows:
                continue
            mc = max(len(r) for r in rows)
            rows = [r + [""] * (mc - len(r)) for r in rows]
            out.append("### Sheet: " + sn)
            out.append("| " + " | ".join(rows[0]) + " |")
            out.append("|" + "---|" * mc)
            for r in rows[1:]:
                out.append("| " + " | ".join(r) + " |")
            out.append("")
        return "\n".join(out)
    except Exception as e:
        return "[Lỗi Excel: " + str(e) + "]"


def _read_csv_safe(data):
    """Đọc file CSV an toàn."""
    try:
        text = None
        for enc in ["utf-8", "utf-8-sig", "latin-1", "cp1258"]:
            try:
                text = data.decode(enc)
                break
            except Exception:
                continue
        if text is None:
            text = data.decode("utf-8", errors="ignore")
        sample = text[:2000]
        delim = "," if sample.count(",") >= sample.count("|") else "|"
        rows = list(csv.reader(io.StringIO(text), delimiter=delim))
        mc = max(len(r) for r in rows)
        rows = [r + [""] * (mc - len(r)) for r in rows]
        out = ["| " + " | ".join(rows[0]) + " |", "|" + "---|" * mc]
        for r in rows[1:]:
            out.append("| " + " | ".join(r) + " |")
        return "\n".join(out)
    except Exception as e:
        return "[Lỗi csv: " + str(e) + "]"


def _read_txt_safe(data):
    """Đọc file text an toàn."""
    for enc in ["utf-8", "utf-8-sig", "latin-1", "cp1258"]:
        try:
            return data.decode(enc)
        except Exception:
            continue
    return data.decode("utf-8", errors="ignore")


def read_file(file, use_ocr=True):
    """Đọc file bất kỳ. Trả về (name, text, size)."""
    if file is None:
        return ("", "", 0)
    try:
        data = file.read()
    except Exception as e:
        return (file.name, "[Lỗi: " + str(e) + "]", 0)
    if not data:
        return (file.name, "[File rỗng]", 0)

    name = file.name
    nl = name.lower()
    magic = data[:8]

    if nl.endswith((".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp")):
        text = ocr_image_with_gemini(data, name) if use_ocr \
            else "[" + name + ": ảnh]"
        return (name, text, len(text))

    if magic[:4] == b"%PDF":
        text = _read_pdf_safe(data)
    elif magic[:4] == b"PK\x03\x04":
        text = _read_excel_safe(data) if nl.endswith((".xlsx", ".xlsm")) \
            else _read_docx_safe(data)
    elif magic[:4] == b"\xd0\xcf\x11\xe0":
        text = "[File .doc cũ. Save As .docx trước]"
    elif nl.endswith(".docx"):
        text = _read_docx_safe(data)
    elif nl.endswith(".pdf"):
        text = _read_pdf_safe(data)
    elif nl.endswith((".xlsx", ".xls")):
        text = _read_excel_safe(data)
    elif nl.endswith(".csv"):
        text = _read_csv_safe(data)
    else:
        text = _read_txt_safe(data)
    return (name, text, len(text))


# ============================================================
# PHẦN C: TÁCH SGK PDF THEO BÀI
# ============================================================

def _slugify(text):
    """Chuẩn hóa tên file: bỏ dấu, lowercase, thay space = _."""
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    text = text.replace("đ", "d").replace("Đ", "D")
    text = text.lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_")


def _render_pdf_page_to_image(pdf_bytes, page_num, dpi=150):
    """Render 1 trang PDF thành ảnh PNG bytes."""
    if not HAS_PYMUPDF:
        return None
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if page_num < 0 or page_num >= len(doc):
            doc.close()
            return None
        page = doc[page_num]
        pix = page.get_pixmap(dpi=dpi)
        img_bytes = pix.tobytes("png")
        doc.close()
        return img_bytes
    except Exception:
        return None


def extract_pdf_page_text(pdf_bytes, page_num):
    """Trích text 1 trang PDF nếu có text layer."""
    if not HAS_PYMUPDF:
        return ""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if page_num < 0 or page_num >= len(doc):
            doc.close()
            return ""
        page = doc[page_num]
        text = page.get_text()
        doc.close()
        return text.strip()
    except Exception:
        return ""


def get_pdf_total_pages(pdf_bytes):
    """Đếm số trang PDF."""
    if not HAS_PYMUPDF:
        return 0
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        n = len(doc)
        doc.close()
        return n
    except Exception:
        return 0


def extract_toc_from_pdf(pdf_bytes):
    """Trích mục lục PDF nếu có."""
    if not HAS_PYMUPDF:
        return []
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        toc = doc.get_toc()
        doc.close()
        return toc
    except Exception:
        return []


def _is_trang_muc_luc(text):
    """Đoán xem trang OCR có phải mục lục không.
    Đặc điểm: nhiều dòng có '.... số'.
    """
    if not text:
        return False
    lines = [ln for ln in text.splitlines() if ln.strip()]
    if len(lines) < 3:
        return False
    so_dong_co_cham = 0
    for ln in lines:
        if re.search(r"\.{3,}\s*\d+\s*$", ln):
            so_dong_co_cham += 1
    return so_dong_co_cham >= max(3, len(lines) * 0.4)


def _tim_tieu_de_bai(text):
    """Tìm tiêu đề bài trong 1 trang OCR.
    Trả về list [{so_bai_raw, ten_bai, dong_goc}] (0-n bài/trang).
    """
    if not text:
        return []

    found = []
    lines = [ln.rstrip() for ln in text.splitlines()]

    patterns = [
        # "Bài 1. Tập hợp" | "BÀI 1: Tập hợp"
        re.compile(
            r"^\s*(?:Bài|BÀI|Bai)\s+(\d+)\s*[\.\:\-–]\s*(.+?)\s*$",
            re.IGNORECASE),
        # "1. Tập hợp" (số + dấu chấm + tên, chữ HOA đầu)
        re.compile(
            r"^\s*(\d{1,2})\s*[\.\:\-–]\s+"
            r"([A-ZĐÀÁẢÃẠĂÂÊÔƠƯ][^\n]{3,80})\s*$"),
        # "Unit 1: My New School"
        re.compile(
            r"^\s*(?:Unit|UNIT)\s+(\d+)\s*[\.\:\-–]\s*(.+?)\s*$",
            re.IGNORECASE),
        # "Lesson 1: ..."
        re.compile(
            r"^\s*(?:Lesson|LESSON)\s+(\d+)\s*[\.\:\-–]\s*(.+?)\s*$",
            re.IGNORECASE),
        # "Chương I. ..."
        re.compile(
            r"^\s*(?:Chương|CHƯƠNG)\s+([IVXLC\d]+)\s*[\.\:\-–]\s*(.+?)\s*$",
            re.IGNORECASE),
        # "CHỦ ĐỀ 1. ..."
        re.compile(
            r"^\s*(?:Chủ đề|CHỦ ĐỀ|Chu de)\s+(\d+)\s*[\.\:\-–]\s*(.+?)\s*$",
            re.IGNORECASE),
    ]

    for i, line in enumerate(lines):
        stripped = line.strip()
        if not stripped:
            continue
        # Bỏ dòng mục lục có "...."
        if re.search(r"\.{3,}\s*\d+\s*$", stripped):
            continue
        # Bỏ dòng quá dài (đoạn văn)
        if len(stripped) > 100:
            continue

        for pat in patterns:
            m = pat.match(stripped)
            if m:
                so_bai = m.group(1)
                ten_bai = m.group(2).strip()
                if len(ten_bai) < 2:
                    continue
                if re.fullmatch(r"[\d\s\.\-]+", ten_bai):
                    continue
                ten_bai = ten_bai.rstrip(".,;: ")
                found.append({
                    "so_bai_raw": so_bai,
                    "ten_bai": ten_bai,
                    "dong_goc": stripped,
                    "line_idx": i,
                })
                break

    return found


def detect_chapter_boundaries(pdf_bytes, max_pages_check=250,
                              progress_cb=None):
    """Quét SGK, tìm trang bắt đầu mỗi bài.

    Quy trình:
    1. OCR từng trang (dùng text layer nếu có)
    2. Bỏ qua trang mục lục
    3. Tìm tiêu đề bài đầu tiên trên trang
    4. Ghi nhận trang đó là trang bắt đầu của bài
    5. Trang kết thúc = trang trước bài kế - 1

    Returns:
        list [{ten_bai, so_bai, page_pdf_start, page_pdf_end, dong_goc}]
    """
    if not HAS_PYMUPDF or not HAS_GEMINI_VISION:
        return []

    total_pages = get_pdf_total_pages(pdf_bytes)
    if total_pages == 0:
        return []

    n_scan = min(max_pages_check, total_pages)
    bai_theo_trang = {}

    for i in range(n_scan):
        if progress_cb:
            progress_cb(int((i + 1) / n_scan * 100),
                        f"Đang quét trang {i+1}/{n_scan}...")

        # 1. Thử text layer trước (nhanh, không tốn token)
        text = extract_pdf_page_text(pdf_bytes, i)

        # 2. Nếu text layer quá ngắn → OCR ảnh
        if not text or len(text.strip()) < 30:
            img_bytes = _render_pdf_page_to_image(pdf_bytes, i, dpi=150)
            if not img_bytes:
                continue
            text = ocr_image_with_gemini(img_bytes, f"page_{i+1}")

        if not text or text.startswith("["):
            continue

        # Bỏ trang mục lục
        if _is_trang_muc_luc(text):
            continue

        # Tìm tiêu đề bài
        ds_bai = _tim_tieu_de_bai(text)
        if ds_bai:
            # Chỉ lấy bài đầu tiên trên trang
            bai_theo_trang[i] = ds_bai[0]

    # Chuyển thành boundaries
    pages_sorted = sorted(bai_theo_trang.keys())
    if not pages_sorted:
        return []

    boundaries = []
    for idx, page_idx in enumerate(pages_sorted):
        b = bai_theo_trang[page_idx]
        if idx + 1 < len(pages_sorted):
            page_end = pages_sorted[idx + 1] - 1
        else:
            page_end = total_pages - 1

        boundaries.append({
            "ten_bai": b["ten_bai"],
            "so_bai": idx + 1,
            "page_pdf_start": page_idx,
            "page_pdf_end": page_end,
            "dong_goc": b["dong_goc"],
        })

    return boundaries


def split_sgk_into_lessons(pdf_bytes, sgk_name, boundaries,
                           progress_cb=None):
    """Tách SGK PDF thành các file bài riêng.

    Args:
        pdf_bytes: bytes PDF gốc
        sgk_name: tên SGK (VD: "SGK_Toan_6")
        boundaries: list [{ten_bai, so_bai, page_pdf_start, page_pdf_end}]
        progress_cb: callback(percent, msg)

    Returns:
        dict metadata {sgk_name, ds_bai: [...]}
    """
    if not HAS_PYMUPDF:
        return None

    sgk_dir = os.path.join(SGK_BAI_DIR, sgk_name)
    os.makedirs(sgk_dir, exist_ok=True)

    total = len(boundaries)
    metadata = {
        "sgk_name": sgk_name,
        "nguon": "auto",
        "so_bai": total,
        "ds_bai": [],
    }

    for idx, b in enumerate(boundaries):
        if progress_cb:
            progress_cb(
                int((idx + 1) / total * 100),
                f"Đang tách bài {idx+1}/{total}: {b['ten_bai'][:40]}")

        safe_ten = _slugify(b["ten_bai"])[:40] or f"bai_{b['so_bai']:02d}"
        fname = f"bai_{b['so_bai']:02d}_{safe_ten}.pdf"
        fpath = os.path.join(sgk_dir, fname)

        try:
            doc_src = fitz.open(stream=pdf_bytes, filetype="pdf")
            doc_new = fitz.open()
            for p in range(b["page_pdf_start"], b["page_pdf_end"] + 1):
                if 0 <= p < len(doc_src):
                    doc_new.insert_pdf(doc_src, from_page=p, to_page=p)
            doc_new.save(fpath)
            doc_new.close()
            doc_src.close()

            metadata["ds_bai"].append({
                "file": fname,
                "ten_bai": b["ten_bai"],
                "so_bai": b["so_bai"],
                "page_pdf_start": b["page_pdf_start"],
                "page_pdf_end": b["page_pdf_end"],
            })
        except Exception:
            continue

    meta_path = os.path.join(sgk_dir, "metadata.json")
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(metadata, f, ensure_ascii=False, indent=2)

    return metadata


# ============================================================
# PHẦN D: ĐỌC / GHI METADATA
# ============================================================

def load_sgk_metadata(sgk_name):
    """Đọc metadata.json của SGK đã tách."""
    meta_path = os.path.join(SGK_BAI_DIR, sgk_name, "metadata.json")
    if not os.path.exists(meta_path):
        return None
    try:
        with open(meta_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def list_sgk_da_tach():
    """Liệt kê các SGK đã tách trong cache."""
    if not os.path.exists(SGK_BAI_DIR):
        return []
    out = []
    for name in os.listdir(SGK_BAI_DIR):
        subdir = os.path.join(SGK_BAI_DIR, name)
        if os.path.isdir(subdir):
            meta = load_sgk_metadata(name)
            if meta:
                out.append({
                    "sgk_name": name,
                    "so_bai": meta.get("so_bai", 0),
                    "nguon": meta.get("nguon", "unknown"),
                })
    return out


def ocr_bai_pdf(sgk_name, bai_file):
    """OCR toàn bộ file PDF của 1 bài → text."""
    fpath = os.path.join(SGK_BAI_DIR, sgk_name, bai_file)
    if not os.path.exists(fpath):
        return ""

    with open(fpath, "rb") as f:
        pdf_bytes = f.read()

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total = len(doc)
    doc.close()

    parts = []
    for i in range(total):
        # 1. Thử text layer
        text = extract_pdf_page_text(pdf_bytes, i)
        if text and len(text.strip()) >= 30:
            parts.append(text)
            continue
        # 2. OCR ảnh
        img_bytes = _render_pdf_page_to_image(pdf_bytes, i, dpi=150)
        if img_bytes:
            text = ocr_image_with_gemini(img_bytes, f"{bai_file}_p{i+1}")
            if text and not text.startswith("["):
                parts.append(text)

    return "\n\n---\n\n".join(parts)


# ============================================================
# PHẦN E: TÌM BÀI THEO TÊN (FUZZY MATCH)
# ============================================================

def _chuan_hoa_tim_kiem(s):
    """Chuẩn hóa chuỗi để so khớp: bỏ dấu, lowercase."""
    if not s:
        return ""
    s = str(s).strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.replace("đ", "d")
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    s = re.sub(r"\s+", " ", s)
    return s


def _do_tuong_dong(a, b):
    """Điểm tương đồng 0..1 giữa 2 chuỗi đã chuẩn hóa."""
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    if a in b or b in a:
        return 0.9
    ta = set(a.split())
    tb = set(b.split())
    if not ta or not tb:
        return 0.0
    inter = ta & tb
    union = ta | tb
    return len(inter) / len(union) if union else 0.0


def find_bai_by_name(sgk_name, ten_bai, threshold=0.45):
    """Tìm bài trong SGK theo TÊN BÀI (fuzzy match).

    Returns:
        dict bài (file, ten_bai, so_bai...) hoặc None.
    """
    meta = load_sgk_metadata(sgk_name)
    if not meta or not meta.get("ds_bai"):
        return None
    ten_can = _chuan_hoa_tim_kiem(ten_bai)
    if not ten_can:
        return None

    # 1. Khớp chính xác
    for b in meta["ds_bai"]:
        if _chuan_hoa_tim_kiem(b.get("ten_bai", "")) == ten_can:
            return b

    # 2. Khớp fuzzy
    best = None
    best_score = 0.0
    for b in meta["ds_bai"]:
        score = _do_tuong_dong(
            _chuan_hoa_tim_kiem(b.get("ten_bai", "")), ten_can)
        if score > best_score:
            best_score = score
            best = b

    if best and best_score >= threshold:
        return best
    return None


def find_bai_by_page(sgk_name, page_pdf):
    """Tìm bài chứa trang PDF (dùng nội bộ)."""
    meta = load_sgk_metadata(sgk_name)
    if not meta:
        return None
    for bai in meta.get("ds_bai", []):
        if bai.get("page_pdf_start", 0) <= page_pdf <= \
           bai.get("page_pdf_end", 0):
            return bai
    return None


def list_ten_bai(sgk_name):
    """Trả về list tên bài của 1 SGK (để hiển thị dropdown)."""
    meta = load_sgk_metadata(sgk_name)
    if not meta or not meta.get("ds_bai"):
        return []
    return [b.get("ten_bai", "") for b in meta["ds_bai"]]


# ============================================================
# PHẦN F: TIỆN ÍCH SGK
# ============================================================

def get_sgk_file_by_name(mon, lop):
    """Lấy file SGK_[Môn]_[Lớp] từ kho Nguồn."""
    if not mon or not lop:
        return (None, None)
    key = "SGK_" + _slugify(mon).title().replace("_", "") + "_" + str(lop)
    sources = st.session_state.get("sources", {})
    for fname in sources.keys():
        base = os.path.splitext(fname)[0]
        if base.lower() == key.lower():
            return (fname, sources[fname])
    for fname in sources.keys():
        base = os.path.splitext(fname)[0]
        if base.lower().startswith("sgk_") and str(lop) in base and \
           _slugify(mon) in _slugify(base):
            return (fname, sources[fname])
    return (None, None)


def suggest_sgk_name(mon, lop):
    """Đề xuất tên SGK chuẩn."""
    return "SGK_" + _slugify(mon).title().replace("_", "") + "_" + str(lop)


def render_danh_sach_bai(sgk_name, expanded=False):
    """Hiển thị danh sách bài của 1 SGK (STT + Tên bài)."""
    meta = load_sgk_metadata(sgk_name)
    if not meta or not meta.get("ds_bai"):
        st.info(f"SGK `{sgk_name}` chưa có bài nào.")
        return
    rows = []
    for b in meta["ds_bai"]:
        rows.append({
            "STT": b.get("so_bai", ""),
            "Tên bài": b.get("ten_bai", ""),
        })
    st.dataframe(rows, use_container_width=True, hide_index=True)


# ============ KẾT THÚC PHẦN 2/6 ============

# ============================================================
# PHẦN 3/6: RULES 12 MÔN + MẪU 5512 + PROMPT
# ============================================================
# Chức năng:
#   - SUBJECT_RULES: bộ quy tắc 12 môn THCS (có Tiếng Anh)
#   - MAU_5512_MAC_DINH: mẫu 5512 hardcode
#   - get_subject_rules, _chuan_hoa_ten_mon
#   - detect_end_of_chapter, detect_luyen_tap_chung
#   - kiem_tra_mau_5512, get_mau_dang_dung
#   - build_exec_prompt
# ============================================================

MAU_5512_MAC_DINH = {
    "ten_mau": "Mẫu CV 5512 chuẩn Bộ GD",
    "che_do": "mac_dinh",
    "nguon": "hardcode",
    "cau_truc": """TÊN BÀI DẠY: ...
Môn học: ...; lớp: ...
Thời gian thực hiện: (... tiết)

I. Mục tiêu
1. Kiến thức: ...
2. Năng lực:
   - Năng lực riêng: ...
   - Năng lực chung: ...
3. Phẩm chất: ...

II. Thiết bị dạy học và học liệu
1. Đối với giáo viên: ...
2. Đối với học sinh: ...

III. Tiến trình dạy học

1. Hoạt động 1: Mở đầu
a) Mục tiêu: ...
b) Nội dung: ...
c) Tổ chức thực hiện và sản phẩm:
[BẢNG 2 CỘT: Hoạt động của GV và HS | Sản phẩm]

2. Hoạt động 2: Giải quyết vấn đề từ Hoạt động 1
a) Mục tiêu: ...
b) Nội dung: ...
c) Tổ chức thực hiện và sản phẩm:
[BẢNG 2 CỘT: Hoạt động của GV và HS | Sản phẩm]

3. Hoạt động 3: Luyện tập
a) Mục tiêu: ...
b) Nội dung: ...
c) Tổ chức thực hiện và sản phẩm:
[BẢNG 2 CỘT: Hoạt động của GV và HS | Sản phẩm]

4. Hoạt động 4: Vận dụng
a) Mục tiêu: ...
b) Nội dung: ...
c) Tổ chức thực hiện và sản phẩm:
[BẢNG 2 CỘT: Hoạt động của GV và HS | Sản phẩm]

* HƯỚNG DẪN VỀ NHÀ
- ...
""",
}


SUBJECT_RULES = {
    "Toán": {
        "ten_goi_khac": ["toan", "math", "toán học"],
        "ky_hieu": "📐 π ∑ √",
        "nang_luc_rieng": [
            "Năng lực tư duy và lập luận toán học",
            "Năng lực giải quyết vấn đề toán học",
            "Năng lực giao tiếp toán học",
            "Năng lực mô hình hóa toán học",
            "Năng lực sử dụng công cụ, phương tiện học toán",
        ],
        "phuong_phap": ["Khám phá quy luật", "Giải quyết vấn đề",
                        "Luyện tập phân hóa"],
        "san_pham": ["Lời giải", "Phép tính, biểu thức", "Hình vẽ",
                     "Lập luận/chứng minh", "Mô hình toán học"],
        "danh_gia": ["Tính chính xác", "Lập luận logic",
                     "Phương pháp giải", "Vận dụng"],
        "khong_nen": ("Chỉ trình bày lý thuyết rồi cho bài tập mà "
                      "không thể hiện quá trình HS hình thành kiến thức."),
    },
    "Ngữ văn": {
        "ten_goi_khac": ["ngu van", "văn", "van", "ngữ văn", "literature"],
        "ky_hieu": "📖 ✍️ 🎭",
        "nang_luc_rieng": ["Năng lực ngôn ngữ", "Năng lực văn học",
                            "Năng lực đọc hiểu", "Năng lực viết",
                            "Năng lực nói và nghe"],
        "phuong_phap": ["Đọc hiểu", "Phân tích, thảo luận",
                        "Thực hành viết", "Đóng vai, trình bày"],
        "san_pham": ["Phiếu học tập đọc hiểu", "Sơ đồ phân tích",
                     "Đoạn văn/bài văn", "Bài thuyết trình"],
        "danh_gia": ["Đọc hiểu", "Phân tích chi tiết, nhân vật",
                     "Viết đoạn/bài văn", "Trình bày, thảo luận"],
        "khong_nen": ("Ép mọi bài vào quy trình phân tích tác phẩm hoặc "
                      "đưa bài văn mẫu để HS học thuộc."),
    },
    "Tiếng Anh": {
        "ten_goi_khac": ["tieng anh", "anh", "english", "ta",
                          "tiếng anh"],
        "ky_hieu": "🇬🇧 🔤 🗣️",
        "nang_luc_rieng": [
            "Năng lực giao tiếp bằng tiếng Anh",
            "Năng lực ngôn ngữ (ngữ âm, từ vựng, ngữ pháp)",
            "Năng lực liên văn hóa",
            "Năng lực tự chủ và tự học ngoại ngữ",
        ],
        "phuong_phap": [
            "Nghe - Nói - Đọc - Viết tích hợp",
            "Luyện tập theo cặp/nhóm",
            "Học qua tình huống giao tiếp",
            "Trò chơi ngôn ngữ, đóng vai",
            "Sử dụng học liệu số, nghe nhìn",
        ],
        "san_pham": [
            "Đoạn hội thoại theo chủ đề",
            "Bài nghe - nói (role-play)",
            "Đoạn văn viết theo gợi ý",
            "Phiếu học tập từ vựng - ngữ pháp",
            "Sản phẩm dự án (poster, video, presentation)",
        ],
        "danh_gia": [
            "Nghe hiểu thông tin chính",
            "Nói theo chủ đề, phát âm rõ",
            "Đọc hiểu văn bản ngắn",
            "Viết đoạn văn đơn giản",
            "Sử dụng từ vựng, ngữ pháp phù hợp",
        ],
        "khong_nen": (
            "Chỉ dạy ngữ pháp và dịch câu mà bỏ qua 4 kỹ năng Nghe - Nói "
            "- Đọc - Viết; hoặc bắt HS học thuộc lòng đoạn hội thoại "
            "mẫu mà không có hoạt động giao tiếp thực tế."
        ),
    },
    "Khoa học tự nhiên": {
        "ten_goi_khac": ["khtn", "khoa hoc tu nhien", "khoa học tự nhiên",
                          "science"],
        "ky_hieu": "🔬 🧪 ⚗️",
        "nang_luc_rieng": ["Năng lực nhận thức KHTN",
                            "Năng lực tìm hiểu tự nhiên",
                            "Năng lực vận dụng kiến thức, kỹ năng"],
        "phuong_phap": ["Quan sát", "Thí nghiệm, thực hành",
                        "Điều tra, khảo sát", "Phân tích dữ liệu"],
        "san_pham": ["Báo cáo thí nghiệm", "Bảng quan sát",
                     "Biểu đồ, sơ đồ", "Bản giải thích hiện tượng",
                     "Kết luận dựa trên bằng chứng"],
        "danh_gia": ["Hiểu khái niệm, quy luật", "Quan sát, đặt câu hỏi",
                     "Dự đoán, giả thuyết", "Thực hiện thí nghiệm",
                     "Phân tích dữ liệu, kết luận"],
        "khong_nen": ("Tạo thí nghiệm không có trong điều kiện thiết bị "
                      "hoặc kết luận không phù hợp dữ liệu."),
    },
    "Lịch sử": {
        "ten_goi_khac": ["lich su", "lịch sử", "history", "ls"],
        "ky_hieu": "🏛️ 📜 ⏳",
        "nang_luc_rieng": ["Năng lực tìm hiểu lịch sử",
                            "Năng lực nhận thức và tư duy lịch sử",
                            "Năng lực vận dụng kiến thức, kỹ năng"],
        "phuong_phap": ["Khai thác tư liệu", "Dòng thời gian",
                        "So sánh sự kiện",
                        "Phân tích nguyên nhân - diễn biến - kết quả"],
        "san_pham": ["Dòng thời gian", "Sơ đồ diễn biến", "Bảng so sánh",
                     "Phiếu khai thác tư liệu", "Bài thuyết trình"],
        "danh_gia": ["Xác định thời gian, không gian",
                     "Nhận biết sự kiện, nhân vật",
                     "Phân tích nguyên nhân, diễn biến, kết quả",
                     "So sánh sự kiện, giai đoạn"],
        "khong_nen": ("Bắt HS học thuộc toàn bộ diễn biến mà không phân "
                      "tích tư liệu, giải thích hoặc so sánh."),
    },
    "Địa lí": {
        "ten_goi_khac": ["dia li", "địa lí", "địa lý", "geography", "dl"],
        "ky_hieu": "🗺️ 🌍 🧭",
        "nang_luc_rieng": ["Năng lực nhận thức khoa học địa lí",
                            "Năng lực tìm hiểu địa lí",
                            "Năng lực vận dụng kiến thức, kỹ năng"],
        "phuong_phap": ["Đọc bản đồ, lược đồ",
                        "Phân tích biểu đồ, số liệu",
                        "So sánh vùng, quốc gia"],
        "san_pham": ["Bản đồ được chú giải", "Biểu đồ", "Bảng thống kê",
                     "Bài nhận xét số liệu", "Sơ đồ phân bố"],
        "danh_gia": ["Xác định vị trí địa lí",
                     "Nhận biết đặc điểm tự nhiên, dân cư, KT-XH",
                     "Đọc và sử dụng bản đồ",
                     "Phân tích biểu đồ, bảng số liệu",
                     "Giải thích sự phân bố"],
        "khong_nen": ("Chỉ yêu cầu đọc lại số liệu hoặc mô tả biểu đồ "
                      "mà không giải thích xu hướng, mối quan hệ địa lí."),
    },
    "Giáo dục công dân": {
        "ten_goi_khac": ["gdcd", "giao duc cong dan", "giáo dục công dân",
                          "civic"],
        "ky_hieu": "⚖️ 🤝 📋",
        "nang_luc_rieng": ["Năng lực điều chỉnh hành vi",
                            "Năng lực phát triển bản thân",
                            "Năng lực tìm hiểu, tham gia hoạt động KT-XH"],
        "phuong_phap": ["Phân tích tình huống", "Xử lý tình huống",
                        "Đóng vai", "Thảo luận nhóm"],
        "san_pham": ["Phiếu xử lý tình huống", "Bảng phân tích hành vi",
                     "Bài trình bày quan điểm", "Kịch bản đóng vai",
                     "Kế hoạch rèn luyện cá nhân"],
        "danh_gia": ["Hiểu khái niệm, chuẩn mực",
                     "Nhận biết quyền, nghĩa vụ, trách nhiệm",
                     "Phân tích hành vi đúng/sai",
                     "Lựa chọn cách ứng xử phù hợp"],
        "khong_nen": ("Chỉ đưa định nghĩa và yêu cầu học thuộc; "
                      "đánh giá phẩm chất HS chỉ dựa 1 câu trả lời."),
    },
    "Mỹ thuật": {
        "ten_goi_khac": ["my thuat", "mỹ thuật", "art"],
        "ky_hieu": "🎨 🖌️ 🖼️",
        "nang_luc_rieng": ["Năng lực quan sát, nhận thức thẩm mỹ",
                            "Năng lực sáng tạo, ứng dụng thẩm mỹ",
                            "Năng lực phân tích, đánh giá thẩm mỹ"],
        "phuong_phap": ["Quan sát", "Thực hành tạo hình", "Thiết kế",
                        "Trưng bày, nhận xét sản phẩm"],
        "san_pham": ["Tranh vẽ", "Sản phẩm tạo hình", "Bản phác thảo",
                     "Sản phẩm trang trí/thiết kế",
                     "Bài giới thiệu, nhận xét tác phẩm"],
        "danh_gia": ["Nhận biết yếu tố tạo hình",
                     "Quan sát, cảm nhận tác phẩm",
                     "Lựa chọn chất liệu, công cụ",
                     "Tạo sản phẩm theo yêu cầu",
                     "Trình bày ý tưởng, nhận xét"],
        "khong_nen": ("Đánh giá chỉ dựa trên đẹp/xấu hoặc yêu cầu mọi HS "
                      "tạo sản phẩm giống hệt nhau."),
    },
    "Âm nhạc": {
        "ten_goi_khac": ["am nhac", "âm nhạc", "music"],
        "ky_hieu": "🎵 🎼 🎤",
        "nang_luc_rieng": ["Năng lực thể hiện âm nhạc",
                            "Năng lực cảm thụ, hiểu biết âm nhạc",
                            "Năng lực ứng dụng, sáng tạo âm nhạc"],
        "phuong_phap": ["Nghe mẫu", "Luyện hát, đọc nhạc",
                        "Thực hành nhạc cụ", "Trò chơi tiết tấu",
                        "Biểu diễn"],
        "san_pham": ["Phần trình diễn bài hát", "Bài đọc nhạc",
                     "Phần biểu diễn nhạc cụ",
                     "Tiết tấu/đoạn nhạc sáng tạo",
                     "Bài cảm nhận âm nhạc"],
        "danh_gia": ["Nhận biết đặc điểm tác phẩm",
                     "Hát đúng giai điệu, lời ca, nhịp điệu",
                     "Nghe và cảm nhận",
                     "Đọc nhạc hoặc sử dụng nhạc cụ",
                     "Hợp tác trong biểu diễn"],
        "khong_nen": ("Chỉ trình bày kiến thức tác phẩm mà bỏ qua hoạt "
                      "động nghe, thực hành, thể hiện âm nhạc."),
    },
    "Giáo dục thể chất": {
        "ten_goi_khac": ["gdtc", "giao duc the chat",
                          "giáo dục thể chất", "pe", "thể dục"],
        "ky_hieu": "⚽ 🏃 🏀",
        "nang_luc_rieng": ["Năng lực vận động cơ bản",
                            "Năng lực thể thao",
                            "Năng lực thể chất và sức khỏe"],
        "phuong_phap": ["Làm mẫu", "Luyện tập", "Sửa lỗi kỹ thuật",
                        "Trò chơi vận động"],
        "san_pham": ["Động tác được thực hiện", "Chuỗi vận động",
                     "Kết quả thực hành", "Bảng theo dõi tiến bộ",
                     "Kết quả trò chơi vận động"],
        "danh_gia": ["Hiểu yêu cầu, tác dụng động tác",
                     "Thực hiện đúng kỹ thuật",
                     "Phát triển tố chất thể lực",
                     "Tự tập luyện, điều chỉnh vận động",
                     "Ý thức an toàn, kỷ luật"],
        "khong_nen": ("Soạn giáo án chỉ có lời giảng, câu hỏi và bài "
                      "tập viết cho nội dung cần thực hành vận động."),
    },
    "Hoạt động trải nghiệm, hướng nghiệp": {
        "ten_goi_khac": ["hdtn", "hoat dong trai nghiem", "hướng nghiệp",
                          "trai nghiem", "experiential"],
        "ky_hieu": "🧭 🌱 🎯",
        "nang_luc_rieng": ["Năng lực thích ứng với cuộc sống",
                            "Năng lực thiết kế, tổ chức hoạt động",
                            "Năng lực định hướng nghề nghiệp"],
        "phuong_phap": ["Trải nghiệm thực tế", "Dự án", "Khảo sát",
                        "Thảo luận nhóm", "Phản hồi, tự đánh giá"],
        "san_pham": ["Nhật ký trải nghiệm", "Kế hoạch cá nhân",
                     "Báo cáo khảo sát", "Poster/video",
                     "Sản phẩm dự án", "Bài trình bày kết quả"],
        "danh_gia": ["Nhận thức bản thân", "Giao tiếp và hợp tác",
                     "Giải quyết vấn đề đời sống",
                     "Trách nhiệm với gia đình, nhà trường, cộng đồng",
                     "Tìm hiểu nghề nghiệp"],
        "khong_nen": ("Biến mọi tiết trải nghiệm thành tiết GV thuyết "
                      "trình về kỹ năng sống."),
    },
    "Tin học": {
        "ten_goi_khac": ["tin hoc", "tin học", "informatics", "it",
                          "cntt"],
        "ky_hieu": "💻 🖥️ ⌨️",
        "nang_luc_rieng": ["Năng lực sử dụng, quản lý phương tiện CNTT",
                            "Năng lực ứng dụng CNTT trong học và tự học",
                            "Năng lực giải quyết vấn đề với CNTT",
                            "Năng lực an toàn, an ninh thông tin"],
        "phuong_phap": ["Thực hành", "Giải quyết vấn đề",
                        "Thử nghiệm", "Kiểm thử, sửa lỗi"],
        "san_pham": ["Tệp văn bản/bảng tính", "Bài trình chiếu",
                     "Sơ đồ thuật toán", "Chương trình máy tính",
                     "Kết quả chạy thử", "Sản phẩm số"],
        "danh_gia": ["Hiểu khái niệm tin học", "Sử dụng công cụ số",
                     "Thực hiện thao tác máy tính",
                     "Phân tích vấn đề, xây dựng thuật toán",
                     "Viết, kiểm thử, sửa chương trình"],
        "khong_nen": ("Chỉ cho HS chép mã chương trình mà không giải "
                      "thích bài toán, thuật toán và cách kiểm tra."),
    },
}


def _chuan_hoa_ten_mon(ten_mon):
    """Chuẩn hóa tên môn: bỏ dấu, lowercase."""
    if not ten_mon:
        return ""
    ten_mon = ten_mon.strip().lower()
    ten_mon = unicodedata.normalize("NFD", ten_mon)
    ten_mon = "".join(c for c in ten_mon
                      if unicodedata.category(c) != "Mn")
    ten_mon = ten_mon.replace("đ", "d").replace("Đ", "D")
    return ten_mon


def get_subject_rules(ten_mon):
    """Tra bộ quy tắc theo tên môn.
    Linh hoạt với viết hoa, dấu, viết tắt.
    Trả về (tên_chuẩn, rules) hoặc (None, None).
    """
    if not ten_mon:
        return (None, None)
    ten_chuan = _chuan_hoa_ten_mon(ten_mon)

    # 1. Khớp chính xác tên môn
    for key, rules in SUBJECT_RULES.items():
        if _chuan_hoa_ten_mon(key) == ten_chuan:
            return (key, rules)

    # 2. Khớp tên gọi khác
    for key, rules in SUBJECT_RULES.items():
        for alias in rules.get("ten_goi_khac", []):
            if _chuan_hoa_ten_mon(alias) == ten_chuan:
                return (key, rules)

    # 3. Khớp mờ
    for key, rules in SUBJECT_RULES.items():
        key_chuan = _chuan_hoa_ten_mon(key)
        if key_chuan in ten_chuan or ten_chuan in key_chuan:
            return (key, rules)
        for alias in rules.get("ten_goi_khac", []):
            alias_chuan = _chuan_hoa_ten_mon(alias)
            if len(alias_chuan) >= 3 and (alias_chuan in ten_chuan or
                                          ten_chuan in alias_chuan):
                return (key, rules)

    return (None, None)


def detect_end_of_chapter(topic):
    """Nhận diện bài cuối chương qua từ khóa."""
    if not topic:
        return False
    t = topic.lower()
    keywords = [
        "bài tập cuối chương", "ôn tập chương", "ôn tập học kì",
        "củng cố, mở rộng", "tổng kết chương", "review",
        "looking back", "trưng bày sản phẩm", "bài tập ôn chương",
        "ôn tập cuối năm", "tổng kết năm học",
    ]
    for kw in keywords:
        if kw in t:
            return True
    return False


def detect_luyen_tap_chung(topic):
    """Nhận diện bài luyện tập chung."""
    if not topic:
        return False
    return "luyện tập chung" in topic.lower()


def kiem_tra_mau_5512(mau_text):
    """Kiểm tra mẫu có đủ cấu trúc 5512 không.
    Trả về (ok, msg).
    """
    if not mau_text:
        return (False, "Mẫu rỗng")

    text_lower = mau_text.lower()
    errors = []

    if "mục tiêu" not in text_lower:
        errors.append("Thiếu 'I. Mục tiêu'")
    if "thiết bị dạy học" not in text_lower and \
       "thiết bị" not in text_lower:
        errors.append("Thiếu 'II. Thiết bị dạy học'")
    if "tiến trình dạy học" not in text_lower and \
       "tiến trình" not in text_lower:
        errors.append("Thiếu 'III. Tiến trình dạy học'")

    for i in [1, 2, 3, 4]:
        if ("hoạt động " + str(i)) not in text_lower:
            errors.append("Thiếu 'Hoạt động " + str(i) + "'")

    if "hoạt động của gv" not in text_lower and \
       "hoạt động của giáo viên" not in text_lower:
        errors.append("Thiếu bảng 'Hoạt động của GV và HS'")

    if errors:
        return (False, " | ".join(errors))
    return (True, "OK")


def get_mau_dang_dung():
    """Lấy mẫu 5512 đang dùng (hardcode hoặc user nạp)."""
    mau = st.session_state.get("mau_5512_dang_dung")
    if mau:
        return mau
    return MAU_5512_MAC_DINH


def build_exec_prompt(plan):
    """Xây prompt 5512 tối ưu token."""
    mem = st.session_state["memory"]

    # 1. Tra môn
    ten_mon_input = plan.get("subject", "") or \
        mem.get("default_subject", "")
    matched_key, rules = get_subject_rules(ten_mon_input)

    if matched_key:
        ten_mon_chuan = matched_key
    else:
        ten_mon_chuan = ten_mon_input or "Chung"
        rules = {
            "ky_hieu": "📚",
            "nang_luc_rieng": ["Năng lực chung", "Năng lực chuyên môn"],
            "phuong_phap": ["Thuyết trình", "Thảo luận", "Thực hành"],
            "san_pham": ["Bài trình bày", "Sản phẩm thực hành"],
            "danh_gia": ["Kiến thức", "Kỹ năng", "Thái độ"],
            "khong_nen": "",
        }

    # 2. Lấy mẫu 5512
    mau = get_mau_dang_dung()
    mau_text = mau.get("cau_truc", "")[:2000]

    # 3. Lấy đoạn SGK (nếu có)
    sgk_text = plan.get("sgk_excerpt", "")[:8000]

    # 4. Xác định HĐ2
    topic = plan.get("topic", "")
    if detect_end_of_chapter(topic):
        hd2_title = "Hệ thống kiến thức"
    elif detect_luyen_tap_chung(topic):
        hd2_title = "Luyện tập"
    else:
        hd2_title = "Giải quyết vấn đề từ Hoạt động 1"

    # 5. Xây prompt
    p = []
    p.append("Bạn là trợ lý soạn KẾ HOẠCH BÀI DẠY THCS chuẩn CV 5512.")
    p.append("")
    p.append("=== THÔNG TIN BÀI DẠY ===")
    p.append("- GV: " + mem.get("teacher_name", "") +
             " | Trường: " + mem.get("school", ""))
    p.append("- Môn: " + ten_mon_chuan + " | Khối: " +
             str(plan.get("grade", "6")))
    p.append("- Tên bài: " + topic)
    p.append("- Số tiết: " + str(plan.get("period", "1")))
    p.append("")

    # Đặc trưng môn
    p.append("=== ĐẶC TRƯNG MÔN " + ten_mon_chuan.upper() + " ===")
    p.append("Ký hiệu: " + rules.get("ky_hieu", ""))
    p.append("Năng lực riêng: " +
             " | ".join(rules.get("nang_luc_rieng", [])))
    p.append("Phương pháp: " +
             " | ".join(rules.get("phuong_phap", [])))
    p.append("Sản phẩm: " + " | ".join(rules.get("san_pham", [])))
    p.append("Đánh giá: " + " | ".join(rules.get("danh_gia", [])))
    if rules.get("khong_nen"):
        p.append("KHÔNG NÊN: " + rules["khong_nen"])
    p.append("")

    # Nguồn SGK
    if sgk_text:
        p.append("=== NGUỒN SGK (BẮT BUỘC DÙNG) ===")
        p.append(sgk_text)
        p.append("")

    # Mẫu 5512
    p.append("=== CẤU TRÚC 5512 (BẮT BUỘC) ===")
    p.append(mau_text)
    p.append("")

    # Tiêu đề bài dạy
    p.append("=== TIÊU ĐỀ BÀI DẠY (DÙNG ĐÚNG FORMAT) ===")
    tieu_de = topic
    tieu_de += " (" + str(plan.get("period", "1")) + " tiết)"
    p.append("**TÊN BÀI DẠY: " + tieu_de + "**")
    p.append("**Môn học: " + ten_mon_chuan + "; lớp: " +
             str(plan.get("grade", "6")) + "**")
    p.append("**Thời gian thực hiện: (" +
             str(plan.get("period", "1")) + " tiết)**")
    p.append("")

    # HĐ2 title
    p.append("=== HOẠT ĐỘNG 2 PHẢI ĐẶT TÊN: ===")
    p.append("**2. Hoạt động 2: " + hd2_title + "**")
    p.append("")

    # Quy tắc nghiêm ngặt
    p.append("=== QUY TẮC BẮT BUỘC ===")
    p.append("1. Mục I (Mục tiêu):")
    p.append("   1. Kiến thức: ...")
    p.append("   2. Năng lực:")
    p.append("      - Năng lực riêng: ...")
    p.append("      - Năng lực chung: ...")
    p.append("   3. Phẩm chất: ...")
    p.append("")
    p.append("2. Mỗi hoạt động có: a) Mục tiêu; b) Nội dung; "
             "c) Tổ chức thực hiện và sản phẩm")
    p.append("")
    p.append("3. Mục c) là BẢNG 2 CỘT: "
             "| Hoạt động của GV và HS | Sản phẩm |")
    p.append("   - Bảng có 4 dòng Bước 1, Bước 2, Bước 3, Bước 4")
    p.append("   - Cột Sản phẩm phải CHI TIẾT, bám sát SGK")
    p.append("")
    p.append("4. Cột Sản phẩm mỗi bước phải có:")
    p.append("   - Kiến thức SGK")
    p.append("   - Các bước giải (nếu có)")
    p.append("   - Ký hiệu toán/đặc trưng môn")
    p.append("   - Mô tả hình ảnh (nếu có)")
    p.append("   - Dự kiến câu trả lời của HS")
    p.append("")
    p.append("5. TUYỆT ĐỐI KHÔNG BỊA kiến thức không có trong SGK.")
    p.append("   Nếu SGK không có → ghi '[SGK không cung cấp]'.")
    p.append("")
    p.append("6. Trong ô bảng dùng <br> để xuống dòng. "
             "Mỗi bước riêng dòng.")
    p.append("")
    p.append("7. Ký hiệu Unicode: x², √, ≤, ≥, π, ∈.")
    p.append("")
    p.append("8. KHÔNG viết 'Người soạn', 'Ký tên', 'Sông Đốc' cuối bài.")
    p.append("")
    p.append("9. Hướng dẫn về nhà cuối bài gồm 3 gạch đầu dòng:")
    p.append("   - Đọc và ghi nhớ nội dung chính")
    p.append("   - Làm bài tập SGK")
    p.append("   - Chuẩn bị bài mới")
    p.append("")

    return "\n".join(p)


# ============ KẾT THÚC PHẦN 3/6 ============

# ============================================================
# PHẦN 4/6: SIDEBAR + 9 TAB + CALL AI
# ============================================================
# Chức năng:
#   - Sidebar: test GitHub, memory, thống kê, dọn dẹp
#   - 9 tab: SGK gốc (có nút tải file đã tách),
#            bài thủ công, kho, mẫu GA, ma trận, đặc tả, đề,
#            sửa tên bài, khớp tên BG
#   - Hàm call_ai() đa provider + fallback
# ============================================================

with st.sidebar:
    st.header("🧠 Bộ não")
    mem = st.session_state["memory"]

    if GITHUB_TOKEN:
        st.success("💾 Đã có GITHUB_TOKEN")
        if st.button("🧪 Test GitHub", use_container_width=True):
            ok, msg = gh_test()
            st.success("✅ " + msg) if ok else st.error("❌ " + msg)
    else:
        st.error("❌ Chưa có GITHUB_TOKEN")

    if GEMINI_KEY:
        st.success("🔍 Gemini Vision sẵn sàng")
    else:
        st.warning("⚠️ Chưa có GEMINI_API_KEY")

    mem["teacher_name"] = st.text_input(
        "Giáo viên", value=mem.get("teacher_name", ""))
    mem["school"] = st.text_input(
        "Trường", value=mem.get("school", ""))
    mem["default_subject"] = st.text_input(
        "Môn mặc định", value=mem.get("default_subject", "Toán"))
    mem["default_grade"] = st.text_input(
        "Khối mặc định", value=mem.get("default_grade", "6"))

    if st.button("💾 Lưu vào bộ nhớ", use_container_width=True):
        ok, msg = save_memory(mem)
        st.success("✅ " + msg) if ok else st.error("❌ " + msg)

    st.divider()
    st.subheader("📊 Thống kê")
    for label, key in [("Mẫu", "templates"), ("Nguồn", "sources"),
                       ("Lịch BG", "schedule")]:
        st.write(label + ": " + str(len(st.session_state[key])) + " file")

    ds_sgk = list_sgk_da_tach()
    st.write("SGK đã tách: " + str(len(ds_sgk)) + " file")
    if ds_sgk:
        with st.expander("📚 Xem SGK đã tách", expanded=False):
            for sgk in ds_sgk:
                st.markdown(f"**{sgk['sgk_name']}** — "
                            f"{sgk['so_bai']} bài ({sgk['nguon']})")
                meta = load_sgk_metadata(sgk["sgk_name"])
                if meta and meta.get("ds_bai"):
                    for b in meta["ds_bai"]:
                        st.caption(
                            f"  • Bài {b.get('so_bai', '?')}: "
                            f"{b.get('ten_bai', '?')}")
                st.markdown("---")

    if st.button("🔄 Tải lại", use_container_width=True):
        st.session_state["templates"] = load_templates()
        st.session_state["sources"] = load_sources()
        st.session_state["schedule"] = load_schedule()
        st.rerun()

    st.write("AI: " + str(len(available_providers)) + " nhà cung cấp")
    st.divider()

    st.subheader("🧹 Dọn dẹp")
    with st.expander("🗑️ Dọn rác và cache", expanded=False):
        st.caption("Chọn mục cần dọn:")
        if st.button("🧼 Xóa phiên làm việc tạm",
                     use_container_width=True, key="clean_session"):
            for k in ["draft", "last_topic", "weekly_results",
                      "lich_bao_giang", "ocr_session_cache"]:
                if k in st.session_state:
                    del st.session_state[k]
            st.success("✅ Đã xóa phiên làm việc tạm")
            st.rerun()

        if st.button("♻️ Xóa cache Streamlit",
                     use_container_width=True, key="clean_cache"):
            st.cache_data.clear()
            st.cache_resource.clear()
            st.success("✅ Đã xóa cache Streamlit")
            st.rerun()

        if st.button("🗑️ Xóa cache OCR local",
                     use_container_width=True, key="clean_ocr"):
            deleted = 0
            if os.path.exists(OCR_CACHE_DIR):
                for f in os.listdir(OCR_CACHE_DIR):
                    try:
                        os.remove(os.path.join(OCR_CACHE_DIR, f))
                        deleted += 1
                    except Exception:
                        pass
            st.session_state["ocr_session_cache"] = {}
            st.success(f"✅ Đã xóa {deleted} file OCR cache")
            st.rerun()

        if st.button("🧹 Xóa ma trận + đặc tả + đề",
                     use_container_width=True, key="clean_exam"):
            _deleted = 0
            for k in list(st.session_state.keys()):
                if k.startswith("matran_") or k in [
                    "dacta_hientai", "de_kiem_tra_hientai",
                    "matran_hientai", "mt_saved_sel", "dt_mt_sel"]:
                    del st.session_state[k]
                    _deleted += 1
            st.success(f"✅ Đã xóa {_deleted} mục.")
            st.rerun()

        if GITHUB_TOKEN:
            if st.button("📜 Xóa lịch sử soạn bài",
                         use_container_width=True, key="clean_history"):
                files = gh_list("data/history")
                deleted = 0
                for f in files:
                    ok, _ = gh_delete(f"data/history/{f}")
                    if ok:
                        deleted += 1
                st.success(f"✅ Đã xóa {deleted} file lịch sử")

    st.divider()
    st.caption("⚠️ **Cảnh báo**: các nút dưới đây xóa VĨNH VIỄN")
    if st.button("🔴 Reset toàn bộ app",
                 use_container_width=True, key="reset_all"):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.success("✅ Đã reset. Đang tải lại...")
        st.rerun()


def call_ai(prompt, system=""):
    """Gọi AI đa provider với fallback."""
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
                t = getattr(resp, "text", None) or \
                    getattr(resp, "content", "")
                if t and len(t.strip()) > 50:
                    return t
            except Exception as e:
                last_err = e
                continue
    try:
        resp = llm.chat(msgs, model="auto")
        return getattr(resp, "text", None) or \
            getattr(resp, "content", "")
    except Exception as e:
        if last_err:
            raise last_err
        raise e


def call_ai_cached(prompt, system=""):
    """Gọi AI có cache theo hash prompt."""
    cache_key = hashlib.md5(
        (system + "||" + prompt).encode("utf-8")).hexdigest()
    cache_store = st.session_state.setdefault("ai_prompt_cache", {})
    if cache_key in cache_store:
        return cache_store[cache_key]
    result = call_ai(prompt, system)
    if result:
        cache_store[cache_key] = result
    return result


# ============================================================
# 9 TAB
# ============================================================
(tab_sgk, tab_manual, tab_kho, tab_mau,
 tab_matrix, tab_despec, tab_exam, tab_edit, tab_match) = st.tabs([
    "📤 Upload SGK gốc",
    "📚 Nạp bài đã tách",
    "📂 Kho tri thức",
    "📐 Nạp mẫu giáo án",
    "🧩 Nạp ma trận đề",
    "📋 Tạo đặc tả đề",
    "📝 Tạo đề + đáp án",
    "✏️ Sửa tên bài SGK",
    "📝 Khớp tên với Lịch BG",
])

# ============================================================
# TAB 1: UPLOAD SGK GỐC + TẢI FILE ĐÃ TÁCH
# ============================================================
with tab_sgk:
    st.markdown("### 📤 Upload SGK gốc (PDF)")
    st.info("💡 Upload file SGK PDF (scan ảnh). App sẽ **tách tự động "
            "theo bài**. Quét từng trang nên tốn token hơn nhưng chính xác.")

    # Xem danh sách SGK đã tách
    _ds_check = list_sgk_da_tach()
    if _ds_check:
        with st.expander(f"📖 Xem {len(_ds_check)} SGK đã tách + tên bài",
                         expanded=False):
            for _sgk in _ds_check:
                st.markdown(f"#### 📘 {_sgk['sgk_name']}")
                render_danh_sach_bai(_sgk["sgk_name"])

    sgk_mon = st.text_input(
        "Môn học (VD: Toán, Tiếng Anh...)",
        value=st.session_state["memory"].get("default_subject", "Toán"),
        key="up_sgk_mon")
    sgk_lop = st.text_input(
        "Lớp (VD: 6, 7, 8, 9)",
        value=st.session_state["memory"].get("default_grade", "6"),
        key="up_sgk_lop")
    sgk_file = st.file_uploader("Chọn file SGK PDF", type=["pdf"],
                                key="up_sgk_pdf")

    if sgk_file:
        sgk_name_goi_y = suggest_sgk_name(sgk_mon, sgk_lop)
        st.caption(f"📝 Tên SGK sẽ lưu: **{sgk_name_goi_y}**")
        existing = load_sgk_metadata(sgk_name_goi_y)
        if existing:
            st.warning(f"⚠️ SGK này đã có {existing.get('so_bai', 0)} "
                       "bài. Upload sẽ ghi đè.")
            ghi_de = st.checkbox("Xác nhận ghi đè", key="overwrite_sgk")
        else:
            ghi_de = True

        if st.button("🚀 Tách SGK tự động", type="primary",
                     use_container_width=True, key="btn_split_sgk"):
            if not HAS_PYMUPDF:
                st.error("❌ Chưa cài pymupdf.")
            elif not GEMINI_KEY:
                st.error("❌ Cần GEMINI_API_KEY để OCR.")
            elif not ghi_de:
                st.warning("⚠️ Bạn chưa xác nhận ghi đè.")
            else:
                pdf_bytes = sgk_file.read()
                total_pages = get_pdf_total_pages(pdf_bytes)
                st.info(f"📄 File có **{total_pages} trang**")
                progress = st.progress(0)
                status = st.empty()
                status.write("🔍 Đang quét SGK để tìm bài...")

                def _cb_scan(percent, msg):
                    progress.progress(min(95, percent))
                    status.write(msg)

                boundaries = detect_chapter_boundaries(
                    pdf_bytes, progress_cb=_cb_scan)
                progress.progress(100)

                if not boundaries:
                    st.warning(
                        "⚠️ Không tìm được bài tự động. "
                        "Bạn có thể dùng tab **📚 Nạp bài đã tách** "
                        "hoặc **✏️ Sửa tên bài SGK** để xử lý thủ công."
                    )
                else:
                    st.success(f"✅ Tìm được {len(boundaries)} bài")
                    with st.expander("👁️ Xem danh sách bài",
                                     expanded=True):
                        for b in boundaries:
                            st.write(
                                f"• **Bài {b['so_bai']}**: "
                                f"{b['ten_bai']} "
                                f"(trang {b['page_pdf_start']}"
                                f"–{b['page_pdf_end']})"
                            )

                    status.write("✂️ Đang tách thành file bài riêng...")
                    progress.progress(50)

                    def _cb(percent, msg):
                        progress.progress(50 + int(percent * 0.4))
                        status.write(msg)

                    metadata = split_sgk_into_lessons(
                        pdf_bytes, sgk_name_goi_y, boundaries, _cb)
                    progress.progress(100)
                    status.write("✅ Hoàn tất!")
                    if metadata:
                        st.success(
                            f"✅ Đã tách {metadata['so_bai']} bài. "
                            "Kiểm tra tên ở tab **✏️ Sửa tên bài SGK**."
                        )
                        st.balloons()
                        st.session_state["_vua_tach_sgk"] = sgk_name_goi_y
                    else:
                        st.error("❌ Tách thất bại.")

    # ---------- TẢI FILE VỪA TÁCH ----------
    st.markdown("---")
    st.markdown("### 📦 Tải file bài đã tách")

    _ds_tai = list_sgk_da_tach()
    if not _ds_tai:
        st.info("Chưa có SGK nào được tách. Upload SGK ở trên để tách.")
    else:
        _sgk_tai = st.selectbox(
            "Chọn SGK để tải",
            [s["sgk_name"] for s in _ds_tai],
            key="tai_sgk_sel")
        _meta_tai = load_sgk_metadata(_sgk_tai)
        if not _meta_tai or not _meta_tai.get("ds_bai"):
            st.warning("SGK này chưa có bài nào.")
        else:
            _ds_bai_tai = _meta_tai["ds_bai"]
            st.caption(f"📚 **{_sgk_tai}** — {len(_ds_bai_tai)} bài")

            # Tải riêng từng bài
            st.markdown("#### 📄 Tải riêng từng bài")
            for _i, _b in enumerate(_ds_bai_tai):
                _c1, _c2 = st.columns([4, 1])
                with _c1:
                    st.write(
                        f"**Bài {_b.get('so_bai', _i+1)}**: "
                        f"{_b.get('ten_bai', '')}"
                    )
                with _c2:
                    _fpath_bai = os.path.join(
                        SGK_BAI_DIR, _sgk_tai, _b.get("file", ""))
                    if os.path.exists(_fpath_bai):
                        with open(_fpath_bai, "rb") as _f:
                            _bytes_bai = _f.read()
                        st.download_button(
                            "⬇️ PDF",
                            _bytes_bai,
                            file_name=_b.get("file", f"bai_{_i+1}.pdf"),
                            mime="application/pdf",
                            key=f"dl_bai_{_i}",
                            use_container_width=True,
                        )
                    else:
                        st.button("❌ Mất file",
                                  key=f"dl_bai_missing_{_i}",
                                  disabled=True,
                                  use_container_width=True)

            # Tải tất cả dạng ZIP
            st.markdown("#### 📦 Tải toàn bộ (ZIP)")
            import zipfile
            _zip_buffer = io.BytesIO()
            _count = 0
            with zipfile.ZipFile(_zip_buffer, "w",
                                 zipfile.ZIP_DEFLATED) as _zf:
                for _b in _ds_bai_tai:
                    _fpath_b = os.path.join(
                        SGK_BAI_DIR, _sgk_tai, _b.get("file", ""))
                    if os.path.exists(_fpath_b):
                        _arcname = _b.get("file", "")
                        _zf.write(_fpath_b, arcname=_arcname)
                        _count += 1
                _meta_path = os.path.join(
                    SGK_BAI_DIR, _sgk_tai, "metadata.json")
                if os.path.exists(_meta_path):
                    _zf.write(_meta_path, arcname="metadata.json")

            _zip_buffer.seek(0)
            st.download_button(
                f"⬇️ Tải tất cả ({_count} file PDF + metadata.json)",
                _zip_buffer.getvalue(),
                file_name=f"{_sgk_tai}_da_tach.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
                key=f"dl_zip_{_sgk_tai}",
            )

            # Tải metadata.json
            _meta_path2 = os.path.join(
                SGK_BAI_DIR, _sgk_tai, "metadata.json")
            if os.path.exists(_meta_path2):
                with open(_meta_path2, "rb") as _f:
                    _meta_bytes = _f.read()
                st.download_button(
                    "⬇️ Tải metadata.json (danh sách bài)",
                    _meta_bytes,
                    file_name=f"{_sgk_tai}_metadata.json",
                    mime="application/json",
                    use_container_width=True,
                    key=f"dl_meta_{_sgk_tai}",
                )

# ============================================================
# TAB 2: NẠP BÀI THỦ CÔNG
# ============================================================
with tab_manual:
    st.markdown("### 📚 Nạp bài đã tách thủ công")
    st.info("💡 Tự tách SGK thành từng file PDF bài riêng → upload lên đây. "
            "Có thể upload file ZIP từ tab **📤 Upload SGK gốc**.")

    # Hỗ trợ nạp ZIP
    _c_nhap1, _c_nhap2 = st.columns([1, 1])
    with _c_nhap1:
        st.markdown("**Cách 1: Nạp nhiều file PDF**")
        sgk_manual_mon = st.text_input(
            "Môn học",
            value=st.session_state["memory"].get("default_subject", "Toán"),
            key="manual_sgk_mon")
        sgk_manual_lop = st.text_input(
            "Lớp",
            value=st.session_state["memory"].get("default_grade", "6"),
            key="manual_sgk_lop")
        sgk_manual_name = suggest_sgk_name(sgk_manual_mon, sgk_manual_lop)
        st.caption(f"📝 SGK đích: **{sgk_manual_name}**")

        manual_files = st.file_uploader(
            "Upload nhiều file PDF bài", type=["pdf"],
            accept_multiple_files=True, key="manual_sgk_files")

        if manual_files:
            st.markdown("**📋 Nhập tên bài cho mỗi file:**")
            manual_info = []
            for i, f in enumerate(manual_files):
                with st.expander("📄 " + f.name, expanded=False):
                    ten = st.text_input(
                        "Tên bài",
                        value=f.name.replace(".pdf", ""),
                        key=f"manual_ten_{i}")
                    manual_info.append({"file_obj": f, "ten_bai": ten})

            if st.button("💾 Lưu bài đã tách", type="primary",
                         use_container_width=True, key="btn_save_manual"):
                sgk_dir = os.path.join(SGK_BAI_DIR, sgk_manual_name)
                os.makedirs(sgk_dir, exist_ok=True)
                meta = load_sgk_metadata(sgk_manual_name) or {
                    "sgk_name": sgk_manual_name, "nguon": "manual",
                    "so_bai": 0, "ds_bai": []}
                saved = 0
                for i, info in enumerate(manual_info):
                    safe = _slugify(info["ten_bai"])[:40] or \
                        f"bai_{i+1:02d}"
                    fname = f"bai_{i+1:02d}_{safe}.pdf"
                    fpath = os.path.join(sgk_dir, fname)
                    try:
                        with open(fpath, "wb") as fo:
                            fo.write(info["file_obj"].read())
                    except Exception:
                        continue
                    meta["ds_bai"].append({
                        "file": fname,
                        "ten_bai": info["ten_bai"],
                        "so_bai": i + 1,
                        "page_pdf_start": 1,
                        "page_pdf_end": 1})
                    saved += 1
                meta["so_bai"] = len(meta["ds_bai"])
                meta["nguon"] = "manual"
                with open(os.path.join(sgk_dir, "metadata.json"), "w",
                          encoding="utf-8") as fo:
                    json.dump(meta, fo, ensure_ascii=False, indent=2)
                st.success(f"✅ Đã nạp {saved} bài vào {sgk_manual_name}")
                st.balloons()

    with _c_nhap2:
        st.markdown("**Cách 2: Nạp file ZIP đã tải**")
        st.caption("Upload file ZIP từ tab **📤 Upload SGK gốc** → "
                   "giải nén tự động và nạp vào cache.")
        sgk_zip_mon = st.text_input(
            "Môn học (ZIP)",
            value=st.session_state["memory"].get("default_subject", "Toán"),
            key="zip_sgk_mon")
        sgk_zip_lop = st.text_input(
            "Lớp (ZIP)",
            value=st.session_state["memory"].get("default_grade", "6"),
            key="zip_sgk_lop")
        sgk_zip_name = suggest_sgk_name(sgk_zip_mon, sgk_zip_lop)
        st.caption(f"📝 SGK đích: **{sgk_zip_name}**")

        zip_file = st.file_uploader(
            "Upload file ZIP đã tách", type=["zip"],
            key="zip_sgk_file")

        if zip_file:
            if st.button("📦 Giải nén và nạp ZIP", type="primary",
                         use_container_width=True, key="btn_unzip"):
                try:
                    import zipfile
                    sgk_dir = os.path.join(SGK_BAI_DIR, sgk_zip_name)
                    os.makedirs(sgk_dir, exist_ok=True)
                    with zipfile.ZipFile(io.BytesIO(zip_file.read()),
                                         "r") as _zf:
                        _zf.extractall(sgk_dir)
                    # Đọc metadata nếu có
                    _meta_zip = load_sgk_metadata(sgk_zip_name)
                    if _meta_zip:
                        st.success(
                            f"✅ Đã giải nén: {_meta_zip.get('so_bai', 0)} "
                            f"bài vào `{sgk_zip_name}`."
                        )
                    else:
                        # Tạo metadata từ các file PDF có trong ZIP
                        _ds_pdf = sorted([
                            f for f in os.listdir(sgk_dir)
                            if f.lower().endswith(".pdf")
                        ])
                        _meta_new = {
                            "sgk_name": sgk_zip_name,
                            "nguon": "zip",
                            "so_bai": len(_ds_pdf),
                            "ds_bai": []
                        }
                        for _i, _f in enumerate(_ds_pdf):
                            _meta_new["ds_bai"].append({
                                "file": _f,
                                "ten_bai": _f.replace(".pdf", ""),
                                "so_bai": _i + 1,
                                "page_pdf_start": 1,
                                "page_pdf_end": 1,
                            })
                        with open(os.path.join(sgk_dir, "metadata.json"),
                                  "w", encoding="utf-8") as _fo:
                            json.dump(_meta_new, _fo,
                                      ensure_ascii=False, indent=2)
                        st.success(
                            f"✅ Đã giải nén: {len(_ds_pdf)} file PDF."
                        )
                    st.balloons()
                except Exception as _e:
                    st.error(f"❌ Lỗi giải nén: {_e}")

# ============================================================
# TAB 3: KHO TRI THỨC
# ============================================================
with tab_kho:
    st.markdown("### 📂 Kho tri thức")
    st.caption("Xem và xóa file đã lưu trên GitHub")

    for label, key, folder in [("📐 Mẫu", "templates", "templates"),
                                ("📚 Nguồn", "sources", "sources"),
                                ("📅 Lịch BG", "schedule", "schedule")]:
        st.subheader(label + " đã lưu")
        data_dict = st.session_state[key]
        if data_dict:
            for name in list(data_dict.keys()):
                col1, col2 = st.columns([5, 1])
                with col1:
                    with st.expander("📄 " + name):
                        content = data_dict[name]
                        st.caption(f"Kích thước: {len(content)} ký tự")
                        st.text(content[:2000])
                with col2:
                    if st.button("🗑️", key=f"del_{key}_{name}"):
                        with st.spinner("Đang xóa..."):
                            ok, msg = gh_delete(f"data/{folder}/" + name)
                            if ok:
                                st.session_state[key] = load_folder(folder)
                                st.rerun()
            if st.button("🗑️ Xóa TẤT CẢ " + label, key=f"del_all_{key}"):
                for name in list(data_dict.keys()):
                    gh_delete(f"data/{folder}/" + name)
                st.session_state[key] = load_folder(folder)
                st.rerun()
        else:
            st.info("Chưa có " + label.lower() + ".")

    st.divider()
    st.subheader("📚 SGK đã tách (local)")
    ds_sgk_local = list_sgk_da_tach()
    if ds_sgk_local:
        for sgk in ds_sgk_local:
            col1, col2 = st.columns([5, 1])
            with col1:
                st.write(f"• **{sgk['sgk_name']}** --- "
                         f"{sgk['so_bai']} bài ({sgk['nguon']})")
                with st.expander("📖 Xem tên bài"):
                    render_danh_sach_bai(sgk["sgk_name"])
            with col2:
                if st.button("🗑️", key="del_sgk_" + sgk["sgk_name"]):
                    import shutil
                    try:
                        shutil.rmtree(
                            os.path.join(SGK_BAI_DIR, sgk["sgk_name"]))
                        st.success("✅ Đã xóa " + sgk["sgk_name"])
                        st.rerun()
                    except Exception as e:
                        st.error("Lỗi: " + str(e))
    else:
        st.info("Chưa có SGK nào được tách.")

# ============================================================
# TAB 4: NẠP MẪU GIÁO ÁN
# ============================================================
with tab_mau:
    st.markdown("### 📐 Nạp mẫu giáo án (CV 5512)")
    st.info("💡 Upload mẫu giáo án của trường. "
            "Hỗ trợ: `.docx`, `.txt`, `.md`.")

    _mau_files = st.file_uploader(
        "Chọn 1 hoặc nhiều file mẫu",
        type=["docx", "txt", "md"],
        accept_multiple_files=True, key="mau_upload")

    if _mau_files:
        for _mf in _mau_files:
            with st.expander(f"📄 {_mf.name}", expanded=True):
                try:
                    _raw = _mf.read()
                    if _mf.name.lower().endswith(".docx"):
                        _mau_text = _read_docx_safe(_raw)
                    else:
                        _mau_text = _raw.decode("utf-8", errors="ignore")
                    _ok, _msg = kiem_tra_mau_5512(_mau_text)
                    _c1, _c2 = st.columns([3, 1])
                    with _c1:
                        st.caption(
                            f"Kích thước: **{len(_mau_text)}** ký tự | "
                            "Kiểm tra 5512: " +
                            ("✅ **OK**" if _ok else f"⚠️ {_msg}"))
                        st.text_area("Xem trước", _mau_text[:3000],
                                     height=200,
                                     key=f"preview_mau_{_mf.name}")
                    with _c2:
                        st.write("")
                        if st.button("✅ Dùng mẫu này",
                                     key=f"use_mau_{_mf.name}",
                                     use_container_width=True,
                                     type="primary"):
                            st.session_state["mau_5512_dang_dung"] = {
                                "ten_mau": _mf.name, "che_do": "user",
                                "nguon": "upload", "cau_truc": _mau_text}
                            if GITHUB_TOKEN:
                                gh_save(f"data/templates/{_mf.name}",
                                        _mau_text,
                                        f"Upload template {_mf.name}")
                                st.session_state["templates"] = \
                                    load_templates()
                            st.success(f"✅ Đã đặt **{_mf.name}** làm mẫu.")
                            st.rerun()
                        if st.button("🗑️ Xóa khỏi kho",
                                     key=f"del_mau_{_mf.name}",
                                     use_container_width=True):
                            if GITHUB_TOKEN:
                                gh_delete(f"data/templates/{_mf.name}")
                                st.session_state["templates"] = \
                                    load_templates()
                                st.success("✅ Đã xóa.")
                                st.rerun()
                except Exception as _e:
                    st.error(f"❌ Lỗi đọc mẫu: {_e}")

    _mau_now = st.session_state.get("mau_5512_dang_dung")
    st.markdown("---")
    if _mau_now:
        st.success(f"📌 **Mẫu đang dùng:** `{_mau_now.get('ten_mau', 'N/A')}`"
                   f" (nguồn: {_mau_now.get('nguon', 'N/A')})")
        if st.button("↩️ Quay lại mẫu mặc định", key="reset_mau",
                     use_container_width=True):
            st.session_state["mau_5512_dang_dung"] = None
            st.success("✅ Đã quay lại mẫu 5512 mặc định.")
            st.rerun()
    else:
        st.info("📌 Đang dùng **mẫu 5512 hardcode**.")

    if st.session_state["templates"]:
        st.markdown("### 📚 Mẫu đã lưu trên GitHub")
        for _tname in list(st.session_state["templates"].keys()):
            _c1, _c2 = st.columns([5, 1])
            with _c1:
                with st.expander(f"📄 {_tname}"):
                    st.text(st.session_state["templates"][_tname][:2000])
            with _c2:
                if st.button("✅ Dùng", key=f"use_saved_mau_{_tname}",
                             use_container_width=True):
                    st.session_state["mau_5512_dang_dung"] = {
                        "ten_mau": _tname, "che_do": "user",
                        "nguon": "github",
                        "cau_truc": st.session_state["templates"][_tname]}
                    st.success(f"✅ Đã dùng mẫu **{_tname}**.")
                    st.rerun()

# ============================================================
# TAB 5: NẠP MA TRẬN ĐỀ
# ============================================================
with tab_matrix:
    st.markdown("### 🧩 Nạp ma trận đề kiểm tra")
    st.info("💡 Upload ma trận đề (Excel/CSV/Word) hoặc nhập tay.")

    _c1, _c2, _c3 = st.columns([1, 1, 1])
    with _c1:
        _mt_mon = st.text_input(
            "Môn",
            value=st.session_state["memory"].get("default_subject", "Toán"),
            key="mt_mon")
    with _c2:
        _mt_lop = st.text_input(
            "Lớp",
            value=st.session_state["memory"].get("default_grade", "6"),
            key="mt_lop")
    with _c3:
        _mt_loai = st.selectbox(
            "Loại đề",
            ["Giữa kì 1", "Cuối kì 1", "Giữa kì 2", "Cuối kì 2",
             "15 phút", "1 tiết", "Khác"],
            key="mt_loai")

    _mt_file = st.file_uploader("Upload file ma trận",
                                type=["xlsx", "xls", "csv", "docx", "txt"],
                                key="mt_upload")
    _mt_text = ""
    if _mt_file:
        try:
            _mt_bytes = _mt_file.read()
            _n = _mt_file.name.lower()
            if _n.endswith((".xlsx", ".xls", ".xlsm")):
                _mt_text = _read_excel_safe(_mt_bytes)
            elif _n.endswith(".csv"):
                _mt_text = _read_csv_safe(_mt_bytes)
            elif _n.endswith(".docx"):
                _mt_text = _read_docx_safe(_mt_bytes)
            else:
                _mt_text = _mt_bytes.decode("utf-8", errors="ignore")
            st.success(f"✅ Đã đọc ma trận ({len(_mt_text)} ký tự).")
        except Exception as _e:
            st.error(f"❌ Lỗi đọc file: {_e}")

    _mt_manual = st.text_area("Hoặc dán ma trận vào đây",
                              value=_mt_text, height=250,
                              key="mt_manual",
                              placeholder="VD:\n"
                              "| Chủ đề | Nhận biết | Thông hiểu | "
                              "Vận dụng | Vận dụng cao |\n"
                              "| Ước chung lớn nhất | 2 câu | 1 câu | "
                              "1 câu | 0 |\n")

    _c_save, _c_view = st.columns([1, 1])
    with _c_save:
        if st.button("💾 Lưu ma trận", type="primary",
                     use_container_width=True, key="mt_save"):
            if not _mt_manual.strip():
                st.warning("⚠️ Chưa có nội dung ma trận.")
            else:
                _mt_key = (f"{_slugify(_mt_mon).title().replace('_', '')}_"
                           f"{_mt_lop}_{_slugify(_mt_loai)}")
                st.session_state[f"matran_{_mt_key}"] = {
                    "mon": _mt_mon, "lop": _mt_lop,
                    "loai": _mt_loai, "noi_dung": _mt_manual}
                if GITHUB_TOKEN:
                    gh_save(f"data/schedule/matran_{_mt_key}.md",
                            _mt_manual, f"Ma trận {_mt_key}")
                    st.session_state["schedule"] = load_schedule()
                st.success(f"✅ Đã lưu ma trận `{_mt_key}`.")
                st.rerun()

    with _c_view:
        _mt_keys = [k.replace("matran_", "")
                    for k in st.session_state.keys()
                    if k.startswith("matran_")
                    and isinstance(st.session_state.get(k), dict)]
        if _mt_keys:
            _default_idx = 0
            _saved_sel = st.session_state.get("mt_saved_sel")
            if _saved_sel in _mt_keys:
                _default_idx = _mt_keys.index(_saved_sel)
            _mt_sel = st.selectbox("Ma trận đã lưu trong phiên",
                                   _mt_keys, index=_default_idx,
                                   key="mt_saved_sel")
            _mt_session_key = f"matran_{_mt_sel}"
            _mt_val = st.session_state.get(_mt_session_key)
            if isinstance(_mt_val, dict):
                st.caption("Xem trước:")
                st.text(str(_mt_val.get("noi_dung", ""))[:500])
            else:
                st.warning("⚠️ Ma trận không đúng định dạng.")

# ============================================================
# TAB 6: TẠO ĐẶC TẢ ĐỀ
# ============================================================
with tab_despec:
    st.markdown("### 📋 Tạo đặc tả đề từ ma trận")
    st.info("💡 Chọn ma trận → AI sinh **bảng đặc tả đề**.")

    _mt_keys_all = [k.replace("matran_", "")
                    for k in st.session_state.keys()
                    if k.startswith("matran_")
                    and isinstance(st.session_state.get(k), dict)]

    if not _mt_keys_all:
        st.warning("⚠️ Chưa có ma trận nào. Vào tab **🧩 Nạp ma trận đề**.")
    else:
        _default_idx = 0
        _saved_dt = st.session_state.get("dt_mt_sel")
        if _saved_dt in _mt_keys_all:
            _default_idx = _mt_keys_all.index(_saved_dt)
        _dt_key = st.selectbox("Chọn ma trận", _mt_keys_all,
                               index=_default_idx, key="dt_mt_sel")
        _dt_session_key = f"matran_{_dt_key}"
        _dt_val = st.session_state.get(_dt_session_key)

        if not isinstance(_dt_val, dict):
            st.error("❌ Ma trận không đúng định dạng. Nạp lại ở tab 🧩.")
        else:
            _dt_mt = _dt_val
            _c1, _c2 = st.columns(2)
            with _c1:
                st.caption(f"📚 Môn: **{_dt_mt.get('mon', '')}** | "
                           f"Lớp: **{_dt_mt.get('lop', '')}** | "
                           f"Loại: **{_dt_mt.get('loai', '')}**")
            with _c2:
                _dt_tg = st.text_input("Thời gian làm bài (phút)",
                                       value="45", key="dt_tg")

            if st.button("🚀 Sinh đặc tả đề", type="primary",
                         use_container_width=True, key="dt_run"):
                if not available_providers:
                    st.error("❌ Chưa có API key.")
                else:
                    _dt_prompt = f"""Bạn là chuyên gia khảo thí THCS.
Sinh BẢNG ĐẶC TẢ ĐỀ KIỂM TRA theo chuẩn CV 7991/BGDĐT-GDTrH.

Môn: {_dt_mt.get('mon', '')}
Lớp: {_dt_mt.get('lop', '')}
Loại: {_dt_mt.get('loai', '')}
Thời gian: {_dt_tg} phút

MA TRẬN GỐC:
\"\"\"
{_dt_mt.get('noi_dung', '')}
\"\"\"

YÊU CẦU:
- Bảng Markdown: TT | Chủ đề | Mạch kiến thức | Yêu cầu cần đạt |
  Số câu (NB/TH/VD/VDC) | Điểm | Câu hỏi minh họa
- Mỗi dòng = 1 đơn vị kiến thức
- Cuối bảng có dòng TỔNG
- KHÔNG viết ngoài bảng
"""
                    with st.spinner("🤖 Đang sinh đặc tả..."):
                        try:
                            _dt_result = call_ai(_dt_prompt)
                            st.session_state["dacta_hientai"] = {
                                "ma_tran_key": _dt_key,
                                "mon": _dt_mt.get("mon", ""),
                                "lop": _dt_mt.get("lop", ""),
                                "loai": _dt_mt.get("loai", ""),
                                "noi_dung": _dt_result}
                            save_history(
                                "Đặc tả đề",
                                f"{_dt_mt.get('mon', '')} "
                                f"{_dt_mt.get('lop', '')} "
                                f"{_dt_mt.get('loai', '')}",
                                _dt_result)
                            st.success("✅ Đã sinh đặc tả.")
                        except Exception as _e:
                            st.error(f"❌ Lỗi AI: {_e}")

            if st.session_state.get("dacta_hientai"):
                _dt_now = st.session_state["dacta_hientai"]
                st.markdown("---")
                st.markdown("### 📄 Đặc tả đề")
                st.text_area("Nội dung (có thể sửa)",
                             _dt_now.get("noi_dung", ""),
                             height=400, key="dt_edit")
                _c1, _c2 = st.columns(2)
                with _c1:
                    st.download_button(
                        "⬇️ Tải đặc tả (Word)",
                        make_doc("Đặc tả đề",
                                 _dt_now.get("noi_dung", ""),
                                 subject_name=_dt_now.get("mon", "")),
                        file_name=f"DacTa_{_slugify(_dt_now.get('mon', ''))}_"
                                  f"{_dt_now.get('lop', '')}.docx",
                        mime="application/vnd.openxmlformats-officedocument."
                             "wordprocessingml.document",
                        type="primary", use_container_width=True)
                with _c2:
                    st.download_button(
                        "⬇️ Tải đặc tả (TXT)",
                        _dt_now.get("noi_dung", "").encode("utf-8"),
                        file_name=f"DacTa_{_slugify(_dt_now.get('mon', ''))}_"
                                  f"{_dt_now.get('lop', '')}.txt",
                        mime="text/plain", use_container_width=True)

# ============================================================
# TAB 7: TẠO ĐỀ + ĐÁP ÁN
# ============================================================
with tab_exam:
    st.markdown("### 📝 Tạo đề kiểm tra + đáp án")
    st.info("💡 Chọn **đặc tả đề** → AI tạo **đề + đáp án + biểu điểm**.")

    if not st.session_state.get("dacta_hientai"):
        st.warning("⚠️ Chưa có đặc tả. Vào tab **📋 Tạo đặc tả đề** trước.")
    else:
        _dt_now = st.session_state["dacta_hientai"]
        _c1, _c2, _c3 = st.columns(3)
        with _c1:
            _ex_hinhthuc = st.selectbox(
                "Hình thức",
                ["Trắc nghiệm 100%", "Tự luận 100%",
                 "Trắc nghiệm + Tự luận"],
                key="ex_hinhthuc")
        with _c2:
            _ex_socau = st.text_input("Tổng số câu", value="20",
                                       key="ex_socau")
        with _c3:
            _ex_thangdiem = st.text_input("Thang điểm", value="10",
                                           key="ex_thangdiem")

        _ex_yeucau = st.text_area("Yêu cầu bổ sung", value="",
                                   height=80, key="ex_yeucau",
                                   placeholder="VD: 4 mã đề, có hình ảnh...")

        if st.button("🚀 Sinh đề + đáp án", type="primary",
                     use_container_width=True, key="ex_run"):
            if not available_providers:
                st.error("❌ Chưa có API key.")
            else:
                _ex_prompt = f"""Bạn là chuyên gia khảo thí THCS.
Từ ĐẶC TẢ ĐỀ dưới đây, sinh ĐỀ KIỂM TRA + ĐÁP ÁN + BIỂU ĐIỂM.

Môn: {_dt_now.get('mon', '')}
Lớp: {_dt_now.get('lop', '')}
Loại: {_dt_now.get('loai', '')}
Hình thức: {_ex_hinhthuc}
Tổng số câu: {_ex_socau}
Thang điểm: {_ex_thangdiem}
Yêu cầu bổ sung: {_ex_yeucau or "Không"}

ĐẶC TẢ:
\"\"\"
{_dt_now.get('noi_dung', '')}
\"\"\"

YÊU CẦU:
### PHẦN A: ĐỀ KIỂM TRA
### PHẦN B: ĐÁP ÁN + BIỂU ĐIỂM
- Unicode toán: x², √, ≤, ≥, π
- Không viết "Người ra đề", "Ký tên"
"""
                with st.spinner("🤖 Đang sinh đề + đáp án..."):
                    try:
                        _ex_result = call_ai(_ex_prompt)
                        st.session_state["de_kiem_tra_hientai"] = {
                            "mon": _dt_now.get("mon", ""),
                            "lop": _dt_now.get("lop", ""),
                            "loai": _dt_now.get("loai", ""),
                            "noi_dung": _ex_result}
                        save_history(
                            "Đề kiểm tra",
                            f"{_dt_now.get('mon', '')} "
                            f"{_dt_now.get('lop', '')}",
                            _ex_result)
                        st.success("✅ Đã sinh đề + đáp án.")
                    except Exception as _e:
                        st.error(f"❌ Lỗi AI: {_e}")

        if st.session_state.get("de_kiem_tra_hientai"):
            _ex_now = st.session_state["de_kiem_tra_hientai"]
            st.markdown("---")
            st.markdown("### 📄 Đề kiểm tra + Đáp án")
            st.text_area("Nội dung (có thể sửa)",
                         _ex_now.get("noi_dung", ""),
                         height=500, key="ex_edit")
            _c1, _c2 = st.columns(2)
            with _c1:
                st.download_button(
                    "⬇️ Tải đề (Word)",
                    make_doc("Đề kiểm tra",
                             _ex_now.get("noi_dung", ""),
                             subject_name=_ex_now.get("mon", "")),
                    file_name=f"DeKiemTra_{_slugify(_ex_now.get('mon', ''))}_"
                              f"{_ex_now.get('lop', '')}.docx",
                    mime="application/vnd.openxmlformats-officedocument."
                         "wordprocessingml.document",
                    type="primary", use_container_width=True)
            with _c2:
                st.download_button(
                    "⬇️ Tải đề (TXT)",
                    _ex_now.get("noi_dung", "").encode("utf-8"),
                    file_name=f"DeKiemTra_{_slugify(_ex_now.get('mon', ''))}_"
                              f"{_ex_now.get('lop', '')}.txt",
                    mime="text/plain", use_container_width=True)

# ============================================================
# TAB 8: SỬA TÊN BÀI SGK
# ============================================================
with tab_edit:
    st.markdown("### ✏️ Sửa tên bài SGK đã tách")
    st.info("💡 Chỉnh sửa **tên bài** trực tiếp. "
            "Bấm **💾 Lưu thay đổi** để ghi vào `metadata.json`.")

    _ds_edit = list_sgk_da_tach()
    if not _ds_edit:
        st.warning("⚠️ Chưa có SGK nào. Vào tab **📤 Upload SGK gốc** trước.")
    else:
        _sgk_edit_name = st.selectbox(
            "Chọn SGK cần sửa",
            [s["sgk_name"] for s in _ds_edit],
            key="edit_sgk_sel")

        _meta_edit = load_sgk_metadata(_sgk_edit_name)
        if not _meta_edit or not _meta_edit.get("ds_bai"):
            st.warning("SGK này chưa có bài nào.")
        else:
            _ds_bai = _meta_edit["ds_bai"]
            st.caption(f"📚 **{_sgk_edit_name}** — {len(_ds_bai)} bài")

            _c1, _c2, _c3 = st.columns([1, 1, 2])
            with _c1:
                _btn_add = st.button("➕ Thêm bài", key="edit_add_bai",
                                     use_container_width=True)
            with _c2:
                _btn_reset = st.button("↩️ Tải lại", key="edit_reload",
                                        use_container_width=True)
            with _c3:
                st.caption("Sửa trực tiếp tên bài → bấm Lưu.")

            if _btn_add:
                _ds_bai.append({
                    "file": "",
                    "ten_bai": "",
                    "so_bai": len(_ds_bai) + 1,
                    "page_pdf_start": 1,
                    "page_pdf_end": 1,
                })
                _meta_edit["ds_bai"] = _ds_bai
                _meta_edit["so_bai"] = len(_ds_bai)
                _save_meta_path = os.path.join(
                    SGK_BAI_DIR, _sgk_edit_name, "metadata.json")
                with open(_save_meta_path, "w", encoding="utf-8") as _f:
                    json.dump(_meta_edit, _f,
                              ensure_ascii=False, indent=2)
                st.success("✅ Đã thêm 1 dòng trống ở cuối.")
                st.rerun()

            if _btn_reset:
                st.rerun()

            st.markdown("---")

            _h1, _h2, _h3 = st.columns([0.5, 5, 0.7])
            with _h1:
                st.markdown("**STT**")
            with _h2:
                st.markdown("**Tên bài**")
            with _h3:
                st.markdown("**Xóa**")

            _deleted_idx = None

            for _i, _b in enumerate(_ds_bai):
                _cols = st.columns([0.5, 5, 0.7])
                with _cols[0]:
                    st.write(f"**{_i+1}**")
                with _cols[1]:
                    _new_ten = st.text_input(
                        "Tên bài",
                        value=_b.get("ten_bai", ""),
                        key=f"edit_ten_{_i}",
                        label_visibility="collapsed",
                        placeholder="Gõ tên bài đúng theo SGK")
                    _b["ten_bai"] = _new_ten
                with _cols[2]:
                    if st.button("🗑️", key=f"edit_del_{_i}",
                                 use_container_width=True):
                        _deleted_idx = _i
                        st.rerun()

            if _deleted_idx is not None:
                _ds_bai.pop(_deleted_idx)
                for _k, _bk in enumerate(_ds_bai):
                    _bk["so_bai"] = _k + 1
                _meta_edit["ds_bai"] = _ds_bai
                _meta_edit["so_bai"] = len(_ds_bai)
                _save_meta_path = os.path.join(
                    SGK_BAI_DIR, _sgk_edit_name, "metadata.json")
                with open(_save_meta_path, "w", encoding="utf-8") as _f:
                    json.dump(_meta_edit, _f,
                              ensure_ascii=False, indent=2)
                st.success("✅ Đã xóa bài.")
                st.rerun()

            st.markdown("---")

            _c1, _c2 = st.columns([1, 1])
            with _c1:
                if st.button("💾 Lưu thay đổi", type="primary",
                             use_container_width=True,
                             key="edit_save_meta"):
                    _errs = []
                    _seen = {}
                    for _k, _bk in enumerate(_ds_bai):
                        ten = _bk.get("ten_bai", "").strip()
                        if not ten:
                            _errs.append(f"Dòng {_k+1}: Thiếu tên bài")
                            continue
                        _ten_chuan = _chuan_hoa_tim_kiem(ten)
                        if _ten_chuan in _seen:
                            _errs.append(
                                f"Dòng {_k+1}: Trùng tên với dòng "
                                f"{_seen[_ten_chuan] + 1}")
                        else:
                            _seen[_ten_chuan] = _k
                        _bk["so_bai"] = _k + 1

                    if _errs:
                        for _e in _errs:
                            st.error("❌ " + _e)
                    else:
                        _meta_edit["ds_bai"] = _ds_bai
                        _meta_edit["so_bai"] = len(_ds_bai)
                        _save_meta_path = os.path.join(
                            SGK_BAI_DIR, _sgk_edit_name, "metadata.json")
                        with open(_save_meta_path, "w",
                                  encoding="utf-8") as _f:
                            json.dump(_meta_edit, _f,
                                      ensure_ascii=False, indent=2)
                        st.success(
                            f"✅ Đã lưu {len(_ds_bai)} tên bài vào "
                            f"`{_sgk_edit_name}/metadata.json`.")
                        st.balloons()

            with _c2:
                if st.button("🔄 Tải lại từ file", use_container_width=True,
                             key="edit_reload_full"):
                    st.rerun()

            st.markdown("---")
            st.download_button(
                "⬇️ Tải metadata.json",
                json.dumps(_meta_edit, ensure_ascii=False, indent=2)
                    .encode("utf-8"),
                file_name=f"{_sgk_edit_name}_metadata.json",
                mime="application/json",
                use_container_width=True)

# ============================================================
# TAB 9: KHỚP TÊN BÀI VỚI LỊCH BÁO GIẢNG
# ============================================================
with tab_match:
    st.markdown("### 📝 Khớp tên bài SGK với Lịch báo giảng")
    st.info("💡 Dán danh sách tên bài từ **Lịch báo giảng** vào ô bên dưới. "
            "App sẽ so khớp với tên bài hiện tại trong SGK và cho phép "
            "gán lại tên chuẩn.")

    _ds_match = list_sgk_da_tach()
    if not _ds_match:
        st.warning("⚠️ Chưa có SGK nào. Vào tab **📤 Upload SGK gốc** trước.")
    else:
        _sgk_match_name = st.selectbox(
            "Chọn SGK cần khớp",
            [s["sgk_name"] for s in _ds_match],
            key="match_sgk_sel")

        _meta_match = load_sgk_metadata(_sgk_match_name)
        if not _meta_match or not _meta_match.get("ds_bai"):
            st.warning("SGK này chưa có bài nào.")
        else:
            _ds_bai_match = _meta_match["ds_bai"]
            st.markdown(f"**📚 {_sgk_match_name}** — "
                        f"{len(_ds_bai_match)} bài trong SGK")

            st.markdown("#### 1️⃣ Dán danh sách tên bài từ Lịch BG")
            st.caption("Mỗi dòng 1 tên bài. Thứ tự dòng phải khớp với "
                       "thứ tự bài trong SGK.")

            _default_match_text = ""
            _lbg_now = st.session_state.get("lich_bao_giang", [])
            if _lbg_now:
                _ten_tu_lbg = [r.get("ten_bai", "").strip()
                               for r in _lbg_now
                               if r.get("ten_bai", "").strip()]
                if _ten_tu_lbg:
                    _default_match_text = "\n".join(_ten_tu_lbg)

            _match_text = st.text_area(
                "Danh sách tên bài (mỗi dòng 1 bài)",
                value=_default_match_text,
                height=250,
                key="match_ten_input",
                placeholder=("VD:\nTập hợp\nCách ghi số tự nhiên\n"
                             "Thứ tự trong tập hợp các số tự nhiên\n..."))

            _c_tk1, _c_tk2 = st.columns([1, 3])
            with _c_tk1:
                if st.button("📥 Lấy từ Lịch BG hiện tại",
                             key="match_load_lbg",
                             use_container_width=True):
                    _lbg_now2 = st.session_state.get("lich_bao_giang", [])
                    if _lbg_now2:
                        _ten2 = [r.get("ten_bai", "").strip()
                                  for r in _lbg_now2
                                  if r.get("ten_bai", "").strip()]
                        if _ten2:
                            st.session_state["match_ten_input"] = \
                                "\n".join(_ten2)
                            st.success(
                                f"✅ Đã lấy {len(_ten2)} tên từ Lịch BG.")
                            st.rerun()
                        else:
                            st.warning("Lịch BG chưa có tên bài nào.")
                    else:
                        st.warning("Chưa có Lịch BG trong phiên.")

            with _c_tk2:
                st.caption("💡 Nếu bạn đã nhập Lịch BG ở phần dưới → "
                           "bấm nút này để tự điền.")

            _ten_moi = [ln.strip() for ln in _match_text.splitlines()
                        if ln.strip()]

            st.markdown("---")
            st.markdown("#### 2️⃣ Đối chiếu và gán tên")

            if not _ten_moi:
                st.info("Chưa có tên bài nào trong ô trên.")
            else:
                st.caption(
                    f"SGK có **{len(_ds_bai_match)}** bài | "
                    f"Danh sách mới có **{len(_ten_moi)}** tên")

                if len(_ten_moi) != len(_ds_bai_match):
                    st.warning(
                        f"⚠️ Số lượng không khớp: SGK có "
                        f"{len(_ds_bai_match)} bài, danh sách có "
                        f"{len(_ten_moi)} tên. App sẽ ghép theo thứ tự, "
                        f"các bài dư sẽ giữ nguyên tên cũ.")

                _h1, _h2, _h3 = st.columns([0.5, 2.5, 2.5])
                with _h1:
                    st.markdown("**STT**")
                with _h2:
                    st.markdown("**Tên hiện tại trong SGK**")
                with _h3:
                    st.markdown("**Tên mới (từ Lịch BG)**")

                _ten_ap_dung = []
                for _i, _b in enumerate(_ds_bai_match):
                    _cols = st.columns([0.5, 2.5, 2.5])
                    with _cols[0]:
                        st.write(f"**{_i+1}**")
                    with _cols[1]:
                        st.text_input(
                            "Tên hiện tại",
                            value=_b.get("ten_bai", ""),
                            key=f"match_old_{_i}",
                            disabled=True,
                            label_visibility="collapsed")
                    with _cols[2]:
                        _default_new = ""
                        if _i < len(_ten_moi):
                            _default_new = _ten_moi[_i]
                        _new_val = st.text_input(
                            "Tên mới",
                            value=_default_new,
                            key=f"match_new_{_i}",
                            label_visibility="collapsed",
                            placeholder="Tên bài đúng")
                        _ten_ap_dung.append(_new_val)

                st.markdown("---")

                _c_ap1, _c_ap2 = st.columns([1, 1])
                with _c_ap1:
                    if st.button(
                        "✅ Áp dụng tên mới vào SGK",
                        type="primary",
                        use_container_width=True,
                        key="match_apply"):
                        _errs = []
                        _seen = {}
                        _count = 0
                        for _k, _bk in enumerate(_ds_bai_match):
                            _ten = _ten_ap_dung[_k].strip() \
                                if _k < len(_ten_ap_dung) else ""
                            if not _ten:
                                _ten = _bk.get("ten_bai", "")
                            _ten_chuan = _chuan_hoa_tim_kiem(_ten)
                            if _ten_chuan in _seen:
                                _errs.append(
                                    f"Dòng {_k+1}: Trùng với dòng "
                                    f"{_seen[_ten_chuan] + 1} ({_ten})")
                            else:
                                _seen[_ten_chuan] = _k
                            if _ten != _bk.get("ten_bai", ""):
                                _bk["ten_bai"] = _ten
                                _count += 1

                        if _errs:
                            for _e in _errs:
                                st.error("❌ " + _e)
                        else:
                            _meta_match["ds_bai"] = _ds_bai_match
                            _save_path = os.path.join(
                                SGK_BAI_DIR, _sgk_match_name,
                                "metadata.json")
                            with open(_save_path, "w",
                                      encoding="utf-8") as _f:
                                json.dump(_meta_match, _f,
                                          ensure_ascii=False, indent=2)
                            st.success(
                                f"✅ Đã cập nhật **{_count}** tên bài. "
                                "Vào **Soạn 1 bài** hoặc **Soạn cả tuần** "
                                "để dùng tên mới.")
                            st.balloons()

                with _c_ap2:
                    if st.button(
                        "🔄 Tải lại từ file",
                        use_container_width=True,
                        key="match_reload"):
                        st.rerun()

                with st.expander("👁️ Xem trước kết quả sau khi áp dụng",
                                 expanded=False):
                    _preview = []
                    for _k, _bk in enumerate(_ds_bai_match):
                        _old = _bk.get("ten_bai", "")
                        _new = _ten_ap_dung[_k] if _k < len(_ten_ap_dung) \
                            else ""
                        _preview.append({
                            "STT": _k + 1,
                            "Tên cũ": _old,
                            "Tên mới": _new if _new else _old,
                            "Thay đổi": "✅" if _new and _new != _old
                                        else "—"})
                    st.dataframe(_preview, use_container_width=True,
                                 hide_index=True)

# ============ KẾT THÚC PHẦN 4/6 ============

# ============================================================
# PHẦN 5/6: WORD HELPERS + SOẠN 1 BÀI
# ============================================================
# Chức năng:
#   - fix_math, _set_run_font, add_para, add_table, _add_heading
#   - write_content_to_doc, make_doc
#   - Nút "Soạn 1 bài" (chỉ dùng TÊN BÀI, bỏ trang PDF)
# ============================================================

def fix_math(text):
    """Chuyển ký hiệu LaTeX đơn giản thành Unicode."""
    text = re.sub(r"\^2\b", "²", text)
    text = re.sub(r"\^3\b", "³", text)
    text = re.sub(r"_2\b", "₂", text)
    text = text.replace("sqrt(", "√(").replace("\\sqrt{", "√")
    text = text.replace("\\times", "×").replace("\\div", "÷")
    text = text.replace("\\leq", "≤").replace("\\geq", "≥")
    text = text.replace("\\pi", "π").replace("\\alpha", "α")
    text = re.sub(r"\$([^\$]+)\$", r"\1", text)
    return text


def _set_run_font(run, size=14, bold=False):
    """Set font Times New Roman cho run."""
    run.bold = bold
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    for at in ["w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"]:
        rfonts.set(qn(at), "Times New Roman")


def add_para(doc, text, style=None):
    """Thêm 1 đoạn văn vào docx."""
    text = text.replace("<br>", "\n").replace("<br/>", "\n")
    text = re.sub(r"^[\-\+\*\•\#]+\s*", "", text)
    text = re.sub(r"^\d+[\.\)]\s+", "", text)
    text = re.sub(r"^[a-z]\)\s+", "", text)
    text = text.replace("**", "").replace("__", "")
    text = re.sub(r"(?<!\*)\*(?!\*)", "", text)
    text = text.strip()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Inches(0.39)
    p.paragraph_format.left_indent = Inches(0)
    p.paragraph_format.right_indent = Inches(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(text)
    _set_run_font(run, 14, bold=False)
    return p


def _tach_dong_trong_o(val):
    """Tách nội dung 1 ô bảng thành nhiều dòng hợp lý."""
    val = val.replace("<br>", "\n").replace("<br/>", "\n")
    val = val.replace("<br />", "\n")
    val = re.sub(r"(?<!^)(?<!\n)(?=Bước\s*\d+\s*:)", "\n", val)

    chu_hoa = (r"[A-ZĐÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬÈÉẺẼẸÊỀẾỂỄỆ"
               r"ÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰ"
               r"ỲÝỶỸỴ]")
    val = re.sub(r"([\.\!\?\:])\s+(?=" + chu_hoa + r")", r"\1\n", val)

    keywords = ["Em hãy", "Vậy ", "Do đó", "Tuy nhiên", "Vì vậy",
                "GV yêu cầu", "GV gọi", "GV đánh giá", "GV chính xác",
                "GV hướng dẫn", "GV đặt", "GV chiếu",
                "HS suy nghĩ", "HS thảo luận", "HS thực hiện",
                "HS báo cáo", "HS nhận xét", "HS trả lời"]
    for kw in keywords:
        val = re.sub(r"(?<=[\.\!\?])\s+" + re.escape(kw), "\n" + kw, val)

    raw_lines = [ln.strip() for ln in val.split("\n") if ln.strip()]
    merged = []
    for ln in raw_lines:
        if merged and len(ln) < 12 and not ln.startswith("Bước"):
            merged[-1] = merged[-1] + " " + ln
        else:
            merged.append(ln)
    return "\n".join(merged)


def add_table(doc, rows):
    """Thêm bảng 2 cột với justify + tách dòng thông minh."""
    if not rows:
        return

    cleaned_rows = []
    for row in rows:
        while len(row) > 2 and len(row[-1].strip()) <= 2 and \
              row[-1].strip() in ("", "=", ".", ";", ",", "a", "b",
                                   "c", "5", "0", "1", "2", "3", "4",
                                   "6", "7", "8", "9", "|"):
            row = row[:-1]
        cleaned_rows.append(row)

    max_n = max(len(r) for r in cleaned_rows)
    if max_n > 2:
        for i, r in enumerate(cleaned_rows):
            if len(r) > 2:
                extra = " ".join(c.strip() for c in r[2:] if c.strip())
                cleaned_rows[i] = [r[0], r[1] + (" " + extra if extra else "")]
        max_n = 2
    rows = cleaned_rows
    n = max_n
    if n < 1:
        return

    t = doc.add_table(rows=len(rows), cols=n)
    t.style = "Table Grid"

    for i, row in enumerate(rows):
        for j in range(n):
            c = t.cell(i, j)
            c.text = ""
            p = c.paragraphs[0]
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.first_line_indent = Inches(0)
            p.paragraph_format.left_indent = Inches(0)
            p.paragraph_format.right_indent = Inches(0)

            is_header = (i == 0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if is_header \
                else WD_ALIGN_PARAGRAPH.JUSTIFY

            val = row[j] if j < len(row) else ""

            if is_header:
                val = re.sub(r"^[\-\+\*\•\#]+\s*", "", val)
                val = re.sub(r"^\d+[\.\)]\s+", "", val)
                val = val.replace("**", "").strip()
                run = p.add_run(val)
                _set_run_font(run, 13, bold=True)
            else:
                val = _tach_dong_trong_o(val)
                lines_clean = []
                for ln in val.split("\n"):
                    ln = re.sub(r"^[\-\+\*\•\#]+\s*", "", ln)
                    ln = re.sub(r"^\d+[\.\)]\s+", "", ln)
                    ln = re.sub(r"\s{2,}", " ", ln).strip()
                    if ln:
                        lines_clean.append(ln)

                first = True
                for line_content in lines_clean:
                    if first:
                        target_p = p
                        first = False
                    else:
                        target_p = c.add_paragraph()
                        target_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        target_p.paragraph_format.line_spacing_rule = \
                            WD_LINE_SPACING.SINGLE
                        target_p.paragraph_format.space_before = Pt(0)
                        target_p.paragraph_format.space_after = Pt(2)
                        target_p.paragraph_format.first_line_indent = Inches(0)
                        target_p.paragraph_format.left_indent = Inches(0)
                        target_p.paragraph_format.right_indent = Inches(0)

                    for part in re.split(r"(\*\*[^\*]+\*\*)", line_content):
                        if part.startswith("**") and part.endswith("**") \
                                and len(part) > 4:
                            run = target_p.add_run(part[2:-2])
                            _set_run_font(run, 13, bold=True)
                        else:
                            run = target_p.add_run(part.replace("**", ""))
                            _set_run_font(run, 13, bold=False)


def _add_heading(doc, text):
    """Thêm heading. Nhận diện 3 dòng tiêu đề 5512."""
    text = re.sub(r"^[\*#\-\+\\]+\s*", "", text)
    text = text.rstrip("*").strip()

    p = doc.add_paragraph()
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)

    # 3 dòng tiêu đề 5512
    is_ten_bai = bool(re.match(r"^TÊN\s*BÀI\s*DẠY\s*:", text, re.I))
    is_mon_hoc = bool(re.match(r"^Môn\s*học\s*:", text, re.I))
    is_thoi_gian = bool(re.match(r"^Thời\s*gian\s*thực\s*hiện\s*:",
                                  text, re.I))

    if is_ten_bai or is_mon_hoc or is_thoi_gian:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Inches(0)
        p.paragraph_format.left_indent = Inches(0)
        p.paragraph_format.right_indent = Inches(0)

        if ":" in text:
            idx = text.index(":")
            label = text[:idx + 1].strip()
            content = text[idx + 1:].strip()
        else:
            label = text
            content = ""

        if is_thoi_gian:
            if content and not content.startswith("("):
                m_so = re.search(r"(\d+)", content)
                if m_so:
                    content = "(" + m_so.group(1) + " tiết)"
                else:
                    content = "(" + content + ")"

        run_label = p.add_run(label + " ")
        _set_run_font(run_label, 14, bold=True)
        if content:
            run_content = p.add_run(content)
            _set_run_font(run_content, 14, bold=False)
        return p

    # Heading a) b) c)
    if re.match(r"^[a-z]\)\s", text):
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Inches(0.39)
        p.paragraph_format.left_indent = Inches(0)
        p.paragraph_format.right_indent = Inches(0)
        if ":" in text:
            idx_colon = text.index(":")
            heading_part = text[:idx_colon + 1]
            content_part = text[idx_colon + 1:]
            heading_part = re.sub(r"^\*\*", "", heading_part).strip()
            heading_part = re.sub(r"\*\*$", "", heading_part).strip()
            content_part = re.sub(r"^\*\*\s*", "", content_part)
            content_part = re.sub(r"\*\*$", "", content_part).strip()
            run1 = p.add_run(heading_part)
            _set_run_font(run1, 14, bold=True)
            run1.italic = True
            if content_part:
                if not content_part.startswith(" "):
                    content_part = " " + content_part
                run2 = p.add_run(content_part)
                _set_run_font(run2, 14, bold=False)
                run2.italic = True
        else:
            clean = re.sub(r"\*\*", "", text).strip()
            run = p.add_run(clean)
            _set_run_font(run, 14, bold=True)
            run.italic = True
        return p

    # Heading thường
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Inches(0.39)
    p.paragraph_format.left_indent = Inches(0)
    p.paragraph_format.right_indent = Inches(0)
    run = p.add_run(text)
    _set_run_font(run, 14, bold=True)
    return p


def write_content_to_doc(d, content):
    """Ghi nội dung vào docx, xử lý heading, bảng, đoạn."""
    content = fix_math(content)
    content = re.sub(r"^\s*BẢNG\s*2\s*CỘT\s*:?\s*$", "",
                     content, flags=re.MULTILINE)
    content = re.sub(r"^\s*Bảng\s*2\s*cột\s*:?\s*$", "",
                     content, flags=re.MULTILINE)
    for tag in ["div", "strong", "span", "p", "em"]:
        content = re.sub(r"<" + tag + r"[^>]*>", "", content)
        content = re.sub(r"</" + tag + r">", "", content)
    content = re.sub(r"^\s*[-=]{3,}\s*$", "", content, flags=re.MULTILINE)

    # Cắt phần thừa
    lines = content.splitlines()
    start_idx = 0
    for idx, line in enumerate(lines):
        s = line.strip()
        if re.match(r"^\*\*\s*TÊN\s*BÀI\s*DẠY", s, re.I):
            start_idx = idx
            break
        if re.match(r"^\*\*\s*I\.\s*(Mục|MỤC)\s*tiêu", s, re.I):
            start_idx = idx
            break
    lines = lines[start_idx:]

    cleaned = []
    for line in lines:
        s = line.strip()
        if re.match(r"^Sông\s+Đốc", s, re.I):
            break
        if re.search(r"NGƯỜI\s*SOẠN", s, re.I):
            break
        if re.search(r"KÝ\s*DUYỆT", s, re.I):
            break
        if re.match(r"^\(?Ký\s*tên\)?\s*$", s, re.I):
            break
        cleaned.append(line)
    content = "\n".join(cleaned)

    lines = content.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        stripped = line.strip()

        if stripped == "":
            if i > 0 and lines[i - 1].strip() != "":
                p = d.add_paragraph("")
                p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
            i += 1
            continue

        # Bảng
        if stripped.startswith("|") and stripped.endswith("|"):
            tl = []
            j = i
            while j < len(lines) and lines[j].strip().startswith("|") \
                    and lines[j].strip().endswith("|"):
                tl.append(lines[j].strip())
                j += 1
            is_valid = False
            if len(tl) >= 2:
                if re.match(r"^\|[\s\-:\|]+\|$", tl[1]):
                    is_valid = True
            if is_valid:
                rows = []
                for t in tl:
                    if re.match(r"^\|[\s\-:\|]+\|$", t):
                        continue
                    cells = [c.strip() for c in t.strip("|").split("|")]
                    if 2 <= len(cells) <= 4:
                        rows.append(cells)
                if rows:
                    add_table(d, rows)
                else:
                    for t in tl:
                        text_line = t.strip("|").strip()
                        if text_line:
                            add_para(d, text_line)
            else:
                for t in tl:
                    text_line = t.strip("|").strip()
                    if text_line:
                        add_para(d, text_line)
            i = j
            continue

        # 3 dòng tiêu đề 5512
        if re.match(r"^\*\*\s*TÊN\s*BÀI\s*DẠY", stripped, re.I):
            _add_heading(d, stripped.replace("**", "").strip())
            i += 1
            continue
        if re.match(r"^\*\*\s*Môn\s*học", stripped, re.I):
            _add_heading(d, stripped.replace("**", "").strip())
            i += 1
            continue
        if re.match(r"^\*\*\s*Thời\s*gian\s*thực\s*hiện", stripped, re.I):
            _add_heading(d, stripped.replace("**", "").strip())
            i += 1
            continue

        # Heading I, II, III
        if re.match(r"^\*\*\s*(I|II|III|IV|V)\.\s+\S", stripped):
            _add_heading(d, stripped.replace("**", "").strip())
            i += 1
            continue

        # Heading 1. 2. 3.
        if re.match(r"^\*\*\s*\d+\.\s+[A-Za-zÀ-ỹ]", stripped) and \
           len(stripped) < 120:
            _add_heading(d, stripped.replace("**", "").strip())
            i += 1
            continue

        # Heading a) b) c)
        if re.match(r"^\*\*\s*[a-z]\)\s+\S", stripped):
            _add_heading(d, stripped.replace("**", "").strip())
            i += 1
            continue

        # Markdown headings
        if stripped.startswith("#### "):
            _add_heading(d, re.sub(r"[\*#\\]", "", stripped[5:].strip()))
        elif stripped.startswith("### "):
            _add_heading(d, re.sub(r"[\*#\\]", "", stripped[4:].strip()))
        elif stripped.startswith("## "):
            _add_heading(d, re.sub(r"[\*#\\]", "", stripped[3:].strip()))
        elif stripped.startswith("# "):
            _add_heading(d, re.sub(r"[\*#\\]", "", stripped[2:].strip()))
        elif re.match(r"^[-\*+\•]\s+", stripped):
            add_para(d, stripped)
        elif re.match(r"^\d+[\.\)]\s+", stripped):
            add_para(d, stripped)
        else:
            text = re.sub(r"^[\*#\-\+\\]+\s*", "", stripped)
            text = re.sub(r"^\*\*\\?\*?\s*", "", text)
            text = re.sub(r"\s*\*\*\\?\*?$", "", text)
            if text:
                add_para(d, text)
        i += 1


def make_doc(title, body, subject_name=""):
    """Tạo file Word với format chuẩn."""
    d = Document()
    s = d.sections[0]
    s.top_margin = Inches(0.7)
    s.bottom_margin = Inches(0.7)
    s.left_margin = Inches(0.8)
    s.right_margin = Inches(0.8)

    style = d.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(14)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    for at in ["w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"]:
        rfonts.set(qn(at), "Times New Roman")
    pf = style.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)

    d.add_heading(title, 0)

    mem = st.session_state.get("memory", {})
    if mem.get("school"):
        d.add_paragraph("Đơn vị: " + mem["school"])
    if mem.get("teacher_name"):
        d.add_paragraph("Người thực hiện: " + mem["teacher_name"])
    d.add_paragraph("Ngày tạo: " + datetime.now().strftime("%d/%m/%Y %H:%M"))
    d.add_paragraph("")

    write_content_to_doc(d, body)

    b = io.BytesIO()
    d.save(b)
    return b.getvalue()


# ============================================================
# SOẠN 1 BÀI THEO LỆNH (CHỈ DÙNG TÊN BÀI)
# ============================================================
st.divider()
st.subheader("💬 Soạn 1 bài theo lệnh")

_danh_sach_mon = list(SUBJECT_RULES.keys())
_danh_sach_sgk = list_sgk_da_tach()

_col1, _col2, _col3 = st.columns([2, 1.5, 1])
with _col1:
    _mon_chon = st.selectbox("📚 Môn học", _danh_sach_mon,
                              key="chon_mon_soan")
with _col2:
    _lop_chon = st.text_input(
        "🎓 Lớp",
        value=st.session_state["memory"].get("default_grade", "6"),
        key="chon_lop_soan")
with _col3:
    _tiet_chon = st.text_input("⏱️ Số tiết", value="1",
                                key="chon_tiet_soan")

# Xem danh sách tên bài để tra cứu
if _danh_sach_sgk:
    with st.expander("🔍 Tra cứu tên bài SGK", expanded=False):
        _sgk_tc = st.selectbox("Chọn SGK",
                                [s["sgk_name"] for s in _danh_sach_sgk],
                                key="tc_sgk_sel")
        _ds_ten_tc = list_ten_bai(_sgk_tc)
        if _ds_ten_tc:
            _filter = st.text_input("Lọc theo tên bài", value="",
                                     key="tc_filter")
            _opts_tc = [""] + [t for t in _ds_ten_tc
                                if not _filter or
                                _filter.lower() in t.lower()]
            _chon_bai = st.selectbox(
                "Chọn bài để điền vào ô Lệnh",
                _opts_tc,
                key="chon_bai_dien")
            if _chon_bai:
                st.session_state["single_cmd"] = _chon_bai
                st.success(f"✅ Đã điền: **{_chon_bai}**")
        else:
            st.info("SGK này chưa có bài nào.")

# Banner môn
_rules = SUBJECT_RULES.get(_mon_chon, {})
if _rules:
    _kh = _rules.get("ky_hieu", "")
    if _kh:
        st.markdown(
            f"<div style='background:#e0f7ff; padding:8px 14px; "
            f"border-radius:8px; border-left:4px solid #0ea5e9; "
            f"margin-bottom:8px;'>"
            f"<b style='font-size:16px;'>{_kh}</b> "
            f"<b>Môn {_mon_chon}</b></div>",
            unsafe_allow_html=True)

cmd = st.text_area(
    "Lệnh", height=100,
    placeholder="VD: Soạn bài Ước chung lớn nhất",
    label_visibility="collapsed",
    key="single_cmd")

_c_btn1, _c_btn2 = st.columns([3, 1])
with _c_btn1:
    _run_btn = st.button("▶️ Thực hiện", type="primary",
                          use_container_width=True)
with _c_btn2:
    if st.button("🗑️ Xóa", use_container_width=True, key="clear_single"):
        st.session_state["draft"] = ""
        st.rerun()

if _run_btn and cmd.strip():
    # === KIỂM TRA ĐẦU VÀO ===
    _errors = []

    # 1. Kiểm tra API
    if not available_providers:
        _errors.append("❌ Chưa có API key.")

    # 2. Kiểm tra mẫu
    _mau = get_mau_dang_dung()
    _ok_mau, _msg_mau = kiem_tra_mau_5512(_mau.get("cau_truc", ""))
    if not _ok_mau:
        _errors.append("❌ Mẫu 5512 không đạt: " + _msg_mau)

    # 3. Kiểm tra SGK đã tách chưa
    _sgk_name_can = suggest_sgk_name(_mon_chon, _lop_chon)
    _meta_sgk = load_sgk_metadata(_sgk_name_can)
    if not _meta_sgk:
        _errors.append(
            f"❌ Chưa tách SGK: **{_sgk_name_can}**. "
            "Vào tab **📤 Upload SGK gốc** hoặc **📚 Nạp bài đã tách**.")

    if _errors:
        for e in _errors:
            st.error(e)
    else:
        # === TÌM BÀI THEO TÊN ===
        sgk_excerpt = ""
        bai_tim_duoc = None

        with st.spinner("📖 Đang tìm bài trong SGK..."):
            bai_tim_duoc = find_bai_by_name(_sgk_name_can, cmd.strip())

        if bai_tim_duoc:
            st.info(
                f"📖 Tìm thấy bài: **{bai_tim_duoc.get('ten_bai', '')}**"
            )
            with st.spinner("📖 Đang lấy nội dung SGK..."):
                text = ocr_bai_pdf(_sgk_name_can, bai_tim_duoc["file"])
                if text:
                    sgk_excerpt = (
                        f"\n=== {bai_tim_duoc.get('ten_bai', '')} ===\n"
                        f"{text}"
                    )
                else:
                    st.warning("⚠️ Không đọc được nội dung bài.")
        else:
            st.warning(
                f"⚠️ Không tìm thấy bài **{cmd}** trong SGK "
                f"`{_sgk_name_can}`. Kiểm tra lại tên bài ở tab "
                f"**✏️ Sửa tên bài SGK** hoặc **📝 Khớp tên với Lịch BG**."
            )

        # === GỌI AI ===
        with st.spinner("🤖 Đang sinh giáo án..."):
            _plan = {
                "task": "Soạn kế hoạch bài dạy",
                "subject": _mon_chon,
                "grade": _lop_chon,
                "topic": cmd,
                "period": _tiet_chon,
                "sgk_excerpt": sgk_excerpt[:8000],
            }
            _prompt = build_exec_prompt(_plan)
            try:
                _result = call_ai(_prompt)
                if _result:
                    st.session_state["draft"] = _result
                    st.session_state["last_topic"] = (
                        _mon_chon + " - " + cmd
                    )
                    save_history("Soạn bài " + _mon_chon, cmd, _result)
                    st.success("✅ Đã soạn xong!")
            except Exception as e:
                st.error("❌ Lỗi AI: " + str(e))

# Hiển thị kết quả
if st.session_state.get("draft"):
    st.divider()
    st.subheader("📄 Kết quả")

    content = st.session_state["draft"]
    safe_name = re.sub(
        r"[^A-Za-z0-9_\-]", "_",
        st.session_state.get("last_topic", "tai_lieu")
    )[:45] or "tai_lieu"

    _d1, _d2 = st.columns(2)
    with _d1:
        st.download_button(
            "⬇️ Tải Word",
            make_doc("Tài liệu", content,
                     subject_name=_mon_chon),
            file_name=safe_name + ".docx",
            mime=("application/vnd.openxmlformats-officedocument."
                  "wordprocessingml.document"),
            type="primary",
            use_container_width=True,
        )
    with _d2:
        st.download_button(
            "⬇️ Tải TXT",
            content.encode("utf-8"),
            file_name=safe_name + ".txt",
            mime="text/plain",
            use_container_width=True,
        )

    st.session_state["draft"] = st.text_area(
        "Nội dung (có thể sửa)", content,
        height=500, key="edit_draft",
    )


# ============ KẾT THÚC PHẦN 5/6 ============
# ============================================================
# PHẦN 6/6: LỊCH BÁO GIẢNG + SOẠN TUẦN + XUẤT WORD
# ============================================================
# Chức năng:
#   - Nạp lịch báo giảng thủ công (Excel/CSV/Word/TXT)
#   - Nhập lịch báo giảng trực tiếp (chỉ TÊN BÀI)
#   - Soạn hàng loạt bài trong tuần (dò theo TÊN)
#   - Xuất Word: giáo án cả tuần + lịch báo giảng
#   - Hiển thị kết quả + tải xuống từng bài
# ============================================================

st.divider()
st.subheader("📅 Soạn cả tuần + Lịch báo giảng")

# ============================================================
# MỤC A: NẠP LỊCH BÁO GIẢNG THỦ CÔNG
# ============================================================

st.markdown("### 📥 Nạp lịch báo giảng có sẵn")

with st.expander("📂 Nạp từ file Excel / CSV / Word / TXT", expanded=False):
    st.caption(
        "💡 Upload file lịch báo giảng đã có sẵn. "
        "App sẽ tự đọc và điền vào bảng bên dưới. "
        "Hỗ trợ: `.xlsx`, `.xls`, `.csv`, `.docx`, `.txt`."
    )

    _lbg_file = st.file_uploader(
        "Chọn file lịch báo giảng",
        type=["xlsx", "xls", "csv", "docx", "txt"],
        key="lbg_upload",
    )

    with st.expander("⚙️ Tùy chỉnh cột (nâng cao)", expanded=False):
        st.caption(
            "Nếu file của bạn có tên cột khác, hãy nhập đúng tên cột "
            "(phân cách bằng dấu phẩy)."
        )
        _col_buoi = st.text_input("Cột Buổi", value="Buổi,Sáng/Chiều",
                                   key="lbg_col_buoi")
        _col_thu = st.text_input("Cột Thứ", value="Thứ,Thứ trong tuần,Ngày",
                                  key="lbg_col_thu")
        _col_tiet = st.text_input("Cột Tiết", value="Tiết,Tiết PPCT,PPCT",
                                   key="lbg_col_tiet")
        _col_mon = st.text_input("Cột Môn", value="Môn,Môn học,Subject",
                                  key="lbg_col_mon")
        _col_lop = st.text_input("Cột Lớp", value="Lớp,Khối,Class",
                                  key="lbg_col_lop")
        _col_bai = st.text_input(
            "Cột Tên bài",
            value="Tên bài,Bài,Tên bài dạy,Chủ đề",
            key="lbg_col_bai")
        _col_sotiet = st.text_input(
            "Cột Số tiết",
            value="Số tiết,Số tiết dạy,Thời lượng",
            key="lbg_col_sotiet")

    def _norm_header(s):
        s = str(s or "").strip().lower()
        s = unicodedata.normalize("NFD", s)
        s = "".join(c for c in s if unicodedata.category(c) != "Mn")
        s = s.replace("đ", "d")
        s = re.sub(r"[^a-z0-9]+", " ", s).strip()
        return s

    def _find_col(headers, candidates_csv):
        cands = [_norm_header(c) for c in candidates_csv.split(",")
                 if c.strip()]
        for i, h in enumerate(headers):
            if _norm_header(h) in cands:
                return i
        for i, h in enumerate(headers):
            hn = _norm_header(h)
            for c in cands:
                if c and (c in hn or hn in c):
                    return i
        return None

    def _parse_lbg_rows(headers, rows):
        idx_buoi = _find_col(headers, _col_buoi)
        idx_thu = _find_col(headers, _col_thu)
        idx_tiet = _find_col(headers, _col_tiet)
        idx_mon = _find_col(headers, _col_mon)
        idx_lop = _find_col(headers, _col_lop)
        idx_bai = _find_col(headers, _col_bai)
        idx_sotiet = _find_col(headers, _col_sotiet)

        out = []
        for r in rows:
            def _get(idx, default=""):
                if idx is None or idx >= len(r):
                    return default
                return str(r[idx] or "").strip()

            ten_bai = _get(idx_bai)
            if not ten_bai:
                continue

            out.append({
                "buoi": _get(idx_buoi, "Sáng") or "Sáng",
                "thu": _get(idx_thu, "Thứ 2") or "Thứ 2",
                "tiet": _get(idx_tiet),
                "mon": _get(idx_mon, "Toán") or "Toán",
                "lop": _get(idx_lop, "6") or "6",
                "ten_bai": ten_bai,
                "so_tiet": _get(idx_sotiet, "1") or "1",
            })
        return out

    def _read_lbg_excel(data):
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        ws = wb[wb.sheetnames[0]]
        all_rows = [
            [str(c) if c is not None else "" for c in row]
            for row in ws.iter_rows(values_only=True)
        ]
        all_rows = [r for r in all_rows if any(c.strip() for c in r)]
        if not all_rows:
            return []
        return _parse_lbg_rows(all_rows[0], all_rows[1:])

    def _read_lbg_csv(data):
        text = None
        for enc in ["utf-8", "utf-8-sig", "latin-1", "cp1258"]:
            try:
                text = data.decode(enc)
                break
            except Exception:
                continue
        if text is None:
            text = data.decode("utf-8", errors="ignore")
        sample = text[:2000]
        delim = "," if sample.count(",") >= sample.count(";") else ";"
        rows = list(csv.reader(io.StringIO(text), delimiter=delim))
        rows = [r for r in rows if any(c.strip() for c in r)]
        if not rows:
            return []
        return _parse_lbg_rows(rows[0], rows[1:])

    def _read_lbg_docx(data):
        d = Document(io.BytesIO(data))
        if d.tables:
            t = d.tables[0]
            all_rows = [[c.text.strip() for c in row.cells]
                        for row in t.rows]
            if len(all_rows) >= 2:
                return _parse_lbg_rows(all_rows[0], all_rows[1:])
        out = []
        for p in d.paragraphs:
            line = p.text.strip()
            if not line:
                continue
            parts = [x.strip() for x in re.split(r"\t|\s{2,}|\|", line)
                     if x.strip()]
            if len(parts) >= 3:
                out.append({
                    "buoi": "Sáng",
                    "thu": parts[0] if len(parts) > 0 else "Thứ 2",
                    "tiet": parts[1] if len(parts) > 1 else "",
                    "mon": parts[2] if len(parts) > 2 else "Toán",
                    "lop": parts[3] if len(parts) > 3 else "6",
                    "ten_bai": parts[4] if len(parts) > 4 else "",
                    "so_tiet": parts[5] if len(parts) > 5 else "1"})
        return out

    def _read_lbg_txt(data):
        text = None
        for enc in ["utf-8", "utf-8-sig", "latin-1", "cp1258"]:
            try:
                text = data.decode(enc)
                break
            except Exception:
                continue
        if text is None:
            text = data.decode("utf-8", errors="ignore")

        lines = [ln.rstrip() for ln in text.splitlines() if ln.strip()]
        if not lines:
            return []

        first_parts = [x.strip() for x in re.split(r"\t|\s{2,}|\|", lines[0])
                       if x.strip()]
        if len(first_parts) >= 3 and any(
            _norm_header(x) in [
                "buoi", "thu", "tiet", "mon", "lop",
                "ten bai", "so tiet"]
            for x in first_parts):
            headers = first_parts
            body_rows = []
            for ln in lines[1:]:
                parts = re.split(r"\t|\s{2,}|\|", ln)
                body_rows.append([x.strip() for x in parts if x.strip()])
            return _parse_lbg_rows(headers, body_rows)

        out = []
        for ln in lines:
            parts = [x.strip() for x in re.split(r"\t|\s{2,}|\|", ln)
                     if x.strip()]
            if len(parts) < 3:
                continue
            out.append({
                "buoi": "Sáng",
                "thu": parts[0],
                "tiet": parts[1] if len(parts) > 1 else "",
                "mon": parts[2] if len(parts) > 2 else "Toán",
                "lop": parts[3] if len(parts) > 3 else "6",
                "ten_bai": parts[4] if len(parts) > 4 else "",
                "so_tiet": parts[5] if len(parts) > 5 else "1"})
        return out

    if _lbg_file is not None:
        _lbg_name = _lbg_file.name.lower()
        try:
            _lbg_bytes = _lbg_file.read()
            _parsed = []

            if _lbg_name.endswith((".xlsx", ".xls", ".xlsm")):
                _parsed = _read_lbg_excel(_lbg_bytes)
            elif _lbg_name.endswith(".csv"):
                _parsed = _read_lbg_csv(_lbg_bytes)
            elif _lbg_name.endswith(".docx"):
                _parsed = _read_lbg_docx(_lbg_bytes)
            elif _lbg_name.endswith(".txt"):
                _parsed = _read_lbg_txt(_lbg_bytes)
            else:
                st.error("❌ Định dạng file không được hỗ trợ.")

            if _parsed:
                st.success(
                    f"✅ Đã đọc được **{len(_parsed)}** dòng "
                    f"từ file `{_lbg_file.name}`."
                )

                with st.expander("👁️ Xem trước", expanded=True):
                    _preview = []
                    for _r in _parsed[:10]:
                        _preview.append({
                            "Buổi": _r["buoi"],
                            "Thứ": _r["thu"],
                            "Tiết": _r["tiet"],
                            "Môn": _r["mon"],
                            "Lớp": _r["lop"],
                            "Tên bài": _r["ten_bai"][:40],
                            "Số tiết": _r["so_tiet"],
                        })
                    st.dataframe(_preview, use_container_width=True)

                _c_mode, _c_btn = st.columns([2, 1])
                with _c_mode:
                    _mode = st.radio(
                        "Cách nạp:",
                        ["Thay thế toàn bộ", "Thêm vào cuối"],
                        horizontal=True,
                        key="lbg_mode",
                    )
                with _c_btn:
                    if st.button(
                        "📥 Nạp vào bảng",
                        type="primary",
                        use_container_width=True,
                        key="lbg_apply",
                    ):
                        if _mode == "Thay thế toàn bộ":
                            st.session_state["lich_bao_giang"] = _parsed
                        else:
                            st.session_state["lich_bao_giang"].extend(_parsed)
                        st.success(
                            f"✅ Đã nạp {len(_parsed)} dòng. "
                            "Xem/sửa ở bảng bên dưới."
                        )
                        st.rerun()
            elif _parsed is not None and not _parsed:
                st.warning(
                    "⚠️ File rỗng hoặc không đọc được dòng nào. "
                    "Kiểm tra lại tên cột trong mục **Tùy chỉnh cột**."
                )
        except Exception as _e:
            st.error(f"❌ Lỗi đọc file: {_e}")

# ============================================================
# MỤC B: KHỞI TẠO SESSION + CẤU HÌNH TUẦN
# ============================================================

if "lich_bao_giang" not in st.session_state:
    st.session_state["lich_bao_giang"] = []
if "weekly_results" not in st.session_state:
    st.session_state["weekly_results"] = []

_c1, _c2, _c3, _c4 = st.columns([1, 1.2, 1, 2.5])
with _c1:
    _tuan_so = st.text_input(
        "Tuần số",
        value=st.session_state.get("tuan_so_val", "1"),
        key="tuan_so_val",
    )
with _c2:
    _ngay_bd = st.date_input(
        "Ngày bắt đầu tuần",
        value=datetime.now(),
        key="ngay_bd_val",
    )
with _c3:
    _so_buoi = st.number_input(
        "Số buổi",
        min_value=1, max_value=10, value=5, step=1,
        key="so_buoi_val",
    )
with _c4:
    st.caption(
        "💡 Nhập lịch dạy từng buổi → bấm **Soạn cả tuần** "
        "để AI soạn toàn bộ."
    )

# ============================================================
# MỤC C: BẢNG NHẬP LỊCH BÁO GIẢNG (CHỈ TÊN BÀI)
# ============================================================

st.markdown("### 📋 Nhập lịch báo giảng")
st.caption("Mỗi dòng = 1 tiết dạy. Cột **Tên bài** dùng dropdown từ SGK.")

_danh_sach_mon_all = list(SUBJECT_RULES.keys())

_c_add, _c_del, _c_clear, _c_auto = st.columns([1, 1, 1, 2])

with _c_add:
    if st.button("➕ Thêm dòng", key="add_row_lich",
                 use_container_width=True):
        st.session_state["lich_bao_giang"].append({
            "buoi": "Sáng",
            "thu": "Thứ 2",
            "tiet": str(len(st.session_state["lich_bao_giang"]) + 1),
            "mon": _danh_sach_mon_all[0] if _danh_sach_mon_all else "Toán",
            "lop": st.session_state["memory"].get("default_grade", "6"),
            "ten_bai": "",
            "so_tiet": "1",
        })
        st.rerun()

with _c_del:
    if st.button("➖ Xóa dòng cuối", key="del_row_lich",
                 use_container_width=True):
        if st.session_state["lich_bao_giang"]:
            st.session_state["lich_bao_giang"].pop()
        st.rerun()

with _c_clear:
    if st.button("🗑️ Xóa tất cả", key="clear_lich",
                 use_container_width=True):
        st.session_state["lich_bao_giang"] = []
        st.session_state["weekly_results"] = []
        st.rerun()

with _c_auto:
    if st.button("⚡ Tạo nhanh 5 buổi sáng (T2→T6)",
                 key="auto_gen_lich", use_container_width=True):
        _thu_list = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"]
        _mon_default = st.session_state["memory"].get("default_subject",
                                                       "Toán")
        _lop_default = st.session_state["memory"].get("default_grade", "6")
        for _i, _t in enumerate(_thu_list):
            st.session_state["lich_bao_giang"].append({
                "buoi": "Sáng",
                "thu": _t,
                "tiet": str(_i + 1),
                "mon": _mon_default,
                "lop": _lop_default,
                "ten_bai": "",
                "so_tiet": "1",
            })
        st.rerun()

# Hiển thị bảng nhập
if st.session_state["lich_bao_giang"]:
    st.markdown("---")

    # Header — KHÔNG CÒN CỘT TRANG SGK
    _h = st.columns([0.8, 0.9, 0.5, 1.2, 0.5, 4.0, 0.6])
    for _col, _label in zip(_h, [
        "**Buổi**", "**Thứ**", "**Tiết**", "**Môn**", "**Lớp**",
        "**Tên bài (chọn từ SGK)**", "**Số tiết**"
    ]):
        _col.markdown(_label)

    # Các dòng nhập
    for _i, _row in enumerate(st.session_state["lich_bao_giang"]):
        _cols = st.columns([0.8, 0.9, 0.5, 1.2, 0.5, 4.0, 0.6])

        with _cols[0]:
            _row["buoi"] = st.selectbox(
                "Buổi", ["Sáng", "Chiều"],
                index=0 if _row.get("buoi") == "Sáng" else 1,
                key=f"buoi_{_i}", label_visibility="collapsed",
            )
        with _cols[1]:
            _thu_opts = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5",
                         "Thứ 6", "Thứ 7", "Chủ nhật"]
            _thu_idx = _thu_opts.index(_row.get("thu", "Thứ 2")) \
                if _row.get("thu") in _thu_opts else 0
            _row["thu"] = st.selectbox(
                "Thứ", _thu_opts, index=_thu_idx,
                key=f"thu_{_i}", label_visibility="collapsed",
            )
        with _cols[2]:
            _row["tiet"] = st.text_input(
                "Tiết", value=_row.get("tiet", ""),
                key=f"tiet_{_i}", label_visibility="collapsed",
            )
        with _cols[3]:
            _mon_idx = _danh_sach_mon_all.index(_row.get("mon", "Toán")) \
                if _row.get("mon") in _danh_sach_mon_all else 0
            _row["mon"] = st.selectbox(
                "Môn", _danh_sach_mon_all, index=_mon_idx,
                key=f"mon_{_i}", label_visibility="collapsed",
            )
        with _cols[4]:
            _row["lop"] = st.text_input(
                "Lớp", value=_row.get("lop", "6"),
                key=f"lop_{_i}", label_visibility="collapsed",
            )
        with _cols[5]:
            # Dropdown tên bài từ SGK tương ứng
            _sgk_row = suggest_sgk_name(_row.get("mon", "Toán"),
                                         _row.get("lop", "6"))
            _ds_ten = list_ten_bai(_sgk_row)
            if _ds_ten:
                _cur = _row.get("ten_bai", "")
                _opts = ["— Chọn từ SGK —"] + _ds_ten
                _sel_idx = 0
                if _cur in _opts:
                    _sel_idx = _opts.index(_cur)
                _chon = st.selectbox(
                    "Tên bài", _opts, index=_sel_idx,
                    key=f"tenbai_{_i}",
                    label_visibility="collapsed")
                if _chon and _chon != "— Chọn từ SGK —":
                    _row["ten_bai"] = _chon
            else:
                _row["ten_bai"] = st.text_input(
                    "Tên bài", value=_row.get("ten_bai", ""),
                    key=f"tenbai_txt_{_i}",
                    label_visibility="collapsed",
                    placeholder="Gõ tên bài (chưa có SGK)")
        with _cols[6]:
            _row["so_tiet"] = st.text_input(
                "Số tiết", value=_row.get("so_tiet", "1"),
                key=f"sotiet_{_i}", label_visibility="collapsed",
            )

    st.markdown("---")

    # Nút soạn cả tuần
    _btn_soan_tuan = st.button(
        "🚀 Soạn cả tuần",
        type="primary", use_container_width=True,
        key="btn_soan_tuan",
    )

    if _btn_soan_tuan:
        _errors_tuan = []
        if not available_providers:
            _errors_tuan.append("❌ Chưa có API key.")
        for _idx, _row in enumerate(st.session_state["lich_bao_giang"]):
            if not _row.get("ten_bai", "").strip():
                _errors_tuan.append(
                    f"❌ Dòng {_idx+1}: Chưa chọn/nhập tên bài."
                )

        if _errors_tuan:
            for _e in _errors_tuan:
                st.error(_e)
        else:
            _progress_tuan = st.progress(0)
            _status_tuan = st.empty()
            _results = []
            _total_rows = len(st.session_state["lich_bao_giang"])

            for _idx, _row in enumerate(st.session_state["lich_bao_giang"]):
                _status_tuan.write(
                    f"🤖 Đang soạn {_idx+1}/{_total_rows}: "
                    f"**{_row['ten_bai'][:50]}**"
                )
                _progress_tuan.progress(int(_idx / _total_rows * 100))

                # Tìm bài theo TÊN
                _sgk_name_row = suggest_sgk_name(_row["mon"], _row["lop"])
                _sgk_excerpt_row = ""
                _ten_bai_row = _row.get("ten_bai", "").strip()

                _bai_row = None
                if _ten_bai_row:
                    _bai_row = find_bai_by_name(_sgk_name_row, _ten_bai_row)

                if _bai_row:
                    _text_row = ocr_bai_pdf(_sgk_name_row, _bai_row["file"])
                    if _text_row:
                        _sgk_excerpt_row = (
                            f"\n=== {_bai_row.get('ten_bai', '')} ===\n"
                            f"{_text_row}"
                        )

                # Build prompt
                _plan_row = {
                    "task": "Soạn kế hoạch bài dạy",
                    "subject": _row["mon"],
                    "grade": _row["lop"],
                    "topic": _row["ten_bai"],
                    "period": _row.get("so_tiet", "1"),
                    "sgk_excerpt": _sgk_excerpt_row[:8000],
                }
                _prompt_row = build_exec_prompt(_plan_row)

                try:
                    _result_row = call_ai(_prompt_row)
                    if _result_row:
                        _results.append({
                            "row": _row, "content": _result_row, "ok": True
                        })
                        save_history(
                            f"Soạn bài {_row['mon']}",
                            _row["ten_bai"],
                            _result_row,
                        )
                    else:
                        _results.append({
                            "row": _row, "content": "", "ok": False,
                            "error": "AI không trả về nội dung",
                        })
                except Exception as _e:
                    _results.append({
                        "row": _row, "content": "", "ok": False,
                        "error": str(_e),
                    })

            _progress_tuan.progress(100)
            _status_tuan.write("✅ Hoàn tất!")
            st.session_state["weekly_results"] = _results

            _ok_n = len([r for r in _results if r["ok"]])
            st.success(f"✅ Đã soạn xong {_ok_n}/{_total_rows} bài.")
            st.balloons()

# ============================================================
# MỤC D: HIỂN THỊ KẾT QUẢ + TẢI XUỐNG
# ============================================================

if st.session_state.get("weekly_results"):
    st.divider()
    st.subheader("📄 Kết quả soạn tuần")

    _weekly = st.session_state["weekly_results"]
    _ok_count = len([r for r in _weekly if r["ok"]])
    _fail_count = len(_weekly) - _ok_count

    st.write(
        f"✅ Thành công: **{_ok_count}** | ❌ Lỗi: **{_fail_count}**"
    )

    _col_dl1, _col_dl2, _col_dl3 = st.columns(3)

    # --- Word gộp cả tuần ---
    with _col_dl1:
        if _ok_count > 0:
            _d_tuan = Document()
            _s_tuan = _d_tuan.sections[0]
            _s_tuan.top_margin = Inches(0.7)
            _s_tuan.bottom_margin = Inches(0.7)
            _s_tuan.left_margin = Inches(0.8)
            _s_tuan.right_margin = Inches(0.8)

            _style_tuan = _d_tuan.styles["Normal"]
            _style_tuan.font.name = "Times New Roman"
            _style_tuan.font.size = Pt(14)
            _rpr_t = _style_tuan.element.get_or_add_rPr()
            _rf_t = _rpr_t.find(qn("w:rFonts"))
            if _rf_t is None:
                _rf_t = _rpr_t.makeelement(qn("w:rFonts"), {})
                _rpr_t.append(_rf_t)
            for _at in ["w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"]:
                _rf_t.set(qn(_at), "Times New Roman")
            _pf_t = _style_tuan.paragraph_format
            _pf_t.line_spacing_rule = WD_LINE_SPACING.SINGLE
            _pf_t.space_before = Pt(0)
            _pf_t.space_after = Pt(0)

            _d_tuan.add_heading(f"KẾ HOẠCH BÀI DẠY TUẦN {_tuan_so}", 0)
            _mem_t = st.session_state.get("memory", {})
            if _mem_t.get("school"):
                _d_tuan.add_paragraph("Đơn vị: " + _mem_t["school"])
            if _mem_t.get("teacher_name"):
                _d_tuan.add_paragraph(
                    "Người thực hiện: " + _mem_t["teacher_name"])
            _d_tuan.add_paragraph(
                "Ngày tạo: " + datetime.now().strftime("%d/%m/%Y %H:%M"))
            _d_tuan.add_paragraph("")

            for _r in _weekly:
                if _r["ok"]:
                    _ri = _r["row"]
                    _d_tuan.add_page_break()
                    _d_tuan.add_heading(
                        f"{_ri['mon']} {_ri['lop']} - {_ri['ten_bai']}", 1)
                    write_content_to_doc(_d_tuan, _r["content"])

            _b_tuan = io.BytesIO()
            _d_tuan.save(_b_tuan)

            st.download_button(
                "⬇️ Tải Word cả tuần",
                _b_tuan.getvalue(),
                file_name=f"KeHoachBaiDay_Tuan{_tuan_so}.docx",
                mime=("application/vnd.openxmlformats-officedocument."
                      "wordprocessingml.document"),
                type="primary",
                use_container_width=True,
            )
        else:
            st.button("⬇️ Tải Word cả tuần", disabled=True,
                      use_container_width=True)

    # --- Word lịch báo giảng ---
    with _col_dl2:
        if st.session_state["lich_bao_giang"]:
            _d_lich = Document()
            _s_lich = _d_lich.sections[0]
            _s_lich.top_margin = Inches(0.7)
            _s_lich.bottom_margin = Inches(0.7)
            _s_lich.left_margin = Inches(0.8)
            _s_lich.right_margin = Inches(0.8)

            _style_l = _d_lich.styles["Normal"]
            _style_l.font.name = "Times New Roman"
            _style_l.font.size = Pt(13)
            _rpr_l = _style_l.element.get_or_add_rPr()
            _rf_l = _rpr_l.find(qn("w:rFonts"))
            if _rf_l is None:
                _rf_l = _rpr_l.makeelement(qn("w:rFonts"), {})
                _rpr_l.append(_rf_l)
            for _at in ["w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"]:
                _rf_l.set(qn(_at), "Times New Roman")

            _d_lich.add_heading(f"LỊCH BÁO GIẢNG TUẦN {_tuan_so}", 0)
            _mem_l = st.session_state.get("memory", {})
            if _mem_l.get("school"):
                _d_lich.add_paragraph("Đơn vị: " + _mem_l["school"])
            if _mem_l.get("teacher_name"):
                _d_lich.add_paragraph(
                    "Giáo viên: " + _mem_l["teacher_name"])
            _d_lich.add_paragraph(
                "Tuần: " + str(_tuan_so)
                + " | Từ ngày: " + _ngay_bd.strftime("%d/%m/%Y"))
            _d_lich.add_paragraph("")

            # Bảng lịch — BỎ CỘT TRANG SGK
            _lich_rows = [[
                "Buổi", "Thứ", "Tiết", "Môn", "Lớp",
                "Tên bài", "Số tiết"
            ]]
            for _row in st.session_state["lich_bao_giang"]:
                _lich_rows.append([
                    _row.get("buoi", ""),
                    _row.get("thu", ""),
                    _row.get("tiet", ""),
                    _row.get("mon", ""),
                    _row.get("lop", ""),
                    _row.get("ten_bai", ""),
                    _row.get("so_tiet", "1"),
                ])

            _t_lich = _d_lich.add_table(rows=len(_lich_rows), cols=7)
            _t_lich.style = "Table Grid"

            for _i, _r in enumerate(_lich_rows):
                for _j in range(7):
                    _cell = _t_lich.cell(_i, _j)
                    _cell.text = ""
                    _p = _cell.paragraphs[0]
                    _p.paragraph_format.line_spacing_rule = \
                        WD_LINE_SPACING.SINGLE
                    _p.paragraph_format.space_before = Pt(0)
                    _p.paragraph_format.space_after = Pt(0)

                    if _i == 0:
                        _p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        _run = _p.add_run(_r[_j])
                        _set_run_font(_run, 13, bold=True)
                    else:
                        _p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        _run = _p.add_run(_r[_j])
                        _set_run_font(_run, 13, bold=False)

            _b_lich = io.BytesIO()
            _d_lich.save(_b_lich)

            st.download_button(
                "⬇️ Tải lịch báo giảng",
                _b_lich.getvalue(),
                file_name=f"LichBaoGiang_Tuan{_tuan_so}.docx",
                mime=("application/vnd.openxmlformats-officedocument."
                      "wordprocessingml.document"),
                use_container_width=True,
            )
        else:
            st.button("⬇️ Tải lịch báo giảng", disabled=True,
                      use_container_width=True)

    # --- TXT cả tuần ---
    with _col_dl3:
        if _ok_count > 0:
            _txt_all = ""
            for _r in _weekly:
                if _r["ok"]:
                    _ri = _r["row"]
                    _txt_all += (
                        f"\n\n{'='*60}\n"
                        f"{_ri['mon']} {_ri['lop']} - {_ri['ten_bai']}\n"
                        f"{'='*60}\n\n"
                        f"{_r['content']}"
                    )

            st.download_button(
                "⬇️ Tải TXT cả tuần",
                _txt_all.encode("utf-8"),
                file_name=f"KeHoachBaiDay_Tuan{_tuan_so}.txt",
                mime="text/plain",
                use_container_width=True,
            )
        else:
            st.button("⬇️ Tải TXT cả tuần", disabled=True,
                      use_container_width=True)

    # --- Chi tiết từng bài ---
    st.markdown("### 📑 Chi tiết từng bài")

    for _idx, _r in enumerate(_weekly):
        _ri = _r["row"]
        _icon = "✅" if _r["ok"] else "❌"

        with st.expander(
            f"{_icon} {_ri['thu']} - Tiết {_ri['tiet']} - "
            f"{_ri['mon']} {_ri['lop']}: {_ri['ten_bai'][:50]}",
            expanded=False,
        ):
            if _r["ok"]:
                _edited = st.text_area(
                    "Nội dung (có thể sửa)",
                    _r["content"],
                    height=300,
                    key=f"edit_weekly_{_idx}",
                )
                _r["content"] = _edited

                _safe_name_row = re.sub(
                    r"[^A-Za-z0-9_\-]", "_", _ri["ten_bai"]
                )[:45] or "bai"

                _col_r1, _col_r2 = st.columns(2)
                with _col_r1:
                    st.download_button(
                        "⬇️ Tải Word bài này",
                        make_doc("Tài liệu", _edited,
                                 subject_name=_ri["mon"]),
                        file_name=f"{_safe_name_row}.docx",
                        mime=("application/vnd.openxmlformats-"
                              "officedocument.wordprocessingml.document"),
                        key=f"dl_word_weekly_{_idx}",
                        use_container_width=True,
                    )
                with _col_r2:
                    st.download_button(
                        "⬇️ Tải TXT bài này",
                        _edited.encode("utf-8"),
                        file_name=f"{_safe_name_row}.txt",
                        mime="text/plain",
                        key=f"dl_txt_weekly_{_idx}",
                        use_container_width=True,
                    )
            else:
                st.error(
                    "Lỗi: " + _r.get("error", "Không rõ nguyên nhân")
                )

else:
    st.info(
        "💡 Bấm **➕ Thêm dòng** hoặc **📥 Nạp từ file** "
        "để bắt đầu nhập lịch báo giảng."
    )

# ============================================================
# MỤC E: FOOTER + THÔNG TIN
# ============================================================

st.divider()

with st.expander("ℹ️ Thông tin ứng dụng", expanded=False):
    st.markdown(
        """
        **🧠 Trợ lý Giáo viên THCS — Phiên bản 2.3**

        - **Mẫu:** CV 5512 chuẩn Bộ GD&ĐT
        - **AI:** Đa nhà cung cấp (Gemini / Groq / OpenRouter / NVIDIA)
        - **Lưu trữ:** GitHub (memory, templates, sources, schedule, history)
        - **Cache:** Local (OCR ảnh, SGK đã tách theo bài)
        - **Môn hỗ trợ:** 12 môn (Toán, Ngữ văn, **Tiếng Anh**,
          KHTN, Lịch sử, Địa lí, GDCD, Mỹ thuật, Âm nhạc, GDTC,
          HĐTN-HN, Tin học)

        **📚 Hướng dẫn nhanh:**

        1. **Upload SGK gốc** → Tách tự động theo bài
        2. **Nạp bài đã tách** → Nếu tự tách thủ công
        3. **Nạp mẫu giáo án** → Mẫu 5512 riêng của trường
        4. **Nạp ma trận đề** → Sinh đặc tả + đề + đáp án
        5. **✏️ Sửa tên bài SGK** → Chỉnh tên nếu OCR sai
        6. **📝 Khớp tên với Lịch BG** → Dán tên từ Lịch BG
        7. **Soạn 1 bài** → Chọn môn, lớp, gõ tên bài
        8. **Soạn cả tuần** → Nhập Lịch BG, chọn bài từ dropdown
        9. **Tải Word** → Xuất file chuẩn 5512

        **⚠️ Lưu ý:**

        - Cần `GEMINI_API_KEY` để OCR mục lục SGK
        - Cần `GITHUB_TOKEN` để lưu trữ lâu dài
        - Soạn bài dò theo **TÊN BÀI** (không dùng trang PDF)
        - Nếu tên bài sai → vào tab **✏️ Sửa tên bài SGK**
        """
    )

st.caption(
    "🧠 Trợ lý Giáo viên THCS | "
    "Bộ não tự học • Mẫu CV 5512 • Sách Kết nối tri thức • Xuất Word"
)

# ============ KẾT THÚC PHẦN 6/6 ============
