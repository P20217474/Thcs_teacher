# ============================================================
# PHẦN 1/6: HẠ TẦNG — IMPORTS + CONFIG + GITHUB API + SESSION
# ============================================================
# Chức năng:
#   - Tự động cài package thiếu
#   - Import toàn bộ thư viện
#   - Cấu hình Streamlit page
#   - Đọc secrets (API keys, GitHub token)
#   - GitHub API: get, save, delete, list, test
#   - Load memory/templates/sources/schedule
#   - Khởi tạo session_state
# ============================================================

import subprocess
import sys
import os
import json
import base64
import urllib.request
from datetime import datetime, timedelta


# ---------- TỰ ĐỘNG CÀI PACKAGE THIẾU ----------
def _ensure(pkg, mod=None):
    """Cài package nếu chưa có."""
    mod = mod or pkg
    try:
        __import__(mod)
    except ImportError:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", pkg, "--quiet"]
        )


_ensure("streamlit")
_ensure("python-docx", "docx")
_ensure("pypdf", "pypdf")
_ensure("pymupdf", "fitz")  # Tách trang PDF


# ---------- IMPORT CHÍNH ----------
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
import io
import re
import csv


# ---------- CÁC FLAG (CÓ HAY KHÔNG) ----------
try:
    from pypdf import PdfReader
    HAS_PYPDF = True
except Exception:
    HAS_PYPDF = False

try:
    import fitz  # pymupdf
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


# ---------- CẤU HÌNH TRANG ----------
st.set_page_config(
    page_title="Trợ lý Giáo viên THCS",
    page_icon="🧠",
    layout="wide",
)
st.title("🧠 Trợ lý Giáo viên THCS")
st.caption("Bộ não tự học • Mẫu CV 5512 • Sách Kết nối tri thức • Xuất Word")


# ---------- ĐỌC SECRETS ----------
def cfg(key, default=""):
    """Đọc secret an toàn."""
    try:
        val = st.secrets.get(key, default)
        return val if val else default
    except Exception:
        return default


# ---------- CẤU HÌNH PROVIDER ----------
PROVIDER_KEYS = {
    "GEMINI_API_KEY": cfg("GEMINI_API_KEY", ""),
    "GROQ_API_KEY": cfg("GROQ_API_KEY", ""),
    "OPENROUTER_API_KEY": cfg("OPENROUTER_API_KEY", ""),
    "NVIDIA_API_KEY": cfg("NVIDIA_API_KEY", ""),
}

for _k, _v in PROVIDER_KEYS.items():
    if _v:
        os.environ[_k] = _v

available_providers = [
    _k for _k, _v in PROVIDER_KEYS.items() if _v
]
GEMINI_KEY = PROVIDER_KEYS.get("GEMINI_API_KEY", "")


# Model ưu tiên từng provider (rẻ + nhanh)
PREFERRED_MODELS = {
    "GEMINI_API_KEY": [
        "gemini-2.5-flash",
        "gemini-2.0-flash",
        "gemini-flash-latest",
    ],
    "GROQ_API_KEY": [
        "llama-3.3-70b-versatile",
        "llama-3.1-70b-versatile",
    ],
    "OPENROUTER_API_KEY": [
        "google/gemini-2.0-flash-exp:free",
    ],
    "NVIDIA_API_KEY": [
        "meta/llama-3.3-70b-instruct",
    ],
}


# ---------- CẤU HÌNH GITHUB ----------
GITHUB_TOKEN = cfg("GITHUB_TOKEN", "")
GITHUB_REPO = cfg("GITHUB_REPO", "P20217474/Thcs_teacher")
GITHUB_BRANCH = cfg("GITHUB_BRANCH", "main")


# ---------- ĐƯỜNG DẪN CACHE LOCAL ----------
CACHE_DIR = "cache"
SGK_BAI_DIR = os.path.join(CACHE_DIR, "sgk_bai")
SGK_PAGES_DIR = os.path.join(CACHE_DIR, "sgk_pages")
OCR_CACHE_DIR = os.path.join(CACHE_DIR, "ocr")

# Tạo sẵn các thư mục
for _d in [CACHE_DIR, SGK_BAI_DIR, SGK_PAGES_DIR, OCR_CACHE_DIR]:
    os.makedirs(_d, exist_ok=True)


# ---------- GITHUB API ----------
def gh_api(method, path, data=None):
    """Gọi GitHub API."""
    if not GITHUB_TOKEN:
        return {"error": "Chưa có GITHUB_TOKEN"}

    url = "https://api.github.com" + path
    headers = {
        "Authorization": "token " + GITHUB_TOKEN,
        "Accept": "application/vnd.github+json",
        "User-Agent": "agent-thcs",
    }
    body = json.dumps(data).encode("utf-8") if data else None

    req = urllib.request.Request(
        url, data=body, headers=headers, method=method
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        return {"error": str(e)}


def gh_get(path):
    """Đọc file từ GitHub. Trả về (content, sha)."""
    res = gh_api(
        "GET",
        f"/repos/{GITHUB_REPO}/contents/{path}?ref={GITHUB_BRANCH}",
    )
    if not res or "error" in res or "content" not in res:
        return (None, None)
    try:
        content = base64.b64decode(res["content"]).decode("utf-8")
        return (content, res["sha"])
    except Exception:
        return (None, None)


def gh_save(path, content, message="Update"):
    """Lưu file lên GitHub. Trả về (ok, msg)."""
    if not GITHUB_TOKEN:
        return (False, "Chưa có GITHUB_TOKEN")

    old, sha = gh_get(path)
    data = {
        "message": message,
        "content": base64.b64encode(
            content.encode("utf-8")
        ).decode("ascii"),
        "branch": GITHUB_BRANCH,
    }
    if sha:
        data["sha"] = sha

    res = gh_api(
        "PUT",
        f"/repos/{GITHUB_REPO}/contents/{path}",
        data,
    )
    if isinstance(res, dict) and "error" in res:
        return (False, "Lỗi: " + str(res["error"]))
    if isinstance(res, dict) and "content" in res:
        return (True, "Đã lưu: " + path)
    return (False, "Phản hồi không hợp lệ")


def gh_delete(path):
    """Xóa file trên GitHub. Trả về (ok, msg)."""
    if not GITHUB_TOKEN:
        return (False, "Chưa có GITHUB_TOKEN")

    old, sha = gh_get(path)
    if not sha:
        return (False, "File không tồn tại")

    res = gh_api(
        "DELETE",
        f"/repos/{GITHUB_REPO}/contents/{path}",
        {
            "message": "Delete " + path,
            "sha": sha,
            "branch": GITHUB_BRANCH,
        },
    )
    if isinstance(res, dict) and "error" in res:
        return (False, "Lỗi: " + str(res["error"]))
    if isinstance(res, dict) and ("content" in res or "commit" in res):
        return (True, "Đã xóa: " + path)
    return (False, "Phản hồi không hợp lệ")


def gh_list(folder):
    """Liệt kê file trong folder GitHub."""
    res = gh_api(
        "GET",
        f"/repos/{GITHUB_REPO}/contents/{folder}?ref={GITHUB_BRANCH}",
    )
    if not res or "error" in res or not isinstance(res, list):
        return []
    return [
        item["name"] for item in res if item["type"] == "file"
    ]


def gh_test():
    """Test GitHub token + quyền ghi. Trả về (ok, msg)."""
    if not GITHUB_TOKEN:
        return (False, "Thiếu GITHUB_TOKEN")

    res = gh_api("GET", "/user")
    if isinstance(res, dict) and "login" in res:
        login = res["login"]
        # Kiểm tra quyền ghi repo
        res2 = gh_api("GET", f"/repos/{GITHUB_REPO}")
        if isinstance(res2, dict) and "permissions" in res2:
            perms = res2["permissions"]
            if perms.get("push") or perms.get("admin"):
                return (True, "OK — " + login + " | Có quyền ghi")
            return (False, "Token thiếu quyền 'repo'")
        return (True, "OK — " + login)
    if isinstance(res, dict) and "error" in res:
        return (False, "Lỗi: " + res["error"])
    return (False, "Token không hợp lệ")


# ---------- LOAD / SAVE DỮ LIỆU ----------
def load_memory():
    """Đọc memory.json từ GitHub."""
    content, _ = gh_get("data/memory.json")
    if content:
        try:
            return json.loads(content)
        except Exception:
            pass
    return {
        "teacher_name": "",
        "school": "",
        "default_subject": "Toán",
        "default_grade": "6",
    }


def save_memory(memory):
    """Lưu memory.json lên GitHub."""
    memory["last_updated"] = datetime.now().strftime("%Y-%m-%d")
    return gh_save(
        "data/memory.json",
        json.dumps(memory, ensure_ascii=False, indent=2),
        "Update memory",
    )


def load_folder(folder):
    """Đọc toàn bộ file trong folder trên GitHub."""
    files = gh_list(f"data/{folder}")
    out = {}
    for f in files:
        c, _ = gh_get(f"data/{folder}/{f}")
        if c:
            out[f] = c
    return out


# Wrapper các folder
def load_templates(): return load_folder("templates")
def load_sources():   return load_folder("sources")
def load_schedule():  return load_folder("schedule")


def save_history(task, topic, content):
    """Lưu lịch sử soạn bài lên GitHub."""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe = re.sub(r"[^a-zA-Z0-9]", "_", topic)[:30]
    return gh_save(
        f"data/history/{ts}_{safe}.md",
        f"# {task}\n\n**Chủ đề:** {topic}\n\n---\n\n{content}",
        f"History {topic}",
    )


# ---------- KHỞI TẠO SESSION STATE ----------
# Reset session nếu version thay đổi
if "app_started_v21" not in st.session_state:
    for k in list(st.session_state.keys()):
        if k.startswith("tuan_") or k.startswith("buoi_") or \
           k.startswith("ngay_bd_"):
            del st.session_state[k]
    st.session_state["app_started_v21"] = True


# Giá trị mặc định
defaults = {
    "memory": load_memory(),
    "templates": load_templates(),
    "sources": load_sources(),
    "schedule": load_schedule(),
    "draft": "",
    "last_topic": "",
    "weekly_results": [],
    "lich_bao_giang": [],
    # Mẫu 5512 đang dùng
    "mau_5512_dang_dung": None,   # None = dùng hardcode
    # Cache OCR trong session
    "ocr_session_cache": {},
}

for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ---------- KẾT THÚC PHẦN 1/6 ----------
# ============================================================
# PHẦN 2/6: ĐỌC FILE + OCR + TÁCH SGK THEO BÀI + CACHE
# ============================================================
# Chức năng:
#   - Đọc file: PDF, Word, Excel, CSV, TXT, ảnh
#   - OCR ảnh bằng Gemini Vision (cache local + session)
#   - Tách SGK PDF theo BÀI (dùng pymupdf)
#   - Gom nhóm trang thành file bài riêng
#   - Tạo metadata.json mapping bài → trang
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

    # 1. Kiểm tra cache session
    cache_key = _ocr_cache_key(data)
    if cache_key in st.session_state.get("ocr_session_cache", {}):
        return st.session_state["ocr_session_cache"][cache_key]

    # 2. Kiểm tra cache local
    local_cache_file = os.path.join(OCR_CACHE_DIR, cache_key + ".txt")
    if os.path.exists(local_cache_file):
        try:
            with open(local_cache_file, "r", encoding="utf-8") as f:
                text = f.read()
            # Lưu vào session cache
            st.session_state["ocr_session_cache"][cache_key] = text
            return text
        except Exception:
            pass

    # 3. Gọi API Gemini Vision
    try:
        genai_vision.configure(api_key=GEMINI_KEY)
        img = Image.open(io.BytesIO(data))
        for mn in ["gemini-2.0-flash", "gemini-2.5-flash", "gemini-flash-latest"]:
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
                    # Lưu cache session
                    st.session_state["ocr_session_cache"][cache_key] = text
                    # Lưu cache local
                    try:
                        with open(local_cache_file, "w", encoding="utf-8") as f:
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
# PHẦN B: ĐỌC FILE AN TOÀN (PDF, WORD, EXCEL, CSV, TXT, ẢNH)
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
                    "| " + " | ".join(c.text.strip() for c in row.cells) + " |"
                )
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

    # Ảnh
    if nl.endswith((".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp")):
        text = ocr_image_with_gemini(data, name) if use_ocr \
            else "[" + name + ": ảnh]"
        return (name, text, len(text))

    # PDF
    if magic[:4] == b"%PDF":
        text = _read_pdf_safe(data)
    # Office Open XML (docx, xlsx)
    elif magic[:4] == b"PK\x03\x04":
        text = _read_excel_safe(data) if nl.endswith((".xlsx", ".xlsm")) \
            else _read_docx_safe(data)
    # DOC cũ
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
# PHẦN C: TÁCH SGK PDF THEO BÀI (DÙNG PYMUPDF)
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
    """Trích mục lục (Table of Contents) của PDF nếu có.
    Trả về list [(level, title, page_num), ...]
    """
    if not HAS_PYMUPDF:
        return []
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        toc = doc.get_toc()  # [[level, title, page], ...]
        doc.close()
        return toc
    except Exception:
        return []


def detect_chapter_boundaries(pdf_bytes, max_pages_check=20):
    """Dò ranh giới chương/bài trong PDF scan ảnh.

    Cách tiếp cận: OCR 20 trang đầu tiên (mục lục + trang đầu),
    tìm các dòng dạng "Bài X." hoặc "Chương X" → đoán bài.

    Trả về list [{ten_bai, page_start, page_end}, ...]
    """
    if not HAS_PYMUPDF or not HAS_GEMINI_VISION:
        return []

    total_pages = get_pdf_total_pages(pdf_bytes)
    if total_pages == 0:
        return []

    # OCR mục lục thường ở 3-10 trang đầu
    boundaries = []
    scanned_text = []

    for i in range(min(max_pages_check, total_pages)):
        img_bytes = _render_pdf_page_to_image(pdf_bytes, i, dpi=150)
        if not img_bytes:
            continue
        text = ocr_image_with_gemini(img_bytes, f"toc_page_{i+1}")
        if text and not text.startswith("["):
            scanned_text.append((i, text))

    # Trích các dòng dạng "Bài N. Tên" hoặc "Chương N"
    pattern = re.compile(
        r"^\s*(?:Bài|BÀI)\s+(\d+)[\.\s]+\s*(.+?)(?:\s*\.{2,}\s*(\d+))?\s*$",
        re.MULTILINE
    )
    pattern_chuong = re.compile(
        r"^\s*(?:Chương|CHƯƠNG)\s+([IVXLC\d]+)\s*[:\.\s]\s*(.+?)$",
        re.MULTILINE
    )

    found_items = []
    for page_idx, text in scanned_text:
        # Tìm "Bài N. Tên bài ... số trang"
        for m in pattern.finditer(text):
            so_bai = int(m.group(1))
            ten_bai = m.group(2).strip()
            page_in_sgk = m.group(3)
            found_items.append({
                "so_bai": so_bai,
                "ten_bai": ten_bai,
                "page_in_sgk": int(page_in_sgk) if page_in_sgk else None,
                "page_pdf_guess": page_idx,
            })

    # Sắp xếp theo số bài
    found_items.sort(key=lambda x: x["so_bai"])

    # Nếu không có mục lục → trả về rỗng
    if not found_items:
        return []

    # Tính offset nếu có page_in_sgk
    offset = None
    for item in found_items:
        if item["page_in_sgk"] is not None:
            offset = item["page_pdf_guess"] - item["page_in_sgk"] + 1
            break

    # Tạo boundaries từ danh sách bài
    for i, item in enumerate(found_items):
        # page_pdf của bài
        if item["page_in_sgk"] is not None and offset is not None:
            page_pdf_start = item["page_in_sgk"] + offset
        else:
            page_pdf_start = item["page_pdf_guess"]

        # page_pdf kết thúc = page_pdf bắt đầu bài tiếp theo - 1
        if i + 1 < len(found_items):
            next_item = found_items[i + 1]
            if next_item["page_in_sgk"] is not None and offset is not None:
                page_pdf_end = next_item["page_in_sgk"] + offset - 1
            else:
                page_pdf_end = next_item["page_pdf_guess"] - 1
        else:
            page_pdf_end = total_pages - 1

        # Đảm bảo trong khoảng hợp lệ
        page_pdf_start = max(0, min(page_pdf_start, total_pages - 1))
        page_pdf_end = max(page_pdf_start, min(page_pdf_end, total_pages - 1))

        boundaries.append({
            "ten_bai": item["ten_bai"],
            "so_bai": item["so_bai"],
            "page_pdf_start": page_pdf_start,
            "page_pdf_end": page_pdf_end,
        })

    return boundaries


def split_sgk_into_lessons(pdf_bytes, sgk_name, boundaries, progress_cb=None):
    """Tách SGK PDF thành các file bài riêng.

    Args:
        pdf_bytes: bytes PDF gốc
        sgk_name: tên SGK (VD: "SGK_Toan_6")
        boundaries: list [{ten_bai, so_bai, page_pdf_start, page_pdf_end}]
        progress_cb: callback(percent, msg) để hiển thị tiến độ

    Returns:
        dict metadata {sgk_name, ds_bai: [...]}
    """
    if not HAS_PYMUPDF:
        return None

    # Tạo thư mục cho SGK
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
                f"Đang tách bài {idx+1}/{total}: {b['ten_bai'][:40]}"
            )

        # Tên file
        safe_ten = _slugify(b["ten_bai"])[:40] or f"bai_{b['so_bai']:02d}"
        fname = f"bai_{b['so_bai']:02d}_{safe_ten}.pdf"
        fpath = os.path.join(sgk_dir, fname)

        # Tạo PDF mới chứa các trang của bài
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
        except Exception as e:
            # Bỏ qua bài lỗi
            continue

    # Lưu metadata.json
    meta_path = os.path.join(sgk_dir, "metadata.json")
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(metadata, f, ensure_ascii=False, indent=2)

    return metadata


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


def find_bai_by_page(sgk_name, page_pdf):
    """Tìm bài chứa trang PDF cho trước."""
    meta = load_sgk_metadata(sgk_name)
    if not meta:
        return None
    for bai in meta.get("ds_bai", []):
        if bai["page_pdf_start"] <= page_pdf <= bai["page_pdf_end"]:
            return bai
    return None


def ocr_bai_pdf(sgk_name, bai_file):
    """OCR toàn bộ file PDF của 1 bài → text.
    Trả về text gộp tất cả trang.
    """
    fpath = os.path.join(SGK_BAI_DIR, sgk_name, bai_file)
    if not os.path.exists(fpath):
        return ""

    with open(fpath, "rb") as f:
        pdf_bytes = f.read()

    # Đọc text layer nếu có
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total = len(doc)
    doc.close()

    parts = []
    for i in range(total):
        # 1. Thử đọc text layer
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


def get_sgk_file_by_name(mon, lop):
    """Lấy ĐÚNG file SGK_[Môn]_[Lớp] từ kho Nguồn.
    Không dò tìm nhiều file.
    """
    if not mon or not lop:
        return (None, None)

    key = "SGK_" + _slugify(mon).title().replace("_", "") + "_" + str(lop)
    sources = st.session_state.get("sources", {})

    # Ưu tiên 1: khớp chính xác "SGK_Toan_6"
    for fname in sources.keys():
        base = os.path.splitext(fname)[0]
        if base.lower() == key.lower():
            return (fname, sources[fname])

    # Ưu tiên 2: khớp không phân biệt hoa thường
    for fname in sources.keys():
        base = os.path.splitext(fname)[0]
        if base.lower().startswith("sgk_") and \
           str(lop) in base and \
           _slugify(mon) in _slugify(base):
            return (fname, sources[fname])

    return (None, None)


def suggest_sgk_name(mon, lop):
    """Đề xuất tên SGK chuẩn."""
    return "SGK_" + _slugify(mon).title().replace("_", "") + "_" + str(lop)


# ---------- KẾT THÚC PHẦN 2/6 ----------
# ============================================================
# PHẦN 3/6: SIDEBAR + UPLOAD + CALL AI
# ============================================================
# Chức năng:
#   - Sidebar: test GitHub, memory, thống kê SGK đã tách, dọn dẹp
#   - Tab Upload SGK gốc: upload PDF → tách tự động
#   - Tab Nạp bài thủ công: upload file bài riêng
#   - Tab Kho tri thức
#   - Hàm call_ai() đa provider
# ============================================================


# ============================================================
# PHẦN A: SIDEBAR
# ============================================================
with st.sidebar:
    st.header("🧠 Bộ não")
    mem = st.session_state["memory"]

    # Trạng thái GitHub
    if GITHUB_TOKEN:
        st.success("💾 Đã có GITHUB_TOKEN")
        if st.button("🧪 Test GitHub", use_container_width=True):
            ok, msg = gh_test()
            st.success("✅ " + msg) if ok else st.error("❌ " + msg)
    else:
        st.error("❌ Chưa có GITHUB_TOKEN")

    # Trạng thái Gemini
    if GEMINI_KEY:
        st.success("🔍 Gemini Vision sẵn sàng")
    else:
        st.warning("⚠️ Chưa có GEMINI_API_KEY")

    # Thông tin giáo viên
    mem["teacher_name"] = st.text_input(
        "Giáo viên", value=mem.get("teacher_name", "")
    )
    mem["school"] = st.text_input(
        "Trường", value=mem.get("school", "")
    )
    mem["default_subject"] = st.text_input(
        "Môn mặc định", value=mem.get("default_subject", "Toán")
    )
    mem["default_grade"] = st.text_input(
        "Khối mặc định", value=mem.get("default_grade", "6")
    )

    if st.button("💾 Lưu vào bộ nhớ", use_container_width=True):
        ok, msg = save_memory(mem)
        st.success("✅ " + msg) if ok else st.error("❌ " + msg)

    st.divider()

    # Thống kê
    st.subheader("📊 Thống kê")
    for label, key in [("Mẫu", "templates"), ("Nguồn", "sources"),
                       ("Lịch BG", "schedule")]:
        st.write(label + ": " + str(len(st.session_state[key])) + " file")

    # Danh sách SGK đã tách
    ds_sgk = list_sgk_da_tach()
    st.write("SGK đã tách: " + str(len(ds_sgk)) + " file")
    if ds_sgk:
        with st.expander("📚 Xem SGK đã tách", expanded=False):
            for sgk in ds_sgk:
                st.write("• " + sgk["sgk_name"] + " — "
                         + str(sgk["so_bai"]) + " bài ("
                         + sgk["nguon"] + ")")

    if st.button("🔄 Tải lại", use_container_width=True):
        st.session_state["templates"] = load_templates()
        st.session_state["sources"] = load_sources()
        st.session_state["schedule"] = load_schedule()
        st.rerun()

    st.write("AI: " + str(len(available_providers)) + " nhà cung cấp")

    st.divider()

    # Dọn dẹp
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
            st.success("✅ Đã xóa " + str(deleted) + " file OCR cache")

        if GITHUB_TOKEN:
            if st.button("📜 Xóa lịch sử soạn bài",
                         use_container_width=True, key="clean_history"):
                files = gh_list("data/history")
                deleted = 0
                for f in files:
                    ok, _ = gh_delete(f"data/history/{f}")
                    if ok:
                        deleted += 1
                st.success("✅ Đã xóa " + str(deleted) + " file lịch sử")

        st.divider()
        st.caption("⚠️ **Cảnh báo**: các nút dưới đây xóa VĨNH VIỄN")

        if st.button("🔴 Reset toàn bộ app",
                     use_container_width=True, key="reset_all"):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.success("✅ Đã reset. Đang tải lại...")
            st.rerun()


# ============================================================
# PHẦN B: HÀM GỌI AI (ĐA PROVIDER + FALLBACK)
# ============================================================
def call_ai(prompt, system=""):
    """Gọi AI đa provider với fallback.
    - Ưu tiên provider theo danh sách available_providers
    - Thử từng model trong PREFERRED_MODELS
    - Nếu tất cả lỗi → thử model="auto"
    """
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
                t = getattr(resp, "text", None) or getattr(resp, "content", "")
                if t and len(t.strip()) > 50:
                    return t
            except Exception as e:
                last_err = e
                continue

    # Fallback cuối
    try:
        resp = llm.chat(msgs, model="auto")
        return getattr(resp, "text", None) or getattr(resp, "content", "")
    except Exception as e:
        if last_err:
            raise last_err
        raise e


def call_ai_cached(prompt, system=""):
    """Gọi AI có cache theo hash prompt."""
    cache_key = hashlib.md5(
        (system + "||" + prompt).encode("utf-8")
    ).hexdigest()

    # Kiểm tra session cache
    cache_store = st.session_state.setdefault("ai_prompt_cache", {})
    if cache_key in cache_store:
        return cache_store[cache_key]

    # Gọi AI
    result = call_ai(prompt, system)
    if result:
        cache_store[cache_key] = result
    return result


# ============================================================
# PHẦN C: UPLOAD — 3 TAB (SGK GỐC, BÀI THỦ CÔNG, KHO)
# ============================================================
tab_sgk, tab_manual, tab_kho = st.tabs([
    "📤 Upload SGK gốc",
    "📚 Nạp bài đã tách",
    "📂 Kho tri thức",
])


# ---------- TAB 1: UPLOAD SGK GỐC + TÁCH TỰ ĐỘNG ----------
with tab_sgk:
    st.markdown("### 📤 Upload SGK gốc (PDF)")
    st.info(
        "💡 Upload file SGK PDF (scan ảnh). App sẽ **tách tự động theo bài** "
        "(không tốn token). Chỉ tốn ~3K token để OCR mục lục."
    )

    sgk_mon = st.text_input(
        "Môn học (VD: Toán, Ngữ văn...)",
        value=st.session_state["memory"].get("default_subject", "Toán"),
        key="up_sgk_mon",
    )
    sgk_lop = st.text_input(
        "Lớp (VD: 6, 7, 8, 9)",
        value=st.session_state["memory"].get("default_grade", "6"),
        key="up_sgk_lop",
    )
    sgk_file = st.file_uploader(
        "Chọn file SGK PDF",
        type=["pdf"],
        key="up_sgk_pdf",
    )

    if sgk_file:
        sgk_name_goi_y = suggest_sgk_name(sgk_mon, sgk_lop)
        st.caption("📝 Tên SGK sẽ lưu: **" + sgk_name_goi_y + "**")

        # Kiểm tra đã tách chưa
        existing = load_sgk_metadata(sgk_name_goi_y)
        if existing:
            st.warning("⚠️ SGK này đã có " + str(existing.get("so_bai", 0))
                       + " bài. Upload sẽ ghi đè.")
            ghi_de = st.checkbox("Xác nhận ghi đè", key="overwrite_sgk")
        else:
            ghi_de = True

        if st.button("🚀 Tách SGK tự động", type="primary",
                     use_container_width=True, key="btn_split_sgk"):
            if not HAS_PYMUPDF:
                st.error("❌ Chưa cài pymupdf. Cài: pip install pymupdf")
            elif not GEMINI_KEY:
                st.error("❌ Cần GEMINI_API_KEY để OCR mục lục.")
            elif not ghi_de:
                st.warning("⚠️ Bạn chưa xác nhận ghi đè.")
            else:
                # Đọc PDF
                pdf_bytes = sgk_file.read()
                total_pages = get_pdf_total_pages(pdf_bytes)
                st.info("📄 File có **" + str(total_pages) + " trang**")

                # Bước 1: Dò ranh giới bài
                progress = st.progress(0)
                status = st.empty()

                status.write("🔍 Đang OCR mục lục để dò ranh giới bài...")
                progress.progress(20)

                boundaries = detect_chapter_boundaries(pdf_bytes)

                if not boundaries:
                    st.warning(
                        "⚠️ Không dò được ranh giới bài tự động. "
                        "Bạn có thể dùng tab **Nạp bài đã tách** để upload thủ công."
                    )
                else:
                    st.success("✅ Tìm được " + str(len(boundaries)) + " bài")

                    with st.expander("👁️ Xem danh sách bài", expanded=True):
                        for b in boundaries:
                            st.write(
                                "• **Bài " + str(b["so_bai"]) + "**: "
                                + b["ten_bai"]
                                + " (trang PDF " + str(b["page_pdf_start"])
                                + "-" + str(b["page_pdf_end"]) + ")"
                            )

                    # Bước 2: Tách file
                    status.write("✂️ Đang tách thành file bài riêng...")
                    progress.progress(50)

                    def _cb(percent, msg):
                        progress.progress(50 + int(percent * 0.4))
                        status.write(msg)

                    metadata = split_sgk_into_lessons(
                        pdf_bytes, sgk_name_goi_y, boundaries, _cb
                    )

                    progress.progress(100)
                    status.write("✅ Hoàn tất!")

                    if metadata:
                        st.success(
                            "✅ Đã tách " + str(metadata["so_bai"])
                            + " bài. Lưu vào: cache/sgk_bai/" + sgk_name_goi_y
                        )
                        st.balloons()
                    else:
                        st.error("❌ Tách thất bại.")


# ---------- TAB 2: NẠP BÀI THỦ CÔNG ----------
with tab_manual:
    st.markdown("### 📚 Nạp bài đã tách thủ công")
    st.info(
        "💡 Bạn tự tách SGK thành từng file PDF bài riêng "
        "(bằng Adobe/Foxit/Smallpdf) → upload lên đây."
    )

    # Chọn SGK đích
    ds_sgk_hien_co = [s["sgk_name"] for s in list_sgk_da_tach()]
    sgk_manual_mon = st.text_input(
        "Môn học",
        value=st.session_state["memory"].get("default_subject", "Toán"),
        key="manual_sgk_mon",
    )
    sgk_manual_lop = st.text_input(
        "Lớp",
        value=st.session_state["memory"].get("default_grade", "6"),
        key="manual_sgk_lop",
    )
    sgk_manual_name = suggest_sgk_name(sgk_manual_mon, sgk_manual_lop)
    st.caption("📝 SGK đích: **" + sgk_manual_name + "**")

    manual_files = st.file_uploader(
        "Upload nhiều file PDF bài (chọn nhiều file cùng lúc)",
        type=["pdf"],
        accept_multiple_files=True,
        key="manual_sgk_files",
    )

    if manual_files:
        st.markdown("**📋 Nhập tên bài + trang cho mỗi file:**")
        manual_info = []
        for i, f in enumerate(manual_files):
            with st.expander("📄 " + f.name, expanded=False):
                c1, c2, c3 = st.columns([3, 1, 1])
                with c1:
                    ten = st.text_input(
                        "Tên bài",
                        value=f.name.replace(".pdf", ""),
                        key=f"manual_ten_{i}",
                    )
                with c2:
                    p_bd = st.number_input(
                        "Trang BĐ", min_value=0, max_value=9999,
                        value=0, step=1, key=f"manual_pbd_{i}",
                    )
                with c3:
                    p_kt = st.number_input(
                        "Trang KT", min_value=0, max_value=9999,
                        value=0, step=1, key=f"manual_pkt_{i}",
                    )
                manual_info.append({
                    "file_obj": f,
                    "ten_bai": ten,
                    "page_bd": p_bd,
                    "page_kt": p_kt,
                })

        if st.button("💾 Lưu bài đã tách", type="primary",
                     use_container_width=True, key="btn_save_manual"):
            # Tạo thư mục SGK
            sgk_dir = os.path.join(SGK_BAI_DIR, sgk_manual_name)
            os.makedirs(sgk_dir, exist_ok=True)

            # Load metadata cũ (nếu có)
            meta = load_sgk_metadata(sgk_manual_name) or {
                "sgk_name": sgk_manual_name,
                "nguon": "manual",
                "so_bai": 0,
                "ds_bai": [],
            }

            saved = 0
            for i, info in enumerate(manual_info):
                # Tên file an toàn
                safe = _slugify(info["ten_bai"])[:40] or f"bai_{i+1:02d}"
                fname = f"bai_{i+1:02d}_{safe}.pdf"
                fpath = os.path.join(sgk_dir, fname)

                # Lưu file
                try:
                    with open(fpath, "wb") as fo:
                        fo.write(info["file_obj"].read())
                except Exception:
                    continue

                # Cập nhật metadata
                meta["ds_bai"].append({
                    "file": fname,
                    "ten_bai": info["ten_bai"],
                    "so_bai": i + 1,
                    "page_pdf_start": info["page_bd"],
                    "page_pdf_end": info["page_kt"],
                })
                saved += 1

            meta["so_bai"] = len(meta["ds_bai"])
            meta["nguon"] = "manual"

            # Lưu metadata
            meta_path = os.path.join(sgk_dir, "metadata.json")
            with open(meta_path, "w", encoding="utf-8") as fo:
                json.dump(meta, fo, ensure_ascii=False, indent=2)

            st.success("✅ Đã nạp " + str(saved) + " bài vào "
                       + sgk_manual_name)
            st.balloons()


# ---------- TAB 3: KHO TRI THỨC ----------
with tab_kho:
    st.markdown("### 📂 Kho tri thức")
    st.caption("Xem và xóa file đã lưu trên GitHub")

    for label, key, folder in [
        ("📐 Mẫu", "templates", "templates"),
        ("📚 Nguồn", "sources", "sources"),
        ("📅 Lịch BG", "schedule", "schedule"),
    ]:
        st.subheader(label + " đã lưu")
        data_dict = st.session_state[key]

        if data_dict:
            for name in list(data_dict.keys()):
                col1, col2 = st.columns([5, 1])
                with col1:
                    with st.expander("📄 " + name):
                        content = data_dict[name]
                        st.caption("Kích thước: " + str(len(content))
                                   + " ký tự")
                        st.text(content[:2000])
                with col2:
                    if st.button("🗑️", key=f"del_{key}_{name}"):
                        with st.spinner("Đang xóa..."):
                            ok, msg = gh_delete(f"data/{folder}/" + name)
                            if ok:
                                st.session_state[key] = load_folder(folder)
                                st.rerun()

            if st.button("🗑️ Xóa TẤT CẢ " + label,
                         key=f"del_all_{key}"):
                for name in list(data_dict.keys()):
                    gh_delete(f"data/{folder}/" + name)
                st.session_state[key] = load_folder(folder)
                st.rerun()
        else:
            st.info("Chưa có " + label.lower() + ".")
        st.divider()

    # SGK đã tách (local)
    st.subheader("📚 SGK đã tách (local)")
    ds_sgk_local = list_sgk_da_tach()
    if ds_sgk_local:
        for sgk in ds_sgk_local:
            col1, col2 = st.columns([5, 1])
            with col1:
                st.write("• **" + sgk["sgk_name"] + "** — "
                         + str(sgk["so_bai"]) + " bài ("
                         + sgk["nguon"] + ")")
            with col2:
                if st.button("🗑️", key="del_sgk_" + sgk["sgk_name"]):
                    import shutil
                    try:
                        shutil.rmtree(
                            os.path.join(SGK_BAI_DIR, sgk["sgk_name"])
                        )
                        st.success("✅ Đã xóa " + sgk["sgk_name"])
                        st.rerun()
                    except Exception as e:
                        st.error("Lỗi: " + str(e))
    else:
        st.info("Chưa có SGK nào được tách.")


# ---------- KẾT THÚC PHẦN 3/6 ----------
# ============================================================
# PHẦN 4/6: RULES 11 MÔN + MẪU 5512 + PROMPT TỐI ƯU
# ============================================================
# Chức năng:
#   - SUBJECT_RULES: bộ quy tắc 11 môn THCS
#   - MAU_5512_MAC_DINH: mẫu 5512 hardcode
#   - get_subject_rules, _chuan_hoa_ten_mon
#   - detect_end_of_chapter: nhận diện bài cuối chương
#   - kiem_tra_mau_5512: kiểm tra mẫu đúng chuẩn
#   - build_exec_prompt: prompt tối ưu token
# ============================================================


# ============================================================
# PHẦN A: MẪU 5512 HARDCODE
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


# ============================================================
# PHẦN B: BỘ QUY TẮC 11 MÔN THCS (TINH GỌN)
# ============================================================
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
        "danh_gia": ["Nhận thức bản thân",
                     "Giao tiếp và hợp tác",
                     "Giải quyết vấn đề đời sống",
                     "Trách nhiệm với gia đình, nhà trường, cộng đồng",
                     "Tìm hiểu nghề nghiệp"],
        "khong_nen": ("Biến mọi tiết trải nghiệm thành tiết GV thuyết "
                      "trình về kỹ năng sống."),
    },
    "Tin học": {
        "ten_goi_khac": ["tin hoc", "tin học", "informatics", "it", "cntt"],
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
        "danh_gia": ["Hiểu khái niệm tin học",
                     "Sử dụng công cụ số",
                     "Thực hiện thao tác máy tính",
                     "Phân tích vấn đề, xây dựng thuật toán",
                     "Viết, kiểm thử, sửa chương trình"],
        "khong_nen": ("Chỉ cho HS chép mã chương trình mà không giải "
                      "thích bài toán, thuật toán và cách kiểm tra."),
    },
}


# ============================================================
# PHẦN C: HÀM TRA CỨU MÔN
# ============================================================
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
            if len(alias_chuan) >= 3 and (
                alias_chuan in ten_chuan or ten_chuan in alias_chuan
            ):
                return (key, rules)

    return (None, None)


# ============================================================
# PHẦN D: NHẬN DIỆN BÀI CUỐI CHƯƠNG / LUYỆN TẬP CHUNG
# ============================================================
def detect_end_of_chapter(topic):
    """Nhận diện bài cuối chương qua từ khóa.
    Trả về True nếu là bài cuối chương (cần đổi HĐ2).
    """
    if not topic:
        return False
    t = topic.lower()
    keywords = [
        "bài tập cuối chương",
        "ôn tập chương",
        "ôn tập học kì",
        "củng cố, mở rộng",
        "tổng kết chương",
        "review",
        "looking back",
        "trưng bày sản phẩm",
        "bài tập ôn chương",
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


# ============================================================
# PHẦN E: KIỂM TRA MẪU 5512
# ============================================================
def kiem_tra_mau_5512(mau_text):
    """Kiểm tra mẫu có đủ cấu trúc 5512 không.
    Trả về (ok, msg).
    """
    if not mau_text:
        return (False, "Mẫu rỗng")

    text_lower = mau_text.lower()
    errors = []

    # Kiểm tra 3 mục chính
    if "mục tiêu" not in text_lower:
        errors.append("Thiếu 'I. Mục tiêu'")
    if "thiết bị dạy học" not in text_lower and \
       "thiết bị" not in text_lower:
        errors.append("Thiếu 'II. Thiết bị dạy học'")
    if "tiến trình dạy học" not in text_lower and \
       "tiến trình" not in text_lower:
        errors.append("Thiếu 'III. Tiến trình dạy học'")

    # Kiểm tra 4 hoạt động
    for i in [1, 2, 3, 4]:
        if ("hoạt động " + str(i)) not in text_lower:
            errors.append("Thiếu 'Hoạt động " + str(i) + "'")

    # Kiểm tra bảng hoạt động
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


# ============================================================
# PHẦN F: BUILD PROMPT TỐI ƯU TOKEN
# ============================================================
def build_exec_prompt(plan):
    """Xây prompt 5512 tối ưu token.
    - Chỉ nhồi rules của môn đang soạn (~200 token)
    - Chỉ nhồi đoạn SGK cần thiết (~2-3K token)
    - Prompt tổng: ~3-5K token
    """
    mem = st.session_state["memory"]

    # 1. Tra môn
    ten_mon_input = plan.get("subject", "") or mem.get("default_subject", "")
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
    mau_text = mau.get("cau_truc", "")[:2000]  # Cắt ngắn

    # 3. Lấy đoạn SGK (nếu có)
    sgk_text = plan.get("sgk_excerpt", "")[:8000]  # Cắt ngắn

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
    p.append("- Môn: " + ten_mon_chuan +
             " | Khối: " + str(plan.get("grade", "6")))
    p.append("- Tên bài: " + topic)
    p.append("- Số tiết: " + str(plan.get("period", "1")))
    p.append("")

    # Đặc trưng môn
    p.append("=== ĐẶC TRƯNG MÔN " + ten_mon_chuan.upper() + " ===")
    p.append("Ký hiệu: " + rules.get("ky_hieu", ""))
    p.append("Năng lực riêng: " + " | ".join(rules.get("nang_luc_rieng", [])))
    p.append("Phương pháp: " + " | ".join(rules.get("phuong_phap", [])))
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
    if plan.get("trang_sgk"):
        tieu_de += " (" + str(plan.get("period", "1")) + " tiết - Tr." + \
                   str(plan["trang_sgk"]) + " SGK)"
    else:
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
    p.append("3. Mục c) là BẢNG 2 CỘT: | Hoạt động của GV và HS | Sản phẩm |")
    p.append("   - Bảng có 4 dòng Bước 1, Bước 2, Bước 3, Bước 4")
    p.append("   - Cột Sản phẩm phải CHI TIẾT, bám sát SGK")
    p.append("")
    p.append("4. Cột Sản phẩm mỗi bước phải có:")
    p.append("   - Kiến thức SGK (kèm số trang)")
    p.append("   - Các bước giải (nếu có)")
    p.append("   - Ký hiệu toán/đặc trưng môn")
    p.append("   - Mô tả hình ảnh (nếu có)")
    p.append("   - Dự kiến câu trả lời của HS")
    p.append("")
    p.append("5. TUYỆT ĐỐI KHÔNG BỊA kiến thức không có trong SGK.")
    p.append("   Nếu SGK không có → ghi '[SGK không cung cấp]'.")
    p.append("")
    p.append("6. Trong ô bảng dùng <br> để xuống dòng. Mỗi bước riêng dòng.")
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


# ---------- KẾT THÚC PHẦN 4/6 ----------
# ============================================================
# PHẦN 5/6: WORD HELPERS + NÚT SOẠN 1 BÀI
# ============================================================
# Chức năng:
#   - fix_math: sửa ký hiệu toán
#   - _set_run_font, add_para, add_table, _add_heading
#   - write_content_to_doc, make_doc
#   - Nút "Soạn 1 bài theo lệnh"
# ============================================================


# ============================================================
# PHẦN A: SỬA KÝ HIỆU TOÁN HỌC
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


# ============================================================
# PHẦN B: WORD HELPERS
# ============================================================
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
    rfonts.set(qn("w:ascii"), "Times New Roman")
    rfonts.set(qn("w:hAnsi"), "Times New Roman")
    rfonts.set(qn("w:eastAsia"), "Times New Roman")
    rfonts.set(qn("w:cs"), "Times New Roman")


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
    """Tách nội dung 1 ô bảng thành nhiều dòng hợp lý để justify không bị dãn xấu."""
    val = val.replace("<br>", "\n").replace("<br/>", "\n")
    val = val.replace("<br />", "\n")
    val = re.sub(r"(?<!^)(?<!\n)(?=Bước\s*\d+\s*:)", "\n", val)

    # Tách sau dấu câu khi theo sau là chữ HOA tiếng Việt
    chu_hoa = (r"[A-ZĐÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬÈÉẺẼẸÊỀẾỂỄỆ"
               r"ÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰ"
               r"ỲÝỶỸỴ]")
    val = re.sub(r"([\.\!\?\:])\s+(?=" + chu_hoa + r")", r"\1\n", val)

    # Tách trước từ khóa phổ biến
    keywords = ["Em hãy", "Vậy ", "Do đó", "Tuy nhiên", "Vì vậy",
                "GV yêu cầu", "GV gọi", "GV đánh giá", "GV chính xác",
                "GV hướng dẫn", "GV đặt", "GV chiếu",
                "HS suy nghĩ", "HS thảo luận", "HS thực hiện",
                "HS báo cáo", "HS nhận xét", "HS trả lời"]
    for kw in keywords:
        val = re.sub(r"(?<=[\.\!\?])\s+" + re.escape(kw), "\n" + kw, val)

    raw_lines = [ln.strip() for ln in val.split("\n") if ln.strip()]

    # Gộp dòng ngắn < 12 ký tự
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

    # Dọn cột rác
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
    is_ten_bai = bool(re.match(r"^TÊN\s*BÀI\s*DẠY\s*:", text, re.IGNORECASE))
    is_mon_hoc = bool(re.match(r"^Môn\s*học\s*:", text, re.IGNORECASE))
    is_thoi_gian = bool(re.match(r"^Thời\s*gian\s*thực\s*hiện\s*:",
                                  text, re.IGNORECASE))

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

        # Chuẩn hóa "Thời gian thực hiện: (X tiết)"
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
        if re.match(r"^\*\*\s*TÊN\s*BÀI\s*DẠY", s, re.IGNORECASE):
            start_idx = idx
            break
        if re.match(r"^\*\*\s*I\.\s*(Mục|MỤC)\s*tiêu", s, re.IGNORECASE):
            start_idx = idx
            break
    lines = lines[start_idx:]

    cleaned = []
    for line in lines:
        s = line.strip()
        if re.match(r"^Sông\s+Đốc", s, re.IGNORECASE):
            break
        if re.search(r"NGƯỜI\s*SOẠN", s, re.IGNORECASE):
            break
        if re.search(r"KÝ\s*DUYỆT", s, re.IGNORECASE):
            break
        if re.match(r"^\(?Ký\s*tên\)?\s*$", s, re.IGNORECASE):
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
        if re.match(r"^\*\*\s*TÊN\s*BÀI\s*DẠY", stripped, re.IGNORECASE):
            _add_heading(d, stripped.replace("**", "").replace("*", "").strip())
            i += 1
            continue
        if re.match(r"^\*\*\s*Môn\s*học", stripped, re.IGNORECASE):
            _add_heading(d, stripped.replace("**", "").replace("*", "").strip())
            i += 1
            continue
        if re.match(r"^\*\*\s*Thời\s*gian\s*thực\s*hiện", stripped, re.IGNORECASE):
            _add_heading(d, stripped.replace("**", "").replace("*", "").strip())
            i += 1
            continue

        # Heading I, II, III
        if re.match(r"^\*\*\s*(I|II|III|IV|V)\.\s+\S", stripped):
            _add_heading(d, stripped.replace("**", "").replace("*", "").strip())
            i += 1
            continue

        # Heading 1. 2. 3.
        if re.match(r"^\*\*\s*\d+\.\s+[A-Za-zÀ-ỹ]", stripped) and \
           len(stripped) < 120:
            _add_heading(d, stripped.replace("**", "").replace("*", "").strip())
            i += 1
            continue

        # Heading a) b) c)
        if re.match(r"^\*\*\s*[a-z]\)\s+\S", stripped):
            _add_heading(d, stripped.replace("**", "").replace("*", "").strip())
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
    rfonts.set(qn("w:ascii"), "Times New Roman")
    rfonts.set(qn("w:hAnsi"), "Times New Roman")
    rfonts.set(qn("w:eastAsia"), "Times New Roman")
    rfonts.set(qn("w:cs"), "Times New Roman")
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
# PHẦN C: NÚT "SOẠN 1 BÀI THEO LỆNH"
# ============================================================
st.divider()
st.subheader("💬 Soạn 1 bài theo lệnh")

# Kiểm tra đầu vào
_danh_sach_mon = list(SUBJECT_RULES.keys())
_danh_sach_sgk = list_sgk_da_tach()
_sgk_names = [s["sgk_name"] for s in _danh_sach_sgk]

_col1, _col2, _col3, _col4 = st.columns([2, 1.5, 0.8, 1])

with _col1:
    _mon_chon = st.selectbox(
        "📚 Môn học", _danh_sach_mon, key="chon_mon_soan",
    )

with _col2:
    _lop_chon = st.text_input(
        "🎓 Lớp",
        value=st.session_state["memory"].get("default_grade", "6"),
        key="chon_lop_soan",
    )

with _col3:
    _tiet_chon = st.text_input(
        "⏱️ Số tiết", value="1", key="chon_tiet_soan",
    )

with _col4:
    _trang_chon = st.text_input(
        "📖 Trang SGK", value="",
        key="chon_trang_soan",
        help="VD: 25-26 (trang PDF của SGK)",
    )

# Hiển thị banner môn
_rules = SUBJECT_RULES.get(_mon_chon, {})
if _rules:
    _kh = _rules.get("ky_hieu", "")
    if _kh:
        st.markdown(
            f"<div style='background:#e0f7ff; padding:8px 14px; "
            f"border-radius:8px; border-left:4px solid #0ea5e9; "
            f"margin-bottom:8px;'>"
            f"<b style='font-size:16px;'>{_kh}</b> "
            f"<b>Môn {_mon_chon}</b>"
            f"</div>",
            unsafe_allow_html=True,
        )

cmd = st.text_area(
    "Lệnh", height=100,
    placeholder="VD: Soạn bài Ước chung lớn nhất — Toán 6",
    label_visibility="collapsed",
    key="single_cmd",
)

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
        _errors.append("❌ Chưa có API key (GEMINI_/GROQ_/...).")

    # 2. Kiểm tra mẫu
    _mau = get_mau_dang_dung()
    _ok_mau, _msg_mau = kiem_tra_mau_5512(_mau.get("cau_truc", ""))
    if not _ok_mau:
        _errors.append("❌ Mẫu 5512 không đạt: " + _msg_mau)

    # 3. Kiểm tra SGK
    _sgk_name_can = suggest_sgk_name(_mon_chon, _lop_chon)
    _meta_sgk = load_sgk_metadata(_sgk_name_can)
    if not _meta_sgk:
        _errors.append(
            "❌ Chưa tách SGK: **" + _sgk_name_can + "**. "
            "Vào tab **Upload SGK gốc** hoặc **Nạp bài đã tách**."
        )

    # 4. Kiểm tra trang SGK
    if not _trang_chon.strip():
        _errors.append("❌ Chưa nhập **Trang SGK**.")

    if _errors:
        for e in _errors:
            st.error(e)
    else:
        # === LẤY NỘI DUNG SGK ===
        sgk_excerpt = ""
        try:
            # Parse trang: "25-26" → (25, 26)
            parts = re.split(r"[-–,;\s]+", _trang_chon.strip())
            pages = [int(p) for p in parts if p.strip().isdigit()]
        except Exception:
            pages = []

        if pages:
            with st.spinner("📖 Đang lấy nội dung SGK..."):
                for pg in pages:
                    bai = find_bai_by_page(_sgk_name_can, pg)
                    if bai:
                        text = ocr_bai_pdf(_sgk_name_can, bai["file"])
                        if text:
                            sgk_excerpt += (
                                "\n=== " + bai["ten_bai"] +
                                " (trang PDF " + str(pg) + ") ===\n" + text
                            )
                    else:
                        # Không tìm thấy bài → OCR trực tiếp
                        sgk_excerpt += (
                            "\n[Không có file bài cho trang " + str(pg) + "]"
                        )
        else:
            st.warning("⚠️ Không parse được trang SGK.")

        # === GỌI AI ===
        with st.spinner("🤖 Đang sinh giáo án..."):
            _plan = {
                "task": "Soạn kế hoạch bài dạy",
                "subject": _mon_chon,
                "grade": _lop_chon,
                "topic": cmd,
                "period": _tiet_chon,
                "trang_sgk": _trang_chon,
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


# ---------- KẾT THÚC PHẦN 5/6 ----------
# ============================================================
# PHẦN 6/6: LỊCH BÁO GIẢNG + SOẠN TUẦN + XUẤT WORD
# ============================================================
# Chức năng:
# - Nhập lịch báo giảng (tuần, buổi, thứ, tiết, môn, bài, trang SGK)
# - Tự động lấy nội dung SGK theo trang
# - Soạn hàng loạt bài trong tuần
# - Xuất Word: giáo án cả tuần + lịch báo giảng
# - Hiển thị kết quả + tải xuống từng bài
# ============================================================

st.divider()
st.subheader("📅 Soạn cả tuần + Lịch báo giảng")

# --- Khởi tạo session ---
if "lich_bao_giang" not in st.session_state:
    st.session_state["lich_bao_giang"] = []
if "weekly_results" not in st.session_state:
    st.session_state["weekly_results"] = []

# --- Cấu hình tuần ---
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
        "để AI soạn toàn bộ. Cột **Trang SGK** giúp AI bám sát nội dung."
    )

# --- Bảng nhập lịch báo giảng ---
st.markdown("### 📋 Nhập lịch báo giảng")
st.caption("Mỗi dòng = 1 tiết dạy. Có thể thêm/xóa dòng tùy ý.")

_danh_sach_mon_all = list(SUBJECT_RULES.keys())

# Nút điều khiển bảng
_c_add, _c_del, _c_clear, _c_auto = st.columns([1, 1, 1, 2])
with _c_add:
    if st.button("➕ Thêm dòng", key="add_row_lich", use_container_width=True):
        st.session_state["lich_bao_giang"].append({
            "buoi": "Sáng",
            "thu": "Thứ 2",
            "tiet": str(len(st.session_state["lich_bao_giang"]) + 1),
            "mon": _danh_sach_mon_all[0] if _danh_sach_mon_all else "Toán",
            "lop": st.session_state["memory"].get("default_grade", "6"),
            "ten_bai": "",
            "trang_sgk": "",
            "so_tiet": "1",
        })
        st.rerun()

with _c_del:
    if st.button("➖ Xóa dòng cuối", key="del_row_lich", use_container_width=True):
        if st.session_state["lich_bao_giang"]:
            st.session_state["lich_bao_giang"].pop()
        st.rerun()

with _c_clear:
    if st.button("🗑️ Xóa tất cả", key="clear_lich", use_container_width=True):
        st.session_state["lich_bao_giang"] = []
        st.session_state["weekly_results"] = []
        st.rerun()

with _c_auto:
    if st.button("⚡ Tạo nhanh 5 buổi sáng (T2→T6)",
                 key="auto_gen_lich", use_container_width=True):
        _thu_list = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"]
        _mon_default = st.session_state["memory"].get("default_subject", "Toán")
        _lop_default = st.session_state["memory"].get("default_grade", "6")
        for _i, _t in enumerate(_thu_list):
            st.session_state["lich_bao_giang"].append({
                "buoi": "Sáng",
                "thu": _t,
                "tiet": str(_i + 1),
                "mon": _mon_default,
                "lop": _lop_default,
                "ten_bai": "",
                "trang_sgk": "",
                "so_tiet": "1",
            })
        st.rerun()

# Hiển thị bảng nhập
if st.session_state["lich_bao_giang"]:
    st.markdown("---")

    # Header
    _h = st.columns([0.8, 0.9, 0.6, 1.3, 0.6, 2.8, 1.3, 0.7])
    for _col, _label in zip(_h, [
        "**Buổi**", "**Thứ**", "**Tiết**", "**Môn**",
        "**Lớp**", "**Tên bài**", "**Trang SGK**", "**Số tiết**"
    ]):
        _col.markdown(_label)

    # Các dòng nhập
    for _i, _row in enumerate(st.session_state["lich_bao_giang"]):
        _cols = st.columns([0.8, 0.9, 0.6, 1.3, 0.6, 2.8, 1.3, 0.7])

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
            _row["ten_bai"] = st.text_input(
                "Tên bài", value=_row.get("ten_bai", ""),
                key=f"tenbai_{_i}", label_visibility="collapsed",
                placeholder="VD: Ước chung lớn nhất",
            )
        with _cols[6]:
            _row["trang_sgk"] = st.text_input(
                "Trang SGK", value=_row.get("trang_sgk", ""),
                key=f"trang_{_i}", label_visibility="collapsed",
                placeholder="VD: 25-26",
            )
        with _cols[7]:
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
                _errors_tuan.append(f"❌ Dòng {_idx+1}: Chưa nhập tên bài.")

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

                # Lấy SGK
                _sgk_name_row = suggest_sgk_name(_row["mon"], _row["lop"])
                _sgk_excerpt_row = ""
                _trang_str = _row.get("trang_sgk", "").strip()

                if _trang_str:
                    try:
                        _parts_row = re.split(r"[-,;\s]+", _trang_str)
                        _pages_row = [
                            int(p) for p in _parts_row if p.strip().isdigit()
                        ]
                    except Exception:
                        _pages_row = []

                    for _pg in _pages_row:
                        _bai_row = find_bai_by_page(_sgk_name_row, _pg)
                        if _bai_row:
                            _text_row = ocr_bai_pdf(
                                _sgk_name_row, _bai_row["file"]
                            )
                            if _text_row:
                                _sgk_excerpt_row += (
                                    f"\n=== {_bai_row['ten_bai']} "
                                    f"(trang PDF {_pg}) ===\n{_text_row}"
                                )

                # Build prompt
                _plan_row = {
                    "task": "Soạn kế hoạch bài dạy",
                    "subject": _row["mon"],
                    "grade": _row["lop"],
                    "topic": _row["ten_bai"],
                    "period": _row.get("so_tiet", "1"),
                    "trang_sgk": _trang_str,
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

    # Hiển thị kết quả
    if st.session_state.get("weekly_results"):
        st.divider()
        st.subheader("📄 Kết quả soạn tuần")

        _weekly = st.session_state["weekly_results"]
        _ok_count = len([r for r in _weekly if r["ok"]])
        _fail_count = len(_weekly) - _ok_count

        st.write(
            f"✅ Thành công: **{_ok_count}** | ❌ Lỗi: **{_fail_count}**"
        )

        # Nút tải xuống
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
                        "Người thực hiện: " + _mem_t["teacher_name"]
                    )
                _d_tuan.add_paragraph(
                    "Ngày tạo: " + datetime.now().strftime("%d/%m/%Y %H:%M")
                )
                _d_tuan.add_paragraph("")

                for _r in _weekly:
                    if _r["ok"]:
                        _ri = _r["row"]
                        _d_tuan.add_page_break()
                        _d_tuan.add_heading(
                            f"{_ri['mon']} {_ri['lop']} - {_ri['ten_bai']}", 1
                        )
                        write_content_to_doc(_d_tuan, _r["content"])

                _b_tuan = io.BytesIO()
                _d_tuan.save(_b_tuan)

                st.download_button(
                    "⬇️ Tải Word cả tuần",
                    _b_tuan.getvalue(),
                    file_name=f"KeHoachBaiDay_Tuan{_tuan_so}.docx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "wordprocessingml.document"
                    ),
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
                    _d_lich.add_paragraph("Giáo viên: " + _mem_l["teacher_name"])
                _d_lich.add_paragraph(
                    "Tuần: " + str(_tuan_so)
                    + " | Từ ngày: " + _ngay_bd.strftime("%d/%m/%Y")
                )
                _d_lich.add_paragraph("")

                _lich_rows = [[
                    "Buổi", "Thứ", "Tiết", "Môn", "Lớp",
                    "Tên bài", "Trang SGK", "Số tiết"
                ]]
                for _row in st.session_state["lich_bao_giang"]:
                    _lich_rows.append([
                        _row.get("buoi", ""),
                        _row.get("thu", ""),
                        _row.get("tiet", ""),
                        _row.get("mon", ""),
                        _row.get("lop", ""),
                        _row.get("ten_bai", ""),
                        _row.get("trang_sgk", ""),
                        _row.get("so_tiet", "1"),
                    ])

                _t_lich = _d_lich.add_table(
                    rows=len(_lich_rows), cols=8
                )
                _t_lich.style = "Table Grid"

                for _i, _r in enumerate(_lich_rows):
                    for _j in range(8):
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
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "wordprocessingml.document"
                    ),
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
                    st.caption(
                        f"Trang SGK: {_ri.get('trang_sgk', '')} | "
                        f"Số tiết: {_ri.get('so_tiet', '1')}"
                    )

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
                            mime=(
                                "application/vnd.openxmlformats-"
                                "officedocument.wordprocessingml.document"
                            ),
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
    st.info("💡 Bấm **➕ Thêm dòng** để bắt đầu nhập lịch báo giảng.")

# ============================================================
# PHẦN D: FOOTER + THÔNG TIN
# ============================================================
st.divider()

with st.expander("ℹ️ Thông tin ứng dụng", expanded=False):
    st.markdown(
        """
        **🧠 Trợ lý Giáo viên THCS — Phiên bản 2.1**

        - **Mẫu:** CV 5512 chuẩn Bộ GD&ĐT
        - **SGK:** Kết nối tri thức với cuộc sống
        - **AI:** Đa nhà cung cấp (Gemini / Groq / OpenRouter / NVIDIA)
        - **Lưu trữ:** GitHub (memory, templates, sources, schedule, history)
        - **Cache:** Local (OCR ảnh, SGK đã tách theo bài)

        **📚 Hướng dẫn nhanh:**

        1. **Upload SGK gốc** → Tách tự động theo bài (chỉ tốn ~3K token OCR mục lục)
        2. **Nạp bài đã tách** → Upload thủ công nếu cần
        3. **Soạn 1 bài** → Nhập lệnh, chọn môn, trang SGK
        4. **Soạn cả tuần** → Nhập lịch báo giảng, bấm soạn
        5. **Tải Word** → Xuất file chuẩn 5512

        **⚠️ Lưu ý:**

        - Cần `GEMINI_API_KEY` để OCR mục lục SGK
        - Cần `GITHUB_TOKEN` để lưu trữ lâu dài
        - Nên nhập **Trang SGK** để AI bám sát nội dung
        - 11 môn hỗ trợ: Toán, Ngữ văn, KHTN, Lịch sử, Địa lí,
          GDCD, Mỹ thuật, Âm nhạc, GDTC, HĐTN-HN, Tin học
        """
    )

st.caption(
    "🧠 Trợ lý Giáo viên THCS | "
    "Bộ não tự học • Mẫu CV 5512 • Sách Kết nối tri thức • Xuất Word"
)

# ============================================================
# KẾT THÚC PHẦN 6/6
# ============================================================
# ============================================================
# TỔNG KẾT: APP HOÀN CHỈNH (6/6 PHẦN)
# ============================================================
# Cấu trúc:
#   Phần 1/6: Hạ tầng (imports, config, GitHub API, session)
#   Phần 2/6: Đọc file + OCR + tách SGK theo bài + cache
#   Phần 3/6: Sidebar + upload + call AI
#   Phần 4/6: Rules 11 môn + mẫu 5512 + prompt tối ưu
#   Phần 5/6: Word helpers + nút soạn 1 bài
#   Phần 6/6: Lịch báo giảng + soạn tuần + xuất Word
# ============================================================
# Hướng dẫn deploy:
#   1. Tạo repo GitHub: P20217474/Thcs_teacher
#   2. Tạo .streamlit/secrets.toml:
#      GEMINI_API_KEY = "..."
#      GROQ_API_KEY = "..."
#      OPENROUTER_API_KEY = "..."
#      NVIDIA_API_KEY = "..."
#      GITHUB_TOKEN = "..."
#      GITHUB_REPO = "P20217474/Thcs_teacher"
#      GITHUB_BRANCH = "main"
#   3. Tạo folder trên GitHub:
#      data/memory.json
#      data/templates/
#      data/sources/
#      data/schedule/
#      data/history/
#   4. Deploy lên Streamlit Cloud
# ============================================================
