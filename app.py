# ============================================================
# TRỢ LÝ GIÁO VIÊN THCS — ĐÃ TỐI ƯU
# Cải tiến: gọn, nhanh, linh hoạt, dễ bảo trì
# ============================================================

import subprocess
import sys
import os
import json
import base64
import urllib.request
import re
import csv
import io
from datetime import datetime, timedelta
from typing import Optional, Dict, List, Tuple, Any, Callable

# ---------- Tự cài thư viện ----------
def _ensure(pkg: str, mod: Optional[str] = None) -> None:
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
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn

# ---------- Tải module tùy chọn ----------
def _try_import(mod: str) -> bool:
    try:
        __import__(mod)
        return True
    except Exception:
        return False

HAS_PYPDF = _try_import("pypdf")
HAS_PIL = _try_import("PIL")
HAS_GEMINI_VISION = _try_import("google.generativeai")
HAS_FREELM = _try_import("freelm")

if HAS_PYPDF:
    from pypdf import PdfReader
if HAS_PIL:
    from PIL import Image
if HAS_GEMINI_VISION:
    import google.generativeai as genai_vision
if _try_import("openpyxl"):
    import openpyxl
else:
    openpyxl = None

# ---------- Cấu hình trang ----------
st.set_page_config(page_title="Trợ lý Giáo viên THCS", page_icon="🧠", layout="wide")

# ============================================================
# LỚP QUẢN LÝ GITHUB — GOM HÀM LẶP LẠI
# ============================================================
class GitHubRepo:
    def __init__(self, token: str, repo: str, branch: str = "main"):
        self.token = token
        self.repo = repo
        self.branch = branch
        self.api_base = "https://api.github.com"

    @property
    def available(self) -> bool:
        return bool(self.token)

    def _req(self, method: str, path: str, data: Optional[Dict] = None) -> Dict:
        url = f"{self.api_base}{path}"
        headers = {
            "Authorization": f"token {self.token}",
            "Accept": "application/vnd.github+json",
            "User-Agent": "agent-thcs",
        }
        body = json.dumps(data).encode("utf-8") if data else None
        req = urllib.request.Request(url, data=body, headers=headers, method=method)
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except Exception as e:
            return {"error": str(e)}

    def get(self, path: str) -> Tuple[Optional[str], Optional[str]]:
        res = self._req("GET", f"/repos/{self.repo}/contents/{path}?ref={self.branch}")
        if not res or "error" in res or "content" not in res:
            return None, None
        try:
            return base64.b64decode(res["content"]).decode("utf-8"), res["sha"]
        except Exception:
            return None, None

    def save(self, path: str, content: str, msg: str = "Update") -> Tuple[bool, str]:
        if not self.available:
            return False, "Chưa có GITHUB_TOKEN"
        _, sha = self.get(path)
        data = {
            "message": msg,
            "content": base64.b64encode(content.encode("utf-8")).decode("ascii"),
            "branch": self.branch,
        }
        if sha:
            data["sha"] = sha
        res = self._req("PUT", f"/repos/{self.repo}/contents/{path}", data)
        if isinstance(res, dict) and "error" in res:
            return False, f"Lỗi: {res['error']}"
        return True, f"Đã lưu: {path}"

    def delete(self, path: str) -> Tuple[bool, str]:
        if not self.available:
            return False, "Chưa có GITHUB_TOKEN"
        _, sha = self.get(path)
        if not sha:
            return False, "File không tồn tại"
        res = self._req("DELETE", f"/repos/{self.repo}/contents/{path}",
                        {"message": f"Delete {path}", "sha": sha, "branch": self.branch})
        if isinstance(res, dict) and "error" in res:
            return False, f"Lỗi: {res['error']}"
        return True, f"Đã xóa: {path}"

    def list_dir(self, folder: str) -> List[str]:
        res = self._req("GET", f"/repos/{self.repo}/contents/{folder}?ref={self.branch}")
        if not res or "error" in res or not isinstance(res, list):
            return []
        return [item["name"] for item in res if item["type"] == "file"]

    def test(self) -> Tuple[bool, str]:
        if not self.available:
            return False, "Thiếu GITHUB_TOKEN"
        res = self._req("GET", "/user")
        if isinstance(res, dict) and "login" in res:
            login = res["login"]
            res2 = self._req("GET", f"/repos/{self.repo}")
            if isinstance(res2, dict) and "permissions" in res2:
                perms = res2["permissions"]
                if perms.get("push") or perms.get("admin"):
                    return True, f"OK — {login} | Có quyền ghi"
                return False, "Token thiếu quyền 'repo'"
            return True, f"OK — {login}"
        if isinstance(res, dict) and "error" in res:
            return False, f"Lỗi: {res['error']}"
        return False, "Token không hợp lệ"

# ---------- Khởi tạo cấu hình & Khóa API ----------
def cfg(key: str, default: str = "") -> str:
    try:
        val = st.secrets.get(key, default)
        return val if val else default
    except Exception:
        return default

GITHUB_TOKEN = cfg("GITHUB_TOKEN", "")
GITHUB_REPO = cfg("GITHUB_REPO", "P20217474/Thcs_teacher")
GITHUB_BRANCH = cfg("GITHUB_BRANCH", "main")
GEMINI_KEY = cfg("GEMINI_API_KEY", "")

gh = GitHubRepo(GITHUB_TOKEN, GITHUB_REPO, GITHUB_BRANCH)

PROVIDER_KEYS = {
    "GEMINI_API_KEY": GEMINI_KEY,
    "GROQ_API_KEY": cfg("GROQ_API_KEY", ""),
    "OPENROUTER_API_KEY": cfg("OPENROUTER_API_KEY", ""),
    "NVIDIA_API_KEY": cfg("NVIDIA_API_KEY", ""),
}
for k, v in PROVIDER_KEYS.items():
    if v:
        os.environ[k] = v
AVAILABLE_PROVIDERS = [k for k, v in PROVIDER_KEYS.items() if v]

PREFERRED_MODELS: Dict[str, List[str]] = {
    "GEMINI_API_KEY": ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-flash-latest"],
    "GROQ_API_KEY": ["llama-3.3-70b-versatile", "llama-3.1-70b-versatile"],
    "OPENROUTER_API_KEY": ["google/gemini-2.0-flash-exp:free"],
    "NVIDIA_API_KEY": ["meta/llama-3.3-70b-instruct"],
}

# ============================================================
# QUẢN LÝ DỮ LIỆU
# ============================================================
PATHS = {
    "memory": "data/memory.json",
    "templates": "data/templates/",
    "sources": "data/sources/",
    "ppct": "data/ppct/",
    "tkb": "data/tkb/",
    "schedule": "data/schedule/",
    "history": "data/history/",
}

def load_json(path: str, default: Any = None) -> Any:
    content, _ = gh.get(path)
    if not content:
        return default or {}
    try:
        return json.loads(content)
    except Exception:
        return default or {}

def save_json(path: str, data: Any, msg: str = "Update") -> Tuple[bool, str]:
    data["last_updated"] = datetime.now().strftime("%Y-%m-%d")
    return gh.save(path, json.dumps(data, ensure_ascii=False, indent=2), msg)

def load_folder(folder: str) -> Dict[str, str]:
    files = gh.list_dir(folder)
    out = {}
    for f in files:
        c, _ = gh.get(f"{folder}/{f}")
        if c:
            out[f] = c
    return out

def save_history_entry(task: str, topic: str, content: str) -> Tuple[bool, str]:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe = re.sub(r"[^a-zA-Z0-9]", "_", topic)[:30]
    path = f"{PATHS['history']}{ts}_{safe}.md"
    body = f"# {task}\n\n**Chủ đề:** {topic}\n\n---\n\n{content}"
    return gh.save(path, body, f"History {topic}")

# ---------- Khởi tạo Session State ----------
DEFAULTS = {
    "memory": lambda: load_json(PATHS["memory"], {
        "teacher_name": "", "school": "",
        "default_subject": "Toan", "default_grade": "6"
    }),
    "templates": lambda: load_folder(PATHS["templates"]),
    "sources": lambda: load_folder(PATHS["sources"]),
    "ppct": lambda: load_folder(PATHS["ppct"]),
    "tkb": lambda: load_folder(PATHS["tkb"]),
    "schedule": lambda: load_folder(PATHS["schedule"]),
    "draft": "",
    "last_topic": "",
    "weekly_results": [],
    "lich_bao_giang": [],
}

def init_session() -> None:
    if "app_started_v16" not in st.session_state:
        for k in list(st.session_state.keys()):
            if k.startswith(("tuan_", "buoi_", "ngay_bd_")):
                del st.session_state[k]
        st.session_state["app_started_v16"] = True

    for k, factory in DEFAULTS.items():
        if k not in st.session_state:
            st.session_state[k] = factory() if callable(factory) else factory

init_session()

# ============================================================
# XỬ LÝ ĐỌC FILE & OCR — TÁCH LỚP GỌN GÀNG
# ============================================================
class FileReader:
    @staticmethod
    def docx(data: bytes) -> str:
        try:
            d = Document(io.BytesIO(data))
            out = [p.text for p in d.paragraphs if p.text.strip()]
            for t in d.tables:
                out.append("")
                for row in t.rows:
                    out.append(f"| {' | '.join(c.text.strip() for c in row.cells)} |")
                out.append("")
            return "\n".join(out)
        except Exception as e:
            return f"[Lỗi docx: {e}]"

    @staticmethod
    def pdf(data: bytes) -> str:
        if not HAS_PYPDF:
            return "[Chưa cài pypdf]"
        try:
            reader = PdfReader(io.BytesIO(data))
            texts = []
            for page in reader.pages:
                try:
                    t = getattr(page, 'extract_text')(extraction_mode="layout") or ""
                except TypeError:
                    t = page.extract_text() or ""
                if t.strip():
                    texts.append(t)
            full = "\n".join(texts)
            if len(full.strip()) < 50:
                return "[PDF scan, không có text]\n" + full
            return full
        except Exception as e:
            return f"[Lỗi pdf: {e}]"

    @staticmethod
    def excel(data: bytes) -> str:
        if not openpyxl:
            return "[Chưa cài openpyxl]"
        try:
            wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
            out = []
            for sn in wb.sheetnames:
                ws = wb[sn]
                rows = [[str(c) if c is not None else "" for c in row]
                        for row in ws.iter_rows(values_only=True)]
                rows = [r for r in rows if any(c.strip() for c in r)]
                if not rows:
                    continue
                mc = max(len(r) for r in rows)
                rows = [r + [""] * (mc - len(r)) for r in rows]
                out.append(f"### Sheet: {sn}")
                out.append(f"| {' | '.join(rows[0])} |")
                out.append("|" + "---|" * mc)
                for r in rows[1:]:
                    out.append(f"| {' | '.join(r)} |")
                out.append("")
            return "\n".join(out)
        except Exception as e:
            return f"[Lỗi Excel: {e}]"

    @staticmethod
    def csv(data: bytes) -> str:
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
            out = [f"| {' | '.join(rows[0])} |", "|" + "---|" * mc]
            for r in rows[1:]:
                out.append(f"| {' | '.join(r)} |")
            return "\n".join(out)
        except Exception as e:
            return f"[Lỗi csv: {e}]"

    @staticmethod
    def ocr_image(data: bytes, fname: str) -> str:
        if not GEMINI_KEY:
            return f"[{fname}: cần GEMINI_API_KEY]"
        if not HAS_GEMINI_VISION or not HAS_PIL:
            return f"[{fname}: thiếu thư viện OCR]"
        try:
            genai_vision.configure(api_key=GEMINI_KEY)
            img = Image.open(io.BytesIO(data))
            prompt = ("Đây là ảnh trang sách giáo khoa. Đọc TOÀN BỘ chữ, giữ cấu trúc, "
                     "dùng Unicode toán: x², √x, ≤, π. Chỉ trả về text.")
            for model_name in PREFERRED_MODELS["GEMINI_API_KEY"]:
                try:
                    model = genai_vision.GenerativeModel(model_name)
                    resp = model.generate_content([prompt, img])
                    t = resp.text.strip()
                    if len(t) >= 20:
                        return t
                except Exception:
                    continue
            return f"[{fname}: AI không đọc được]"
        except Exception as e:
            return f"[{fname}: Lỗi OCR — {e}]"

    @classmethod
    def read(cls, file, use_ocr: bool = True) -> Tuple[str, str, int]:
        if file is None:
            return "", "", 0
        try:
            data = file.read()
        except Exception as e:
            return file.name, f"[Lỗi: {e}]", 0
        if not data:
            return file.name, "[File rỗng]", 0

        name = file.name
        nl = name.lower()
        magic = data[:8]

        if nl.endswith((".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp")):
            return name, cls.ocr_image(data, name) if use_ocr else f"[{name}: ảnh]", 0
        if magic[:4] == b"%PDF":
            text = cls.pdf(data)
        elif magic[:4] == b"PK\x03\x04":
            text = cls.excel(data) if nl.endswith((".xlsx", ".xlsm")) else cls.docx(data)
        elif nl.endswith(".docx"):
            text = cls.docx(data)
        elif nl.endswith(".csv"):
            text = cls.csv(data)
        else:
            for enc in ["utf-8", "utf-8-sig", "latin-1", "cp1258"]:
                try:
                    text = data.decode(enc)
                    break
                except Exception:
                    text = data.decode("utf-8", errors="ignore")
        return name, text, len(text)

# ============================================================
# GỌI AI — ĐƠN GIẢN HƠN, BÁO LỖI RÕ RÀNG
# ============================================================
def call_ai(prompt: str, system: str = "") -> str:
    if not AVAILABLE_PROVIDERS:
        raise RuntimeError("Chưa có API key nào được cấu hình")
    if not HAS_FREELM:
        raise RuntimeError("Thiếu thư viện 'freelm' — cài: pip install freelm")

    import freelm
    llm = freelm.FreeLLM.from_env()
    msgs = []
    if system:
        msgs.append({"role": "system", "content": system})
    msgs.append({"role": "user", "content": prompt})

    last_err: Optional[Exception] = None
    for provider in AVAILABLE_PROVIDERS:
        for model in PREFERRED_MODELS.get(provider, []):
            try:
                resp = llm.chat(msgs, model=model)
                t = getattr(resp, "text", None) or getattr(resp, "content", "")
                if t and len(t.strip()) > 50:
                    return t.strip()
            except Exception as e:
                last_err = e
                continue
    try:
        resp = llm.chat(msgs, model="auto")
        t = getattr(resp, "text", None) or getattr(resp, "content", "")
        if not t:
            raise ValueError("AI trả về nội dung trống")
        return t.strip()
    except Exception:
        if last_err:
            raise last_err
        raise RuntimeError("Tất cả nhà cung cấp AI đều không phản hồi")

# ============================================================
# XỬ LÝ TEXT & WORD — TÁCH HELPER GỌN
# ============================================================
def fix_math(text: str) -> str:
    patterns = [
        (r"\^2\b", "²"), (r"\^3\b", "³"), (r"_2\b", "₂"),
        (r"sqrt\(", "√("), (r"\\sqrt\{", "√"),
        (r"\\times", "×"), (r"\\div", "÷"),
        (r"\\leq", "≤"), (r"\\geq", "≥"),
        (r"\\pi", "π"), (r"\\alpha", "α"),
        (r"\$([^\$]+)\$", r"\1"),
    ]
    for p, repl in patterns:
        text = re.sub(p, repl, text)
    return text

def set_font(run, size: int = 14, bold: bool = False) -> None:
    run.bold = bold
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), "Times New Roman")

def clean_text_for_doc(text: str) -> str:
    text = text.replace("<br>", "\n").replace("<br/>", "\n")
    text = re.sub(r"^[\-\+\*\•\#]+\s*", "", text, flags=re.MULTILINE)
    text = re.sub(r"^\d+[\.\)]\s+", "", text, flags=re.MULTILINE)
    text = re.sub(r"^[a-z]\)\s+", "", text, flags=re.MULTILINE)
    text = text.replace("**", "").replace("__", "")
    text = re.sub(r"(?<!\*)\*(?!\*)", "", text)
    return text.strip()

def split_cell_content(val: str) -> str:
    val = re.sub(r"<br\s*/?>", "\n", val, flags=re.IGNORECASE)
    val = re.sub(r"(?<!^)(?<!\n)(?=Bước\s*\d+\s*:)", "\n", val)
    chu_hoa = r"[A-ZĐÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴ]"
    val = re.sub(r"([\.\!\?\:])\s+(?=" + chu_hoa + r")", r"\1\n", val)
    keywords = ["Em hãy", "Vậy ", "Do đó", "Tuy nhiên", "Vì vậy",
                "GV yêu cầu", "GV gọi", "GV đánh giá", "GV chính xác",
                "GV hướng dẫn", "GV đặt", "HS suy nghĩ", "HS thảo luận",
                "HS thực hiện", "HS báo cáo", "HS nhận xét", "HS trả lời",
                "HS chú ý", "HS hoàn thành"]
    for kw in keywords:
        val = re.sub(r"(?<=[\.\!\?])\s+" + re.escape(kw), r"\n" + kw, val)
    lines = [ln.strip() for ln in val.split("\n") if ln.strip()]
    merged = []
    for ln in lines:
        if merged and len(ln) < 12 and not ln.startswith("Bước"):
            merged[-1] = f"{merged[-1]} {ln}"
        else:
            merged.append(ln)
    return "\n".join(merged)

def add_para(doc, text: str) -> Any:
    text = clean_text_for_doc(text)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Inches(0.39)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    run = p.add_run(text)
    set_font(run, 14)
    return p

def add_table(doc, rows: List[List[str]]) -> None:
    if not rows:
        return
    cleaned = []
    for row in rows:
        while len(row) > 2 and row[-1].strip() in ("", "=", ".", ";", ",", "a", "b", "c", "|"):
            row = row[:-1]
        cleaned.append(row)
    max_cols = max(len(r) for r in cleaned)
    if max_cols > 2:
        for i, r in enumerate(cleaned):
            if len(r) > 2:
                extra = " ".join(c.strip() for c in r[2:] if c.strip())
                cleaned[i] = [r[0], f"{r[1]} {extra}".strip()]
        max_cols = 2

    t = doc.add_table(rows=len(cleaned), cols=max_cols)
    t.style = "Table Grid"
    for i, row in enumerate(cleaned):
        for j in range(max_cols):
            c = t.cell(i, j)
            p = c.paragraphs[0]
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            val = row[j] if j < len(row) else ""
            is_header = i == 0
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if is_header else WD_ALIGN_PARAGRAPH.JUSTIFY
            if is_header:
                val = clean_text_for_doc(val)
                run = p.add_run(val)
                set_font(run, 13, bold=True)
            else:
                val = split_cell_content(val)
                for line_idx, line_content in enumerate(val.split("\n")):
                    target_p = p if line_idx == 0 else c.add_paragraph()
                    if line_idx > 0:
                        target_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    for part in re.split(r"(\*\*[^\*]+\*\*)", line_content):
                        if part.startswith("**") and part.endswith("**") and len(part) > 4:
                            run = target_p.add_run(part[2:-2])
                            set_font(run, 13, bold=True)
                        elif part:
                            run = target_p.add_run(part.replace("**", ""))
                            set_font(run, 13)

def add_heading(doc, text: str) -> Any:
    text = re.sub(r"^[\*#\-\+\\]+\s*", "", text).rstrip("*").strip()
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    uc = text.upper()
    if uc.startswith("TÊN BÀI DẠY") or text.startswith("Môn học") or text.startswith("Thời gian"):
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        set_font(run, 14, bold=True)
        return p
    if re.match(r"^[a-z]\)\s", text):
        if ":" in text:
            head, cont = text.split(":", 1)
            run1 = p.add_run(head.strip() + ":")
            set_font(run1, 14, bold=True, italic=True)
            run2 = p.add_run(" " + cont.strip())
            set_font(run2, 14, italic=True)
        return p
    run = p.add_run(text)
    set_font(run, 14, bold=True)
    return p

def write_content(doc, content: str) -> None:
    content = fix_math(content)
    content = re.sub(r"^\s*BẢNG\s*2\s*CỘT\s*:?\s*$", "", content, flags=re.MULTILINE | re.IGNORECASE)
    content = re.sub(r"<\/?(?:div|span|p|em|strong)[^>]*>", "", content)
    lines = content.splitlines()
    start = 0
    for i, ln in enumerate(lines):
        s = ln.strip()
        if re.match(r"^\*\*\s*(TÊN BÀI DẠY|I\.\s*MỤC)", s, re.IGNORECASE):
            start = i
            break
    lines = lines[start:]
    cleaned = []
    for ln in lines:
        s = ln.strip()
        if re.match(r"^(Sông Đốc|Người soạn|Ký tên|Ký duyệt)", s, re.IGNORECASE):
            break
        cleaned.append(ln)
    lines = cleaned

    i = 0
    while i < len(lines):
        s = lines[i].rstrip()
        stripped = s.strip()
        if not stripped:
            i += 1
            continue
        if stripped.startswith("|") and stripped.endswith("|"):
            table_rows = []
            while i < len(lines) and lines[i].strip().startswith("|") and lines[i].strip().endswith("|"):
                ln_strip = lines[i].strip()
                if not re.match(r"^\|[\s\-:\|]+\|$", ln_strip):
                    cells = [c.strip() for c in ln_strip.strip("|").split("|")]
                    if 2 <= len(cells) <= 4:
                        table_rows.append(cells)
                i += 1
            if table_rows:
                add_table(doc, table_rows)
            continue
        if re.match(r"^\*\*?\*?\s*(TÊN BÀI DẠY|Môn học|Thời gian|I\.|II\.|III\.)", stripped):
            add_heading(doc, stripped.replace("*", "").strip())
            i += 1
            continue
        if re.match(r"^\*\*\s*\d+\.\s+[A-Z]", stripped) or re.match(r"^\*\*\s*[a-z]\)\s+", stripped):
            add_heading(doc, stripped.replace("*", "").strip())
            i += 1
            continue
        add_para(doc, stripped)
        i += 1

def make_doc(title: str, body: str) -> bytes:
    d = Document()
    s = d.sections[0]
    s.top_margin = Inches(0.7)
    s.bottom_margin = Inches(0.7)
    s.left_margin = Inches(0.8)
    s.right_margin = Inches(0.8)
    style = d.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(14)
    pf = style.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE

    mem = st.session_state["memory"]
    d.add_heading(title, 0)
    if mem.get("school"):
        d.add_paragraph(f"Đơn vị: {mem['school']}")
    if mem.get("teacher_name"):
        d.add_paragraph(f"Người thực hiện: {mem['teacher_name']}")
    d.add_paragraph(f"Ngày tạo: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    d.add_paragraph("")
    write_content(d, body)
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()

def add_header_footer(doc, lich: List[Dict], tuan_start: Any, buoi: str, tuan_num: int) -> None:
    mem = st.session_state["memory"]
    end_date = tuan_start + timedelta(days=6)
    htable = doc.add_table(rows=1, cols=2)
    htable.autofit = True
    c1 = htable.cell(0, 0)
    p = c1.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("UBND XÃ SÔNG ĐỐC")
    set_font(run, 13)
    p2 = c1.add_paragraph()
    run = p2.add_run("TRƯỜNG THCS SÔNG ĐỐC")
    set_font(run, 13, bold=True)

    c2 = htable.cell(0, 1)
    p = c2.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    set_font(run, 13, bold=True)
    p2 = c2.add_paragraph()
    run = p2.add_run("Độc lập — Tự do — Hạnh phúc")
    set_font(run, 13, bold=True)

    p_date = doc.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p_date.add_run(f"Sông Đốc, ngày {datetime.now():%d} tháng {datetime.now():%m} năm {datetime.now():%Y}")
    set_font(run, 13, italic=True)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_title.add_run("PHIẾU BÁO GIẢNG ")
    set_font(run, 14, bold=True)
    run = p_title.add_run(buoi)
    set_font(run, 14, bold=True)
    run.font.color.rgb = RGBColor(0xFF, 0, 0)

    p_tuan = doc.add_paragraph()
    p_tuan.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_tuan.add_run(f"Tuần {tuan_num}, năm học 2026-2027 (Từ {tuan_start:%d/%m/%Y} đến {end_date:%d/%m/%Y})")
    set_font(run, 13, italic=True)
    doc.add_paragraph("")

    table = doc.add_table(rows=1, cols=7)
    table.style = "Table Grid"
    widths = [Cm(2.5), Cm(1.2), Cm(1.5), Cm(1.5), Cm(1.3), Cm(9.0), Cm(1.5)]
    headers = [["Thứ", "ngày, tháng"], ["Tiết", "TKB"], ["Môn"], ["Lớp"], ["Tiết", "PP CT"], ["Tên bài dạy"], ["Ghi", "chú"]]
    for i, lines in enumerate(headers):
        c = table.rows[0].cells[i]
        for j, line in enumerate(lines):
            p = c.paragraphs[0] if j == 0 else c.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(line)
            set_font(run, 13, bold=True)
        if i == 6:
            for p in c.paragraphs:
                for r in p.runs:
                    r.font.color.rgb = RGBColor(0xFF, 0, 0)
        c.width = widths[i]

    # Nhóm & thêm dữ liệu
    thu_order = {"thứ 2": 2, "thứ 3": 3, "thứ 4": 4, "thứ 5": 5, "thứ 6": 6, "thứ 7": 7}
    grouped: Dict[str, List[Dict]] = {}
    for r in lich:
        key = f"{r['thu'].strip().lower()}|{r.get('ngay','').strip()}"
        grouped.setdefault(key, []).append(r)

    for key, items in grouped.items():
        thu_disp = items[0]["thu"]
        ngay_disp = items[0].get("ngay", "")
        for idx, r in enumerate(items):
            row = table.add_row().cells
            if idx == 0:
                p = row[0].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(thu_disp)
                set_font(run, 13)
                if ngay_disp:
                    p2 = row[0].add_paragraph()
                    run2 = p2.add_run(ngay_disp)
                    set_font(run2, 13)
            for c_idx, key_f in [(1, "tiet"), (2, "mon"), (3, "lop"), (4, "stt_ppct"), (5, "ten_bai")]:
                val = str(r.get(key_f, ""))
                if key_f == "tiet":
                    val = val.replace("Tiết ", "")
                if key_f == "stt_ppct":
                    m = re.search(r"\d+", val)
                    val = m.group(0) if m else val
                p = row[c_idx].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx != 5 else WD_ALIGN_PARAGRAPH.LEFT
                run = p.add_run(val)
                set_font(run, 13)
            row[6].text = ""

    doc.add_paragraph("")

# ============================================================
# PROMPT XÂY DỰNG KẾ HOẠCH BÀI DẠY 5512
# ============================================================
def build_lesson_plan_prompt(plan: Dict) -> str:
    mem = st.session_state["memory"]
    tpl_text = "".join(
        f"\n=== FILE MẪU ===\n{st.session_state['templates'][f][:20000]}"
        for f in plan.get("template_files", []) if f in st.session_state["templates"]
    )
    src_text = "".join(
        f"\n=== NGUỒN SGK ===\n{st.session_state['sources'][f][:25000]}"
        for f in plan.get("source_files", []) if f in st.session_state["sources"]
    )

    topic = plan.get("topic", "")
    subject = plan.get("subject", "")
    grade = plan.get("grade", "")
    period = plan.get("period", "1")

    return f"""Bạn là trợ lý chuyên môn cao cấp cho giáo viên THCS Việt Nam.

NHIỆM VỤ: Soạn KẾ HOẠCH BÀI DẠY theo ĐÚNG mẫu Công văn 5512.

THÔNG TIN:
- GV: {mem.get('teacher_name', '')} | Trường: {mem.get('school', '')}
- Môn: {subject} | Khối: {grade}
- Tên bài: {topic}
- Số tiết: {period} tiết

{tpl_text}
{src_text}

CẤU TRÚC BẮT BUỘC:

**TÊN BÀI DẠY: {topic}**
**Môn học: {subject}; lớp: {grade}**
**Thời gian thực hiện: {period} tiết**

**I. Mục tiêu**
1. Kiến thức:
[2-4 đoạn văn mô tả, không dùng gạch đầu dòng]
2. Năng lực:
Năng lực riêng: [mô tả 1-2 đoạn]
Năng lực chung: Năng lực tư duy và lập luận; giao tiếp toán học; tự học; giải quyết vấn đề; sáng tạo; hợp tác.
3. Phẩm chất:
Bồi dưỡng hứng thú học tập, ý thức làm việc nhóm, tìm tòi, độc lập, tự tin, tự chủ.

**II. Thiết bị dạy học và học liệu**
**1. Giáo viên:** Kế hoạch bài dạy, SGK, SGV, phiếu học tập, thước kẻ, phấn màu, tivi/máy vi tính nếu cần.
**2. Học sinh:** SGK, vở ghi, đồ dùng, ôn kiến thức cũ.

**III. Tiến trình dạy học**

**1. Hoạt động 1: Mở đầu**
**a) Mục tiêu:** [2 câu ngắn]
**b) Nội dung:** [1-2 câu]
**c) Tổ chức thực hiện và sản phẩm:**

| Hoạt động của GV và HS | Sản phẩm |
|---|---|
| Bước 1: Chuyển giao nhiệm vụ:<br>GV đặt vấn đề...<br>Câu hỏi: ... | [Sản phẩm dự kiến] |
| Bước 2: Thực hiện nhiệm vụ:<br>HS suy nghĩ, thảo luận... | |
| Bước 3: Báo cáo, thảo luận:<br>Đại diện trả lời... | |
| Bước 4: Kết luận:<br>GV đánh giá, dẫn vào bài mới | |

**2. Hoạt động 2: Giải quyết vấn đề**
**a) Mục tiêu:** [mô tả]
**b) Nội dung:** [mô tả]
**c) Tổ chức thực hiện và sản phẩm:**

| Hoạt động của GV và HS | Sản phẩm |
|---|---|
| Bước 1: Chuyển giao nhiệm vụ:<br>GV hướng dẫn yêu cầu HS thực hiện... | [Kiến thức trọng tâm]<br>Ví dụ 1: ...<br>Ví dụ 2: ... |
| Bước 2: Thực hiện nhiệm vụ:<br>HS lắng nghe, thực hiện... | |
| Bước 3: Báo cáo, thảo luận:<br>Thảo luận, nhận xét, bổ sung... | |
| Bước 4: Kết luận:<br>GV chính xác hóa, gọi 1 HS nhắc lại nội dung chính | |

**3. Hoạt động 3: Luyện tập**
**a) Mục tiêu:** Củng cố kiến thức qua bài tập.
**b) Nội dung:** Vận dụng làm BT.
**c) Tổ chức thực hiện và sản phẩm:**

| Hoạt động của GV và HS | Sản phẩm |
|---|---|
| GV yêu cầu HS hoàn thành BT [số] (tr[XX]-SGK). | Bài [số]:<br>[Lời giải chi tiết] |

**4. Hoạt động 4: Vận dụng**
**a) Mục tiêu:** Củng cố & áp dụng thực tiễn.
**b) Nội dung:** Vận dụng kiến thức giải quyết vấn đề thực tế.
**c) Tổ chức thực hiện và sản phẩm:**

| Hoạt động của GV và HS | Sản phẩm |
|---|---|
| GV cho tìm hiểu, yêu cầu hoàn thành... | Bài [số]: [Lời giải chi tiết] |

**HƯỚNG DẪN VỀ NHÀ:**
- Ghi nhớ nội dung chính.
- Làm bài [số] (SGK-tr[XX]).
- Chuẩn bị bài mới "[tên bài tiếp theo]".

QUY TẮC:
1. Lấy nội dung từ nguồn SGK, không bịa
2. Không viết "máy chiếu" → dùng "tivi" hoặc "máy vi tính"
3. Ký hiệu toán Unicode: x², √, ≤, ≥, π — không dùng LaTeX
4. Phân số dạng a/b
5. Trong ô bảng dùng <br> để xuống dòng
6. Độ dài ≥ 1500 từ
7. Không dùng --- , === , ####
8. Không thêm "Người soạn", "Ký tên" ở cuối
9. Không dùng gạch đầu dòng trong đoạn văn
"""

# ============================================================
# GIAO DIỆN SIDEBAR
# ============================================================
with st.sidebar:
    st.header("🧠 Bộ não")
    if gh.available:
        st.success("💾 Đã có GITHUB_TOKEN")
        if st.button("🧪 Test GitHub", use_container_width=True):
            ok, msg = gh.test()
            st.success(f"✅ {msg}" if ok else f"❌ {msg}")
    else:
        st.error("❌ Chưa có GITHUB_TOKEN")

    if GEMINI_KEY:
        st.success("🔍 Gemini Vision sẵn sàng")
    else:
        st.warning("⚠️ Chưa có GEMINI_API_KEY")

    mem = st.session_state["memory"]
    mem["teacher_name"] = st.text_input("Giáo viên", value=mem.get("teacher_name", ""))
    mem["school"] = st.text_input("Trường", value=mem.get("school", ""))
    mem["default_subject"] = st.text_input("Môn mặc định", value=mem.get("default_subject", "Toan"))
    mem["default_grade"] = st.text_input("Khối mặc định", value=mem.get("default_grade", "6"))

    if st.button("💾 Lưu vào bộ nhớ", use_container_width=True):
        ok, msg = save_json(PATHS["memory"], mem, "Update memory")
        st.success(f"✅ {msg}" if ok else f"❌ {msg}")

    st.divider()
    st.subheader("📊 Thống kê")
    for label, key in [("Mẫu", "templates"), ("Nguồn", "sources"), ("PPCT", "ppct"), ("TKB", "tkb"), ("Lịch BG", "schedule")]:
        st.write(f"{label}: {len(st.session_state[key])} file")

    if st.button("🔄 Tải lại", use_container_width=True):
        for k, folder in [("templates", "templates"), ("sources", "sources"), ("ppct", "ppct"), ("tkb", "tkb"), ("schedule", "schedule")]:
            st.session_state[k] = load_folder(PATHS[folder])
        st.rerun()

    st.write(f"AI: {len(AVAILABLE_PROVIDERS)} nhà cung cấp")

    # Dọn dẹp — rút gọn logic
    st.divider()
    st.subheader("🧹 Dọn dẹp")
    with st.expander("🗑️ Dọn rác & cache"):
        if st.button("🧼 Xóa phiên tạm", use_container_width=True):
            for k in ["draft", "last_topic", "weekly_results", "lich_bao_giang"]:
                st.session_state.pop(k, None)
            st.success("✅ Đã xóa phiên làm việc tạm")
            st.rerun()

        if st.button("♻️ Xóa cache Streamlit", use_container_width=True):
            st.cache_data.clear()
            st.cache_resource.clear()
            st.success("✅ Đã xóa cache")
            st.rerun()

        if gh.available:
            if st.button("📜 Xóa lịch sử soạn bài", use_container_width=True):
                deleted = sum(1 for f in gh.list_dir(PATHS["history"]) if gh.delete(f"{PATHS['history']}{f}")[0])
                st.success(f"✅ Đã xóa {deleted} file lịch sử")

            if st.button("🗑️ Xóa file rác (file_*.md)", use_container_width=True):
                deleted = 0
                for folder in ["templates", "sources", "ppct", "tkb", "schedule"]:
                    for f in gh.list_dir(PATHS[folder]):
                        if f.startswith("file_"):
                            if gh.delete(f"data/{folder}/{f}")[0]:
                                deleted += 1
                st.success(f"✅ Đã xóa {deleted} file rác")
                st.rerun()

            if st.button("📅 Xóa tất cả lịch đã lưu", use_container_width=True):
                deleted = sum(1 for f in gh.list_dir(PATHS["schedule"]) if gh.delete(f"{PATHS['schedule']}{f}")[0])
                st.success(f"✅ Đã xóa {deleted} file lịch")

        st.divider()
        st.caption("⚠️ Dưới đây xóa dữ liệu VĨNH VIỄN")
        if st.button("🔴 Reset toàn bộ app", use_container_width=True):
            for k in list(st.session_state.keys()):
                del st.session_state
