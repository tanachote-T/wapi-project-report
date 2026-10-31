import streamlit as st
import requests
import io
import os
import re
import math
import shutil
import subprocess
import tempfile
import base64
from docx import Document
from docx.shared import Inches, Pt, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ตรวจสอบไลบรารีจัดการรูปภาพ (Pillow) สำหรับครอปอัตราส่วน 4:3 และ 3:4
try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# ตรวจสอบไลบรารีตัดคำภาษาไทย
try:
    from pythainlp.tokenize import word_tokenize
    THAI_TOKENIZER_AVAILABLE = True
except ImportError:
    THAI_TOKENIZER_AVAILABLE = False

# ตรวจสอบไลบรารีเชื่อมต่อ Microsoft Word (สำหรับ Windows)
try:
    import pythoncom
    import win32com.client
    WORD_COM_AVAILABLE = True
except ImportError:
    WORD_COM_AVAILABLE = False

try:
    from docx2pdf import convert
    DOCX2PDF_AVAILABLE = True
except ImportError:
    DOCX2PDF_AVAILABLE = False

# ตรวจสอบโปรแกรม LibreOffice (สำหรับแปลง PDF บน Streamlit Cloud / Linux)
LIBREOFFICE_PATH = shutil.which("soffice") or shutil.which("libreoffice")
PDF_SUPPORT = WORD_COM_AVAILABLE or DOCX2PDF_AVAILABLE or (LIBREOFFICE_PATH is not None)

# ==========================================================
# ส่วนที่ 1: ตัวแปรตั้งค่าพื้นฐาน (คุณครูสามารถปรับเปลี่ยนได้)
# ==========================================================
DEFAULT_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwFr3O-GcrOnIGnxKZ-0owTxEdORla-THc49pOfhhs9F8RyRel9cufZVH6tGIFObE6uEA/exec"
SCHOOL_LOGO_URL = "https://i.ibb.co/qMb0pmBV/logo.png"

DEFAULT_SEMESTER = "2"
DEFAULT_ACADEMIC_YEAR = "2569"

# ตั้งค่าฟอนต์และขนาดตัวอักษรมาตรฐานในเอกสาร Word
REPORT_FONT_NAME = "TH SarabunPSK"
COVER_FONT_SIZE = 20
CHAPTER_TITLE_SIZE = 20
SECTION_HEADING_SIZE = 18
BODY_FONT_SIZE = 16
CAPTION_FONT_SIZE = 14
PAGE_NUMBER_FONT_SIZE = 14
LINE_SPACING_VAL = 1.0

# ระยะขอบของเอกสารทุกด้าน เท่ากับ 25.4 มิลลิเมตร
PAGE_MARGIN_MM = 25.4

# ค่าเริ่มต้นความกว้างคอลัมน์บนหน้าปก: ชื่อ-นามสกุล (7.5 ซม.) และ เลขที่ (23 มิลลิเมตร = 2.3 ซม.)
DEFAULT_COVER_NAME_COL_CM = 7.5
DEFAULT_COVER_NUM_COL_CM = 2.3

# จำนวนบรรทัดขั้นต่ำของเนื้อหาที่ต้องอยู่คู่กับหัวข้อหลักในหน้าเดียวกัน
MIN_LINES_WITH_HEADING = 3

# รหัสสีหัวตาราง: สีมะกอก ส่วนที่ถูกเน้น 3 อ่อนลง 80% (#EBF1DE)
OLIVE_ACCENT3_LIGHTER80_HEX = "EBF1DE"

# ขนาดและอัตราส่วนรูปภาพมาตรฐาน
LANDSCAPE_IMG_HEIGHT_CM = 7.0
LANDSCAPE_IMG_WIDTH_CM = round(7.0 * (4.0 / 3.0), 2)

PORTRAIT_IMG_WIDTH_CM = 3.5
PORTRAIT_IMG_HEIGHT_CM = round(3.5 * (4.0 / 3.0), 2)

# ระยะย่อหน้าในหน้าประวัติคณะผู้จัดทำ (หน่วยมิลลิเมตร)
BIO_LABEL_INDENT_MM = 30.0
BIO_DATA_INDENT_MM = 80.0

st.set_page_config(
    page_title="ระบบสร้างเอกสารรายงานโครงงานอาชีพ",
    page_icon="📘",
    layout="wide"
)

# ==========================================================
# ส่วนที่ 1.5: ฟังก์ชันตกแต่งหน้าตาระบบด้วยฟอนต์ทางการ (IBM Plex Sans Thai)
# ==========================================================
def inject_formal_academic_css(ui_font_family="IBM Plex Sans Thai"):
    """ตกแต่งหน้าเว็บด้วยฟอนต์ทางการ คมชัด และเหมาะกับการเรียนการสอน"""
    st.markdown(
        f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Bai+Jamjuree:wght@400;500;600;700&family=IBM+Plex+Sans+Thai:wght@400;500;600;700&family=Noto+Serif+Thai:wght@500;600;700&display=swap');

        html, body, .stApp, .stMarkdown, .stText,
        div.stButton > button, div.stDownloadButton > button,
        div[data-baseweb="select"], div[data-baseweb="input"],
        label, p, h1, h2, h3, h4, h5, h6, th, td, li {{
            font-family: '{ui_font_family}', 'Bai Jamjuree', -apple-system, BlinkMacSystemFont, sans-serif !important;
            -webkit-font-smoothing: antialiased;
            -moz-osx-font-smoothing: grayscale;
        }}

        span.material-symbols-rounded, i.material-icons {{
            font-family: 'Material Symbols Rounded' !important;
        }}

        .stApp {{
            background-color: #f0f4f8;
            color: #1e293b;
        }}

        .edu-hero-banner {{
            background: linear-gradient(135deg, #0f2942 0%, #1b4332 65%, #2d6a4f 100%);
            color: #ffffff;
            padding: 26px 32px;
            border-radius: 14px;
            border-bottom: 4px solid #d4af37;
            box-shadow: 0 6px 20px rgba(15, 41, 66, 0.14);
            display: flex;
            align-items: center;
            gap: 24px;
            margin-bottom: 26px;
        }}
        .edu-hero-logo {{
            width: 84px;
            height: 84px;
            object-fit: contain;
            background: #ffffff;
            border-radius: 50%;
            padding: 6px;
            border: 2px solid #d4af37;
            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.2);
            flex-shrink: 0;
        }}
        .edu-hero-title {{
            font-family: 'Noto Serif Thai', '{ui_font_family}', serif !important;
            font-size: 1.7rem;
            font-weight: 700;
            margin: 0 0 6px 0;
            letter-spacing: 0.4px;
            color: #ffffff;
        }}
        .edu-hero-subtitle {{
            font-size: 1.0rem;
            font-weight: 400;
            margin: 0;
            color: #e2e8f0;
            letter-spacing: 0.2px;
        }}

        .edu-info-card {{
            background: #ffffff;
            border-radius: 12px;
            padding: 22px 26px;
            border: 1px solid #cbd5e1;
            border-top: 4px solid #0f2942;
            box-shadow: 0 4px 12px rgba(15, 23, 42, 0.05);
            margin-bottom: 22px;
            line-height: 1.7;
        }}
        .edu-card-heading {{
            font-size: 1.15rem;
            font-weight: 700;
            color: #0f2942;
            padding-bottom: 10px;
            margin-bottom: 12px;
            border-bottom: 1px solid #e2e8f0;
        }}
        .edu-badge {{
            display: inline-block;
            background-color: #f8fafc;
            color: #0f2942;
            padding: 5px 14px;
            border-radius: 6px;
            font-weight: 500;
            font-size: 0.92rem;
            margin-right: 8px;
            margin-bottom: 8px;
            border: 1px solid #cbd5e1;
        }}

        div.stButton > button[kind="primary"] {{
            background: linear-gradient(135deg, #0f2942 0%, #1b4332 100%) !important;
            color: #ffffff !important;
            border: 1px solid #0f2942 !important;
            border-radius: 10px !important;
            padding: 12px 22px !important;
            font-size: 1.02rem !important;
            font-weight: 600 !important;
            letter-spacing: 0.2px !important;
            box-shadow: 0 4px 12px rgba(15, 41, 66, 0.2) !important;
        }}
        div.stButton > button[kind="primary"]:hover {{
            background: linear-gradient(135deg, #163a5f 0%, #245540 100%) !important;
        }}

        div.stDownloadButton > button {{
            border-radius: 10px !important;
            padding: 12px 18px !important;
            font-size: 0.98rem !important;
            font-weight: 600 !important;
            border: 1.5px solid #0f2942 !important;
            color: #0f2942 !important;
            background-color: #ffffff !important;
        }}
        div.stDownloadButton > button:hover {{
            background-color: #f1f5f9 !important;
        }}

        .edu-footer {{
            margin-top: 48px;
            padding: 18px 24px;
            background: #ffffff;
            border-top: 3px solid #0f2942;
            border-radius: 10px;
            text-align: center;
            color: #1e293b;
            font-size: 1.0rem;
            font-weight: 600;
            letter-spacing: 0.3px;
            box-shadow: 0 -2px 10px rgba(0, 0, 0, 0.03);
        }}
        </style>
        """,
        unsafe_allow_html=True
    )


def render_hero_header():
    """แสดงส่วนหัวของระบบพร้อมตราโรงเรียนและชื่อระบบ"""
    st.markdown(
        f"""
        <div class="edu-hero-banner">
            <img src="{SCHOOL_LOGO_URL}" class="edu-hero-logo" alt="ตราโรงเรียนวาปีปทุม">
            <div>
                <div class="edu-hero-title">ระบบสร้างเอกสารรายงานโครงงานอาชีพ</div>
                <div class="edu-hero-subtitle">
                    รายวิชา การงานอาชีพเพิ่มเติม (โครงงานอาชีพ) รหัสวิชา ง30202 | โรงเรียนวาปีปทุม จังหวัดมหาสารคาม
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


def render_footer():
    """แสดงข้อความผู้จัดทำที่ด้านล่างของระบบ"""
    st.markdown(
        """
        <div class="edu-footer">
            จัดทำโดย นายธนโชติ จันทรังษี ครูโรงเรียนวาปีปทุม
        </div>
        """,
        unsafe_allow_html=True
    )

# ==========================================================
# ส่วนที่ 2: ฟังก์ชันสร้างชื่อไฟล์ดาวน์โหลด ดึงข้อมูล และระบบตัดคำภาษาไทย
# ==========================================================
def extract_room_number_only(room_raw):
    """
    สกัดเฉพาะตัวเลขห้องเรียน เช่น:
    - '5/12' หรือ 'ม.5/12' -> '12'
    - '5/5' -> '5'
    - '12' -> '12'
    """
    r = str(room_raw or "").strip()
    if not r or r == "-":
        return "ไม่ระบุห้อง"
    if "/" in r:
        parts = r.split("/")
        last_part = parts[-1].strip()
        return last_part if last_part else r
    # กรณีไม่มีเครื่องหมาย / ให้ดึงตัวเลขออกมา
    nums = re.findall(r"\d+", r)
    return nums[-1] if nums else r


def build_export_filename(group_data, extension="pdf"):
    """
    ข้อ 1 (ใหม่): สร้างชื่อไฟล์ดาวน์โหลด (ทั้ง Word และ PDF) ในรูปแบบ:
    'โครงงานอาชีพ เรื่อง (ชื่อโครงงาน) ชั้นมัธยมศึกษาปีที่ 5 ห้อง (เลขห้อง).นามสกุลไฟล์'
    เช่น 'โครงงานอาชีพ เรื่อง การศึกษาและทำขนมพระพาย ชั้นมัธยมศึกษาปีที่ 5 ห้อง 12.pdf'
    """
    progress_steps = group_data.get("progress", {}).get("steps", {})
    step1_data = (progress_steps.get("1") or progress_steps.get(1) or {}).get("data", {})

    raw_title = (
        step1_data.get("projectTitle")
        or group_data.get("projectTitle")
        or group_data.get("groupName")
        or "ไม่ระบุชื่อเรื่อง"
    )
    clean_title = clean_single_line_text(raw_title)
    # ตัดคำว่า "เรื่อง" ด้านหน้าออกหากนักเรียนเผลอพิมพ์คำว่า "เรื่อง" ซ้ำมาในช่องชื่อโครงงาน
    clean_title = re.sub(r"^เรื่อง\s*", "", clean_title).strip()

    members = group_data.get("members", [])
    first_room_raw = members[0].get("room", "") if members else ""
    room_num_str = extract_room_number_only(first_room_raw)

    base_name = f"โครงงานอาชีพ เรื่อง {clean_title} ชั้นมัธยมศึกษาปีที่ 5 ห้อง {room_num_str}"
    # ลบอักขระต้องห้ามของชื่อไฟล์ในระบบปฏิบัติการ Windows / macOS / iPadOS
    safe_base_name = re.sub(r'[\\/*?:"<>|]', "", base_name).strip()
    clean_ext = str(extension).lstrip(".")
    return f"{safe_base_name}.{clean_ext}"


@st.cache_data(ttl=3600)
def download_url_bytes(url):
    """ดาวน์โหลดข้อมูลรูปภาพจาก URL มาเก็บไว้ในหน่วยความจำ"""
    try:
        resp = requests.get(url, timeout=15)
        resp.raise_for_status()
        return resp.content
    except Exception:
        return None


def load_image_bytes_from_source(img_src, target_aspect_ratio=None):
    """แปลงแหล่งข้อมูลรูปภาพพร้อมครอปตัดจากกึ่งกลางให้ได้อัตราส่วน 4:3 หรือ 3:4"""
    if not img_src or not isinstance(img_src, str):
        return None
    src = img_src.strip()
    if not src:
        return None

    raw_bytes = None
    try:
        if src.startswith("data:image"):
            parts = src.split(",", 1)
            if len(parts) == 2:
                raw_bytes = base64.b64decode(parts[1])
        elif src.startswith("http://") or src.startswith("https://"):
            raw_bytes = download_url_bytes(src)
    except Exception:
        raw_bytes = None

    if not raw_bytes:
        return None

    if PIL_AVAILABLE:
        try:
            with Image.open(io.BytesIO(raw_bytes)) as img:
                if img.mode in ("RGBA", "P"):
                    img = img.convert("RGB")
                
                if target_aspect_ratio is not None and target_aspect_ratio > 0:
                    w, h = img.size
                    current_ratio = w / float(h)
                    if abs(current_ratio - target_aspect_ratio) > 0.02:
                        if current_ratio > target_aspect_ratio:
                            new_w = int(h * target_aspect_ratio)
                            left = (w - new_w) // 2
                            img = img.crop((left, 0, left + new_w, h))
                        else:
                            new_h = int(w / target_aspect_ratio)
                            top = (h - new_h) // 2
                            img = img.crop((0, top, w, top + new_h))

                out_buf = io.BytesIO()
                img.save(out_buf, format="PNG")
                return out_buf.getvalue()
        except Exception:
            return raw_bytes
    return raw_bytes


def format_calc_number(val):
    """หากตัวเลขการคำนวณเป็นทศนิยม ให้แสดงเป็นทศนิยม 2 ตำแหน่งเสมอ"""
    if val is None or str(val).strip() == "" or str(val).strip() == "-":
        return "-"
    try:
        cleaned_str = str(val).replace(",", "").strip()
        num = float(cleaned_str)
        if num.is_integer():
            return f"{int(num):,}"
        return f"{num:,.2f}"
    except Exception:
        return str(val).strip()


def format_school_line(school_raw):
    """จัดรูปแบบชื่อโรงเรียนให้ขึ้นต้นด้วยคำว่า 'โรงเรียน' เสมอโดยไม่ซ้ำคำ"""
    s = clean_single_line_text(school_raw)
    if not s or s == "-":
        return "โรงเรียน........................................"
    if s.startswith("โรงเรียน"):
        return s
    return f"โรงเรียน{s}"


def count_thai_base_chars(text):
    """นับจำนวนตัวอักษรภาษาไทยเฉพาะตัวฐานเพื่อคำนวณความยาวบรรทัด"""
    if not text:
        return 0
    no_zwsp = str(text).replace("\u200b", "")
    base_only = re.sub(r"[\u0E31\u0E34-\u0E3A\u0E47-\u0E4E]", "", no_zwsp)
    return len(base_only)


def estimate_paragraph_lines(text, first_line_indent_cm=1.5, left_indent_cm=0.0):
    """คำนวณจำนวนบรรทัดของย่อหน้าเมื่อแสดงผลบนกระดาษ A4 (ขอบ 25.4 มม. พื้นที่พิมพ์กว้าง 15.92 ซม.)"""
    base_len = count_thai_base_chars(text)
    if base_len == 0:
        return 1
    usable_width_first = max(4.0, 15.92 - left_indent_cm - first_line_indent_cm)
    usable_width_next = max(4.0, 15.92 - left_indent_cm)

    chars_first_line = int(usable_width_first * 4.8)
    chars_next_line = int(usable_width_next * 4.8)

    if base_len <= chars_first_line:
        return 1
    remaining = base_len - chars_first_line
    return 1 + math.ceil(remaining / float(chars_next_line))


@st.cache_data(ttl=120, show_spinner=False)
def fetch_all_groups_from_gas(web_app_url):
    """ดึงข้อมูลทุกกลุ่มจาก Google Sheets ผ่านพารามิเตอร์ ?page=api (จำข้อมูลไว้ 2 นาทีเพื่อความรวดเร็วเมื่อใช้หลายคน)"""
    clean_url = web_app_url.strip()
    if not clean_url:
        return {"success": False, "message": "กรุณาวาง URL ของเว็บแอป Google Apps Script ในแถบด้านซ้ายมือก่อนครับ"}
    
    separator = "&" if "?" in clean_url else "?"
    api_url = f"{clean_url}{separator}page=api"
    
    try:
        response = requests.get(api_url, timeout=25)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        return {"success": False, "message": f"ไม่สามารถดึงข้อมูลได้: {e}"}


def repair_known_typos(text):
    """ซ่อมแซมคำที่พยัญชนะต้นอาจตกหล่นในฐานข้อมูล เช่น 'ัญหา' -> 'ปัญหา', 'วยตวง' -> 'ถ้วยตวง'"""
    if not text:
        return ""
    s = str(text)
    s = re.sub(r"(?<![ป])ัญหา", "ปัญหา", s)
    s = re.sub(r"(?<![ถ้])วยตวง", "ถ้วยตวง", s)
    return s


def clean_single_line_text(text):
    """ลบการขึ้นบรรทัดใหม่ (Enter) ช่องว่างซ้ำซ้อน และซ่อมคำตกหล่น"""
    if text is None:
        return ""
    cleaned = re.sub(r"[\r\n\t]+", " ", str(text))
    cleaned = re.sub(r"\s{2,}", " ", cleaned).strip()
    return repair_known_typos(cleaned)


def strip_leading_number(text):
    """ลบตัวเลขลำดับข้อด้านหน้าออกเพื่อป้องกันเลขข้อซ้ำ"""
    cleaned = clean_single_line_text(text)
    stripped = re.sub(r"^[\d\.]+\s*[\.\)]?\s*", "", cleaned).strip()
    return repair_known_typos(stripped)


def insert_thai_word_breaks(text):
    """แทรกอักขระจุดตัดคำที่มองไม่เห็น (Zero-Width Space: \u200b) ระหว่างคำและพยางค์ภาษาไทย"""
    clean_txt = clean_single_line_text(text)
    if not clean_txt:
        return ""

    zwsp = "\u200b"

    if THAI_TOKENIZER_AVAILABLE:
        tokens = word_tokenize(clean_txt, engine="newmm", keep_whitespace=True)
        out = []
        for i, tok in enumerate(tokens):
            out.append(tok)
            if i < len(tokens) - 1:
                next_tok = tokens[i + 1]
                if not tok.strip() or not next_tok.strip():
                    continue
                if re.search(r"[\u0E00-\u0E7F]$", tok) and re.search(r"^[\u0E00-\u0E7F]", next_tok):
                    out.append(zwsp)
        return "".join(out)

    fallback_words = [
        "เอกสาร", "รายงาน", "โครงงาน", "อาชีพ", "ฉบับนี้", "เป็นส่วนหนึ่ง", "ของรายวิชา",
        "ภาคเรียนที่", "ปีการศึกษา", "โดยมี", "วัตถุประสงค์", "เพื่อศึกษา", "ประวัติ",
        "และแหล่งที่มา", "ของ", "พัฒนา", "ทักษะ", "การทำ", "และสร้าง", "รายได้",
        "ระหว่างเรียน", "และใช้", "เวลาว่าง", "ให้เกิด", "ประโยชน์", "จัดทำขึ้น",
        "เนื่องจาก", "เล็งเห็น", "ความสำคัญ", "ของทักษะ", "การใช้เวลา", "ตลอดจน",
        "การพัฒนา", "ต่อยอด", "ภูมิปัญญา", "หรือ", "ความต้องการ", "ของตลาด", "ในปัจจุบัน",
        "ในการจัดทำ", "ครั้งนี้", "คณะผู้จัดทำ", "ได้รับ", "ความช่วยเหลือ", "คำแนะนำ",
        "และการสนับสนุน", "อย่างดียิ่ง", "จาก", "คุณครู", "ครูที่ปรึกษา", "ผู้ปกครอง", "และผู้ที่เกี่ยวข้อง",
        "ทุกท่าน", "ซึ่งมีส่วน", "ทำให้", "การดำเนินงาน", "สำเร็จ", "ลุล่วง", "ไปได้ด้วยดี",
        "ขอขอบพระคุณ", "หวังเป็นอย่างยิ่งว่า", "จะเป็นประโยชน์", "ต่อผู้ที่สนใจ", "ศึกษาหาข้อมูล",
        "หรือนำไป", "เป็นแนวทาง", "ในการประกอบอาชีพ", "ต่อไป", "หากมี", "ข้อผิดพลาด",
        "ประการใด", "ต้องขออภัย", "ผลิตภัณฑ์", "กระบวนการ", "กลุ่มเป้าหมาย", "จัดจำหน่าย",
        "ระยะเวลา", "ได้ศึกษา", "ทฤษฎี", "หรือแนวคิด", "ที่เกี่ยวข้อง", "โดยสามารถ", "สรุป",
        "สาระสำคัญ", "ตามลำดับ", "ดังนี้", "งานวิจัย", "ในระหว่าง", "พบปัญหา", "และอุปสรรค",
        "ข้อเสนอแนะ", "ในการพัฒนา", "ปรับปรุง", "และแก้ไข", "ได้มีการใช้", "ในการดำเนิน",
        "โดยมีรายละเอียด", "การคำนวณ", "ต้นทุน", "ดังแสดง", "ในตารางที่", "ดังต่อไปนี้",
        "ได้คำนวณ", "รายละเอียด", "ราคาจัดจำหน่าย", "และกำไร",
        "ปัญหา", "ในการเตรียม", "วัตถุดิบ", "และการทำ", "ขนม", "เช่น", "การกวนไส้", "ถั่วเหลือง",
        "ต้องใช้", "เวลา", "และความชำนาญ", "ถ้ากวนไม่ดี", "อาจไหม้", "หรือเนื้อไส้", "ไม่เนียน",
        "ในการจำหน่าย", "และการตลาด", "การโปรโมท", "สินค้า", "ในช่วงแรก", "อาจเข้าถึง", "ลูกค้า",
        "ได้น้อย", "ในการทำงาน", "เป็นทีม", "และการบริหาร", "การแบ่งหน้าที่", "ภายในทีม",
        "ไม่ชัดเจน", "ทำให้เกิด", "ความล่าช้า", "ให้มีความ", "หลากหลาย", "การออกแบบ",
        "บรรจุภัณฑ์", "ที่สวยงาม", "และเป็นมิตร", "กับสิ่งแวดล้อม", "และปรับใช้"
    ]
    result = clean_txt
    for w in fallback_words:
        result = result.replace(w, f"{zwsp}{w}{zwsp}")

    result = re.sub(r"(?<=[\u0E00-\u0E7F])(?=[เแโไใ])", zwsp, result)
    result = re.sub(r"(?<=[\u0E00-\u0E7F])(?=(การ|ความ|และ|หรือ|เพื่อ|ของ|ใน|จาก|โดย|ตาม|เป็น|มี|ได้|ให้|ไม่|อาจ|ต้อง))", zwsp, result)

    result = re.sub(f"{zwsp}+", zwsp, result)
    result = result.replace(f" {zwsp}", " ").replace(f"{zwsp} ", " ")
    return result.strip(zwsp)


def format_class_room_display(room_raw):
    """แปลงข้อมูลชั้น/ห้องให้เป็นรูปแบบ 5/X เสมอ"""
    r = str(room_raw or "").strip()
    if not r or r == "-":
        return "5/..."
    if "/" in r:
        return r.replace("ม.", "").strip()
    return f"5/{r}"


def format_objectives_for_preface(objectives_list):
    """จัดรูปแบบข้อความวัตถุประสงค์จากปุ่มที่ 2 สำหรับใส่ในย่อหน้าที่ 1 ของคำนำ"""
    if not objectives_list or not isinstance(objectives_list, list):
        return "เพื่อ......................................................................"

    cleaned_items = []
    for raw_obj in objectives_list:
        text = strip_leading_number(raw_obj)
        if not text:
            continue
        if text.startswith("เพื่อ"):
            text = text[len("เพื่อ"):].strip()
        if text:
            cleaned_items.append(text)

    if not cleaned_items:
        return "เพื่อ......................................................................"

    if len(cleaned_items) == 1:
        return f"เพื่อ{cleaned_items[0]}"

    formatted_parts = []
    for idx, item in enumerate(cleaned_items):
        if idx == 0:
            formatted_parts.append(f"เพื่อ{item}")
        elif idx == len(cleaned_items) - 1:
            formatted_parts.append(f"และ{item}")
        else:
            formatted_parts.append(item)

    return " ".join(formatted_parts)


def format_advisors_for_preface(advisors_list, strip_honorific=True):
    """แปลงรายชื่อครูที่ปรึกษาโครงงานที่ระบุไว้ในปุ่มที่ 1 สำหรับใส่ในหน้าคำนำ ย่อหน้าที่ 2"""
    if not advisors_list or not isinstance(advisors_list, list):
        return "คุณครูธนโชติ จันทรังษี"

    formatted_names = []
    for raw_adv in advisors_list:
        name = clean_single_line_text(raw_adv)
        if not name:
            continue
        name = re.sub(r"^(คุณครู|ครู)\s*", "", name).strip()
        if strip_honorific:
            name = re.sub(r"^(นาย|นางสาว|นาง|น\.ส\.)\s*", "", name).strip()
        if name:
            formatted_names.append(f"คุณครู{name}")

    if not formatted_names:
        return "คุณครูธนโชติ จันทรังษี"
    if len(formatted_names) == 1:
        return formatted_names[0]
    if len(formatted_names) == 2:
        return f"{formatted_names[0]} และ{formatted_names[1]}"
    return " ".join(formatted_names[:-1]) + f" และ{formatted_names[-1]}"


def format_theory_citation_prefix(p):
    """สร้างข้อความอ้างอิงด้านหน้าเนื้อหาย่อหน้าในบทที่ 2 ส่วนที่ 1 ('ชื่อ นามสกุล (ปีที่พิมพ์) ')"""
    if not p or not isinstance(p, dict):
        return ""
    ref_type = p.get("refType", "none")
    if not ref_type or ref_type == "none":
        return ""

    author_str = ""
    year_str = ""

    if ref_type == "book":
        count = str(p.get("bookAuthorCount", "1"))
        authors = p.get("bookAuthors", [])
        a1 = authors[0] if len(authors) > 0 and isinstance(authors[0], dict) else {}
        full1 = f"{str(a1.get('firstName', '')).strip()} {str(a1.get('lastName', '')).strip()}".strip()

        if not full1:
            author_str = ""
        elif count == "1":
            author_str = full1
        elif count == "2":
            a2 = authors[1] if len(authors) > 1 and isinstance(authors[1], dict) else {}
            full2 = f"{str(a2.get('firstName', '')).strip()} {str(a2.get('lastName', '')).strip()}".strip()
            author_str = f"{full1} และ {full2}" if full2 else full1
        else:
            author_str = f"{full1} และคณะ"

        if p.get("bookYearType") == "no_year":
            year_str = "ม.ป.ป."
        else:
            year_str = str(p.get("bookYear", "")).strip() or "ม.ป.ป."

    elif ref_type == "website":
        if p.get("webHasAuthor") == "no":
            author_str = str(p.get("webArticleTitleAlt", "")).strip()
        else:
            first_w = str(p.get("webAuthorFirstName", "")).strip()
            last_w = str(p.get("webAuthorLastName", "")).strip()
            author_str = f"{first_w} {last_w}".strip()

        if p.get("webYearType") == "no_year":
            year_str = "ม.ป.ป."
        else:
            y = str(p.get("webYear", "")).strip() or "ม.ป.ป."
            dm = str(p.get("webDayMonth", "")).strip()
            year_str = f"{y}, {dm}" if dm else y

    if not author_str:
        return ""
    return f"{author_str} ({year_str}) "


def format_research_citation_prefix(r_obj):
    """สร้างข้อความอ้างอิงด้านหน้าเนื้อหาในบทที่ 2 ส่วนที่ 2 ('ชื่อ นามสกุล (ปีที่พิมพ์) ')"""
    if not r_obj or not isinstance(r_obj, dict):
        return ""
    count = str(r_obj.get("authorCount", "1"))
    authors = r_obj.get("authors", [])
    a1 = authors[0] if len(authors) > 0 and isinstance(authors[0], dict) else {}
    full1 = f"{str(a1.get('firstName', '')).strip()} {str(a1.get('lastName', '')).strip()}".strip()

    if not full1:
        return ""

    if count == "1":
        author_str = full1
    elif count == "2":
        a2 = authors[1] if len(authors) > 1 and isinstance(authors[1], dict) else {}
        full2 = f"{str(a2.get('firstName', '')).strip()} {str(a2.get('lastName', '')).strip()}".strip()
        author_str = f"{full1} และ {full2}" if full2 else full1
    else:
        author_str = f"{full1} และคณะ"

    if r_obj.get("yearType") == "no_year":
        year_str = "ม.ป.ป."
    else:
        year_str = str(r_obj.get("year", "")).strip() or "ม.ป.ป."

    return f"{author_str} ({year_str}) "

# ==========================================================
# ส่วนที่ 3: ฟังก์ชันจัดการเลขหน้า ขนาด 14 pt ระยะขอบ 25.4 มม. และย่อหน้า
# ==========================================================
def set_rPr_thai_font_and_size(rPr, font_name=REPORT_FONT_NAME, size_pt=BODY_FONT_SIZE, bold=False, italic=False):
    """ฝังค่าฟอนต์และขนาดตัวอักษรทั้งฝั่ง Latin (w:sz) และฝั่งภาษาไทย (w:szCs)"""
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:cs'), font_name)

    half_pt_str = str(int(round(size_pt * 2)))
    for sz_tag in ('w:sz', 'w:szCs'):
        sz_el = rPr.find(qn(sz_tag))
        if sz_el is None:
            sz_el = OxmlElement(sz_tag)
            rPr.append(sz_el)
        sz_el.set(qn('w:val'), half_pt_str)

    for b_tag in ('w:b', 'w:bCs'):
        b_el = rPr.find(qn(b_tag))
        if bold:
            if b_el is None:
                b_el = OxmlElement(b_tag)
                rPr.append(b_el)
            b_el.set(qn('w:val'), '1')
        else:
            if b_el is not None:
                rPr.remove(b_el)

    for i_tag in ('w:i', 'w:iCs'):
        i_el = rPr.find(qn(i_tag))
        if italic:
            if i_el is None:
                i_el = OxmlElement(i_tag)
                rPr.append(i_el)
            i_el.set(qn('w:val'), '1')
        else:
            if i_el is not None:
                rPr.remove(i_el)


def set_thai_font(run, font_name=REPORT_FONT_NAME, size_pt=BODY_FONT_SIZE, bold=False, italic=False):
    """ตั้งค่าฟอนต์ภาษาไทย (TH SarabunPSK) ขนาด ความหนา และตัวเอียงใน Run"""
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    
    rPr = run._element.get_or_add_rPr()
    set_rPr_thai_font_and_size(rPr, font_name=font_name, size_pt=size_pt, bold=bold, italic=italic)


def set_paragraph_default_rPr(paragraph, font_name=REPORT_FONT_NAME, size_pt=BODY_FONT_SIZE, bold=False):
    """ตั้งค่าฟอนต์เริ่มต้นในระดับย่อหน้า (w:pPr/w:rPr)"""
    pPr = paragraph._p.get_or_add_pPr()
    rPr = pPr.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        pPr.append(rPr)
    set_rPr_thai_font_and_size(rPr, font_name=font_name, size_pt=size_pt, bold=bold, italic=False)


def configure_document_thai_settings(doc, font_name=REPORT_FONT_NAME):
    """ตั้งค่าระดับเอกสารและสไตล์พื้นฐาน (Normal และ Header)"""
    doc._lines_needed_with_heading = 0
    doc._bookmark_id_counter = 1
    settings = doc.settings.element
    compat = settings.find(qn('w:compat'))
    if compat is None:
        compat = OxmlElement('w:compat')
        settings.append(compat)
    if compat.find(qn('w:doNotExpandShiftReturn')) is None:
        dont_expand = OxmlElement('w:doNotExpandShiftReturn')
        compat.append(dont_expand)

    if settings.find(qn('w:updateFields')) is None:
        update_fields = OxmlElement('w:updateFields')
        update_fields.set(qn('w:val'), 'true')
        settings.append(update_fields)

    for style_name, target_sz in [('Normal', BODY_FONT_SIZE), ('Header', PAGE_NUMBER_FONT_SIZE)]:
        if style_name in doc.styles:
            st_obj = doc.styles[style_name]
            st_obj.font.name = font_name
            st_obj.font.size = Pt(target_sz)
            st_obj.font.bold = False
            rPr = st_obj.element.get_or_add_rPr()
            set_rPr_thai_font_and_size(rPr, font_name=font_name, size_pt=target_sz, bold=False, italic=False)


def apply_single_spacing(paragraph, space_before_pt=0, space_after_pt=0):
    """กำหนดระยะห่างบรรทัดเป็น 1.0 พร้อมตั้งค่าระยะห่างก่อนและหลังย่อหน้า"""
    paragraph.paragraph_format.line_spacing = LINE_SPACING_VAL
    paragraph.paragraph_format.space_before = Pt(space_before_pt)
    paragraph.paragraph_format.space_after = Pt(space_after_pt)
    paragraph.paragraph_format.widow_control = True


def apply_standard_page_setup(section):
    """ปรับระยะขอบของเอกสารทุกด้าน (บน, ล่าง, ซ้าย, ขวา) เท่ากับ 25.4 มิลลิเมตร"""
    section.page_width = Mm(210.0)
    section.page_height = Mm(297.0)
    section.top_margin = Mm(PAGE_MARGIN_MM)
    section.bottom_margin = Mm(PAGE_MARGIN_MM)
    section.left_margin = Mm(PAGE_MARGIN_MM)
    section.right_margin = Mm(PAGE_MARGIN_MM)
    section.header_distance = Mm(12.7)
    section.footer_distance = Mm(12.7)


def clear_header_content(header):
    """ล้างเนื้อหาใน Header (สำหรับหน้าที่ไม่แสดงเลขหน้า)"""
    header.is_linked_to_previous = False
    for p in header.paragraphs:
        p.text = ""


def add_page_number_field_to_header(header, font_name=REPORT_FONT_NAME, size_pt=PAGE_NUMBER_FONT_SIZE):
    """สร้างเลขหน้าอัตโนมัติที่มุมบนขวา (Header) ด้วยฟอนต์ TH SarabunPSK ขนาด 14 pt บาง"""
    header.is_linked_to_previous = False
    p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    p.text = ""
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    apply_single_spacing(p, space_before_pt=0, space_after_pt=0)
    set_paragraph_default_rPr(p, font_name=font_name, size_pt=size_pt, bold=False)

    r_begin = p.add_run()
    set_thai_font(r_begin, font_name=font_name, size_pt=size_pt, bold=False)
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    r_begin._r.append(fldChar1)

    r_instr = p.add_run()
    set_thai_font(r_instr, font_name=font_name, size_pt=size_pt, bold=False)
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('w:space'), 'preserve')
    instrText.text = "PAGE"
    r_instr._r.append(instrText)

    r_sep = p.add_run()
    set_thai_font(r_sep, font_name=font_name, size_pt=size_pt, bold=False)
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    r_sep._r.append(fldChar2)

    r_val = p.add_run("1")
    set_thai_font(r_val, font_name=font_name, size_pt=size_pt, bold=False)

    r_end = p.add_run()
    set_thai_font(r_end, font_name=font_name, size_pt=size_pt, bold=False)
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    r_end._r.append(fldChar3)


def configure_section_page_numbering(section, fmt="decimal", start_num=None, show_on_first_page=True, show_on_other_pages=True, font_name=REPORT_FONT_NAME):
    """ตั้งค่ารูปแบบเลขหน้าของแต่ละ Section (ก-ฮ หรือ 1-99) พร้อมซ่อนเลขหน้าในหน้าแรกของบทและหน้าคั่นภาคผนวก"""
    apply_standard_page_setup(section)

    sectPr = section._sectPr
    pgNumType = sectPr.find(qn('w:pgNumType'))
    if pgNumType is None:
        pgNumType = OxmlElement('w:pgNumType')
        sectPr.append(pgNumType)
    pgNumType.set(qn('w:fmt'), fmt)

    if start_num is not None:
        pgNumType.set(qn('w:start'), str(start_num))
    else:
        if qn('w:start') in pgNumType.attrib:
            del pgNumType.attrib[qn('w:start')]

    if not show_on_first_page and not show_on_other_pages:
        section.different_first_page_header_footer = False
        clear_header_content(section.header)
    elif not show_on_first_page and show_on_other_pages:
        section.different_first_page_header_footer = True
        clear_header_content(section.first_page_header)
        add_page_number_field_to_header(section.header, font_name=font_name, size_pt=PAGE_NUMBER_FONT_SIZE)
    else:
        section.different_first_page_header_footer = False
        add_page_number_field_to_header(section.header, font_name=font_name, size_pt=PAGE_NUMBER_FONT_SIZE)


def add_bookmark_to_paragraph(doc, paragraph, bookmark_name):
    """ฝัง Bookmark ลงในย่อหน้าเพื่อให้หน้าสารบัญดึงเลขหน้าจริงไปแสดงผลได้อัตโนมัติ"""
    if not bookmark_name:
        return
    bm_id = str(getattr(doc, "_bookmark_id_counter", 1))
    doc._bookmark_id_counter = int(bm_id) + 1

    bm_start = OxmlElement('w:bookmarkStart')
    bm_start.set(qn('w:id'), bm_id)
    bm_start.set(qn('w:name'), bookmark_name)

    bm_end = OxmlElement('w:bookmarkEnd')
    bm_end.set(qn('w:id'), bm_id)

    paragraph._p.insert(0, bm_start)
    paragraph._p.append(bm_end)


def add_pageref_run(paragraph, bookmark_name, fallback_page_str="1", font_name=REPORT_FONT_NAME, size_pt=BODY_FONT_SIZE):
    """สร้างฟิลด์ PAGEREF ในหน้าสารบัญ โดยบังคับให้ตัวอักษรและตัวเลขระบุหน้ามีขนาด 16 บาง (bold=False) เสมอ"""
    set_paragraph_default_rPr(paragraph, font_name=font_name, size_pt=size_pt, bold=False)

    r_begin = paragraph.add_run()
    set_thai_font(r_begin, font_name=font_name, size_pt=size_pt, bold=False)
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    r_begin._r.append(fldChar1)

    r_instr = paragraph.add_run()
    set_thai_font(r_instr, font_name=font_name, size_pt=size_pt, bold=False)
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('w:space'), 'preserve')
    instrText.text = f"PAGEREF {bookmark_name} \\h"
    r_instr._r.append(instrText)

    r_sep = paragraph.add_run()
    set_thai_font(r_sep, font_name=font_name, size_pt=size_pt, bold=False)
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    r_sep._r.append(fldChar2)

    r_val = paragraph.add_run(str(fallback_page_str))
    set_thai_font(r_val, font_name=font_name, size_pt=size_pt, bold=False)

    r_end = paragraph.add_run()
    set_thai_font(r_end, font_name=font_name, size_pt=size_pt, bold=False)
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    r_end._r.append(fldChar3)


def apply_min_3_lines_guard(doc, paragraph, text_str, first_line_indent_cm=1.5, left_indent_cm=0.0):
    """ควบคุมไม่ให้หัวข้อหลักเหลือข้อความอยู่ท้ายหน้าเดิมน้อยกว่า 3 บรรทัด"""
    needed = getattr(doc, "_lines_needed_with_heading", 0)
    if needed <= 0:
        return

    est_lines = estimate_paragraph_lines(text_str, first_line_indent_cm=first_line_indent_cm, left_indent_cm=left_indent_cm)
    if est_lines <= 5:
        paragraph.paragraph_format.keep_together = True

    remaining_needed = needed - est_lines
    if remaining_needed > 0:
        paragraph.paragraph_format.keep_with_next = True
        doc._lines_needed_with_heading = remaining_needed
    else:
        doc._lines_needed_with_heading = 0


def set_thai_distribute(paragraph):
    """ตั้งค่าการจัดวางย่อหน้าเป็นแบบ 'กระจายแบบไทย' (Thai Distribute)"""
    pPr = paragraph._p.get_or_add_pPr()
    
    jc = pPr.find(qn('w:jc'))
    if jc is None:
        jc = OxmlElement('w:jc')
        pPr.append(jc)
    jc.set(qn('w:val'), 'thaiDistribute')

    wordWrap = pPr.find(qn('w:wordWrap'))
    if wordWrap is None:
        wordWrap = OxmlElement('w:wordWrap')
        pPr.append(wordWrap)
    wordWrap.set(qn('w:val'), '1')


def add_paragraph_in_cell(cell, text, align=WD_ALIGN_PARAGRAPH.CENTER, font_name=REPORT_FONT_NAME, size_pt=COVER_FONT_SIZE, bold=True, space_before_pt=0, space_after_pt=0):
    """เพิ่มย่อหน้าข้อความลงในเซลล์ตาราง"""
    p = cell.add_paragraph()
    p.alignment = align
    apply_single_spacing(p, space_before_pt=space_before_pt, space_after_pt=space_after_pt)
    run = p.add_run(text)
    set_thai_font(run, font_name=font_name, size_pt=size_pt, bold=bold)
    return p


def add_blank_line_16pt(doc, font_name=REPORT_FONT_NAME):
    """สร้างบรรทัดว่าง 1 บรรทัด ขนาดฟอนต์ 16 pt"""
    doc._lines_needed_with_heading = 0
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    apply_single_spacing(p, space_before_pt=0, space_after_pt=0)
    run = p.add_run("")
    set_thai_font(run, font_name=font_name, size_pt=BODY_FONT_SIZE, bold=False)
    return p


def set_cell_shading(cell, fill_hex=OLIVE_ACCENT3_LIGHTER80_HEX):
    """แรเงาพื้นหลังเซลล์หัวตารางด้วย 'สีมะกอก ส่วนที่ถูกเน้น 3 อ่อนลง 80%' (#EBF1DE)"""
    tcPr = cell._tc.get_or_add_tcPr()
    existing_shd = tcPr.find(qn('w:shd'))
    if existing_shd is not None:
        tcPr.remove(existing_shd)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    shd.set(qn('w:themeFill'), 'accent3')
    shd.set(qn('w:themeFillTint'), '33')
    tcPr.append(shd)


def set_cell_top_or_bottom_border(cell, edge_name, sz_eighth_pts=12):
    """บังคับความหนาเส้นขอบบนหรือเส้นขอบล่างของเซลล์ให้เท่ากับ 1.5 pt (sz=12)"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    existing = tcBorders.find(qn(f'w:{edge_name}'))
    if existing is not None:
        tcBorders.remove(existing)
    border = OxmlElement(f'w:{edge_name}')
    border.set(qn('w:val'), 'single')
    border.set(qn('w:sz'), str(sz_eighth_pts))
    border.set(qn('w:space'), '0')
    border.set(qn('w:color'), '000000')
    tcBorders.append(border)


def set_cell_properties_and_text(
    cell,
    text,
    width_cm,
    h_align=WD_ALIGN_PARAGRAPH.LEFT,
    v_align=WD_ALIGN_VERTICAL.TOP,
    font_name=REPORT_FONT_NAME,
    size_pt=BODY_FONT_SIZE,
    bold=False,
    shade_header=False
):
    """ตั้งค่าความกว้างของเซลล์ การจัดแนวตั้ง/แนวนอน ฟอนต์ และการแรเงาสีหัวตาราง"""
    cell.width = Cm(width_cm)
    cell.vertical_alignment = v_align
    set_cell_margins(cell, top=60, bottom=60, left=100, right=100)
    if shade_header:
        set_cell_shading(cell, fill_hex=OLIVE_ACCENT3_LIGHTER80_HEX)

    p = cell.paragraphs[0]
    p.alignment = h_align
    apply_single_spacing(p)
    run = p.add_run(str(text))
    set_thai_font(run, font_name=font_name, size_pt=size_pt, bold=bold)
    return p


def add_chapter_header(doc, chapter_num_str, chapter_title_str, font_name=REPORT_FONT_NAME, bookmark_name=None):
    """สร้างหัวข้อประจำบท 2 บรรทัด (ขนาด 20 หนา กึ่งกลาง เว้นระยะห่างหลัง 12 pt) พร้อมฝัง Bookmark"""
    doc._lines_needed_with_heading = 0
    p_num = doc.add_paragraph()
    p_num.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_single_spacing(p_num, space_before_pt=0, space_after_pt=0)
    p_num.paragraph_format.keep_with_next = True
    r_num = p_num.add_run(chapter_num_str)
    set_thai_font(r_num, font_name=font_name, size_pt=CHAPTER_TITLE_SIZE, bold=True)
    if bookmark_name:
        add_bookmark_to_paragraph(doc, p_num, bookmark_name)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_single_spacing(p_title, space_before_pt=0, space_after_pt=12)
    p_title.paragraph_format.keep_with_next = True
    r_title = p_title.add_run(chapter_title_str)
    set_thai_font(r_title, font_name=font_name, size_pt=CHAPTER_TITLE_SIZE, bold=True)


def add_main_section_heading(doc, text, font_name=REPORT_FONT_NAME, size_pt=SECTION_HEADING_SIZE, bold=True, bookmark_name=None):
    """สร้างหัวข้อหลักของแต่ละบท (เว้นระยะห่างก่อน 6 pt เสมอ) พร้อมผูกเนื้อหาอย่างน้อย 3 บรรทัดและฝัง Bookmark"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    apply_single_spacing(p, space_before_pt=6, space_after_pt=0)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.keep_together = True
    run = p.add_run(text)
    set_thai_font(run, font_name=font_name, size_pt=size_pt, bold=bold)
    if bookmark_name:
        add_bookmark_to_paragraph(doc, p, bookmark_name)

    doc._lines_needed_with_heading = MIN_LINES_WITH_HEADING
    return p


def add_sub_heading_left(doc, text, font_name=REPORT_FONT_NAME, size_pt=BODY_FONT_SIZE, bold=True):
    """สร้างหัวข้อย่อยชิดซ้าย (เช่น หัวข้อสาระสำคัญในบทที่ 2 ขนาด 16 หนา)"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    apply_single_spacing(p, space_before_pt=0, space_after_pt=0)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.keep_together = True
    run = p.add_run(text)
    set_thai_font(run, font_name=font_name, size_pt=size_pt, bold=bold)

    doc._lines_needed_with_heading = max(getattr(doc, "_lines_needed_with_heading", 0), MIN_LINES_WITH_HEADING)
    return p


def add_centered_table_title(doc, title_text, font_name=REPORT_FONT_NAME, size_pt=BODY_FONT_SIZE, bold=True, bookmark_name=None, page_break_before=False):
    """ชื่อตารางที่ 1 และตารางที่ 2 จัดอยู่กึ่งกลาง พร้อมรองรับการขึ้นหน้าใหม่ด้วย page_break_before"""
    doc._lines_needed_with_heading = 0
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_single_spacing(p, space_before_pt=6, space_after_pt=3)
    p.paragraph_format.keep_with_next = True
    if page_break_before:
        p.paragraph_format.page_break_before = True
    run = p.add_run(title_text)
    set_thai_font(run, font_name=font_name, size_pt=size_pt, bold=bold)
    if bookmark_name:
        add_bookmark_to_paragraph(doc, p, bookmark_name)
    return p


def add_subheading_indented(doc, text, font_name=REPORT_FONT_NAME, size_pt=BODY_FONT_SIZE, bold=True, left_indent_cm=1.0):
    """สร้างหัวข้อย่อยที่มีการย่อหน้าเข้าไปด้านใน (ขนาด 16 หนา)"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    apply_single_spacing(p)
    p.paragraph_format.left_indent = Cm(left_indent_cm)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.keep_together = True
    run = p.add_run(text)
    set_thai_font(run, font_name=font_name, size_pt=size_pt, bold=bold)

    doc._lines_needed_with_heading = max(getattr(doc, "_lines_needed_with_heading", 0), 2)
    return p


def add_indented_body_paragraph(doc, text, font_name=REPORT_FONT_NAME, size_pt=BODY_FONT_SIZE, indent_cm=1.5, use_thai_distribute=True, bold=False):
    """สร้างย่อหน้าเนื้อหาขนาด 16 บาง พร้อมแทรกจุดตัดคำภาษาไทยเสมอทุกย่อหน้า"""
    p = doc.add_paragraph()
    apply_single_spacing(p)
    p.paragraph_format.left_indent = Cm(0.0)
    p.paragraph_format.first_line_indent = Cm(indent_cm)

    final_text = insert_thai_word_breaks(text)

    if use_thai_distribute:
        set_thai_distribute(p)
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pPr = p._p.get_or_add_pPr()
        wordWrap = pPr.find(qn('w:wordWrap'))
        if wordWrap is None:
            wordWrap = OxmlElement('w:wordWrap')
            pPr.append(wordWrap)
        wordWrap.set(qn('w:val'), '1')

    run = p.add_run(final_text)
    set_thai_font(run, font_name=font_name, size_pt=size_pt, bold=bold)

    if bold:
        p.paragraph_format.keep_with_next = True
        doc._lines_needed_with_heading = max(getattr(doc, "_lines_needed_with_heading", 0), 2)
    else:
        apply_min_3_lines_guard(doc, p, text, first_line_indent_cm=indent_cm, left_indent_cm=0.0)
    return p


def add_landscape_image_with_caption(doc, img_src, caption_full_text, font_name=REPORT_FONT_NAME, bookmark_name=None, page_break_before=False):
    """แทรกรูปภาพแนวนอน 4:3 สูง 7.0 ซม. พร้อมคำบรรยายใต้ภาพ ขนาด 14 บาง เว้นระยะห่างหลัง 12 pt"""
    doc._lines_needed_with_heading = 0
    img_bytes = load_image_bytes_from_source(img_src, target_aspect_ratio=(4.0 / 3.0))
    if img_bytes:
        try:
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_single_spacing(p_img, space_before_pt=0, space_after_pt=0)
            p_img.paragraph_format.keep_with_next = True
            p_img.paragraph_format.keep_together = True
            if page_break_before:
                p_img.paragraph_format.page_break_before = True
                page_break_before = False
            run_img = p_img.add_run()
            run_img.add_picture(
                io.BytesIO(img_bytes),
                width=Cm(LANDSCAPE_IMG_WIDTH_CM),
                height=Cm(LANDSCAPE_IMG_HEIGHT_CM)
            )
        except Exception:
            pass

    if caption_full_text:
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_single_spacing(p_cap, space_before_pt=0, space_after_pt=12)
        p_cap.paragraph_format.keep_together = True
        if page_break_before:
            p_cap.paragraph_format.page_break_before = True
        r_cap = p_cap.add_run(clean_single_line_text(caption_full_text))
        set_thai_font(r_cap, font_name=font_name, size_pt=CAPTION_FONT_SIZE, bold=False)
        if bookmark_name:
            add_bookmark_to_paragraph(doc, p_cap, bookmark_name)


def remove_all_table_borders(table):
    """ซ่อนเส้นขอบตารางทั้งหมด"""
    tblPr = table._tbl.tblPr
    existing_borders = tblPr.find(qn('w:tblBorders'))
    if existing_borders is not None:
        tblPr.remove(existing_borders)
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        tblBorders.append(border)
    tblPr.append(tblBorders)


def set_row_repeat_header(row):
    """ตั้งค่าให้แถวหัวตารางแสดงซ้ำอัตโนมัติเมื่อตารางยาวข้ามหน้า"""
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn('w:tblHeader')) is None:
        tblHeader = OxmlElement('w:tblHeader')
        trPr.append(tblHeader)


def set_table_grid_borders(table):
    """เส้นขอบบนและล่างของตารางใช้น้ำหนักปากกา 1 1/2 pt (sz=12)"""
    tblPr = table._tbl.tblPr
    existing_borders = tblPr.find(qn('w:tblBorders'))
    if existing_borders is not None:
        tblPr.remove(existing_borders)

    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        if border_name in ('top', 'bottom'):
            border.set(qn('w:sz'), '12')
        else:
            border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tblBorders.append(border)
    tblPr.append(tblBorders)


def finalize_table_top_bottom_thick_borders(table):
    """ย้ำเส้นขอบบนสุดของแถวแรกและเส้นขอบล่างสุดของแถวสุดท้ายให้เป็น 1.5 pt (sz=12)"""
    if not table.rows:
        return
    for cell in table.rows[0].cells:
        set_cell_top_or_bottom_border(cell, 'top', sz_eighth_pts=12)
    for cell in table.rows[-1].cells:
        set_cell_top_or_bottom_border(cell, 'bottom', sz_eighth_pts=12)


def set_table_row_height(row, height_cm):
    """กำหนดความสูงของแถวตาราง"""
    trPr = row._tr.get_or_add_trPr()
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), str(int(height_cm * 567)))
    trHeight.set(qn('w:hRule'), 'exact')
    trPr.append(trHeight)


def set_cell_margins(cell, top=0, bottom=0, left=100, right=100):
    """ปรับระยะห่างภายในเซลล์ตาราง"""
    tcPr = cell._tc.get_or_add_tcPr()
    existing_mar = tcPr.find(qn('w:tcMar'))
    if existing_mar is not None:
        tcPr.remove(existing_mar)
    tcMar = OxmlElement('w:tcMar')
    for m_name, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m_name}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def render_paginated_table(
    doc,
    base_title,
    headers,
    col_widths,
    col_alignments,
    rows_data,
    current_y_cm,
    usable_page_height_cm=24.2,
    font_name=REPORT_FONT_NAME,
    max_rows_per_page=18,
    bookmark_name=None
):
    """สร้างตารางพร้อมระบบตัดหน้า '(ต่อ)' อัตโนมัติโดยไม่ทำให้เกิดหน้าว่าง"""
    title_plus_header_height = 2.0
    start_first_chunk_on_new_page = False

    if current_y_cm + title_plus_header_height + 1.8 > usable_page_height_cm:
        start_first_chunk_on_new_page = True
        current_y_cm = 0.0

    def create_new_table_chunk(is_continued=False, force_new_page=False):
        display_title = f"{base_title} (ต่อ)" if is_continued else base_title
        bm_to_use = None if is_continued else bookmark_name
        add_centered_table_title(
            doc,
            display_title,
            font_name=font_name,
            size_pt=BODY_FONT_SIZE,
            bold=True,
            bookmark_name=bm_to_use,
            page_break_before=(is_continued or force_new_page)
        )
        
        tbl = doc.add_table(rows=1, cols=len(headers))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        set_table_grid_borders(tbl)

        for c_i, w_cm in enumerate(col_widths):
            tbl.columns[c_i].width = Cm(w_cm)

        hdr_cells = tbl.rows[0].cells
        for c_i, h_text in enumerate(headers):
            set_cell_properties_and_text(
                hdr_cells[c_i],
                h_text,
                width_cm=col_widths[c_i],
                h_align=WD_ALIGN_PARAGRAPH.CENTER,
                v_align=WD_ALIGN_VERTICAL.CENTER,
                font_name=font_name,
                size_pt=BODY_FONT_SIZE,
                bold=True,
                shade_header=True
            )
        return tbl

    active_table = create_new_table_chunk(is_continued=False, force_new_page=start_first_chunk_on_new_page)
    current_y_cm += title_plus_header_height
    rows_in_current_chunk = 0

    for row_info in rows_data:
        cells_text = row_info.get("cells", [])
        is_bold_row = row_info.get("bold", False)
        merge_first_n = row_info.get("merge_first_n", 0)
        custom_aligns = row_info.get("aligns", col_alignments)

        max_lines = 1
        for c_i, txt in enumerate(cells_text):
            w_cm = col_widths[c_i] if c_i < len(col_widths) else 3.5
            chars_per_line = max(10, int(w_cm * 4.5))
            est_lines = math.ceil(len(str(txt)) / float(chars_per_line))
            if est_lines > max_lines:
                max_lines = est_lines
        row_height_cm = max(0.85, max_lines * 0.65 + 0.2)

        if rows_in_current_chunk > 0 and (
            current_y_cm + row_height_cm > usable_page_height_cm or rows_in_current_chunk >= max_rows_per_page
        ):
            finalize_table_top_bottom_thick_borders(active_table)
            current_y_cm = 0.0
            active_table = create_new_table_chunk(is_continued=True, force_new_page=True)
            current_y_cm += title_plus_header_height
            rows_in_current_chunk = 0

        r_cells = active_table.add_row().cells
        if merge_first_n and merge_first_n > 1:
            r_cells[0].merge(r_cells[merge_first_n - 1])
            merged_w = sum(col_widths[:merge_first_n])
            set_cell_properties_and_text(
                r_cells[0],
                cells_text[0],
                width_cm=merged_w,
                h_align=custom_aligns[0],
                v_align=WD_ALIGN_VERTICAL.TOP,
                font_name=font_name,
                size_pt=BODY_FONT_SIZE,
                bold=is_bold_row
            )
            for c_i in range(merge_first_n, len(col_widths)):
                set_cell_properties_and_text(
                    r_cells[c_i],
                    cells_text[c_i],
                    width_cm=col_widths[c_i],
                    h_align=custom_aligns[c_i],
                    v_align=WD_ALIGN_VERTICAL.TOP,
                    font_name=font_name,
                    size_pt=BODY_FONT_SIZE,
                    bold=is_bold_row
                )
        else:
            for c_i in range(len(col_widths)):
                txt_val = cells_text[c_i] if c_i < len(cells_text) else ""
                set_cell_properties_and_text(
                    r_cells[c_i],
                    txt_val,
                    width_cm=col_widths[c_i],
                    h_align=custom_aligns[c_i],
                    v_align=WD_ALIGN_VERTICAL.TOP,
                    font_name=font_name,
                    size_pt=BODY_FONT_SIZE,
                    bold=is_bold_row
                )

        current_y_cm += row_height_cm
        rows_in_current_chunk += 1

    finalize_table_top_bottom_thick_borders(active_table)
    add_blank_line_16pt(doc, font_name=font_name)
    current_y_cm += 0.8
    return current_y_cm

# ==========================================================
# ส่วนที่ 4: ฟังก์ชันสร้างหน้าสารบัญ สารบัญตาราง และสารบัญภาพประกอบ
# ==========================================================
def add_main_toc_row(
    tbl,
    title_text,
    bookmark_name,
    fallback_page_str,
    font_name=REPORT_FONT_NAME,
    title_bold=False,
    left_indent_cm=0.0,
    space_before_pt=0
):
    """เพิ่มแถวในตาราง 'สารบัญ' (2 คอลัมน์ไร้เส้นขอบ)"""
    row_cells = tbl.add_row().cells
    c_title, c_page = row_cells[0], row_cells[1]
    c_title.width = Cm(14.4)
    c_page.width = Cm(1.5)
    c_title.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    c_page.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    set_cell_margins(c_title, top=0, bottom=0, left=0, right=80)
    set_cell_margins(c_page, top=0, bottom=0, left=0, right=0)

    p_title = c_title.paragraphs[0]
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    apply_single_spacing(p_title, space_before_pt=space_before_pt, space_after_pt=0)
    if left_indent_cm > 0:
        p_title.paragraph_format.left_indent = Cm(left_indent_cm)
    r_t = p_title.add_run(clean_single_line_text(title_text))
    set_thai_font(r_t, font_name=font_name, size_pt=BODY_FONT_SIZE, bold=title_bold)

    p_page = c_page.paragraphs[0]
    p_page.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_single_spacing(p_page, space_before_pt=space_before_pt, space_after_pt=0)
    set_paragraph_default_rPr(p_page, font_name=font_name, size_pt=BODY_FONT_SIZE, bold=False)

    if bookmark_name and WORD_COM_AVAILABLE:
        add_pageref_run(p_page, bookmark_name, fallback_page_str=fallback_page_str, font_name=font_name, size_pt=BODY_FONT_SIZE)
    else:
        r_pg = p_page.add_run(str(fallback_page_str))
        set_thai_font(r_pg, font_name=font_name, size_pt=BODY_FONT_SIZE, bold=False)


def add_numbered_toc_row(
    tbl,
    seq_num_str,
    item_title_text,
    bookmark_name,
    fallback_page_str,
    col1_width_cm,
    col2_width_cm,
    col3_width_cm=1.5,
    font_name=REPORT_FONT_NAME
):
    """เพิ่มแถวในตาราง 'สารบัญตาราง' หรือ 'สารบัญภาพประกอบ' (3 คอลัมน์ไร้เส้นขอบ)"""
    row_cells = tbl.add_row().cells
    c_num, c_desc, c_page = row_cells[0], row_cells[1], row_cells[2]
    c_num.width = Cm(col1_width_cm)
    c_desc.width = Cm(col2_width_cm)
    c_page.width = Cm(col3_width_cm)

    for c in (c_num, c_desc, c_page):
        c.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    set_cell_margins(c_num, top=0, bottom=0, left=0, right=0)
    set_cell_margins(c_desc, top=0, bottom=0, left=80, right=80)
    set_cell_margins(c_page, top=0, bottom=0, left=0, right=0)

    p_num = c_num.paragraphs[0]
    p_num.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_single_spacing(p_num, space_before_pt=0, space_after_pt=0)
    r_num = p_num.add_run(str(seq_num_str))
    set_thai_font(r_num, font_name=font_name, size_pt=BODY_FONT_SIZE, bold=False)

    p_desc = c_desc.paragraphs[0]
    p_desc.alignment = WD_ALIGN_PARAGRAPH.LEFT
    apply_single_spacing(p_desc, space_before_pt=0, space_after_pt=0)
    r_desc = p_desc.add_run(clean_single_line_text(item_title_text))
    set_thai_font(r_desc, font_name=font_name, size_pt=BODY_FONT_SIZE, bold=False)

    p_page = c_page.paragraphs[0]
    p_page.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_single_spacing(p_page, space_before_pt=0, space_after_pt=0)
    set_paragraph_default_rPr(p_page, font_name=font_name, size_pt=BODY_FONT_SIZE, bold=False)

    if bookmark_name and WORD_COM_AVAILABLE:
        add_pageref_run(p_page, bookmark_name, fallback_page_str=fallback_page_str, font_name=font_name, size_pt=BODY_FONT_SIZE)
    else:
        r_pg = p_page.add_run(str(fallback_page_str))
        set_thai_font(r_pg, font_name=font_name, size_pt=BODY_FONT_SIZE, bold=False)


def build_all_toc_pages(doc, group_data, font_name=REPORT_FONT_NAME):
    """สร้างหน้าสารบัญ, สารบัญตาราง, และสารบัญภาพประกอบ"""
    progress_steps = group_data.get("progress", {}).get("steps", {})
    step4_data = (progress_steps.get("4") or progress_steps.get(4) or {}).get("data", {})
    step5_data = (progress_steps.get("5") or progress_steps.get(5) or {}).get("data", {})
    step7_data = (progress_steps.get("7") or progress_steps.get(7) or {}).get("data", {})

    mat_type = clean_single_line_text(step4_data.get("materialType", "วัตถุดิบ")) or "วัตถุดิบ"

    # 1. หน้าสารบัญ
    doc.add_page_break()
    p_toc_title = doc.add_paragraph()
    p_toc_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_single_spacing(p_toc_title, space_before_pt=0, space_after_pt=12)
    r_toc_title = p_toc_title.add_run("สารบัญ")
    set_thai_font(r_toc_title, font_name=font_name, size_pt=CHAPTER_TITLE_SIZE, bold=True)
    add_bookmark_to_paragraph(doc, p_toc_title, "bm_toc_main")

    toc_main_tbl = doc.add_table(rows=1, cols=2)
    toc_main_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    toc_main_tbl.autofit = False
    remove_all_table_borders(toc_main_tbl)
    set_row_repeat_header(toc_main_tbl.rows[0])

    hdr_main = toc_main_tbl.rows[0].cells
    hdr_main[0].width = Cm(14.4)
    hdr_main[1].width = Cm(1.5)
    set_cell_margins(hdr_main[0], top=0, bottom=60, left=0, right=80)
    set_cell_margins(hdr_main[1], top=0, bottom=60, left=0, right=0)

    p_hm0 = hdr_main[0].paragraphs[0]
    p_hm0.alignment = WD_ALIGN_PARAGRAPH.LEFT
    apply_single_spacing(p_hm0)
    r_hm0 = p_hm0.add_run("เรื่อง")
    set_thai_font(r_hm0, font_name=font_name, size_pt=BODY_FONT_SIZE, bold=True)

    p_hm1 = hdr_main[1].paragraphs[0]
    p_hm1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_single_spacing(p_hm1)
    r_hm1 = p_hm1.add_run("หน้า")
    set_thai_font(r_hm1, font_name=font_name, size_pt=BODY_FONT_SIZE, bold=True)

    front_items = [
        ("คำนำ", "bm_preface", "ก"),
        ("สารบัญ", "bm_toc_main", "ข"),
        ("สารบัญตาราง", "bm_toc_tables", "ค"),
        ("สารบัญภาพประกอบ", "bm_toc_figures", "ง"),
    ]
    for f_title, f_bm, f_pg in front_items:
        add_main_toc_row(
            toc_main_tbl, f_title, f_bm, f_pg, font_name=font_name, title_bold=True, left_indent_cm=0.0
        )

    chapters_toc_structure = [
        ("บทที่ 1 บทนำ", "bm_ch1", "1", [
            ("ที่มาและความสำคัญของโครงงาน", "bm_ch1_1", ""),
            ("วัตถุประสงค์ของโครงงาน", "bm_ch1_2", "2"),
            ("ขอบเขตของโครงงาน", "bm_ch1_3", ""),
            ("ประโยชน์ที่คาดว่าจะได้รับ", "bm_ch1_4", ""),
        ]),
        ("บทที่ 2 เอกสารและงานที่เกี่ยวข้อง", "bm_ch2", "3", [
            ("ทฤษฎีหรือแนวคิดที่เกี่ยวข้องกับโครงงาน", "bm_ch2_1", ""),
            ("งานวิจัยและโครงงานที่เกี่ยวข้อง", "bm_ch2_2", "4"),
            ("สรุป", "bm_ch2_3", ""),
        ]),
        ("บทที่ 3 ขั้นตอนการดำเนินโครงงาน", "bm_ch3", "5", [
            ("วิธีการดำเนินงาน", "bm_ch3_1", ""),
            (f"{mat_type}และอุปกรณ์", "bm_ch3_2", ""),
            ("ขั้นตอนการทำ", "bm_ch3_3", "6"),
        ]),
        ("บทที่ 4 การอภิปรายผลโครงงาน", "bm_ch4", "7", [
            ("ขั้นตอนการดำเนินงานตามขั้นตอนการทำงาน", "bm_ch4_1", ""),
            ("การคำนวณต้นทุน", "bm_ch4_2", "8"),
            ("การคำนวณการจัดจำหน่าย", "bm_ch4_3", "9"),
        ]),
        ("บทที่ 5 สรุปผลโครงงาน", "bm_ch5", "10", [
            ("ผลการดำเนินงานตามวัตถุประสงค์", "bm_ch5_1", ""),
            ("ปัญหาและอุปสรรค", "bm_ch5_2", ""),
            ("ข้อเสนอแนะ", "bm_ch5_3", "11"),
        ]),
    ]

    for ch_title, ch_bm, ch_fb, sub_items in chapters_toc_structure:
        add_main_toc_row(
            toc_main_tbl, ch_title, ch_bm, ch_fb, font_name=font_name, title_bold=True, left_indent_cm=0.0, space_before_pt=3
        )
        for sub_title, sub_bm, sub_fb in sub_items:
            add_main_toc_row(
                toc_main_tbl, sub_title, sub_bm, sub_fb, font_name=font_name, title_bold=False, left_indent_cm=1.0, space_before_pt=0
            )

    back_items = [
        ("บรรณานุกรม", "bm_bibliography", "12"),
        ("ภาคผนวก", "bm_appendix", "13"),
        ("ประวัติคณะผู้จัดทำ", "bm_biography", "15"),
    ]
    for b_title, b_bm, b_pg in back_items:
        add_main_toc_row(
            toc_main_tbl, b_title, b_bm, b_pg, font_name=font_name, title_bold=True, left_indent_cm=0.0, space_before_pt=3
        )

    # 2. หน้าสารบัญตาราง
    doc.add_page_break()
    p_tbl_title = doc.add_paragraph()
    p_tbl_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_single_spacing(p_tbl_title, space_before_pt=0, space_after_pt=12)
    r_tbl_title = p_tbl_title.add_run("สารบัญตาราง")
    set_thai_font(r_tbl_title, font_name=font_name, size_pt=CHAPTER_TITLE_SIZE, bold=True)
    add_bookmark_to_paragraph(doc, p_tbl_title, "bm_toc_tables")

    toc_tbl_table = doc.add_table(rows=1, cols=3)
    toc_tbl_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    toc_tbl_table.autofit = False
    remove_all_table_borders(toc_tbl_table)
    set_row_repeat_header(toc_tbl_table.rows[0])

    hdr_t = toc_tbl_table.rows[0].cells
    hdr_t[0].width = Cm(2.0)
    hdr_t[1].width = Cm(12.4)
    hdr_t[2].width = Cm(1.5)
    set_cell_margins(hdr_t[0], top=0, bottom=60, left=0, right=0)
    set_cell_margins(hdr_t[1], top=0, bottom=60, left=80, right=80)
    set_cell_margins(hdr_t[2], top=0, bottom=60, left=0, right=0)

    p_ht0 = hdr_t[0].paragraphs[0]
    p_ht0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_single_spacing(p_ht0)
    r_ht0 = p_ht0.add_run("ตารางที่")
    set_thai_font(r_ht0, font_name=font_name, size_pt=BODY_FONT_SIZE, bold=True)

    p_ht1 = hdr_t[1].paragraphs[0]
    p_ht1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    apply_single_spacing(p_ht1)

    p_ht2 = hdr_t[2].paragraphs[0]
    p_ht2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_single_spacing(p_ht2)
    r_ht2 = p_ht2.add_run("หน้า")
    set_thai_font(r_ht2, font_name=font_name, size_pt=BODY_FONT_SIZE, bold=True)

    add_numbered_toc_row(
        toc_tbl_table,
        seq_num_str="1",
        item_title_text="การแสดงการคำนวณต้นทุนในการดำเนินการ",
        bookmark_name="bm_table_1",
        fallback_page_str="8",
        col1_width_cm=2.0,
        col2_width_cm=12.4,
        col3_width_cm=1.5,
        font_name=font_name
    )
    add_numbered_toc_row(
        toc_tbl_table,
        seq_num_str="2",
        item_title_text="การแสดงการคำนวณการจัดจำหน่าย",
        bookmark_name="bm_table_2",
        fallback_page_str="9",
        col1_width_cm=2.0,
        col2_width_cm=12.4,
        col3_width_cm=1.5,
        font_name=font_name
    )

    # 3. หน้าสารบัญภาพประกอบ
    doc.add_page_break()
    p_fig_title = doc.add_paragraph()
    p_fig_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_single_spacing(p_fig_title, space_before_pt=0, space_after_pt=12)
    r_fig_title = p_fig_title.add_run("สารบัญภาพประกอบ")
    set_thai_font(r_fig_title, font_name=font_name, size_pt=CHAPTER_TITLE_SIZE, bold=True)
    add_bookmark_to_paragraph(doc, p_fig_title, "bm_toc_figures")

    toc_fig_table = doc.add_table(rows=1, cols=3)
    toc_fig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    toc_fig_table.autofit = False
    remove_all_table_borders(toc_fig_table)
    set_row_repeat_header(toc_fig_table.rows[0])

    hdr_f = toc_fig_table.rows[0].cells
    hdr_f[0].width = Cm(2.8)
    hdr_f[1].width = Cm(11.6)
    hdr_f[2].width = Cm(1.5)
    set_cell_margins(hdr_f[0], top=0, bottom=60, left=0, right=0)
    set_cell_margins(hdr_f[1], top=0, bottom=60, left=80, right=80)
    set_cell_margins(hdr_f[2], top=0, bottom=60, left=0, right=0)

    p_hf0 = hdr_f[0].paragraphs[0]
    p_hf0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_single_spacing(p_hf0)
    r_hf0 = p_hf0.add_run("ภาพประกอบที่")
    set_thai_font(r_hf0, font_name=font_name, size_pt=BODY_FONT_SIZE, bold=True)

    p_hf1 = hdr_f[1].paragraphs[0]
    p_hf1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    apply_single_spacing(p_hf1)

    p_hf2 = hdr_f[2].paragraphs[0]
    p_hf2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_single_spacing(p_hf2)
    r_hf2 = p_hf2.add_run("หน้า")
    set_thai_font(r_hf2, font_name=font_name, size_pt=BODY_FONT_SIZE, bold=True)

    work_steps = step5_data.get("workSteps", [])
    valid_work_steps = [
        ws for ws in work_steps
        if isinstance(ws, dict) and (clean_single_line_text(ws.get("stepText")) or clean_single_line_text(ws.get("captionText")) or ws.get("imageUrl"))
    ]

    fig_entries = []
    for idx, ws in enumerate(valid_work_steps, 1):
        cap_txt = clean_single_line_text(ws.get("captionText", "")) or clean_single_line_text(ws.get("stepText", ""))
        cap_txt = re.sub(r"^ภาพประกอบที่\s*\d+\s*", "", cap_txt).strip()
        est_page = str(7 + ((idx - 1) // 2))
        fig_entries.append((str(idx), cap_txt, f"bm_fig_{idx}", est_page))

    step5_img_count = len(valid_work_steps)
    appendix_items = step7_data.get("appendixItems", [])
    valid_appendix = [
        a for a in appendix_items
        if isinstance(a, dict) and (clean_single_line_text(a.get("captionText")) or a.get("imageUrl"))
    ]

    for idx, app_item in enumerate(valid_appendix, 1):
        fig_num = app_item.get("figureNumber") or (step5_img_count + idx)
        cap_text = clean_single_line_text(app_item.get("captionText", ""))
        cap_text = re.sub(r"^ภาพประกอบที่\s*\d+\s*", "", cap_text).strip()
        est_page = str(14 + ((idx - 1) // 2))
        fig_entries.append((str(fig_num), cap_text, f"bm_app_fig_{idx}", est_page))

    if fig_entries:
        last_fb_pg = None
        for seq_str, cap_only, fig_bm, fig_fb in fig_entries:
            fb_to_show = "" if (not WORD_COM_AVAILABLE and fig_fb == last_fb_pg) else fig_fb
            last_fb_pg = fig_fb
            add_numbered_toc_row(
                toc_fig_table,
                seq_num_str=seq_str,
                item_title_text=cap_only,
                bookmark_name=fig_bm,
                fallback_page_str=fb_to_show,
                col1_width_cm=2.8,
                col2_width_cm=11.6,
                col3_width_cm=1.5,
                font_name=font_name
            )
    else:
        add_numbered_toc_row(
            toc_fig_table,
            seq_num_str="1",
            item_title_text="................................................................",
            bookmark_name=None,
            fallback_page_str="-",
            col1_width_cm=2.8,
            col2_width_cm=11.6,
            col3_width_cm=1.5,
            font_name=font_name
        )

# ==========================================================
# ส่วนที่ 5: ฟังก์ชันสร้างเอกสารแต่ละส่วน (หน้าปก ถึง ประวัติคณะผู้จัดทำ)
# ==========================================================
def build_cover_page(
    doc,
    group_data,
    font_name=REPORT_FONT_NAME,
    font_size=COVER_FONT_SIZE,
    logo_width_cm=3.2,
    name_col_width_cm=DEFAULT_COVER_NAME_COL_CM,
    num_col_width_cm=DEFAULT_COVER_NUM_COL_CM
):
    """สร้างหน้าที่ 1: หน้าปก (คอลัมน์เลขที่กว้าง 23 มิลลิเมตร = 2.3 ซม.)"""
    sec_cover = doc.sections[0]
    configure_section_page_numbering(sec_cover, fmt="decimal", start_num=None, show_on_first_page=False, show_on_other_pages=False, font_name=font_name)

    progress_steps = group_data.get("progress", {}).get("steps", {})
    step1_dict = progress_steps.get("1") or progress_steps.get(1) or {}
    step1_data = step1_dict.get("data", {})
    
    project_title = clean_single_line_text(
        step1_data.get("projectTitle")
        or group_data.get("projectTitle")
        or "...................................................."
    )

    members = group_data.get("members", [])
    first_room = members[0].get("room", "") if members else ""
    class_room_str = format_class_room_display(first_room)

    layout_table = doc.add_table(rows=3, cols=1)
    layout_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    layout_table.autofit = False
    remove_all_table_borders(layout_table)

    for row in layout_table.rows:
        row.cells[0].width = Cm(15.9)

    set_table_row_height(layout_table.rows[0], 7.6)
    set_table_row_height(layout_table.rows[1], 10.4)
    set_table_row_height(layout_table.rows[2], 6.4)

    cell_top = layout_table.rows[0].cells[0]
    cell_top.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    p_logo = cell_top.paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_single_spacing(p_logo, space_before_pt=0, space_after_pt=12)

    logo_bytes = load_image_bytes_from_source(SCHOOL_LOGO_URL, target_aspect_ratio=None)
    if logo_bytes:
        run_logo = p_logo.add_run()
        run_logo.add_picture(io.BytesIO(logo_bytes), width=Cm(logo_width_cm))
    else:
        run_fallback = p_logo.add_run("[ตราสัญลักษณ์โรงเรียนวาปีปทุม]")
        set_thai_font(run_fallback, font_name=font_name, size_pt=font_size, bold=True)

    add_paragraph_in_cell(cell_top, "โครงงานอาชีพ", align=WD_ALIGN_PARAGRAPH.CENTER, font_name=font_name, size_pt=font_size, bold=True)
    add_paragraph_in_cell(cell_top, f"เรื่อง {project_title}", align=WD_ALIGN_PARAGRAPH.CENTER, font_name=font_name, size_pt=font_size, bold=True)

    cell_mid = layout_table.rows[1].cells[0]
    cell_mid.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    p_maker_title = cell_mid.paragraphs[0]
    p_maker_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_single_spacing(p_maker_title)
    r_maker = p_maker_title.add_run("คณะผู้จัดทำ")
    set_thai_font(r_maker, font_name=font_name, size_pt=font_size, bold=True)

    member_table = cell_mid.add_table(rows=0, cols=2)
    member_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    member_table.autofit = False
    remove_all_table_borders(member_table)

    if members:
        for m in members:
            full_name = str(m.get("fullName", "-")).strip()
            number_str = str(m.get("number", "-")).strip()
            
            m_row = member_table.add_row()
            c_name, c_num = m_row.cells[0], m_row.cells[1]
            c_name.width = Cm(name_col_width_cm)
            c_num.width = Cm(num_col_width_cm)
            set_cell_margins(c_name, top=0, bottom=0, left=60, right=100)
            set_cell_margins(c_num, top=0, bottom=0, left=100, right=60)

            p_n = c_name.paragraphs[0]
            p_n.alignment = WD_ALIGN_PARAGRAPH.LEFT
            apply_single_spacing(p_n)
            r_n = p_n.add_run(full_name)
            set_thai_font(r_n, font_name=font_name, size_pt=font_size, bold=True)

            p_no = c_num.paragraphs[0]
            p_no.alignment = WD_ALIGN_PARAGRAPH.LEFT
            apply_single_spacing(p_no)
            r_no = p_no.add_run(f"เลขที่ {number_str}")
            set_thai_font(r_no, font_name=font_name, size_pt=font_size, bold=True)
    else:
        m_row = member_table.add_row()
        p_empty = m_row.cells[0].paragraphs[0]
        apply_single_spacing(p_empty)
        r_empty = p_empty.add_run("(ยังไม่มีรายชื่อสมาชิกในกลุ่ม)")
        set_thai_font(r_empty, font_name=font_name, size_pt=font_size, bold=True)

    add_paragraph_in_cell(
        cell_mid,
        f"ชั้นมัธยมศึกษาปีที่ {class_room_str}",
        align=WD_ALIGN_PARAGRAPH.CENTER,
        font_name=font_name,
        size_pt=font_size,
        bold=True
    )

    cell_bot = layout_table.rows[2].cells[0]
    cell_bot.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM

    footer_lines = [
        "โครงงานอาชีพฉบับนี้",
        "เป็นส่วนหนึ่งของรายวิชา การงานอาชีพเพิ่มเติม (โครงงานอาชีพ)",
        "รหัสวิชา ง30202",
        "โรงเรียนวาปีปทุม สำนักงานเขตพื้นที่การศึกษามัธยมศึกษามหาสารคาม",
        "สำนักงานคณะกรรมการการศึกษาขั้นพื้นฐาน กระทรวงศึกษาธิการ"
    ]

    for idx, line_text in enumerate(footer_lines):
        if idx == 0:
            p_f = cell_bot.paragraphs[0]
            p_f.alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_single_spacing(p_f)
            r_f = p_f.add_run(line_text)
            set_thai_font(r_f, font_name=font_name, size_pt=font_size, bold=True)
        else:
            add_paragraph_in_cell(
                cell_bot,
                line_text,
                align=WD_ALIGN_PARAGRAPH.CENTER,
                font_name=font_name,
                size_pt=font_size,
                bold=True
            )


def build_preface_page(doc, group_data, font_name=REPORT_FONT_NAME, semester=DEFAULT_SEMESTER, academic_year=DEFAULT_ACADEMIC_YEAR, first_line_indent_cm=1.5, strip_advisor_honorific=True):
    """สร้างหน้าคำนำ"""
    sec_front = doc.add_section(WD_SECTION.NEW_PAGE)
    configure_section_page_numbering(
        sec_front, fmt="thaiLetters", start_num=1, show_on_first_page=True, show_on_other_pages=True, font_name=font_name
    )
    doc._lines_needed_with_heading = 0

    progress_steps = group_data.get("progress", {}).get("steps", {})
    step1_data = (progress_steps.get("1") or progress_steps.get(1) or {}).get("data", {})
    step2_data = (progress_steps.get("2") or progress_steps.get(2) or {}).get("data", {})

    project_title = clean_single_line_text(
        step1_data.get("projectTitle")
        or group_data.get("projectTitle")
        or "...................................................."
    )

    objectives_list = step2_data.get("objectives", [])
    formatted_objectives = format_objectives_for_preface(objectives_list)

    advisors_list = step1_data.get("advisors", ["นายธนโชติ จันทรังษี"])
    formatted_advisors = format_advisors_for_preface(advisors_list, strip_honorific=strip_advisor_honorific)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_single_spacing(p_title, space_before_pt=0, space_after_pt=12)
    run_title = p_title.add_run("คำนำ")
    set_thai_font(run_title, font_name=font_name, size_pt=20, bold=True)
    add_bookmark_to_paragraph(doc, p_title, "bm_preface")

    raw_para1 = (
        f"เอกสารรายงานโครงงานอาชีพ เรื่อง {project_title} "
        f"ฉบับนี้เป็นส่วนหนึ่งของรายวิชา โครงงานอาชีพ รหัสวิชา ง30202 "
        f"ภาคเรียนที่ {semester} ปีการศึกษา {academic_year} "
        f"โดยมีวัตถุประสงค์{formatted_objectives} "
        f"จัดทำขึ้นเนื่องจากเล็งเห็นความสำคัญของทักษะอาชีพ การใช้เวลาให้เกิดประโยชน์ "
        f"ตลอดจนการพัฒนาต่อยอดภูมิปัญญาหรือความต้องการของตลาดในปัจจุบัน"
    )

    raw_para2 = (
        f"ในการจัดทำโครงงานครั้งนี้ คณะผู้จัดทำได้รับความช่วยเหลือ คำแนะนำ "
        f"และการสนับสนุนอย่างดียิ่งจาก ครูที่ปรึกษาโครงงาน {formatted_advisors} ผู้ปกครอง และผู้ที่เกี่ยวข้องทุกท่าน "
        f"ซึ่งมีส่วนทำให้การดำเนินงานสำเร็จลุล่วงไปได้ด้วยดี คณะผู้จัดทำขอขอบพระคุณมา ณ โอกาสนี้"
    )

    raw_para3 = (
        "คณะผู้จัดทำหวังเป็นอย่างยิ่งว่า รายงานโครงงานอาชีพฉบับนี้จะเป็นประโยชน์ต่อผู้ที่สนใจ "
        "ศึกษาหาข้อมูล หรือนำไปเป็นแนวทางในการประกอบอาชีพต่อไป "
        "หากมีข้อผิดพลาดประการใด คณะผู้จัดทำต้องขออภัยมา ณ ที่นี้ด้วย"
    )

    for raw_text in [raw_para1, raw_para2, raw_para3]:
        add_indented_body_paragraph(
            doc,
            raw_text,
            font_name=font_name,
            size_pt=16,
            indent_cm=first_line_indent_cm,
            use_thai_distribute=True
        )

    add_blank_line_16pt(doc, font_name=font_name)

    sig_table = doc.add_table(rows=2, cols=1)
    sig_table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    sig_table.autofit = False
    remove_all_table_borders(sig_table)

    sig_lines = [
        "คณะผู้จัดทำ",
        f"ปีการศึกษา {academic_year}"
    ]

    for row_idx, sig_text in enumerate(sig_lines):
        cell = sig_table.rows[row_idx].cells[0]
        cell.width = Cm(5.5)
        set_cell_margins(cell, top=0, bottom=0, left=60, right=0)
        p_sig = cell.paragraphs[0]
        p_sig.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_single_spacing(p_sig)
        r_sig = p_sig.add_run(sig_text)
        set_thai_font(r_sig, font_name=font_name, size_pt=16, bold=False)


def build_chapter1_page(doc, group_data, font_name=REPORT_FONT_NAME, first_line_indent_cm=1.5, sub_left_indent_cm=1.0, use_thai_distribute=True):
    """สร้างบทที่ 1 บทนำ"""
    sec_ch1 = doc.add_section(WD_SECTION.NEW_PAGE)
    configure_section_page_numbering(
        sec_ch1, fmt="decimal", start_num=1, show_on_first_page=False, show_on_other_pages=True, font_name=font_name
    )

    progress_steps = group_data.get("progress", {}).get("steps", {})
    step2_data = (progress_steps.get("2") or progress_steps.get(2) or {}).get("data", {})

    add_chapter_header(doc, "บทที่ 1", "บทนำ", font_name=font_name, bookmark_name="bm_ch1")

    add_main_section_heading(doc, "1. ที่มาและความสำคัญของโครงงาน", font_name=font_name, size_pt=SECTION_HEADING_SIZE, bold=True, bookmark_name="bm_ch1_1")

    bg_paras = step2_data.get("backgroundParagraphs", [])
    if not bg_paras or not isinstance(bg_paras, list):
        raw_bg = str(step2_data.get("background") or "").strip()
        bg_paras = [p.strip() for p in re.split(r"\n+", raw_bg) if p.strip()]
    else:
        bg_paras = [str(p).strip() for p in bg_paras if str(p).strip()]

    if not bg_paras:
        bg_paras = ["........................................................................................................................"]

    for para_text in bg_paras:
        add_indented_body_paragraph(
            doc,
            para_text,
            font_name=font_name,
            size_pt=BODY_FONT_SIZE,
            indent_cm=first_line_indent_cm,
            use_thai_distribute=use_thai_distribute
        )

    add_main_section_heading(doc, "2. วัตถุประสงค์ของโครงงาน", font_name=font_name, size_pt=SECTION_HEADING_SIZE, bold=True, bookmark_name="bm_ch1_2")

    raw_objectives = step2_data.get("objectives", [])
    clean_objectives = [strip_leading_number(o) for o in raw_objectives if strip_leading_number(o)]
    if not clean_objectives:
        clean_objectives = ["........................................................................................................................"]

    for idx, obj_text in enumerate(clean_objectives, 1):
        add_indented_body_paragraph(
            doc,
            f"{idx}. {obj_text}",
            font_name=font_name,
            size_pt=BODY_FONT_SIZE,
            indent_cm=first_line_indent_cm,
            use_thai_distribute=False
        )

    add_main_section_heading(doc, "3. ขอบเขตของโครงงาน", font_name=font_name, size_pt=SECTION_HEADING_SIZE, bold=True, bookmark_name="bm_ch1_3")

    scopes = [
        ("3.1 ขอบเขตด้านผลิตภัณฑ์", step2_data.get("scopeProduct", "")),
        ("3.2 ขอบเขตด้านการผลิต", step2_data.get("scopeProduction", "")),
        ("3.3 ขอบเขตด้านการตลาดและกลุ่มเป้าหมาย", step2_data.get("scopeMarket", "")),
        ("3.4 ขอบเขตด้านระยะเวลา", step2_data.get("scopeTime", ""))
    ]

    for sub_title, sub_content in scopes:
        add_subheading_indented(
            doc,
            sub_title,
            font_name=font_name,
            size_pt=BODY_FONT_SIZE,
            bold=True,
            left_indent_cm=sub_left_indent_cm
        )
        content_str = clean_single_line_text(sub_content) or "........................................................................................................................"
        add_indented_body_paragraph(
            doc,
            content_str,
            font_name=font_name,
            size_pt=BODY_FONT_SIZE,
            indent_cm=first_line_indent_cm,
            use_thai_distribute=use_thai_distribute
        )

    add_main_section_heading(doc, "4. ประโยชน์ที่คาดว่าจะได้รับ", font_name=font_name, size_pt=SECTION_HEADING_SIZE, bold=True, bookmark_name="bm_ch1_4")

    raw_benefits = step2_data.get("benefits", [])
    clean_benefits = [strip_leading_number(b) for b in raw_benefits if strip_leading_number(b)]
    if not clean_benefits:
        clean_benefits = ["........................................................................................................................"]

    for idx, ben_text in enumerate(clean_benefits, 1):
        add_indented_body_paragraph(
            doc,
            f"{idx}. {ben_text}",
            font_name=font_name,
            size_pt=BODY_FONT_SIZE,
            indent_cm=first_line_indent_cm,
            use_thai_distribute=False
        )


def build_chapter2_page(doc, group_data, font_name=REPORT_FONT_NAME, first_line_indent_cm=1.5, use_thai_distribute=True, use_sub_numbering=False):
    """สร้างบทที่ 2 เอกสารและงานที่เกี่ยวข้อง"""
    sec_ch2 = doc.add_section(WD_SECTION.NEW_PAGE)
    configure_section_page_numbering(
        sec_ch2, fmt="decimal", start_num=None, show_on_first_page=False, show_on_other_pages=True, font_name=font_name
    )

    progress_steps = group_data.get("progress", {}).get("steps", {})
    step1_data = (progress_steps.get("1") or progress_steps.get(1) or {}).get("data", {})
    step3_data = (progress_steps.get("3") or progress_steps.get(3) or {}).get("data", {})

    project_title = clean_single_line_text(
        step1_data.get("projectTitle")
        or group_data.get("projectTitle")
        or "...................................................."
    )

    add_chapter_header(doc, "บทที่ 2", "เอกสารและงานที่เกี่ยวข้อง", font_name=font_name, bookmark_name="bm_ch2")

    add_main_section_heading(doc, "1. ทฤษฎีหรือแนวคิดที่เกี่ยวข้องกับโครงงาน", font_name=font_name, size_pt=SECTION_HEADING_SIZE, bold=True, bookmark_name="bm_ch2_1")

    intro_text = (
        f"โครงงานอาชีพ เรื่อง {project_title} "
        f"คณะผู้จัดทำได้ศึกษาทฤษฎีหรือแนวคิดต่างๆ ที่เกี่ยวข้องกับโครงงานอาชีพ "
        f"โดยสามารถสรุปสาระสำคัญตามลำดับ ดังนี้"
    )
    add_indented_body_paragraph(
        doc,
        intro_text,
        font_name=font_name,
        size_pt=BODY_FONT_SIZE,
        indent_cm=first_line_indent_cm,
        use_thai_distribute=use_thai_distribute
    )

    theories_list = step3_data.get("theories", [])
    valid_theories = [
        th for th in theories_list
        if isinstance(th, dict) and (
            clean_single_line_text(th.get("topic"))
            or clean_single_line_text(th.get("summary"))
            or any(clean_single_line_text(p.get("text")) for p in th.get("paragraphs", []) if isinstance(p, dict))
        )
    ]

    if valid_theories:
        for idx, th in enumerate(valid_theories, 1):
            topic_name = strip_leading_number(th.get("topic", "")) or "...................................................."
            add_indented_body_paragraph(
                doc,
                f"{idx}. {topic_name}",
                font_name=font_name,
                size_pt=BODY_FONT_SIZE,
                indent_cm=first_line_indent_cm,
                use_thai_distribute=False
            )
    else:
        add_indented_body_paragraph(
            doc,
            "1. ........................................................................................................................",
            font_name=font_name,
            size_pt=BODY_FONT_SIZE,
            indent_cm=first_line_indent_cm,
            use_thai_distribute=False
        )

    for idx, th in enumerate(valid_theories, 1):
        topic_name = strip_leading_number(th.get("topic", "")) or f"สาระสำคัญที่ {idx}"
        prefix_num = f"1.{idx}" if use_sub_numbering else f"{idx}."
        
        add_sub_heading_left(
            doc,
            f"{prefix_num} {topic_name}",
            font_name=font_name,
            size_pt=BODY_FONT_SIZE,
            bold=True
        )

        paragraphs = th.get("paragraphs", [])
        for p_obj in paragraphs:
            if not isinstance(p_obj, dict):
                continue
            body_txt = clean_single_line_text(p_obj.get("text", ""))
            if not body_txt:
                continue
            citation_prefix = format_theory_citation_prefix(p_obj)
            full_para_text = f"{citation_prefix}{body_txt}"
            add_indented_body_paragraph(
                doc,
                full_para_text,
                font_name=font_name,
                size_pt=BODY_FONT_SIZE,
                indent_cm=first_line_indent_cm,
                use_thai_distribute=use_thai_distribute
            )

        summary_txt = clean_single_line_text(th.get("summary", ""))
        if summary_txt:
            add_indented_body_paragraph(
                doc,
                summary_txt,
                font_name=font_name,
                size_pt=BODY_FONT_SIZE,
                indent_cm=first_line_indent_cm,
                use_thai_distribute=use_thai_distribute
            )

    add_main_section_heading(doc, "2. งานวิจัยและโครงงานที่เกี่ยวข้อง", font_name=font_name, size_pt=SECTION_HEADING_SIZE, bold=True, bookmark_name="bm_ch2_2")

    researches_list = step3_data.get("researches", [])
    valid_researches = [
        r for r in researches_list
        if isinstance(r, dict) and clean_single_line_text(r.get("content", ""))
    ]

    if valid_researches:
        for r_obj in valid_researches:
            res_prefix = format_research_citation_prefix(r_obj)
            res_content = clean_single_line_text(r_obj.get("content", ""))
            full_res_text = f"{res_prefix}{res_content}"
            add_indented_body_paragraph(
                doc,
                full_res_text,
                font_name=font_name,
                size_pt=BODY_FONT_SIZE,
                indent_cm=first_line_indent_cm,
                use_thai_distribute=use_thai_distribute
            )
    else:
        add_indented_body_paragraph(
            doc,
            "........................................................................................................................",
            font_name=font_name,
            size_pt=BODY_FONT_SIZE,
            indent_cm=first_line_indent_cm,
            use_thai_distribute=False
        )

    add_main_section_heading(doc, "3. สรุป", font_name=font_name, size_pt=SECTION_HEADING_SIZE, bold=True, bookmark_name="bm_ch2_3")

    overall_summary = clean_single_line_text(step3_data.get("overallSummary", ""))
    if not overall_summary:
        overall_summary = "........................................................................................................................"

    add_indented_body_paragraph(
        doc,
        overall_summary,
        font_name=font_name,
        size_pt=BODY_FONT_SIZE,
        indent_cm=first_line_indent_cm,
        use_thai_distribute=use_thai_distribute
    )


def build_chapter3_page(doc, group_data, font_name=REPORT_FONT_NAME, first_line_indent_cm=1.5, sub_left_indent_cm=1.0, use_thai_distribute=True):
    """สร้างบทที่ 3 ขั้นตอนการดำเนินโครงงาน"""
    sec_ch3 = doc.add_section(WD_SECTION.NEW_PAGE)
    configure_section_page_numbering(
        sec_ch3, fmt="decimal", start_num=None, show_on_first_page=False, show_on_other_pages=True, font_name=font_name
    )

    progress_steps = group_data.get("progress", {}).get("steps", {})
    step1_data = (progress_steps.get("1") or progress_steps.get(1) or {}).get("data", {})
    step4_data = (progress_steps.get("4") or progress_steps.get(4) or {}).get("data", {})

    project_title = clean_single_line_text(
        step1_data.get("projectTitle")
        or group_data.get("projectTitle")
        or "...................................................."
    )

    add_chapter_header(doc, "บทที่ 3", "ขั้นตอนการดำเนินโครงงาน", font_name=font_name, bookmark_name="bm_ch3")

    add_main_section_heading(doc, "1. วิธีการดำเนินงาน", font_name=font_name, size_pt=SECTION_HEADING_SIZE, bold=True, bookmark_name="bm_ch3_1")

    intro_ch3 = f"คณะผู้จัดทำ โครงงานอาชีพ เรื่อง {project_title} ได้มีกระบวนการดำเนินโครงงาน ดังต่อไปนี้"
    add_indented_body_paragraph(
        doc,
        intro_ch3,
        font_name=font_name,
        size_pt=BODY_FONT_SIZE,
        indent_cm=first_line_indent_cm,
        use_thai_distribute=use_thai_distribute
    )

    methods = [strip_leading_number(m) for m in step4_data.get("operationMethods", []) if strip_leading_number(m)]
    if not methods:
        methods = ["........................................................................................................................"]

    for idx, m_txt in enumerate(methods, 1):
        full_m_line = f"{idx}. {m_txt}"
        is_multi_line = estimate_paragraph_lines(full_m_line, first_line_indent_cm=first_line_indent_cm) > 1
        add_indented_body_paragraph(
            doc,
            full_m_line,
            font_name=font_name,
            size_pt=BODY_FONT_SIZE,
            indent_cm=first_line_indent_cm,
            use_thai_distribute=(use_thai_distribute and is_multi_line)
        )

    mat_type = clean_single_line_text(step4_data.get("materialType", "วัตถุดิบ")) or "วัตถุดิบ"
    add_main_section_heading(doc, f"2. {mat_type}และอุปกรณ์", font_name=font_name, size_pt=SECTION_HEADING_SIZE, bold=True, bookmark_name="bm_ch3_2")

    add_subheading_indented(doc, f"2.1 {mat_type}", font_name=font_name, size_pt=BODY_FONT_SIZE, bold=True, left_indent_cm=sub_left_indent_cm)
    materials = step4_data.get("materials", [])
    valid_mats = [m for m in materials if isinstance(m, dict) and clean_single_line_text(m.get("name"))]

    if valid_mats:
        for idx, m in enumerate(valid_mats, 1):
            m_name = strip_leading_number(m.get("name", ""))
            m_qty = clean_single_line_text(m.get("qty", ""))
            m_unit = clean_single_line_text(m.get("unit", ""))
            qty_part = f" จำนวน {m_qty} {m_unit}".rstrip() if m_qty else ""
            add_indented_body_paragraph(
                doc,
                f"2.1.{idx} {m_name}{qty_part}",
                font_name=font_name,
                size_pt=BODY_FONT_SIZE,
                indent_cm=first_line_indent_cm,
                use_thai_distribute=False
            )
    else:
        add_indented_body_paragraph(
            doc,
            "2.1.1 ........................................................................................................................",
            font_name=font_name,
            size_pt=BODY_FONT_SIZE,
            indent_cm=first_line_indent_cm,
            use_thai_distribute=False
        )

    add_subheading_indented(doc, "2.2 เครื่องมือและอุปกรณ์", font_name=font_name, size_pt=BODY_FONT_SIZE, bold=True, left_indent_cm=sub_left_indent_cm)
    equipments = [strip_leading_number(e) for e in step4_data.get("equipments", []) if strip_leading_number(e)]
    if not equipments:
        equipments = ["........................................................................................................................"]

    for idx, eq_txt in enumerate(equipments, 1):
        add_indented_body_paragraph(
            doc,
            f"2.2.{idx} {eq_txt}",
            font_name=font_name,
            size_pt=BODY_FONT_SIZE,
            indent_cm=first_line_indent_cm,
            use_thai_distribute=False
        )

    add_main_section_heading(doc, "3. ขั้นตอนการทำ", font_name=font_name, size_pt=SECTION_HEADING_SIZE, bold=True, bookmark_name="bm_ch3_3")
    making_steps = [strip_leading_number(s) for s in step4_data.get("makingSteps", []) if strip_leading_number(s)]
    if not making_steps:
        making_steps = ["........................................................................................................................"]

    for idx, st_txt in enumerate(making_steps, 1):
        full_st_line = f"3.{idx} {st_txt}"
        is_multi_line = estimate_paragraph_lines(full_st_line, first_line_indent_cm=first_line_indent_cm) > 1
        add_indented_body_paragraph(
            doc,
            full_st_line,
            font_name=font_name,
            size_pt=BODY_FONT_SIZE,
            indent_cm=first_line_indent_cm,
            use_thai_distribute=(use_thai_distribute and is_multi_line)
        )


def build_chapter4_page(doc, group_data, font_name=REPORT_FONT_NAME, first_line_indent_cm=1.5, use_thai_distribute=True, max_table_rows_per_page=18):
    """สร้างบทที่ 4 การอภิปรายผลโครงงาน"""
    sec_ch4 = doc.add_section(WD_SECTION.NEW_PAGE)
    configure_section_page_numbering(
        sec_ch4, fmt="decimal", start_num=None, show_on_first_page=False, show_on_other_pages=True, font_name=font_name
    )

    usable_h_cm = 24.2
    current_y_cm = 0.0

    progress_steps = group_data.get("progress", {}).get("steps", {})
    step1_data = (progress_steps.get("1") or progress_steps.get(1) or {}).get("data", {})
    step4_data = (progress_steps.get("4") or progress_steps.get(4) or {}).get("data", {})
    step5_data = (progress_steps.get("5") or progress_steps.get(5) or {}).get("data", {})

    project_title = clean_single_line_text(
        step1_data.get("projectTitle")
        or group_data.get("projectTitle")
        or "...................................................."
    )
    mat_type = clean_single_line_text(step4_data.get("materialType", "วัตถุดิบ")) or "วัตถุดิบ"

    add_chapter_header(doc, "บทที่ 4", "การอภิปรายผลโครงงาน", font_name=font_name, bookmark_name="bm_ch4")
    current_y_cm += 2.2

    intro_ch4 = f"ในการดำเนินการตามโครงงานอาชีพ เรื่อง {project_title} คณะผู้จัดทำขออภิปรายผลโครงงานตามรายละเอียด ดังต่อไปนี้"
    add_indented_body_paragraph(
        doc,
        intro_ch4,
        font_name=font_name,
        size_pt=BODY_FONT_SIZE,
        indent_cm=first_line_indent_cm,
        use_thai_distribute=use_thai_distribute
    )
    current_y_cm += 1.3

    add_main_section_heading(doc, "1. ขั้นตอนการดำเนินงานตามขั้นตอนการทำงาน", font_name=font_name, size_pt=SECTION_HEADING_SIZE, bold=True, bookmark_name="bm_ch4_1")
    current_y_cm += 1.0

    work_steps = step5_data.get("workSteps", [])
    valid_work_steps = [
        ws for ws in work_steps
        if isinstance(ws, dict) and (clean_single_line_text(ws.get("stepText")) or clean_single_line_text(ws.get("captionText")) or ws.get("imageUrl"))
    ]

    if valid_work_steps:
        for idx, ws in enumerate(valid_work_steps, 1):
            step_desc = strip_leading_number(ws.get("stepText", ""))
            if step_desc:
                full_ws_line = f"1.{idx} {step_desc}"
                is_multi_line = estimate_paragraph_lines(full_ws_line, first_line_indent_cm=first_line_indent_cm) > 1
                p_step = add_indented_body_paragraph(
                    doc,
                    full_ws_line,
                    font_name=font_name,
                    size_pt=BODY_FONT_SIZE,
                    indent_cm=first_line_indent_cm,
                    use_thai_distribute=(use_thai_distribute and is_multi_line)
                )
                p_step.paragraph_format.keep_with_next = True
                p_step.paragraph_format.keep_together = True
            cap_txt = clean_single_line_text(ws.get("captionText", ""))
            full_cap = f"ภาพประกอบที่ {idx} {cap_txt}".strip()
            add_landscape_image_with_caption(
                doc,
                ws.get("imageUrl", ""),
                full_cap,
                font_name=font_name,
                bookmark_name=f"bm_fig_{idx}"
            )
        num_imgs = len(valid_work_steps)
        if num_imgs % 2 == 1:
            current_y_cm = (4.5 + 9.2) if num_imgs == 1 else 9.2
        else:
            current_y_cm = 22.5
    else:
        add_indented_body_paragraph(
            doc,
            "........................................................................................................................",
            font_name=font_name,
            size_pt=BODY_FONT_SIZE,
            indent_cm=first_line_indent_cm,
            use_thai_distribute=False
        )
        current_y_cm += 0.8

    p_cost_heading = add_main_section_heading(doc, "2. การคำนวณต้นทุน", font_name=font_name, size_pt=SECTION_HEADING_SIZE, bold=True, bookmark_name="bm_ch4_2")
    if current_y_cm + 5.5 > usable_h_cm:
        p_cost_heading.paragraph_format.page_break_before = True
        current_y_cm = 0.0
    current_y_cm += 1.0

    cost_intro_text = (
        f"คณะผู้จัดทำ โครงงานอาชีพ เรื่อง {project_title} "
        f"ได้มีการใช้{mat_type}ในการดำเนินโครงงาน โดยมีรายละเอียดการคำนวณต้นทุน "
        f"ดังแสดงในตารางที่ 1 ดังต่อไปนี้"
    )
    p_cost_intro = add_indented_body_paragraph(
        doc,
        cost_intro_text,
        font_name=font_name,
        size_pt=BODY_FONT_SIZE,
        indent_cm=first_line_indent_cm,
        use_thai_distribute=use_thai_distribute
    )
    p_cost_intro.paragraph_format.keep_with_next = True
    current_y_cm += 1.4

    cost_col_widths = [5.5, 3.5, 3.5, 3.5]
    cost_col_aligns = [
        WD_ALIGN_PARAGRAPH.LEFT,
        WD_ALIGN_PARAGRAPH.RIGHT,
        WD_ALIGN_PARAGRAPH.RIGHT,
        WD_ALIGN_PARAGRAPH.RIGHT
    ]
    cost_headers = [
        f"รายละเอียด{mat_type}",
        "ปริมาณที่ใช้ไป",
        "ราคาต่อหน่วยที่ซื้อ",
        "ราคาต่อหน่วยที่ใช้ (บาท)"
    ]

    cost_rows_data = []
    cost_items = step5_data.get("costItems", [])
    for idx, c_item in enumerate(cost_items, 1):
        if not isinstance(c_item, dict):
            continue
        c_name = strip_leading_number(c_item.get("name", "-"))
        used_qty_str = format_calc_number(c_item.get("usedQty", ""))
        used_unit_str = clean_single_line_text(c_item.get("usedUnit", ""))
        c_used = f"{used_qty_str} {used_unit_str}".strip()

        u_price_str = format_calc_number(c_item.get("unitPrice", ""))
        p_unit_str = clean_single_line_text(c_item.get("priceUnit", ""))
        c_price = f"{u_price_str} บาท/{p_unit_str}".strip() if p_unit_str else f"{u_price_str} บาท"

        c_calc = format_calc_number(c_item.get("calculatedCost", 0))
        cost_rows_data.append({
            "cells": [f"{idx}. {c_name}", c_used, c_price, c_calc],
            "bold": False
        })

    cost_rows_data.append({
        "cells": ["รวมเงินทั้งสิ้น", "", "", f"{format_calc_number(step5_data.get('totalCost', 0))} บาท"],
        "bold": True,
        "merge_first_n": 3,
        "aligns": [
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT
        ]
    })

    current_y_cm = render_paginated_table(
        doc=doc,
        base_title="ตารางที่ 1 การแสดงการคำนวณต้นทุนในการดำเนินการ",
        headers=cost_headers,
        col_widths=cost_col_widths,
        col_alignments=cost_col_aligns,
        rows_data=cost_rows_data,
        current_y_cm=current_y_cm,
        usable_page_height_cm=usable_h_cm,
        font_name=font_name,
        max_rows_per_page=max_table_rows_per_page,
        bookmark_name="bm_table_1"
    )

    p_sales_heading = add_main_section_heading(doc, "3. การคำนวณการจัดจำหน่าย", font_name=font_name, size_pt=SECTION_HEADING_SIZE, bold=True, bookmark_name="bm_ch4_3")
    if current_y_cm + 5.5 > usable_h_cm:
        p_sales_heading.paragraph_format.page_break_before = True
        current_y_cm = 0.0
    current_y_cm += 1.0

    sales_intro_text = (
        f"คณะผู้จัดทำ โครงงานอาชีพ เรื่อง {project_title} "
        f"คณะผู้จัดทำ ได้คำนวณการจัดจำหน่าย รายละเอียดต้นทุน ราคาจัดจำหน่าย และกำไร "
        f"ดังแสดงในตารางที่ 2 ดังต่อไปนี้"
    )
    p_sales_intro = add_indented_body_paragraph(
        doc,
        sales_intro_text,
        font_name=font_name,
        size_pt=BODY_FONT_SIZE,
        indent_cm=first_line_indent_cm,
        use_thai_distribute=use_thai_distribute
    )
    p_sales_intro.paragraph_format.keep_with_next = True
    current_y_cm += 1.4

    sales_headers = ["รายการ", "รายละเอียด"]
    sales_col_widths = [11.0, 5.0]
    sales_col_aligns = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT]

    raw_sales_pairs = [
        ("ราคาต้นทุน", f"{format_calc_number(step5_data.get('totalCost', 0))} บาท"),
        ("จำนวนที่ได้จากการจัดทำ", f"{format_calc_number(step5_data.get('unitsProduced', '-'))} ชิ้น/กล่อง"),
        ("ราคาต่อชิ้น", f"{format_calc_number(step5_data.get('costPerUnit', 0))} บาท/ชิ้น"),
        ("ราคาที่จัดจำหน่าย", f"{format_calc_number(step5_data.get('sellingPrice', '-'))} บาท/ชิ้น"),
        ("กำไรต่อชิ้น", f"{format_calc_number(step5_data.get('profitPerUnit', 0))} บาท/ชิ้น"),
        ("กำไรที่ได้จากการจัดทำในครั้งนี้", f"{format_calc_number(step5_data.get('totalProfit', 0))} บาท")
    ]

    sales_rows_data = []
    for s_idx, (s_lbl, s_val) in enumerate(raw_sales_pairs):
        sales_rows_data.append({
            "cells": [s_lbl, s_val],
            "bold": (s_idx == len(raw_sales_pairs) - 1)
        })

    current_y_cm = render_paginated_table(
        doc=doc,
        base_title="ตารางที่ 2 การแสดงการคำนวณการจัดจำหน่าย",
        headers=sales_headers,
        col_widths=sales_col_widths,
        col_alignments=sales_col_aligns,
        rows_data=sales_rows_data,
        current_y_cm=current_y_cm,
        usable_page_height_cm=usable_h_cm,
        font_name=font_name,
        max_rows_per_page=max_table_rows_per_page,
        bookmark_name="bm_table_2"
    )


def build_chapter5_page(doc, group_data, font_name=REPORT_FONT_NAME, first_line_indent_cm=1.5, use_thai_distribute=True):
    """สร้างบทที่ 5 สรุปผลโครงงาน"""
    sec_ch5 = doc.add_section(WD_SECTION.NEW_PAGE)
    configure_section_page_numbering(
        sec_ch5, fmt="decimal", start_num=None, show_on_first_page=False, show_on_other_pages=True, font_name=font_name
    )

    progress_steps = group_data.get("progress", {}).get("steps", {})
    step1_data = (progress_steps.get("1") or progress_steps.get(1) or {}).get("data", {})
    step6_data = (progress_steps.get("6") or progress_steps.get(6) or {}).get("data", {})

    project_title = clean_single_line_text(
        step1_data.get("projectTitle")
        or group_data.get("projectTitle")
        or "...................................................."
    )

    add_chapter_header(doc, "บทที่ 5", "สรุปผลโครงงาน", font_name=font_name, bookmark_name="bm_ch5")

    intro_ch5 = f"คณะผู้จัดทำ โครงงานอาชีพ เรื่อง {project_title} ได้สรุปผลการดำเนินโครงงาน ได้ดังนี้"
    add_indented_body_paragraph(
        doc,
        intro_ch5,
        font_name=font_name,
        size_pt=BODY_FONT_SIZE,
        indent_cm=first_line_indent_cm,
        use_thai_distribute=use_thai_distribute
    )

    add_main_section_heading(doc, "1. ผลการดำเนินงานตามวัตถุประสงค์", font_name=font_name, size_pt=SECTION_HEADING_SIZE, bold=True, bookmark_name="bm_ch5_1")
    sub_intro_1 = f"คณะผู้จัดทำโครงงานอาชีพ เรื่อง {project_title} มีผลการดำเนินงานตามวัตถุประสงค์ ดังต่อไปนี้"
    add_indented_body_paragraph(doc, sub_intro_1, font_name=font_name, size_pt=BODY_FONT_SIZE, indent_cm=first_line_indent_cm, use_thai_distribute=use_thai_distribute)

    obj_results = step6_data.get("objectiveResults", [])
    if obj_results:
        for idx, item in enumerate(obj_results, 1):
            if not isinstance(item, dict):
                continue
            obj_txt = strip_leading_number(item.get("objectiveText", ""))
            res_para = clean_single_line_text(item.get("resultParagraph", "")) or "........................................................................................................................"
            add_indented_body_paragraph(doc, f"1.{idx} {obj_txt}", font_name=font_name, size_pt=BODY_FONT_SIZE, indent_cm=first_line_indent_cm, use_thai_distribute=False, bold=True)
            add_indented_body_paragraph(doc, res_para, font_name=font_name, size_pt=BODY_FONT_SIZE, indent_cm=first_line_indent_cm, use_thai_distribute=use_thai_distribute)
    else:
        add_indented_body_paragraph(doc, "........................................................................................................................", font_name=font_name, size_pt=BODY_FONT_SIZE, indent_cm=first_line_indent_cm, use_thai_distribute=False)

    add_main_section_heading(doc, "2. ปัญหาและอุปสรรค", font_name=font_name, size_pt=SECTION_HEADING_SIZE, bold=True, bookmark_name="bm_ch5_2")
    sub_intro_2 = f"คณะผู้จัดทำโครงงานอาชีพ เรื่อง {project_title} ในระหว่างดำเนินโครงงาน คณะผู้จัดทำได้พบปัญหาและอุปสรรคต่อการดำเนินงาน ดังต่อไปนี้"
    add_indented_body_paragraph(doc, sub_intro_2, font_name=font_name, size_pt=BODY_FONT_SIZE, indent_cm=first_line_indent_cm, use_thai_distribute=use_thai_distribute)

    problems = [strip_leading_number(p) for p in step6_data.get("problems", []) if strip_leading_number(p)]
    if not problems:
        problems = ["........................................................................................................................"]
    for idx, p_txt in enumerate(problems, 1):
        full_prob_line = f"2.{idx} {p_txt}"
        is_multi_line = estimate_paragraph_lines(full_prob_line, first_line_indent_cm=first_line_indent_cm) > 1
        add_indented_body_paragraph(
            doc,
            full_prob_line,
            font_name=font_name,
            size_pt=BODY_FONT_SIZE,
            indent_cm=first_line_indent_cm,
            use_thai_distribute=(use_thai_distribute and is_multi_line)
        )

    add_main_section_heading(doc, "3. ข้อเสนอแนะ", font_name=font_name, size_pt=SECTION_HEADING_SIZE, bold=True, bookmark_name="bm_ch5_3")
    sub_intro_3 = f"คณะผู้จัดทำโครงงานอาชีพ เรื่อง {project_title} มีข้อเสนอแนะในการพัฒนา ปรับปรุง และแก้ไขจากการดำเนินโครงงานในครั้งนี้ โดยมีข้อเสนอแนะดังต่อไปนี้"
    add_indented_body_paragraph(doc, sub_intro_3, font_name=font_name, size_pt=BODY_FONT_SIZE, indent_cm=first_line_indent_cm, use_thai_distribute=use_thai_distribute)

    suggestions = [strip_leading_number(s) for s in step6_data.get("suggestions", []) if strip_leading_number(s)]
    if not suggestions:
        suggestions = ["........................................................................................................................"]
    for idx, s_txt in enumerate(suggestions, 1):
        full_sug_line = f"3.{idx} {s_txt}"
        is_multi_line = estimate_paragraph_lines(full_sug_line, first_line_indent_cm=first_line_indent_cm) > 1
        add_indented_body_paragraph(
            doc,
            full_sug_line,
            font_name=font_name,
            size_pt=BODY_FONT_SIZE,
            indent_cm=first_line_indent_cm,
            use_thai_distribute=(use_thai_distribute and is_multi_line)
        )


def add_apa7_bibliography_paragraph(doc, html_bib_str, font_name=REPORT_FONT_NAME):
    """สร้างย่อหน้าบรรณานุกรม APA 7 พร้อมแปลงแท็ก <i>...</i> เป็นตัวเอียงใน Word"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    apply_single_spacing(p)
    p.paragraph_format.left_indent = Cm(1.5)
    p.paragraph_format.first_line_indent = Cm(-1.5)

    tokens = re.split(r"(<i>.*?</i>)", str(html_bib_str))
    for tok in tokens:
        if not tok:
            continue
        if tok.startswith("<i>") and tok.endswith("</i>"):
            clean_italic = tok[3:-4]
            r = p.add_run(clean_italic)
            set_thai_font(r, font_name=font_name, size_pt=BODY_FONT_SIZE, bold=False, italic=True)
        else:
            r = p.add_run(tok)
            set_thai_font(r, font_name=font_name, size_pt=BODY_FONT_SIZE, bold=False, italic=False)


def add_appendix_cover_page(doc, font_name=REPORT_FONT_NAME):
    """สร้างหน้าคั่นคำว่า 'ภาคผนวก' กึ่งกลางกระดาษ ขนาด 20 หนา (ไม่แสดงเลขหน้าในหน้าคั่นนี้)"""
    sec_app = doc.add_section(WD_SECTION.NEW_PAGE)
    configure_section_page_numbering(
        sec_app, fmt="decimal", start_num=None, show_on_first_page=False, show_on_other_pages=True, font_name=font_name
    )
    doc._lines_needed_with_heading = 0

    cover_tbl = doc.add_table(rows=1, cols=1)
    cover_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cover_tbl.autofit = False
    remove_all_table_borders(cover_tbl)

    cover_tbl.rows[0].cells[0].width = Cm(15.9)
    set_table_row_height(cover_tbl.rows[0], 21.0)

    cell = cover_tbl.rows[0].cells[0]
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_single_spacing(p, space_before_pt=0, space_after_pt=0)
    run = p.add_run("ภาคผนวก")
    set_thai_font(run, font_name=font_name, size_pt=CHAPTER_TITLE_SIZE, bold=True)
    add_bookmark_to_paragraph(doc, p, "bm_appendix")


def build_bibliography_and_appendix_pages(doc, group_data, font_name=REPORT_FONT_NAME):
    """สร้างหน้าบรรณานุกรม -> หน้าคั่นคำว่า 'ภาคผนวก' -> หน้ารูปภาพภาคผนวก"""
    progress_steps = group_data.get("progress", {}).get("steps", {})
    step5_data = (progress_steps.get("5") or progress_steps.get(5) or {}).get("data", {})
    step7_data = (progress_steps.get("7") or progress_steps.get(7) or {}).get("data", {})

    sec_bib = doc.add_section(WD_SECTION.NEW_PAGE)
    configure_section_page_numbering(
        sec_bib, fmt="decimal", start_num=None, show_on_first_page=True, show_on_other_pages=True, font_name=font_name
    )
    doc._lines_needed_with_heading = 0

    p_bib_title = doc.add_paragraph()
    p_bib_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_single_spacing(p_bib_title, space_before_pt=0, space_after_pt=12)
    p_bib_title.paragraph_format.keep_with_next = True
    r_bib_title = p_bib_title.add_run("บรรณานุกรม")
    set_thai_font(r_bib_title, font_name=font_name, size_pt=CHAPTER_TITLE_SIZE, bold=True)
    add_bookmark_to_paragraph(doc, p_bib_title, "bm_bibliography")

    bib_html_list = step7_data.get("bibliographyHtmlList", [])
    bib_plain_list = step7_data.get("bibliographyPlainList", [])
    source_bibs = bib_html_list if bib_html_list else bib_plain_list

    if source_bibs:
        for bib_item in source_bibs:
            add_apa7_bibliography_paragraph(doc, bib_item, font_name=font_name)
    else:
        add_indented_body_paragraph(doc, "........................................................................................................................", font_name=font_name, size_pt=BODY_FONT_SIZE, indent_cm=0, use_thai_distribute=False)

    add_appendix_cover_page(doc, font_name=font_name)

    work_steps = step5_data.get("workSteps", [])
    valid_work_steps = [
        ws for ws in work_steps
        if isinstance(ws, dict) and (clean_single_line_text(ws.get("stepText")) or clean_single_line_text(ws.get("captionText")) or ws.get("imageUrl"))
    ]
    step5_img_count = len(valid_work_steps)

    appendix_items = step7_data.get("appendixItems", [])
    valid_appendix = [
        a for a in appendix_items
        if isinstance(a, dict) and (clean_single_line_text(a.get("captionText")) or a.get("imageUrl"))
    ]

    if valid_appendix:
        for idx, app_item in enumerate(valid_appendix, 1):
            fig_num = app_item.get("figureNumber") or (step5_img_count + idx)
            cap_text = clean_single_line_text(app_item.get("captionText", ""))
            full_cap = f"ภาพประกอบที่ {fig_num} {cap_text}".strip()
            add_landscape_image_with_caption(
                doc,
                app_item.get("imageUrl", ""),
                full_cap,
                font_name=font_name,
                bookmark_name=f"bm_app_fig_{idx}",
                page_break_before=(idx == 1)
            )
    else:
        p_empty_app = add_indented_body_paragraph(
            doc, "........................................................................................................................",
            font_name=font_name, size_pt=BODY_FONT_SIZE, indent_cm=0, use_thai_distribute=False
        )
        p_empty_app.paragraph_format.page_break_before = True


def add_bio_hanging_row(doc, label_text, value_text, font_name=REPORT_FONT_NAME):
    """จัดข้อมูลในหน้าประวัติคณะผู้จัดทำ (หัวข้อ 30 มม. 16 หนา | ข้อมูล 80 มม. กั้นหน้าลอย 16 บาง)"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    apply_single_spacing(p, space_before_pt=0, space_after_pt=0)

    p.paragraph_format.left_indent = Mm(BIO_DATA_INDENT_MM)
    p.paragraph_format.first_line_indent = Mm(BIO_LABEL_INDENT_MM - BIO_DATA_INDENT_MM)
    p.paragraph_format.tab_stops.add_tab_stop(Mm(BIO_DATA_INDENT_MM))

    r_lbl = p.add_run(clean_single_line_text(label_text))
    set_thai_font(r_lbl, font_name=font_name, size_pt=BODY_FONT_SIZE, bold=True)

    r_val = p.add_run("\t" + clean_single_line_text(value_text))
    set_thai_font(r_val, font_name=font_name, size_pt=BODY_FONT_SIZE, bold=False)
    return p


def add_bio_continuation_line(doc, value_text, font_name=REPORT_FONT_NAME):
    """สร้างบรรทัดข้อมูลต่อเนื่องในหน้าประวัติคณะผู้จัดทำ ที่ย่อเข้ามาตรงแนวกั้นหน้าลอย 80 มิลลิเมตร (ขนาด 16 บาง)"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    apply_single_spacing(p, space_before_pt=0, space_after_pt=0)
    p.paragraph_format.left_indent = Mm(BIO_DATA_INDENT_MM)
    p.paragraph_format.first_line_indent = Mm(0)

    r_val = p.add_run(clean_single_line_text(value_text))
    set_thai_font(r_val, font_name=font_name, size_pt=BODY_FONT_SIZE, bold=False)
    return p


def build_biography_page(doc, group_data, font_name=REPORT_FONT_NAME, middle_edu_label="ระดับมัธยมศึกษาตอนต้น"):
    """สร้างหน้าประวัติคณะผู้จัดทำ (แสดงเลขหน้าอารบิกต่อเนื่องจนถึงคนสุดท้าย และแยกคนละ 1 หน้า)"""
    sec_bio = doc.add_section(WD_SECTION.NEW_PAGE)
    configure_section_page_numbering(
        sec_bio, fmt="decimal", start_num=None, show_on_first_page=True, show_on_other_pages=True, font_name=font_name
    )
    doc._lines_needed_with_heading = 0

    progress_steps = group_data.get("progress", {}).get("steps", {})
    step8_data = (progress_steps.get("8") or progress_steps.get(8) or {}).get("data", {})
    profiles = step8_data.get("memberProfiles", [])

    if not profiles:
        p_bio_title = doc.add_paragraph()
        p_bio_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_single_spacing(p_bio_title, space_before_pt=0, space_after_pt=12)
        r_bio_title = p_bio_title.add_run("ประวัติคณะผู้จัดทำ")
        set_thai_font(r_bio_title, font_name=font_name, size_pt=CHAPTER_TITLE_SIZE, bold=True)
        add_bookmark_to_paragraph(doc, p_bio_title, "bm_biography")
        add_bio_hanging_row(doc, "ชื่อ - นามสกุล", "........................................................", font_name=font_name)
        return

    for idx, mp in enumerate(profiles, 1):
        p_bio_title = doc.add_paragraph()
        p_bio_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        apply_single_spacing(p_bio_title, space_before_pt=0, space_after_pt=12)
        if idx > 1:
            p_bio_title.paragraph_format.page_break_before = True
        r_bio_title = p_bio_title.add_run("ประวัติคณะผู้จัดทำ")
        set_thai_font(r_bio_title, font_name=font_name, size_pt=CHAPTER_TITLE_SIZE, bold=True)
        if idx == 1:
            add_bookmark_to_paragraph(doc, p_bio_title, "bm_biography")

        photo_bytes = load_image_bytes_from_source(mp.get("photoUrl", ""), target_aspect_ratio=(3.0 / 4.0))
        if photo_bytes:
            try:
                p_ph = doc.add_paragraph()
                p_ph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                apply_single_spacing(p_ph, space_before_pt=0, space_after_pt=12)
                r_ph = p_ph.add_run()
                r_ph.add_picture(
                    io.BytesIO(photo_bytes),
                    width=Cm(PORTRAIT_IMG_WIDTH_CM),
                    height=Cm(PORTRAIT_IMG_HEIGHT_CM)
                )
            except Exception:
                pass

        addr = mp.get("address", {}) if isinstance(mp.get("address"), dict) else {}
        edu = mp.get("education", {}) if isinstance(mp.get("education"), dict) else {}
        p_edu = edu.get("primary", {}) if isinstance(edu.get("primary"), dict) else {}
        j_edu = edu.get("juniorHigh", {}) if isinstance(edu.get("juniorHigh"), dict) else {}
        s_edu = edu.get("seniorHigh", {}) if isinstance(edu.get("seniorHigh"), dict) else {}

        full_name_str = clean_single_line_text(mp.get("fullName", "-")) or "-"
        dob_str = clean_single_line_text(mp.get("dob", "-")) or "-"
        address_str = (
            f"บ้านเลขที่ {clean_single_line_text(addr.get('houseNo', '-'))} "
            f"หมู่ที่ {clean_single_line_text(addr.get('moo', '-'))} "
            f"หมู่บ้าน{clean_single_line_text(addr.get('village', '-'))} "
            f"ตำบล{clean_single_line_text(addr.get('subDistrict', '-'))} "
            f"อำเภอ{clean_single_line_text(addr.get('district', '-'))} "
            f"จังหวัด{clean_single_line_text(addr.get('province', '-'))} "
            f"รหัสไปรษณีย์ {clean_single_line_text(addr.get('postalCode', '-'))}"
        )
        phone_str = clean_single_line_text(mp.get("phone", "-")) or "-"

        add_bio_hanging_row(doc, "ชื่อ - นามสกุล", full_name_str, font_name=font_name)
        add_bio_hanging_row(doc, "วัน เดือน ปีเกิด", dob_str, font_name=font_name)
        add_bio_hanging_row(doc, "ที่อยู่", address_str, font_name=font_name)

        prim_school = format_school_line(p_edu.get("schoolName", ""))
        prim_loc = f"อำเภอ{clean_single_line_text(p_edu.get('district', '.............'))} จังหวัด{clean_single_line_text(p_edu.get('province', '.............'))}"

        jh_school = format_school_line(j_edu.get("schoolName", ""))
        jh_loc = f"อำเภอ{clean_single_line_text(j_edu.get('district', '.............'))} จังหวัด{clean_single_line_text(j_edu.get('province', '.............'))}"

        sh_school = format_school_line(s_edu.get("schoolName", ""))
        sh_loc = f"อำเภอ{clean_single_line_text(s_edu.get('district', '.............'))} จังหวัด{clean_single_line_text(s_edu.get('province', '.............'))}"

        add_bio_hanging_row(doc, "ประวัติการศึกษา", "ระดับประถมศึกษา", font_name=font_name)
        add_bio_continuation_line(doc, prim_school, font_name=font_name)
        add_bio_continuation_line(doc, prim_loc, font_name=font_name)

        add_bio_continuation_line(doc, middle_edu_label, font_name=font_name)
        add_bio_continuation_line(doc, jh_school, font_name=font_name)
        add_bio_continuation_line(doc, jh_loc, font_name=font_name)

        add_bio_continuation_line(doc, "ปัจจุบันกำลังศึกษาอยู่ระดับมัธยมศึกษาตอนปลาย", font_name=font_name)
        add_bio_continuation_line(doc, sh_school, font_name=font_name)
        add_bio_continuation_line(doc, sh_loc, font_name=font_name)

        add_bio_hanging_row(doc, "เบอร์โทรศัพท์", phone_str, font_name=font_name)


def build_report_docx(
    group_data,
    font_name,
    cover_size,
    logo_width_cm,
    cover_name_col_w_cm,
    cover_num_col_w_cm,
    semester,
    academic_year,
    first_line_indent_cm,
    sub_left_indent_cm,
    body_thai_distribute,
    ch2_sub_numbering,
    max_table_rows_per_page,
    middle_edu_label,
    strip_advisor_honorific
):
    """สร้างไฟล์เอกสาร Word (.docx) ครบทั้งเล่ม รวมสารบัญทั้ง 3 หน้า และระบบเลขหน้าอัตโนมัติ"""
    doc = Document()
    configure_document_thai_settings(doc, font_name=font_name)
    
    normal_style = doc.styles['Normal']
    normal_style.paragraph_format.line_spacing = LINE_SPACING_VAL
    normal_style.paragraph_format.space_before = Pt(0)
    normal_style.paragraph_format.space_after = Pt(0)

    build_cover_page(
        doc,
        group_data,
        font_name=font_name,
        font_size=cover_size,
        logo_width_cm=logo_width_cm,
        name_col_width_cm=cover_name_col_w_cm,
        num_col_width_cm=cover_num_col_w_cm
    )
    
    build_preface_page(
        doc, group_data, font_name=font_name, semester=semester,
        academic_year=academic_year, first_line_indent_cm=first_line_indent_cm,
        strip_advisor_honorific=strip_advisor_honorific
    )

    build_all_toc_pages(doc, group_data, font_name=font_name)

    build_chapter1_page(
        doc, group_data, font_name=font_name, first_line_indent_cm=first_line_indent_cm,
        sub_left_indent_cm=sub_left_indent_cm, use_thai_distribute=body_thai_distribute
    )

    build_chapter2_page(
        doc, group_data, font_name=font_name, first_line_indent_cm=first_line_indent_cm,
        use_thai_distribute=body_thai_distribute, use_sub_numbering=ch2_sub_numbering
    )

    build_chapter3_page(
        doc, group_data, font_name=font_name, first_line_indent_cm=first_line_indent_cm,
        sub_left_indent_cm=sub_left_indent_cm, use_thai_distribute=body_thai_distribute
    )

    build_chapter4_page(
        doc, group_data, font_name=font_name, first_line_indent_cm=first_line_indent_cm,
        use_thai_distribute=body_thai_distribute, max_table_rows_per_page=max_table_rows_per_page
    )

    build_chapter5_page(
        doc, group_data, font_name=font_name, first_line_indent_cm=first_line_indent_cm,
        use_thai_distribute=body_thai_distribute
    )

    build_bibliography_and_appendix_pages(doc, group_data, font_name=font_name)

    build_biography_page(doc, group_data, font_name=font_name, middle_edu_label=middle_edu_label)

    return doc


def post_process_word_toc_and_headers(wdoc, font_name=REPORT_FONT_NAME):
    """จัดการฟอนต์เลขหน้า 14 pt และตัดเลขหน้าซ้ำในตารางสารบัญทั้ง 3 ตาราง"""
    for sec in wdoc.Sections:
        for hdr_idx in (1, 2):
            try:
                hdr_range = sec.Headers(hdr_idx).Range
                hdr_range.Font.Name = font_name
                hdr_range.Font.NameBi = font_name
                hdr_range.Font.Size = PAGE_NUMBER_FONT_SIZE
                hdr_range.Font.SizeBi = PAGE_NUMBER_FONT_SIZE
                hdr_range.Font.Bold = False
                hdr_range.Font.BoldBi = False
            except Exception:
                pass

    for tbl in wdoc.Tables:
        try:
            if tbl.Rows.Count < 2:
                continue
            num_cols = tbl.Columns.Count
            first_cell_txt = tbl.Cell(1, 1).Range.Text.replace('\r', '').replace('\x07', '').strip()
            last_cell_txt = tbl.Cell(1, num_cols).Range.Text.replace('\r', '').replace('\x07', '').strip()

            if first_cell_txt in ("เรื่อง", "ตารางที่", "ภาพประกอบที่") and last_cell_txt == "หน้า":
                last_seen_page = None
                for r_idx in range(2, tbl.Rows.Count + 1):
                    pg_cell = tbl.Cell(r_idx, num_cols)
                    try:
                        pg_cell.Range.Fields.Unlink()
                    except Exception:
                        pass

                    pg_text = pg_cell.Range.Text.replace('\r', '').replace('\x07', '').strip()
                    if pg_text and pg_text == last_seen_page:
                        pg_cell.Range.Text = ""
                    else:
                        if pg_text:
                            last_seen_page = pg_text
                        pg_cell.Range.Text = pg_text
                        pg_cell.Range.ParagraphFormat.Alignment = 1
                        pg_cell.Range.Font.Name = font_name
                        pg_cell.Range.Font.NameBi = font_name
                        pg_cell.Range.Font.Size = BODY_FONT_SIZE
                        pg_cell.Range.Font.SizeBi = BODY_FONT_SIZE
                        pg_cell.Range.Font.Bold = False
                        pg_cell.Range.Font.BoldBi = False
        except Exception:
            pass


@st.cache_resource
def setup_linux_fonts_once():
    """
    ค้นหาไฟล์ฟอนต์ .ttf (เช่น THSarabunPSK.ttf) ที่วางอยู่ในโฟลเดอร์โปรเจ็กต์บน GitHub
    แล้วติดตั้งเข้าสู่ระบบ Linux ของ Streamlit Cloud อัตโนมัติ เพื่อให้ PDF แสดงฟอนต์ตรงตามต้นฉบับ
    """
    if WORD_COM_AVAILABLE:
        return True
    try:
        user_font_dir = os.path.expanduser("~/.local/share/fonts")
        os.makedirs(user_font_dir, exist_ok=True)
        
        # ค้นหาไฟล์ .ttf ในโฟลเดอร์หลัก หรือโฟลเดอร์ fonts/
        search_dirs = [".", "fonts"]
        copied_any = False
        for s_dir in search_dirs:
            if os.path.isdir(s_dir):
                for fname in os.listdir(s_dir):
                    if fname.lower().endswith(".ttf"):
                        src_path = os.path.join(s_dir, fname)
                        dst_path = os.path.join(user_font_dir, fname)
                        if not os.path.exists(dst_path):
                            shutil.copy2(src_path, dst_path)
                            copied_any = True
        if copied_any:
            subprocess.run(["fc-cache", "-f", user_font_dir], check=False, timeout=20)
        return True
    except Exception:
        return False


def convert_docx_to_pdf_linux(docx_bytes):
    """
    แปลงไฟล์ Word (.docx) เป็น PDF (.pdf) บน Streamlit Cloud (Linux) ด้วย LibreOffice
    โดยแยกโฟลเดอร์โปรไฟล์ชั่วคราว (-env:UserInstallation) เพื่อรองรับนักเรียนใช้งานพร้อมกันหลายคน
    """
    soffice_bin = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice_bin:
        return None

    setup_linux_fonts_once()

    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, "report_full.docx")
        pdf_path = os.path.join(tmpdir, "report_full.pdf")
        profile_dir = os.path.join(tmpdir, "lo_profile")
        profile_uri = f"file://{profile_dir}"

        with open(docx_path, "wb") as f:
            f.write(docx_bytes)

        cmd = [
            soffice_bin,
            f"-env:UserInstallation={profile_uri}",
            "--headless",
            "--nologo",
            "--nofirststartwizard",
            "--convert-to",
            "pdf:writer_pdf_Export",
            "--outdir",
            tmpdir,
            docx_path
        ]
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=90)

        if os.path.exists(pdf_path):
            with open(pdf_path, "rb") as f:
                return f.read()
    return None


def generate_docx_and_pdf_bytes(
    group_data,
    font_name,
    cover_size,
    logo_width_cm,
    cover_name_col_w_cm,
    cover_num_col_w_cm,
    semester,
    academic_year,
    first_line_indent_cm,
    sub_left_indent_cm,
    body_thai_distribute,
    ch2_sub_numbering,
    max_table_rows_per_page,
    middle_edu_label,
    strip_advisor_honorific
):
    """
    บันทึกเป็นไฟล์ Word (.docx) และแปลงเป็นไฟล์ PDF (.pdf)
    รองรับทั้ง Windows (Microsoft Word) และ Streamlit Community Cloud (Linux LibreOffice)
    """
    doc = build_report_docx(
        group_data, font_name, cover_size, logo_width_cm,
        cover_name_col_w_cm, cover_num_col_w_cm,
        semester, academic_year, first_line_indent_cm, sub_left_indent_cm,
        body_thai_distribute, ch2_sub_numbering, max_table_rows_per_page,
        middle_edu_label, strip_advisor_honorific
    )
    
    docx_buffer = io.BytesIO()
    doc.save(docx_buffer)
    docx_bytes = docx_buffer.getvalue()
    
    pdf_bytes = None
    if WORD_COM_AVAILABLE:
        try:
            pythoncom.CoInitialize()
            with tempfile.TemporaryDirectory() as tmpdir:
                docx_path = os.path.abspath(os.path.join(tmpdir, "report_full.docx"))
                pdf_path = os.path.abspath(os.path.join(tmpdir, "report_full.pdf"))
                with open(docx_path, "wb") as f:
                    f.write(docx_bytes)

                word_app = win32com.client.DispatchEx("Word.Application")
                word_app.Visible = False
                word_app.DisplayAlerts = 0
                try:
                    wdoc = word_app.Documents.Open(docx_path)
                    wdoc.Repaginate()
                    wdoc.Fields.Update()
                    for story_range in wdoc.StoryRanges:
                        story_range.Fields.Update()

                    post_process_word_toc_and_headers(wdoc, font_name=font_name)

                    wdoc.Save()
                    wdoc.ExportAsFixedFormat(pdf_path, 17)
                    wdoc.Close(SaveChanges=False)
                finally:
                    word_app.Quit()

                if os.path.exists(docx_path):
                    with open(docx_path, "rb") as f:
                        docx_bytes = f.read()
                if os.path.exists(pdf_path):
                    with open(pdf_path, "rb") as f:
                        pdf_bytes = f.read()
        except Exception as e:
            st.warning(f"สร้างไฟล์ Word สำเร็จ แต่ไม่สามารถอัปเดตผ่าน Word อัตโนมัติได้ (สาเหตุ: {e})")
    elif (shutil.which("soffice") or shutil.which("libreoffice")) is not None:
        try:
            pdf_bytes = convert_docx_to_pdf_linux(docx_bytes)
        except Exception as e:
            st.warning(f"สร้างไฟล์ Word สำเร็จ แต่ไม่สามารถแปลง PDF บน Cloud ได้ (สาเหตุ: {e})")
    elif DOCX2PDF_AVAILABLE:
        try:
            pythoncom.CoInitialize()
            with tempfile.TemporaryDirectory() as tmpdir:
                docx_path = os.path.join(tmpdir, "report_full.docx")
                pdf_path = os.path.join(tmpdir, "report_full.pdf")
                with open(docx_path, "wb") as f:
                    f.write(docx_bytes)
                convert(docx_path, pdf_path)
                if os.path.exists(pdf_path):
                    with open(pdf_path, "rb") as f:
                        pdf_bytes = f.read()
        except Exception as e:
            st.warning(f"สร้างไฟล์ Word สำเร็จ แต่ไม่สามารถแปลงเป็น PDF ได้ (สาเหตุ: {e})")
            
    return docx_bytes, pdf_bytes


def display_pdf_preview(pdf_bytes):
    """แสดงหน้าต่าง Preview PDF บนหน้าเว็บ"""
    b64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")
    pdf_html = f'<iframe src="data:application/pdf;base64,{b64_pdf}" width="100%" height="850" style="border-radius: 12px; border: 1px solid #cbd5e1;" type="application/pdf"></iframe>'
    st.markdown(pdf_html, unsafe_allow_html=True)

# ==========================================================
# ส่วนที่ 6: หน้าจอผู้ใช้งาน (Streamlit Web UI)
# ==========================================================
with st.sidebar:
    st.header("🌐 เชื่อมต่อฐานข้อมูล")
    web_app_url_input = st.text_input(
        "ลิงก์ Web App (Google Apps Script):",
        value=DEFAULT_WEB_APP_URL
    )
    if st.button("🔄 ดึงข้อมูลกลุ่มจาก Google Sheets", type="primary", use_container_width=True):
        fetch_all_groups_from_gas.clear()
        with st.spinner("กำลังดึงข้อมูลออนไลน์จาก Google Sheets..."):
            res = fetch_all_groups_from_gas(web_app_url_input)
            if res.get("success"):
                st.session_state["all_groups_data"] = res.get("groups", [])
                st.success(f"ดึงข้อมูลสำเร็จ ({len(st.session_state['all_groups_data'])} กลุ่ม)")
            else:
                st.error(res.get("message", "เกิดข้อผิดพลาด"))

    st.divider()
    st.subheader("📐 ปรับแต่งรายชื่อหน้าปก")
    col_sem, col_yr = st.columns(2)
    with col_sem:
        semester_input = st.text_input("ภาคเรียนที่:", value=DEFAULT_SEMESTER)
    with col_yr:
        academic_year_input = st.text_input("ปีการศึกษา:", value=DEFAULT_ACADEMIC_YEAR)

    selected_font = st.selectbox("เลือกชื่อฟอนต์เอกสาร Word/PDF:", ["TH SarabunPSK", "TH Sarabun New"], index=0)
    cover_name_col_w = st.slider(
        "ความกว้างคอลัมน์ชื่อ-นามสกุล บนหน้าปก (ซม.):",
        min_value=4.5,
        max_value=11.0,
        value=DEFAULT_COVER_NAME_COL_CM,
        step=0.25
    )

    with st.expander("⚙️ ตั้งค่าเอกสาร", expanded=False):
        ui_font_choice = st.selectbox(
            "ฟอนต์แสดงผลหน้าเว็บระบบ (Web UI Font):",
            ["IBM Plex Sans Thai", "Bai Jamjuree", "Noto Serif Thai"],
            index=0
        )
        selected_logo_w = st.slider(
            "ความกว้างตราโรงเรียนบนหน้าปก (ซม.):",
            min_value=2.0,
            max_value=5.0,
            value=3.2,
            step=0.1
        )
        cover_num_col_w = st.slider(
            "ความกว้างคอลัมน์เลขที่ บนหน้าปก (ซม.):",
            min_value=1.5,
            max_value=5.0,
            value=DEFAULT_COVER_NUM_COL_CM,
            step=0.1
        )
        indent_cm = st.slider(
            "ระยะย่อหน้าบรรทัดแรก (ซม.):",
            min_value=0.0,
            max_value=2.5,
            value=1.5,
            step=0.25
        )
        sub_indent_cm = st.slider(
            "ระยะย่อหน้าหัวข้อย่อย (ซม.):",
            min_value=0.0,
            max_value=2.5,
            value=1.0,
            step=0.25
        )
        max_rows_page = st.slider(
            "จำนวนแถวตารางสูงสุดต่อหน้า (ก่อนตัดขึ้นตาราง (ต่อ)):",
            min_value=5,
            max_value=25,
            value=16,
            step=1
        )
        mid_edu_label_opt = st.selectbox(
            "หัวข้อการศึกษาบล็อกที่ 2 ในหน้าประวัติผู้จัดทำ:",
            ["ระดับมัธยมศึกษาตอนต้น", "ระดับประถมศึกษา"],
            index=0
        )
        strip_hon_opt = st.checkbox(
            "ตัดคำนำหน้า (นาย/นาง/นางสาว) หลังคำว่า 'คุณครู' ในคำนำ",
            value=True
        )
        body_distribute_opt = st.checkbox(
            "จัดเนื้อหาย่อหน้าแบบ 'กระจายแบบไทย'",
            value=True
        )
        ch2_subnum_opt = st.checkbox(
            "ใช้เลขหัวข้อทฤษฎีในบทที่ 2 แบบ 1.1, 1.2",
            value=False
        )

inject_formal_academic_css(ui_font_family=ui_font_choice)
render_hero_header()

if "all_groups_data" not in st.session_state and DEFAULT_WEB_APP_URL:
    with st.spinner("กำลังเชื่อมต่อฐานข้อมูล Google Sheets อัตโนมัติ..."):
        init_res = fetch_all_groups_from_gas(DEFAULT_WEB_APP_URL)
        if init_res.get("success"):
            st.session_state["all_groups_data"] = init_res.get("groups", [])

all_groups = st.session_state.get("all_groups_data", [])

if not all_groups:
    st.info("👈 กดปุ่ม **'🔄 ดึงข้อมูลกลุ่มจาก Google Sheets'** ที่แถบด้านซ้ายมือ เพื่อเริ่มต้นสร้างรูปเล่มรายงานครับ")
else:
    col1, col2 = st.columns([1, 2])
    rooms_set = sorted(list({str(g["members"][0].get("room", "-")).strip() for g in all_groups if g.get("members")}))
    with col1:
        room_filter = st.selectbox("🏫 เลือกชั้น/ห้องเรียน:", ["แสดงทุกห้อง"] + rooms_set)
    
    filtered_groups = [
        g for g in all_groups
        if room_filter == "แสดงทุกห้อง" or (g.get("members") and str(g["members"][0].get("room", "-")).strip() == room_filter)
    ]

    with col2:
        group_options = [
            f"{g.get('groupName', '-')} — เรื่อง: {g.get('projectTitle') or '(ยังไม่ระบุชื่อเรื่อง)'}"
            for g in filtered_groups
        ]
        chosen_idx = st.selectbox("👥 เลือกกลุ่มโครงงาน:", range(len(group_options)), format_func=lambda i: group_options[i])

    if filtered_groups and chosen_idx is not None:
        active_group = filtered_groups[chosen_idx]
        progress_steps = active_group.get("progress", {}).get("steps", {})
        s1_data = (progress_steps.get("1") or progress_steps.get(1) or {}).get("data", {})
        project_title_disp = clean_single_line_text(
            s1_data.get("projectTitle") or active_group.get("projectTitle") or "(ยังไม่ได้ระบุชื่อโครงงาน)"
        )
        members_list = active_group.get("members", [])
        advisors_disp = ", ".join(s1_data.get("advisors", ["นายธนโชติ จันทรังษี"]))

        # ตัวอย่างชื่อไฟล์ที่จะได้เมื่อกดดาวน์โหลด
        preview_pdf_filename = build_export_filename(active_group, "pdf")

        members_badges_html = "".join([
            f'<span class="edu-badge">👤 {m.get("fullName", "-")} (เลขที่ {m.get("number", "-")})</span>'
            for m in members_list
        ])

        st.markdown(
            f"""
            <div class="edu-info-card">
                <div class="edu-card-heading">📌 ข้อมูลกลุ่มโครงงาน: {active_group.get('groupName', '-')}</div>
                <div style="margin-bottom: 6px;"><strong>ชื่อเรื่องโครงงานอาชีพ:</strong> <span style="color:#0f2942; font-weight:600;">{project_title_disp}</span></div>
                <div style="margin-bottom: 6px;"><strong>ครูที่ปรึกษาโครงงาน:</strong> {advisors_disp}</div>
                <div style="margin-bottom: 10px;"><strong>ชื่อไฟล์ที่จะดาวน์โหลด:</strong> <code style="color:#1b4332; font-weight:600;">{preview_pdf_filename}</code></div>
                <div><strong>สมาชิกในกลุ่ม ({len(members_list)} คน):</strong><br><div style="margin-top:6px;">{members_badges_html}</div></div>
            </div>
            """,
            unsafe_allow_html=True
        )

        if st.button("🚀 สร้างเอกสารรายงานโครงงานอาชีพ (Word & PDF)", type="primary", use_container_width=True):
            with st.spinner("กำลังจัดหน้าเอกสารรายงานโครงงานอาชีพ และแปลงเป็นไฟล์ PDF..."):
                docx_b, pdf_b = generate_docx_and_pdf_bytes(
                    active_group,
                    font_name=selected_font,
                    cover_size=COVER_FONT_SIZE,
                    logo_width_cm=selected_logo_w,
                    cover_name_col_w_cm=cover_name_col_w,
                    cover_num_col_w_cm=cover_num_col_w,
                    semester=semester_input,
                    academic_year=academic_year_input,
                    first_line_indent_cm=indent_cm,
                    sub_left_indent_cm=sub_indent_cm,
                    body_thai_distribute=body_distribute_opt,
                    ch2_sub_numbering=ch2_subnum_opt,
                    max_table_rows_per_page=max_rows_page,
                    middle_edu_label=mid_edu_label_opt,
                    strip_advisor_honorific=strip_hon_opt
                )
                st.session_state["report_docx"] = docx_b
                st.session_state["report_pdf"] = pdf_b
                st.session_state["export_docx_filename"] = build_export_filename(active_group, "docx")
                st.session_state["export_pdf_filename"] = build_export_filename(active_group, "pdf")

        if "report_docx" in st.session_state:
            docx_fname = st.session_state.get("export_docx_filename", "โครงงานอาชีพ.docx")
            pdf_fname = st.session_state.get("export_pdf_filename", "โครงงานอาชีพ.pdf")

            if st.session_state.get("report_pdf"):
                dl_col1, dl_col2 = st.columns(2)
                with dl_col1:
                    st.download_button(
                        label=f"📥 ดาวน์โหลด Word: {docx_fname}",
                        data=st.session_state["report_docx"],
                        file_name=docx_fname,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                with dl_col2:
                    st.download_button(
                        label=f"📕 ดาวน์โหลด PDF: {pdf_fname}",
                        data=st.session_state["report_pdf"],
                        file_name=pdf_fname,
                        mime="application/pdf",
                        use_container_width=True
                    )
                st.subheader("🔍 ตัวอย่างเอกสารรายงานโครงงานอาชีพ (PDF Preview)")
                display_pdf_preview(st.session_state["report_pdf"])
            else:
                st.success("✅ สร้างไฟล์เอกสารรายงานโครงงานอาชีพเรียบร้อยแล้ว! กดปุ่มด้านล่างเพื่อดาวน์โหลดไฟล์ได้เลยครับ")
                st.download_button(
                    label=f"📥 ดาวน์โหลดเอกสารรายงาน Word (.docx): {docx_fname}",
                    data=st.session_state["report_docx"],
                    file_name=docx_fname,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            
            if st.session_state.get("report_pdf"):
                st.subheader("🔍 ตัวอย่างเอกสารรายงานโครงงานอาชีพ (PDF Preview)")
                display_pdf_preview(st.session_state["report_pdf"])

render_footer()
