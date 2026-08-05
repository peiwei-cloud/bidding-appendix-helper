# -*- coding: utf-8 -*-
"""
備標人員附錄整理、證照檢核與 CV 結構化轉檔系統
==================================================
功能：
1. 依人員名單 Excel 排序，將每位同仁的 CV / 學歷 / 證照 / 技師執業執照 /
   技師會員證 / 投保證明，依標準 6 分類順序合併成單一附錄 PDF。
2. 自動掃描技師執業執照與公會會員證內文字，判讀民國年效期，產出「已過期 /
   即將過期 / 有效」檢核報告。
3. 呼叫 Gemini AI 解析每位同仁 CV 全文，結構化萃取為符合
   Template_Staffing.xlsx 標準格式的 13 個欄位。

部署需求（Streamlit Community Cloud）：
- 本機 / 雲端皆需安裝 LibreOffice（將 .doc/.docx 轉為 PDF 供合併，以及供
  文字擷取備援）。若部署到 Streamlit Cloud，請在 repo 根目錄新增
  `packages.txt`，內容如下（本工具已一併產生於輸出資料夾）：
      libreoffice
      tesseract-ocr
      tesseract-ocr-chi-tra
      poppler-utils
- Gemini API Key：於 App 的 Settings -> Secrets 貼上
      GEMINI_API_KEY = "你的 Gemini API Key"
  或於側邊欄手動輸入（僅存於本次瀏覽器工作階段）。

本機執行：
    pip install -r requirements.txt
    streamlit run app.py
"""

import difflib
import io
import json
import os
import random
import re
import subprocess
import tempfile
import time
import zipfile
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import openpyxl

try:
    from docx import Document as DocxDocument
except Exception:  # pragma: no cover
    DocxDocument = None

try:
    from pypdf import PdfReader, PdfWriter
except Exception:  # pragma: no cover
    PdfReader = None
    PdfWriter = None

try:
    import pdfplumber
except Exception:  # pragma: no cover
    pdfplumber = None

try:
    from PIL import Image
except Exception:  # pragma: no cover
    Image = None

try:
    from google import genai
    GENAI_AVAILABLE = True
except Exception:  # pragma: no cover
    GENAI_AVAILABLE = False
    genai = None


# ============================================================
# 常數設定
# ============================================================

CATEGORY_ORDER = [
    "1_CV",
    "2_學歷",
    "3_證照",
    "4_技師執業執照",
    "5_技師會員證",
    "6_投保證明",
]

CATEGORY_LABELS = {
    "1_CV": "CV／履歷",
    "2_學歷": "學歷證明",
    "3_證照": "一般證照",
    "4_技師執業執照": "技師執業執照",
    "5_技師會員證": "技師公會會員證",
    "6_投保證明": "投保證明",
}

CATEGORY_KEYWORDS = {
    "6_投保證明": ["投保", "被保險人", "被保险人", "勞保", "劳保", "勞退"],
    "2_學歷": ["畢業", "毕业", "學歷", "学历", "學位", "学位"],
}

# CV 檔名關鍵字（擴充：學經歷／經歷表）
CV_FILENAME_KEYWORDS = ["cv", "履歷", "履历", "學經歷", "学经历", "經歷表", "经历表"]

# 三段式雙軌配對機制信心度門檻
HIGH_CONFIDENCE_THRESHOLD = 0.80
MEDIUM_CONFIDENCE_THRESHOLD = 0.60

TEMPLATE_COLUMNS = [
    "Layer", "Role", "GroupName", "Name", "Title", "Company", "Badges",
    "PhotoName", "YearsOfExp", "Degree", "JobDescription", "Expertise",
    "BioNarrative",
]

LAYER_OPTIONS = ["Top", "SubTop", "Advisor", "Middle", "GroupLeader",
                  "GroupMember", "Subcontractor"]

BADGE_SORT_ORDER = ["技", "碩", "博", "品", "安", "採", "乙", "甲", "景", "土", "水"]

DEFAULT_MODEL_CHAIN = [
    "gemini-3.6-flash",
    "gemini-3.5-flash",
    "gemini-2.5-flash",
    "gemini-flash-latest",
]

EXCLUDE_DIR_TOKENS = ["目前不需要", "不需要", "__macosx", ".git"]


# ============================================================
# 工具函式：檔案分類與人名比對
# ============================================================

def classify_category_by_filename_hint(fname_lower: str) -> str:
    """僅用於投保證明／學歷證明（內容結構化程度低，維持以檔名關鍵字判斷）。
    找不到關鍵字則回傳 None，交由主分類流程繼續判斷。"""
    for cat in ["6_投保證明", "2_學歷"]:
        for kw in CATEGORY_KEYWORDS[cat]:
            if kw.lower() in fname_lower:
                return cat
    return None


# ------------------------------------------------------------
# 三段式雙軌配對機制（姓名 + 英文姓名，容錯拼寫誤差）
# ------------------------------------------------------------

def normalize_name(s: str) -> str:
    """轉小寫並移除常見分隔符號，供比對使用。"""
    s = str(s or "").lower()
    s = re.sub(r"[_\-()（）\s,，、.。]", "", s)
    return s


def _windowed_ratio(alias_norm: str, stem_norm: str) -> float:
    """計算 alias 與檔名（正規化後）的相似度，錨定於檔名最左側區段，
    並輔以全字串比對，兼顧「姓名在最左側」與「姓名夾雜於中段」兩種情況。"""
    if not alias_norm or not stem_norm:
        return 0.0
    if alias_norm in stem_norm:
        return 1.0

    best = 0.0
    lo = max(1, len(alias_norm) - 3)
    hi = min(len(stem_norm), len(alias_norm) + 6)
    for end in range(lo, hi + 1):
        window = stem_norm[:end]
        ratio = difflib.SequenceMatcher(None, alias_norm, window).ratio()
        if ratio > best:
            best = ratio

    full_ratio = difflib.SequenceMatcher(None, alias_norm, stem_norm).ratio()
    return max(best, full_ratio)


def build_alias_list(df_people: pd.DataFrame) -> list:
    """回傳 [(alias_norm, alias_original, person_name), ...]，
    姓名與英文姓名皆納入比對清單。"""
    aliases = []
    for _, row in df_people.iterrows():
        person = str(row["姓名"]).strip()
        for col in ["姓名", "英文姓名"]:
            val = row.get(col, "") if col in df_people.columns else ""
            if val is None or (isinstance(val, float) and pd.isna(val)):
                continue
            val = str(val).strip()
            if val and val.lower() != "nan":
                aliases.append((normalize_name(val), val, person))
    return aliases


def match_person_candidates(filename: str, aliases: list) -> list:
    """回傳依信心分數（0~1）由高到低排序的
    [(person_name, score, matched_alias_text), ...]。"""
    stem = Path(filename).stem
    stem_norm = normalize_name(stem)

    best_per_person = {}
    for alias_norm, alias_text, person in aliases:
        score = _windowed_ratio(alias_norm, stem_norm)
        if person not in best_per_person or score > best_per_person[person][0]:
            best_per_person[person] = (score, alias_text)

    ranked = sorted(
        [(person, score, alias_text) for person, (score, alias_text) in best_per_person.items()],
        key=lambda x: x[1],
        reverse=True,
    )
    return ranked


# ============================================================
# 工具函式：格式轉換
# ============================================================

def convert_to_pdf(filepath: str, workdir: str):
    """將 docx/doc/圖片轉為 PDF；已是 PDF 則直接回傳原路徑。失敗回傳 None。"""
    ext = Path(filepath).suffix.lower()
    if ext == ".pdf":
        return filepath

    if ext in (".doc", ".docx"):
        try:
            subprocess.run(
                ["libreoffice", "--headless", "--convert-to", "pdf",
                 "--outdir", workdir, filepath],
                check=True, timeout=180, capture_output=True,
            )
            out_path = Path(workdir) / (Path(filepath).stem + ".pdf")
            return str(out_path) if out_path.exists() else None
        except Exception:
            return None

    if ext in (".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif") and Image is not None:
        try:
            img = Image.open(filepath).convert("RGB")
            out_path = Path(workdir) / (Path(filepath).stem + ".pdf")
            img.save(out_path, "PDF")
            return str(out_path)
        except Exception:
            return None

    return None


def extract_docx_text(filepath: str, workdir: str) -> str:
    """讀取 Word CV 全文（含表格）。.doc 會先透過 LibreOffice 轉為 .docx。"""
    ext = Path(filepath).suffix.lower()
    target = filepath

    if ext == ".doc":
        try:
            subprocess.run(
                ["libreoffice", "--headless", "--convert-to", "docx",
                 "--outdir", workdir, filepath],
                check=True, timeout=180, capture_output=True,
            )
            converted = Path(workdir) / (Path(filepath).stem + ".docx")
            if converted.exists():
                target = str(converted)
        except Exception:
            pass

    if DocxDocument is None:
        return ""

    try:
        doc = DocxDocument(target)
        parts = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        parts.append(cell.text.strip())
        return "\n".join(parts)
    except Exception:
        return ""


def extract_cv_text(filepath: str, workdir: str) -> str:
    """依副檔名分派 CV 文字擷取方式：.pdf 走 PDF 文字/OCR 擷取，
    .doc/.docx 走 Word 段落/表格擷取。確保 PDF 格式 CV 不再被跳過。"""
    ext = Path(filepath).suffix.lower()
    if ext == ".pdf":
        return extract_pdf_text(filepath)
    if ext in (".doc", ".docx"):
        return extract_docx_text(filepath, workdir)
    return ""


def extract_pdf_text(pdf_path: str) -> str:
    """從 PDF 擷取文字，若為掃描檔（無文字層）則嘗試 OCR 備援。"""
    text = ""
    if pdfplumber is not None:
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t:
                        text += t + "\n"
        except Exception:
            pass

    if not text.strip():
        try:
            from pdf2image import convert_from_path
            import pytesseract
            images = convert_from_path(pdf_path, dpi=200)
            for img in images:
                text += pytesseract.image_to_string(img, lang="chi_tra+eng") + "\n"
        except Exception:
            pass

    return text


def classify_file_content(filepath: str, workdir: str):
    """內文導向（Content-Driven）分類。回傳 (category, extracted_text)。

    - 1_CV：以檔名關鍵字（CV/履歷/學經歷/經歷表，不分大小寫）或 Word 副檔名
      （.doc/.docx，不分大小寫）快速判定，不需額外 OCR。
    - 4_技師執業執照：需先擷取 PDF/圖片內文字（含 OCR 備援），內文包含
      「技師執業執照」或「執業執照」才成立。
    - 5_技師會員證：內文「同時」包含「會員證」與「技師公會」才成立
      （例如：台北市水利技師公會、台灣省水利技師公會），避免景觀學會等
      一般協會會員證被誤判。
    - 6_投保證明 / 2_學歷：內容結構化程度低，維持以檔名關鍵字判斷。
    - 其餘：歸為 3_證照，不進行過期告警。

    extracted_text 僅在實際執行過 PDF/OCR 擷取時回傳（CV／未擷取則為 None），
    供效期檢核複用，避免對同一檔案重複 OCR。
    """
    fname = os.path.basename(filepath)
    fname_lower = fname.lower()
    ext = Path(filepath).suffix.lower()

    if any(kw.lower() in fname_lower for kw in CV_FILENAME_KEYWORDS):
        return "1_CV", None
    if ext in (".doc", ".docx"):
        return "1_CV", None

    pdf_path = filepath if ext == ".pdf" else convert_to_pdf(filepath, workdir)
    text = extract_pdf_text(pdf_path) if pdf_path else ""

    if "技師執業執照" in text or "執業執照" in text or "执业执照" in text:
        return "4_技師執業執照", text

    has_member_kw = ("會員證" in text) or ("会员证" in text)
    has_guild_kw = ("技師公會" in text) or ("技师公会" in text)
    if has_member_kw and has_guild_kw:
        return "5_技師會員證", text

    hint = classify_category_by_filename_hint(fname_lower)
    if hint:
        return hint, text

    return "3_證照", text


# ============================================================
# 證照效期檢核
# ============================================================

DATE_RANGE_PATTERN = re.compile(
    r"(?:自)?民國\s*(\d{2,3})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日\s*"
    r"(?:止|至)\s*(?:民國\s*)?(\d{2,3})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日"
)

# 「115年會員證」／「115年度會員證」等只標示單一年度、無明確起訖日的格式：
# 推算至該年12/31止（僅供 5_技師會員證 使用）
YEAR_ONLY_PATTERN = re.compile(r"(\d{2,3})\s*年度?\s*(?:會員證|会员证)")


def parse_license_expiry(category: str, text: str) -> dict:
    """依「已擷取好」的內文文字解析證照效期，僅供 4_技師執業執照 與
    5_技師會員證 呼叫；不重複進行 PDF/OCR 擷取。日期一律以呼叫當下的
    date.today() 動態計算，避免因常數在應用程式啟動時凍結而產生誤判。

    優先順序：明確起訖日期 > 單一年度（僅會員證，推算至當年12/31）。
    """
    if not text or not text.strip():
        return {"狀態": "⚠️ 無法自動判讀（掃描檔/OCR不可用）", "起始日": "", "截止日": "",
                "備註": "請人工確認"}

    start = end = None

    m = DATE_RANGE_PATTERN.search(text)
    if m:
        ry1, rm1, rd1, ry2, rm2, rd2 = map(int, m.groups())
        try:
            start = date(ry1 + 1911, rm1, rd1)
            end = date(ry2 + 1911, rm2, rd2)
        except ValueError:
            return {"狀態": "⚠️ 日期格式錯誤", "起始日": "", "截止日": "", "備註": "請人工確認"}
    elif category == "5_技師會員證":
        m2 = YEAR_ONLY_PATTERN.search(text)
        if m2:
            ry = int(m2.group(1))
            try:
                start = date(ry + 1911, 1, 1)
                end = date(ry + 1911, 12, 31)
            except ValueError:
                return {"狀態": "⚠️ 日期格式錯誤", "起始日": "", "截止日": "", "備註": "請人工確認"}

    if start is None or end is None:
        return {"狀態": "⚠️ 未偵測到效期文字", "起始日": "", "截止日": "", "備註": "請人工確認"}

    today = date.today()
    days_left = (end - today).days
    if days_left < 0:
        status = "🔴 已過期"
        remark = f"逾期 {-days_left} 天"
    elif days_left <= 90:
        status = "🟡 即將過期（90天內）"
        remark = f"剩 {days_left} 天"
    else:
        status = "🟢 有效"
        remark = f"剩 {days_left} 天"

    return {"狀態": status, "起始日": start.isoformat(), "截止日": end.isoformat(),
            "備註": remark}


# ============================================================
# Badge / Layer 判定
# ============================================================

def rule_based_badges(cats: dict) -> set:
    badges = set()
    if cats.get("4_技師執業執照") or cats.get("5_技師會員證"):
        badges.add("技")

    all_names_text = " ".join(
        os.path.basename(f) for files in cats.values() for f in files
    )
    keyword_map = {
        "碩": ["碩士", "碩"],
        "博": ["博士"],
        "採": ["採購"],
        "品": ["品質"],
        "安": ["安全衛生", "勞安"],
        "乙": ["乙級"],
        "甲": ["甲級"],
        "景": ["景觀"],
        "土": ["土木"],
        "水": ["水利"],
    }
    for badge, kws in keyword_map.items():
        if any(kw in all_names_text for kw in kws):
            badges.add(badge)
    return badges


def merge_and_sort_badges(ai_badges_str: str, rule_badges: set) -> str:
    ai_set = set()
    if ai_badges_str:
        for b in re.split(r"[,，、/\s]+", str(ai_badges_str)):
            b = b.strip()
            if b:
                ai_set.add(b)
    merged = ai_set | rule_badges
    ordered = [b for b in BADGE_SORT_ORDER if b in merged]
    extra = [b for b in merged if b not in BADGE_SORT_ORDER]
    return ",".join(ordered + extra)


def determine_layer(role_text: str) -> str:
    role_text = str(role_text or "")
    if "計畫顧問" in role_text or "品質督導" in role_text:
        return "Advisor"
    if "計畫主持人" in role_text or "協同主持人" in role_text or "代表廠商" in role_text:
        return "Top"
    if "設計負責人" in role_text:
        return "SubTop"
    if "計畫經理" in role_text:
        return "Middle"
    if "協力廠商" in role_text or "分包" in role_text:
        return "Subcontractor"
    if "組長" in role_text:
        return "GroupLeader"
    return "GroupMember"


# ============================================================
# Gemini AI 結構化萃取（含模型自動退避機制）
# ============================================================

def build_model_chain(selected_model: str) -> list:
    chain = [selected_model] + DEFAULT_MODEL_CHAIN
    seen = []
    for m in chain:
        if m and m not in seen:
            seen.append(m)
    return seen


def gemini_extract_person(client, model_chain, name, role, group, cv_text, badge_hints):
    prompt = f"""你是專業的標案人員履歷分析師。請閱讀以下人員的 CV 全文，並嚴格以 JSON
格式輸出以下欄位（僅輸出 JSON，不要有任何其他文字、不要使用 markdown code block）：

{{
  "Title": "公司職稱（如：資深協理）",
  "Company": "目前服務公司全名",
  "Badges": "從CV中辨識出的證照縮寫代碼，逗號分隔，可選值：技,碩,博,品,安,採,乙,甲,景,土,水",
  "YearsOfExp": 從CV推算的總工作年資（僅數字，無法判斷則填0）,
  "Degree": "最高學歷校名與科系",
  "JobDescription": "依據此人擔任的團隊職務（{role}），歸納其在本專案中擬任的工作內容，50字以內",
  "Expertise": "專長關鍵字，以斜線分隔，最多6項",
  "BioNarrative": "200字以內、具評審說服力的專業經歷敘述，需包含資歷年數、代表性經驗與專長"
}}

人員姓名：{name}
團隊職務：{role}
組別：{group}
已知證照線索（來自證照檔案分類，可供 Badges 參考）：{badge_hints}

CV全文如下：
---
{cv_text[:8000]}
---
"""
    last_err = None
    for model in model_chain:
        for attempt in range(3):
            try:
                resp = client.models.generate_content(model=model, contents=prompt)
                raw = (resp.text or "").strip()
                raw = re.sub(r"^```json\s*|\s*```$", "", raw, flags=re.MULTILINE).strip("`").strip()
                data = json.loads(raw)
                return data, model
            except json.JSONDecodeError as e:
                last_err = e
                break  # 解析失敗，換下一個模型，不重試同一模型
            except Exception as e:
                msg = str(e)
                if "404" in msg or "NOT_FOUND" in msg.upper():
                    last_err = e
                    break
                if "503" in msg or "UNAVAILABLE" in msg.upper() or "overloaded" in msg.lower():
                    last_err = e
                    time.sleep((2 ** attempt) + random.random())
                    continue
                last_err = e
                break
    raise RuntimeError(f"所有候選模型皆呼叫失敗：{last_err}")


# ============================================================
# PDF 合併
# ============================================================

def merge_person_pdfs(ordered_names: list, files_by_person: dict, workdir: str) -> bytes:
    if PdfWriter is None:
        raise RuntimeError("pypdf 套件未安裝，無法合併 PDF")

    writer = PdfWriter()
    page_cursor = 0

    for person in ordered_names:
        cats = files_by_person.get(person, {})
        person_start_page = page_cursor
        any_page = False

        for cat in CATEGORY_ORDER:
            filelist = sorted(cats.get(cat, []))
            for fp in filelist:
                pdf_path = convert_to_pdf(fp, workdir)
                if not pdf_path:
                    continue
                try:
                    reader = PdfReader(pdf_path)
                    for page in reader.pages:
                        writer.add_page(page)
                        page_cursor += 1
                        any_page = True
                except Exception:
                    continue

        if any_page:
            try:
                writer.add_outline_item(person, person_start_page)
            except Exception:
                pass

    buffer = io.BytesIO()
    writer.write(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# Excel 輸出
# ============================================================

def build_template_excel(df: pd.DataFrame) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Staffing"
    wb.calculation.fullCalcOnLoad = True

    header_font = Font(name="微軟正黑體", bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    body_font = Font(name="微軟正黑體")

    for col_idx, col_name in enumerate(TEMPLATE_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    safe_df = df.copy()
    for col in TEMPLATE_COLUMNS:
        if col not in safe_df.columns:
            safe_df[col] = ""
    safe_df = safe_df[TEMPLATE_COLUMNS]

    for row_idx, row in enumerate(safe_df.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = body_font
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    widths = [10, 14, 12, 10, 14, 20, 10, 16, 10, 22, 28, 28, 45]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def build_license_report_excel(df_license: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_license.to_excel(writer, index=False, sheet_name="證照效期檢核")
        wb = writer.book
        wb.calculation.fullCalcOnLoad = True
        ws = writer.sheets["證照效期檢核"]

        header_font = Font(name="微軟正黑體", bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill

        red_fill = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
        yellow_fill = PatternFill(start_color="FCF3CF", end_color="FCF3CF", fill_type="solid")

        status_col_idx = None
        for idx, cell in enumerate(ws[1], start=1):
            if cell.value == "狀態":
                status_col_idx = idx
                break

        if status_col_idx and ws.max_row >= 2:
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                val = row[status_col_idx - 1].value
                if val and "已過期" in str(val):
                    for c in row:
                        c.fill = red_fill
                elif val and "即將過期" in str(val):
                    for c in row:
                        c.fill = yellow_fill

        for column_cells in ws.columns:
            length = max((len(str(c.value)) if c.value else 0) for c in column_cells)
            col_letter = column_cells[0].column_letter
            ws.column_dimensions[col_letter].width = min(max(length + 2, 10), 40)

    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# 主流程（Phase 1：掃描與比對 / Phase 2：確認後執行）
# ============================================================

def scan_and_match_files(excel_file, zip_file) -> dict:
    """Phase 1：讀取人員名單、解壓 ZIP，並以三段式雙軌配對機制為每個檔案
    產生高/中/低信心度配對結果。不執行 AI 呼叫、PDF 合併等重工作，
    以便使用者能先確認/修正配對結果。"""
    df_people = pd.read_excel(excel_file)
    required_cols = ["順序", "姓名", "部門", "團隊職務"]
    missing = [c for c in required_cols if c not in df_people.columns]
    if missing:
        raise ValueError(f"人員名單 Excel 缺少必要欄位：{'、'.join(missing)}")
    if "英文姓名" not in df_people.columns:
        df_people["英文姓名"] = ""

    df_people = df_people.sort_values("順序").reset_index(drop=True)
    names = [str(n) for n in df_people["姓名"].tolist()]
    aliases = build_alias_list(df_people)

    workdir = tempfile.mkdtemp(prefix="staffing_")
    extract_dir = os.path.join(workdir, "extracted")
    convert_workdir = os.path.join(workdir, "converted")
    os.makedirs(extract_dir, exist_ok=True)
    os.makedirs(convert_workdir, exist_ok=True)

    with zipfile.ZipFile(zip_file) as z:
        z.extractall(extract_dir)

    all_files = []
    for root, _dirs, files in os.walk(extract_dir):
        root_lower = root.lower()
        if any(token.lower() in root_lower for token in EXCLUDE_DIR_TOKENS):
            continue
        for f in files:
            if f.startswith("."):
                continue
            all_files.append(os.path.join(root, f))

    high_matches = {}      # filepath -> (person, score, matched_alias_text)
    pending_matches = []   # 中信心度，待使用者確認
    low_matches = []       # 低信心度，待使用者手動指定

    for fp in all_files:
        fname = os.path.basename(fp)
        ranked = match_person_candidates(fname, aliases)
        if ranked:
            top_person, top_score, top_alias = ranked[0]
        else:
            top_person, top_score, top_alias = None, 0.0, ""

        if top_score >= HIGH_CONFIDENCE_THRESHOLD:
            high_matches[fp] = (top_person, top_score, top_alias)
        elif top_score >= MEDIUM_CONFIDENCE_THRESHOLD:
            pending_matches.append({
                "file": fp, "fname": fname,
                "candidates": ranked[:5], "default": top_person,
            })
        else:
            low_matches.append({
                "file": fp, "fname": fname, "candidates": ranked[:5],
            })

    return {
        "df_people": df_people,
        "names": names,
        "workdir": workdir,
        "extract_dir": extract_dir,
        "convert_workdir": convert_workdir,
        "high_matches": high_matches,
        "pending_matches": pending_matches,
        "low_matches": low_matches,
    }


def finalize_processing(scan: dict, pending_selection: dict, low_selection: dict,
                         api_key: str, model_name: str) -> dict:
    """Phase 2：整合使用者確認後的人員配對結果，以「內文導向」對每個檔案
    進行分類與 OCR 掃描（含進度條），執行證照效期檢核、AI 履歷結構化萃取，
    並產出合併附錄 PDF。"""
    df_people = scan["df_people"]
    names = scan["names"]
    convert_workdir = scan["convert_workdir"]

    final_mapping = {fp: info[0] for fp, info in scan["high_matches"].items()}
    for item in scan["pending_matches"]:
        person = pending_selection.get(item["file"])
        if person:
            final_mapping[item["file"]] = person

    still_unmatched = []
    for item in scan["low_matches"]:
        person = low_selection.get(item["file"])
        if person:
            final_mapping[item["file"]] = person
        else:
            still_unmatched.append(item["fname"])

    # ------------------------------------------------------------
    # Phase 2a：內文導向分類 + OCR 掃描（含動態進度條）
    # ------------------------------------------------------------
    files_by_person = {}
    license_rows = []

    file_items = list(final_mapping.items())
    total_files = max(len(file_items), 1)
    ocr_status = st.status("正在進行內文與 OCR 掃描...", expanded=True)
    ocr_progress = st.progress(0.0)

    for idx, (fp, person) in enumerate(file_items, start=1):
        fname = os.path.basename(fp)
        ocr_status.write(f"正在進行內文與 OCR 掃描：{fname} ({idx}/{total_files})")
        ocr_progress.progress(idx / total_files)

        category, text = classify_file_content(fp, convert_workdir)
        files_by_person.setdefault(person, {}).setdefault(category, []).append(fp)

        if category in ("4_技師執業執照", "5_技師會員證"):
            check = parse_license_expiry(category, text or "")
            license_rows.append({
                "姓名": person,
                "文件類別": "技師執業執照" if category == "4_技師執業執照" else "技師公會會員證",
                "檔名": fname,
                **check,
            })

    ocr_progress.progress(1.0)
    ocr_progress.empty()
    ocr_status.update(label=f"內文與 OCR 掃描完成（共 {total_files} 個檔案）", state="complete")

    # ------------------------------------------------------------
    # Phase 2b：CV 文字擷取 + Gemini AI 結構化萃取
    # ------------------------------------------------------------
    client = None
    model_chain = build_model_chain(model_name)
    if api_key and GENAI_AVAILABLE:
        try:
            client = genai.Client(api_key=api_key)
        except Exception:
            client = None

    staffing_rows = []
    no_cv_people = []

    total = max(len(df_people), 1)
    progress = st.progress(0.0, text="開始進行 AI 履歷結構化...")

    for i, prow in df_people.iterrows():
        name = str(prow["姓名"])
        role = str(prow.get("團隊職務", "") or "")
        group = str(prow.get("部門", "") or "")
        progress.progress(i / total, text=f"AI 履歷結構化：{name}")

        cats = files_by_person.get(name, {})

        # CV 文字擷取：PDF（含大寫 .PDF）與 Word（含大寫 .DOC/.DOCX）皆支援
        cv_files = cats.get("1_CV", [])
        cv_text = extract_cv_text(cv_files[0], convert_workdir) if cv_files else ""
        if not cv_files:
            no_cv_people.append(name)

        badge_hints = rule_based_badges(cats)

        ai_data = {}
        if client and cv_text.strip():
            try:
                ai_data, _used_model = gemini_extract_person(
                    client, model_chain, name, role, group, cv_text, ",".join(sorted(badge_hints))
                )
            except Exception as e:
                st.warning(f"⚠️ 「{name}」的 AI 履歷解析失敗，相關欄位將留空供手動填寫：{e}")
                ai_data = {}
        elif cv_files and not cv_text.strip():
            st.warning(f"⚠️ 「{name}」的 CV 檔案無法擷取到文字內容（可能為純掃描影像），"
                       "請確認 OCR 環境或改由手動填寫。")

        layer = determine_layer(role)
        badges_final = merge_and_sort_badges(ai_data.get("Badges", ""), badge_hints)

        staffing_rows.append({
            "Layer": layer,
            "Role": role,
            "GroupName": group if group and group.lower() != "nan" else "—",
            "Name": name,
            "Title": ai_data.get("Title", ""),
            "Company": ai_data.get("Company", ""),
            "Badges": badges_final,
            "PhotoName": f"{name}.jpg",
            "YearsOfExp": ai_data.get("YearsOfExp", 0) or 0,
            "Degree": ai_data.get("Degree", ""),
            "JobDescription": ai_data.get("JobDescription", ""),
            "Expertise": ai_data.get("Expertise", ""),
            "BioNarrative": ai_data.get("BioNarrative") or "[資料待補]",
        })

    progress.progress(1.0, text="完成")
    progress.empty()

    if no_cv_people:
        st.warning("⚠️ 以下人員找不到 CV 檔案，AI 結構化欄位留空，請手動補充：" + "、".join(no_cv_people))
    if not api_key:
        st.info("ℹ️ 未提供 Gemini API Key，CV 結構化欄位（Title/Company/JobDescription 等）將留空，"
                "您可於下方表格中手動填寫。")
    elif not client:
        st.warning("⚠️ Gemini Client 初始化失敗，請確認 API Key 是否正確，或 google-genai 套件是否已安裝。")

    df_staffing = pd.DataFrame(staffing_rows, columns=TEMPLATE_COLUMNS)
    df_license = pd.DataFrame(
        license_rows,
        columns=["姓名", "文件類別", "檔名", "狀態", "起始日", "截止日", "備註"],
    )

    merged_pdf_bytes = b""
    try:
        merged_pdf_bytes = merge_person_pdfs(names, files_by_person, convert_workdir)
    except Exception as e:
        st.error(f"附錄 PDF 合併發生錯誤：{e}")

    return {
        "df_staffing": df_staffing,
        "df_license": df_license,
        "merged_pdf": merged_pdf_bytes,
        "unmatched_files": still_unmatched,
    }


# ============================================================
# Streamlit 介面
# ============================================================

st.set_page_config(page_title="備標人員附錄整理、證照檢核與 CV 結構化轉檔系統", layout="wide")

st.title("📋 備標人員附錄整理、證照檢核與 CV 結構化轉檔系統")
st.caption("上傳人員名單 Excel 與證明文件 ZIP，自動排序合併附錄 PDF、檢核證照效期，"
           "並以 AI 產出可直接餵給組織圖產生器的 Template_Staffing.xlsx")

for key, default in [
    ("scan", None),
    ("pending_selection", {}),
    ("low_selection", {}),
    ("df_staffing", None),
    ("df_license", None),
    ("merged_pdf_bytes", None),
    ("unmatched_files", []),
]:
    if key not in st.session_state:
        st.session_state[key] = default

with st.sidebar:
    st.header("① 檔案上傳")
    excel_file = st.file_uploader(
        "人員名單 Excel（需含：順序／姓名／部門／團隊職務，建議另含「英文姓名」欄）",
        type=["xlsx"],
    )
    zip_file = st.file_uploader("人員證明文件 ZIP（CV / 學歷 / 證照 / 投保等 PDF・Word）", type=["zip"])

    st.header("② Gemini AI 設定")
    secrets_key = None
    try:
        secrets_key = st.secrets.get("GEMINI_API_KEY", None)
    except Exception:
        secrets_key = None

    if secrets_key:
        st.success("🔒 已從 Streamlit Secrets 自動讀取 API Key")
        api_key = secrets_key
    else:
        api_key = st.text_input("Gemini API Key（未設定 Secrets 時使用）", type="password")

    model_options = ["gemini-3.6-flash", "gemini-3.5-flash", "gemini-2.5-flash",
                      "gemini-flash-latest", "自訂模型"]
    selected_choice = st.selectbox("Gemini 模型選取", model_options)
    if selected_choice == "自訂模型":
        selected_model = st.text_input("自訂 Model ID", value="gemini-3.6-flash")
    else:
        selected_model = selected_choice

    if not GENAI_AVAILABLE:
        st.warning("⚠️ 尚未安裝 google-genai 套件，AI 結構化功能將無法使用。")

    st.divider()
    scan_btn = st.button("① 掃描並比對人員檔案", use_container_width=True)

# ------------------------------------------------------------
# Step 1：掃描並比對（三段式雙軌配對機制）
# ------------------------------------------------------------

if scan_btn:
    if not excel_file or not zip_file:
        st.error("請先於左側上傳人員名單 Excel 與證明文件 ZIP。")
    else:
        with st.spinner("掃描檔案並比對人員中，請稍候..."):
            try:
                scan = scan_and_match_files(excel_file, zip_file)
                st.session_state.scan = scan
                st.session_state.pending_selection = {
                    item["file"]: item["default"] for item in scan["pending_matches"]
                }
                st.session_state.low_selection = {
                    item["file"]: None for item in scan["low_matches"]
                }
                # 重置尚未產生的最終結果
                st.session_state.df_staffing = None
                st.session_state.df_license = None
                st.session_state.merged_pdf_bytes = None
                st.session_state.unmatched_files = []
                st.success(
                    f"✅ 掃描完成：高信心度自動配對 {len(scan['high_matches'])} 個檔案、"
                    f"待確認 {len(scan['pending_matches'])} 個、"
                    f"低信心度待手動指定 {len(scan['low_matches'])} 個。"
                )
            except Exception as e:
                st.error(f"❌ 掃描失敗：{e}")

scan = st.session_state.scan

if scan is not None:
    st.subheader("Step 1｜檔案配對結果")

    with st.expander(f"✅ 已自動配對（高信心度 ≥ {HIGH_CONFIDENCE_THRESHOLD:.0%}）："
                      f"{len(scan['high_matches'])} 個檔案", expanded=False):
        if scan["high_matches"]:
            high_rows = [
                {
                    "檔案": os.path.basename(fp),
                    "配對依據": alias_text,
                    "配對人員": person,
                    "信心度": f"{score:.0%}",
                }
                for fp, (person, score, alias_text) in scan["high_matches"].items()
            ]
            st.dataframe(pd.DataFrame(high_rows), use_container_width=True, hide_index=True)
            for row in high_rows:
                st.caption(f"已自動配對：{row['配對依據']} → {row['配對人員']}")
        else:
            st.write("（無）")

    if scan["pending_matches"]:
        st.markdown(f"### ⚠️ 待確認配對（中信心度 {MEDIUM_CONFIDENCE_THRESHOLD:.0%}～"
                    f"{HIGH_CONFIDENCE_THRESHOLD:.0%}）：{len(scan['pending_matches'])} 個檔案")
        st.caption("已預設選中最高機率的候選人員，請確認或修正後再進行 Step 2。")
        for item in scan["pending_matches"]:
            option_names = [c[0] for c in item["candidates"]]
            for n in scan["names"]:
                if n not in option_names:
                    option_names.append(n)
            options = ["— 略過此檔案 —"] + option_names
            default_person = st.session_state.pending_selection.get(item["file"], item["default"])
            default_idx = options.index(default_person) if default_person in options else 1
            score_hint = "、".join(f"{p}({s:.0%})" for p, s, _ in item["candidates"][:3])
            choice = st.selectbox(
                f"⚠️ {item['fname']}　－　候選：{score_hint}",
                options, index=default_idx, key=f"pending_{item['file']}",
            )
            st.session_state.pending_selection[item["file"]] = (
                None if choice == "— 略過此檔案 —" else choice
            )

    if scan["low_matches"]:
        with st.expander(f"❓ 低信心度（< {MEDIUM_CONFIDENCE_THRESHOLD:.0%}），"
                          f"請手動指定或忽略：{len(scan['low_matches'])} 個檔案", expanded=True):
            for item in scan["low_matches"]:
                options = ["— 不指定（略過）—"] + scan["names"]
                choice = st.selectbox(
                    item["fname"], options, index=0, key=f"low_{item['file']}",
                )
                st.session_state.low_selection[item["file"]] = (
                    None if choice.startswith("—") else choice
                )

    st.divider()
    process_btn = st.button("🚀 Step 2｜確認配對並開始處理", type="primary")

    if process_btn:
        with st.spinner("處理中，請稍候（含 PDF 轉檔、效期檢核與 AI 呼叫，可能需數分鐘）..."):
            try:
                result = finalize_processing(
                    scan, st.session_state.pending_selection, st.session_state.low_selection,
                    api_key, selected_model,
                )
                st.session_state.df_staffing = result["df_staffing"]
                st.session_state.df_license = result["df_license"]
                st.session_state.merged_pdf_bytes = result["merged_pdf"]
                st.session_state.unmatched_files = result["unmatched_files"]
                st.success("✅ 處理完成！請於下方分頁查看結果並下載檔案。")
            except Exception as e:
                st.error(f"❌ 處理失敗：{e}")
else:
    st.info("請先於左側上傳人員名單 Excel 與證明文件 ZIP，並點擊「① 掃描並比對人員檔案」。")

tab1, tab2, tab3 = st.tabs(["📜 附錄證照效期檢核", "🧾 CV 結構化預覽與編輯", "⬇️ 下載"])

with tab1:
    if st.session_state.df_license is not None and not st.session_state.df_license.empty:
        st.dataframe(st.session_state.df_license, use_container_width=True, hide_index=True)
        status_series = st.session_state.df_license["狀態"].astype(str)
        n_expired = status_series.str.contains("已過期").sum()
        n_soon = status_series.str.contains("即將過期").sum()
        n_unknown = status_series.str.contains("⚠️").sum() - n_expired - n_soon
        if n_expired > 0:
            st.error(f"🔴 共有 {n_expired} 筆證照已過期，請優先處理！")
        if n_soon > 0:
            st.warning(f"🟡 共有 {n_soon} 筆證照即將於 90 天內到期。")
        if n_unknown > 0:
            st.info(f"⚠️ 共有 {n_unknown} 筆證照無法自動判讀效期，請人工確認。")
    elif st.session_state.df_license is not None:
        st.info("未偵測到任何技師執業執照或會員證檔案。")
    else:
        st.info("請先於左側上傳檔案並依序完成 Step 1 掃描配對與 Step 2 開始處理。")

with tab2:
    if st.session_state.df_staffing is not None:
        st.caption("可直接於表格中編輯任一欄位，或於最下方新增協力廠商（Subcontractor）資料列。")
        edited_df = st.data_editor(
            st.session_state.df_staffing,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Layer": st.column_config.SelectboxColumn(options=LAYER_OPTIONS),
                "YearsOfExp": st.column_config.NumberColumn(min_value=0, step=1),
            },
            key="staffing_editor",
        )
        st.session_state.df_staffing = edited_df

        if st.session_state.unmatched_files:
            with st.expander(f"⚠️ 有 {len(st.session_state.unmatched_files)} 個檔案無法比對到人員名單"):
                for f in st.session_state.unmatched_files:
                    st.write(f"- {f}")
    else:
        st.info("請先於左側上傳檔案並依序完成 Step 1 掃描配對與 Step 2 開始處理。")

with tab3:
    if st.session_state.df_staffing is not None:
        col1, col2, col3 = st.columns(3)

        with col1:
            try:
                excel_bytes = build_template_excel(st.session_state.df_staffing)
                st.download_button(
                    "⬇️ 下載 Template_Staffing.xlsx",
                    data=excel_bytes,
                    file_name="Template_Staffing.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Excel 產生失敗：{e}")

        with col2:
            if st.session_state.merged_pdf_bytes:
                st.download_button(
                    "⬇️ 下載 附錄_人員資格證明檔_Merged.pdf",
                    data=st.session_state.merged_pdf_bytes,
                    file_name="附錄_人員資格證明檔_Merged.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            else:
                st.button("⬇️ 附錄 PDF（尚無資料）", disabled=True, use_container_width=True)

        with col3:
            try:
                license_excel = build_license_report_excel(
                    st.session_state.df_license if st.session_state.df_license is not None
                    else pd.DataFrame(columns=["姓名", "文件類別", "檔名", "狀態", "起始日", "截止日", "備註"])
                )
                st.download_button(
                    "⬇️ 下載 附錄人員資格與證照效期檢核報告.xlsx",
                    data=license_excel,
                    file_name="附錄人員資格與證照效期檢核報告.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"檢核報告產生失敗：{e}")
    else:
        st.info("請先於左側上傳檔案並依序完成 Step 1 掃描配對與 Step 2 開始處理。")

st.divider()
with st.expander("ℹ️ 使用說明與注意事項"):
    st.markdown(f"""
- **人員名單 Excel** 必須包含欄位：`順序`、`姓名`、`部門`、`團隊職務`；
  建議另增 `英文姓名`（如 `Joe Lim`、`RAY HSU`）以提升檔名比對準確率，
  未提供也不影響程式運作。
- **三段式雙軌配對機制**：以「姓名」與「英文姓名」為比對清單，比對前一律轉小寫、
  忽略大小寫，並移除 `_` `-` `()` `（）` 空格等符號差異，容許少量拼字誤差
  （如 Brain/Brian）：
  - 信心度 ≥ {HIGH_CONFIDENCE_THRESHOLD:.0%}：自動配對，於「已自動配對」清單顯示
    （如「已自動配對：Brian Lo → 羅翊軒」）。
  - 信心度 {MEDIUM_CONFIDENCE_THRESHOLD:.0%}～{HIGH_CONFIDENCE_THRESHOLD:.0%}：預設選中最高機率候選人，
    可於畫面上一鍵確認或修正。
  - 信心度 < {MEDIUM_CONFIDENCE_THRESHOLD:.0%}：需手動由下拉選單指定人員，或選擇略過。
- **內文導向（Content-Driven）分類**：
  - CV：檔名（不分大小寫）含「CV／履歷／學經歷／經歷表」，或副檔名為
    `.doc`／`.docx`（含大寫 `.DOC`）→ 直接判定為 CV，不需 OCR。
  - 技師執業執照：需先對 PDF／圖片進行文字擷取（含 OCR 備援），**內文**包含
    「技師執業執照」或「執業執照」才成立，不再依賴檔名。
  - 技師公會會員證：**內文同時**包含「會員證」與「技師公會」（如：台北市水利
    技師公會、台灣省水利技師公會）才成立，避免景觀學會等一般協會會員證被
    誤判。
  - 投保證明／學歷證明：內容結構化程度低，維持以檔名關鍵字判斷。
  - 其餘（一般協會會員證、電子證書等）：一律歸為一般證照，不進行過期告警。
- **內文與 OCR 掃描進度**：Step 2 開始處理時，會先以進度條與狀態訊息顯示
  「正在進行內文與 OCR 掃描：[檔名] (X/Y)」，掃描完成後才進入 AI 履歷結構化階段。
- **CV 文字擷取**：PDF（含大寫 `.PDF`，內建 OCR 備援）與 Word（含大寫
  `.DOC`／`.DOCX`）皆會完整擷取全文，並完整傳送給 Gemini AI 進行 13 欄位
  結構化萃取；輸出的 Template_Staffing.xlsx 維持原 13 欄標準格式。
- **合併 PDF** 會依 Excel 的「順序」欄排列人員，同一人內再依
  CV → 學歷 → 證照 → 執業執照 → 會員證 → 投保證明 排序，並加入書籤方便導覽。
- **證照效期檢核**：僅針對「技師執業執照」與「技師會員證」兩類進行，其餘證照
  類別不判斷過期；支援「執照有效期間：自民國X年X月X日至X年X月X日止」、
  「有效期限：民國X年X月X日至X年X月X日」，以及「115年會員證／115年度會員證」
  （自動推算至該年12月31日，如115+1911=2026年12月31日）等格式；日期一律
  以當下系統日期（`date.today()`）動態計算。
- 部署到 **Streamlit Community Cloud** 時，請將本工具一併產生的 `packages.txt`
  放在 repo 根目錄，以安裝 LibreOffice 與 Tesseract 等系統套件。
""")
