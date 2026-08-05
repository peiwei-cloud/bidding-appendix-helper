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

TODAY = date.today()

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
    "4_技師執業執照": ["執業執照", "执业执照"],
    "5_技師會員證": ["會員證", "会员证", "公會", "公会"],
    "6_投保證明": ["投保", "被保險人", "被保险人", "勞保", "劳保", "勞退"],
    "2_學歷": ["畢業", "毕业", "學歷", "学历", "學位", "学位"],
    "1_CV": ["cv", "履歷", "履历"],
}

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

def classify_category(filename: str) -> str:
    """依檔名判斷屬於 6 分類中的哪一類，找不到明確關鍵字則歸為一般證照。"""
    lower = filename.lower()
    ext = Path(filename).suffix.lower()

    for cat in ["4_技師執業執照", "5_技師會員證", "6_投保證明", "2_學歷"]:
        for kw in CATEGORY_KEYWORDS[cat]:
            if kw.lower() in lower:
                return cat

    if ext in (".doc", ".docx"):
        return "1_CV"
    if any(kw.lower() in lower for kw in CATEGORY_KEYWORDS["1_CV"]):
        return "1_CV"

    return "3_證照"


def match_person(filename: str, names: list) -> str:
    """在檔名中尋找最長匹配的人員姓名，避免姓名互為子字串造成誤判。"""
    base = Path(filename).stem
    base_clean = re.sub(r"[_\-\s（）()]", "", base)
    candidates = []
    for name in names:
        name_clean = re.sub(r"[_\-\s（）()]", "", str(name))
        if name_clean and name_clean in base_clean:
            candidates.append(name)
    if not candidates:
        return None
    candidates.sort(key=len, reverse=True)
    return candidates[0]


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


# ============================================================
# 證照效期檢核
# ============================================================

DATE_RANGE_PATTERN = re.compile(
    r"(?:自)?民國\s*(\d{2,3})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日\s*"
    r"(?:止|至)\s*(?:民國\s*)?(\d{2,3})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日"
)


def check_license_expiry(pdf_path: str) -> dict:
    if not pdf_path or not os.path.exists(pdf_path):
        return {"狀態": "⚠️ 無法轉檔判讀", "起始日": "", "截止日": "", "備註": "請人工確認"}

    text = extract_pdf_text(pdf_path)
    if not text.strip():
        return {"狀態": "⚠️ 無法自動判讀（掃描檔/OCR不可用）", "起始日": "", "截止日": "",
                "備註": "請人工確認"}

    m = DATE_RANGE_PATTERN.search(text)
    if not m:
        return {"狀態": "⚠️ 未偵測到效期文字", "起始日": "", "截止日": "", "備註": "請人工確認"}

    ry1, rm1, rd1, ry2, rm2, rd2 = map(int, m.groups())
    try:
        start = date(ry1 + 1911, rm1, rd1)
        end = date(ry2 + 1911, rm2, rd2)
    except ValueError:
        return {"狀態": "⚠️ 日期格式錯誤", "起始日": "", "截止日": "", "備註": "請人工確認"}

    days_left = (end - TODAY).days
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
# 主流程
# ============================================================

def process_all(excel_file, zip_file, api_key: str, model_name: str) -> dict:
    df_people = pd.read_excel(excel_file)
    required_cols = ["順序", "姓名", "部門", "團隊職務"]
    missing = [c for c in required_cols if c not in df_people.columns]
    if missing:
        raise ValueError(f"人員名單 Excel 缺少必要欄位：{'、'.join(missing)}")

    df_people = df_people.sort_values("順序").reset_index(drop=True)
    names = [str(n) for n in df_people["姓名"].tolist()]

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

    files_by_person = {}
    unmatched_files = []
    for fp in all_files:
        fname = os.path.basename(fp)
        person = match_person(fname, names)
        if not person:
            unmatched_files.append(fname)
            continue
        category = classify_category(fname)
        files_by_person.setdefault(person, {}).setdefault(category, []).append(fp)

    client = None
    model_chain = build_model_chain(model_name)
    if api_key and GENAI_AVAILABLE:
        try:
            client = genai.Client(api_key=api_key)
        except Exception:
            client = None

    staffing_rows = []
    license_rows = []

    total = max(len(df_people), 1)
    progress = st.progress(0.0, text="開始處理人員資料...")

    no_cv_people = []

    for i, prow in df_people.iterrows():
        name = str(prow["姓名"])
        role = str(prow.get("團隊職務", "") or "")
        group = str(prow.get("部門", "") or "")
        progress.progress(i / total, text=f"處理中：{name}")

        cats = files_by_person.get(name, {})

        for cat_key, cat_label in [("4_技師執業執照", "技師執業執照"),
                                     ("5_技師會員證", "技師公會會員證")]:
            for fp in cats.get(cat_key, []):
                pdf_path = fp if fp.lower().endswith(".pdf") else convert_to_pdf(fp, convert_workdir)
                check = check_license_expiry(pdf_path)
                license_rows.append({
                    "姓名": name,
                    "文件類別": cat_label,
                    "檔名": os.path.basename(fp),
                    **check,
                })

        cv_files = cats.get("1_CV", [])
        cv_text = extract_docx_text(cv_files[0], convert_workdir) if cv_files else ""
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
        "unmatched_files": unmatched_files,
    }


# ============================================================
# Streamlit 介面
# ============================================================

st.set_page_config(page_title="備標人員附錄整理、證照檢核與 CV 結構化轉檔系統", layout="wide")

st.title("📋 備標人員附錄整理、證照檢核與 CV 結構化轉檔系統")
st.caption("上傳人員名單 Excel 與證明文件 ZIP，自動排序合併附錄 PDF、檢核證照效期，"
           "並以 AI 產出可直接餵給組織圖產生器的 Template_Staffing.xlsx")

for key, default in [
    ("df_staffing", None),
    ("df_license", None),
    ("merged_pdf_bytes", None),
    ("unmatched_files", []),
]:
    if key not in st.session_state:
        st.session_state[key] = default

with st.sidebar:
    st.header("① 檔案上傳")
    excel_file = st.file_uploader("人員名單 Excel（需含：順序／姓名／部門／團隊職務）", type=["xlsx"])
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
    start_btn = st.button("🚀 開始處理", type="primary", use_container_width=True)

if start_btn:
    if not excel_file or not zip_file:
        st.error("請先於左側上傳人員名單 Excel 與證明文件 ZIP。")
    else:
        with st.spinner("處理中，請稍候（首次執行含 PDF 轉檔與 AI 呼叫，可能需數分鐘）..."):
            try:
                result = process_all(excel_file, zip_file, api_key, selected_model)
                st.session_state.df_staffing = result["df_staffing"]
                st.session_state.df_license = result["df_license"]
                st.session_state.merged_pdf_bytes = result["merged_pdf"]
                st.session_state.unmatched_files = result["unmatched_files"]
                st.success("✅ 處理完成！請於下方分頁查看結果並下載檔案。")
            except Exception as e:
                st.error(f"❌ 處理失敗：{e}")

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
        st.info("請先於左側上傳檔案並點擊「開始處理」。")

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
        st.info("請先於左側上傳檔案並點擊「開始處理」。")

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
        st.info("請先於左側上傳檔案並點擊「開始處理」。")

st.divider()
with st.expander("ℹ️ 使用說明與注意事項"):
    st.markdown("""
- **人員名單 Excel** 必須包含欄位：`順序`、`姓名`、`部門`、`團隊職務`。
- **證明文件 ZIP** 內的檔名需包含人員姓名（可含底線／空格／括號），系統會自動比對並分類：
  含「執業執照」→ 技師執業執照；含「會員證」或「公會」→ 技師會員證；
  含「投保」或「被保險人」→ 投保證明；含「畢業」或「學歷」→ 學歷證明；
  Word 檔或含「CV／履歷」→ CV；其餘 PDF 預設歸為一般證照。
- **合併 PDF** 會依 Excel 的「順序」欄排列人員，同一人內再依
  CV → 學歷 → 證照 → 執業執照 → 會員證 → 投保證明 排序，並加入書籤方便導覽。
- **證照效期檢核** 會嘗試從 PDF 文字層直接判讀民國年效期；若為掃描檔且環境已安裝
  Tesseract OCR，會自動 OCR 備援；仍無法判讀則標示為「請人工確認」。
- 部署到 **Streamlit Community Cloud** 時，請將本工具一併產生的 `packages.txt`
  放在 repo 根目錄，以安裝 LibreOffice 與 Tesseract 等系統套件。
""")
