import os
import re
import io
import json
import zipfile
import tempfile
from datetime import datetime
from typing import Dict, Any, List, Tuple, Optional

import streamlit as st
import pandas as pd
from PIL import Image


# ============================================================
# Page config + professional styling (Smallpdf-like clean UI)
# ============================================================
st.set_page_config(page_title="Pro Document Converter", layout="wide")

st.markdown(
    """
    <style>
      .block-container { padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1200px; }
      .muted { color: rgba(255,255,255,0.65); font-size: 0.92rem; }
      .card {
        border: 1px solid rgba(255,255,255,0.10);
        border-radius: 16px;
        padding: 14px 16px;
        background: rgba(255,255,255,0.03);
      }
      .kpi {
        font-size: 0.9rem;
        color: rgba(255,255,255,0.75);
      }
      .badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 999px;
        border: 1px solid rgba(255,255,255,0.14);
        background: rgba(255,255,255,0.04);
        font-size: 0.85rem;
        margin-right: 6px;
      }
      .divider { height: 1px; background: rgba(255,255,255,0.10); border: none; margin: 14px 0; }
      .step {
        display: inline-block;
        padding: 6px 10px;
        border-radius: 10px;
        border: 1px solid rgba(255,255,255,0.12);
        background: rgba(255,255,255,0.03);
        font-size: 0.9rem;
        margin-right: 8px;
      }
      .btnrow > div { display: inline-block; margin-right: 10px; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🧰 Pro Document Converter")
st.caption("Upload → Choose → Convert → Download • Web-only, fast, and reliable conversions (PDF / Image / Word / Excel / PPT).")


# ============================================================
# Helpers
# ============================================================
def now_stamp() -> str:
    return datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")


def now_file_stamp() -> str:
    return datetime.utcnow().strftime("%Y-%m-%d_%H%M%S_UTC")


def safe_filename(base: str) -> str:
    base = re.sub(r"[^A-Za-z0-9._-]+", "_", base).strip("_")
    return base[:120] if base else "output"


def build_zip(files: Dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as z:
        for path, data in files.items():
            z.writestr(path, data)
    return buf.getvalue()


def is_pdf(name: str) -> bool:
    return name.lower().endswith(".pdf")


def is_image(name: str) -> bool:
    return name.lower().endswith((".png", ".jpg", ".jpeg", ".webp"))


def is_docx(name: str) -> bool:
    return name.lower().endswith(".docx")


def is_xlsx(name: str) -> bool:
    return name.lower().endswith((".xlsx", ".xlsm"))


def is_pptx(name: str) -> bool:
    return name.lower().endswith(".pptx")


def detect_type(filename: str) -> str:
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".pdf":
        return "PDF"
    if ext in [".png", ".jpg", ".jpeg", ".webp"]:
        return "IMAGE"
    if ext == ".docx":
        return "WORD"
    if ext in [".xlsx", ".xlsm"]:
        return "EXCEL"
    if ext == ".pptx":
        return "PPT"
    return "UNKNOWN"


def to_displayable_image(obj):
    try:
        if isinstance(obj, Image.Image):
            return obj.convert("RGB")
    except Exception:
        pass
    try:
        im = obj.convert("RGB")
        if isinstance(im, Image.Image):
            return im
    except Exception:
        pass
    try:
        import numpy as np
        return np.array(obj)
    except Exception:
        return None


# ============================================================
# Text cleanup for OCR/table extraction
# ============================================================
def _collapse_spaces(s: str) -> str:
    return re.sub(r"[ \t]+", " ", s).strip()


def normalize_cell_text_raw(val):
    if val is None:
        return val
    s = str(val).replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s


def normalize_cell_text_clean(val):
    if val is None:
        return val
    s = str(val)

    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"\n+", "\n", s).replace("\n", " ")
    s = s.replace("\u00a0", " ")
    s = _collapse_spaces(s)

    s = re.sub(r"^\|\s*", "", s)

    def join_spaced_letters(m):
        return m.group(0).replace(" ", "")

    s = re.sub(r"(?:\b[A-Za-z]\b(?:\s+|$)){4,}", join_spaced_letters, s)
    s = re.sub(r"(?:\b\d\b\s+){3,}\b\d\b", lambda m: m.group(0).replace(" ", ""), s)

    for _ in range(2):
        s = re.sub(r"\b([A-Za-z])\s+([A-Za-z]{2,})\b", r"\1\2", s)

    s = re.sub(r"\s*([,/:\.\-\+])\s*", r"\1", s)
    s = re.sub(r"\b(GB(?:/T)?)\s*([0-9])", r"\1 \2", s, flags=re.IGNORECASE)

    return _collapse_spaces(s)


# ============================================================
# PDF / OCR / Tables core logic (same as before, organized)
# ============================================================
def pdf_metadata_to_dict(pdf_bytes: bytes) -> Dict[str, Any]:
    from pypdf import PdfReader
    r = PdfReader(io.BytesIO(pdf_bytes))
    md = r.metadata or {}
    out = {"page_count": len(r.pages)}
    for k, v in md.items():
        out[str(k)] = str(v) if v is not None else None
    return out


def pdf_textlayer_extract(pdf_bytes: bytes, max_pages: int) -> List[str]:
    import pdfplumber
    texts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for p in pdf.pages[:max_pages]:
            texts.append(p.extract_text() or "")
    return texts


def pdf_render_pages_to_images(pdf_bytes: bytes, dpi: int, max_pages: int):
    from pdf2image import convert_from_bytes
    return convert_from_bytes(pdf_bytes, dpi=dpi)[:max_pages]


def ocr_image_to_text(pil_img: Image.Image, lang: str = "eng") -> str:
    import pytesseract
    return (pytesseract.image_to_string(pil_img, lang=lang) or "").strip()


def pdf_hybrid_text_extract(pdf_bytes: bytes, max_pages: int, lang: str, dpi: int = 260) -> List[str]:
    layer_texts = pdf_textlayer_extract(pdf_bytes, max_pages=max_pages)

    needs_ocr = []
    for t in layer_texts:
        t2 = re.sub(r"\s+", "", t or "")
        needs_ocr.append(len(t2) < 40)

    if not any(needs_ocr):
        return layer_texts

    images = pdf_render_pages_to_images(pdf_bytes, dpi=dpi, max_pages=max_pages)
    out = []
    for i, base_text in enumerate(layer_texts):
        if i < len(images) and needs_ocr[i]:
            try:
                ocr_t = ocr_image_to_text(images[i], lang=lang)
                out.append(ocr_t if ocr_t else base_text)
            except Exception:
                out.append(base_text)
        else:
            out.append(base_text)
    return out


def extract_tables_pdf_textlayer(pdf_bytes: bytes, max_pages: int) -> List[pd.DataFrame]:
    import pdfplumber
    dfs: List[pd.DataFrame] = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages[:max_pages]:
            tbls = page.extract_tables()
            if not tbls:
                continue
            for t in tbls:
                if t:
                    dfs.append(pd.DataFrame(t))
    return dfs


def flatten_img2table_tables(tables_obj) -> List:
    if tables_obj is None:
        return []
    if isinstance(tables_obj, list):
        return tables_obj
    if isinstance(tables_obj, dict):
        out = []
        for _, v in tables_obj.items():
            if isinstance(v, list):
                out.extend(v)
        return out
    return []


def table_to_df_safe(table) -> Optional[pd.DataFrame]:
    if table is None:
        return None
    if isinstance(table, pd.DataFrame):
        return table
    if hasattr(table, "df"):
        try:
            df = table.df
            if isinstance(df, pd.DataFrame):
                return df
        except Exception:
            return None
    return None


def df_to_json_records(df: pd.DataFrame) -> List[Dict[str, Any]]:
    df2 = df.copy()

    cols = []
    for i, c in enumerate(df2.columns):
        name = str(c).strip() if c is not None else ""
        if not name or name.lower() in {"nan", "none"}:
            name = f"col_{i+1}"
        cols.append(name)
    df2.columns = cols

    seen = {}
    new_cols = []
    for c in df2.columns:
        if c not in seen:
            seen[c] = 1
            new_cols.append(c)
        else:
            seen[c] += 1
            new_cols.append(f"{c}_{seen[c]}")
    df2.columns = new_cols

    df2 = df2.where(pd.notnull(df2), None)
    return df2.to_dict(orient="records")


def build_tables_bundle(tables: List[pd.DataFrame], normalizer, base_root: str) -> Dict[str, bytes]:
    from openpyxl.styles import Alignment

    cleaned = [df.applymap(normalizer) for df in tables]
    files: Dict[str, bytes] = {}

    # Excel multi-sheet
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        for i, df in enumerate(cleaned, start=1):
            df.to_excel(writer, sheet_name=f"Table_{i}"[:31], index=False)
        wb = writer.book
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=False, vertical="top")
    files[f"{base_root}.xlsx"] = excel_buf.getvalue()

    combined_csv_parts = []
    combined_json = {"tables": []}

    for i, df in enumerate(cleaned, start=1):
        files[f"{base_root}/tables/table_{i}.csv"] = df.to_csv(index=False).encode("utf-8")
        combined_csv_parts.append(f"# --- Table {i} ---\n")
        combined_csv_parts.append(df.to_csv(index=False))

        one = {"table_index": i, "rows": df_to_json_records(df)}
        files[f"{base_root}/tables/table_{i}.json"] = json.dumps(one, ensure_ascii=False, indent=2).encode("utf-8")
        combined_json["tables"].append(one)

    files[f"{base_root}/tables/combined.csv"] = "".join(combined_csv_parts).encode("utf-8")
    files[f"{base_root}/tables/combined.json"] = json.dumps(combined_json, ensure_ascii=False, indent=2).encode("utf-8")

    files[f"{base_root}/manifest.json"] = json.dumps(
        {"type": "tables_export", "table_count": len(cleaned), "created": now_stamp()},
        ensure_ascii=False,
        indent=2
    ).encode("utf-8")

    return files


def pdf_to_images_zip(pdf_bytes: bytes, max_pages: int, dpi: int = 220) -> Tuple[bytes, int]:
    images = pdf_render_pages_to_images(pdf_bytes, dpi=dpi, max_pages=max_pages)
    files = {}
    for i, im in enumerate(images, start=1):
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        files[f"pdf_pages/page_{i:03d}.png"] = buf.getvalue()

    files["manifest.json"] = json.dumps(
        {"type": "pdf_to_images", "page_count": len(images), "dpi": dpi, "created": now_stamp()},
        indent=2
    ).encode("utf-8")

    return build_zip(files), len(images)


# ============================================================
# Office conversions: Excel <-> Word + more outputs
# ============================================================
def text_to_pdf_bytes(title: str, text: str) -> bytes:
    # Web-safe text-only PDF (readable text, selectable)
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    margin = 40
    y = height - margin
    c.setFont("Helvetica-Bold", 14)
    c.drawString(margin, y, title[:95])
    y -= 24

    c.setFont("Helvetica", 10)
    for line in (text or "").splitlines():
        line = line.rstrip()

        while len(line) > 120:
            c.drawString(margin, y, line[:120])
            line = line[120:]
            y -= 14
            if y < margin:
                c.showPage()
                c.setFont("Helvetica", 10)
                y = height - margin

        c.drawString(margin, y, line)
        y -= 14
        if y < margin:
            c.showPage()
            c.setFont("Helvetica", 10)
            y = height - margin

    c.save()
    return buf.getvalue()


def excel_to_pdf_bytes(xlsx_bytes: bytes, max_rows: int = 80, max_cols: int = 12) -> bytes:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet

    styles = getSampleStyleSheet()
    story = []

    xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    for si, sheet in enumerate(xls.sheet_names, start=1):
        df = xls.parse(sheet).copy()
        df = df.iloc[:max_rows, :max_cols]

        story.append(Paragraph(f"Sheet: {sheet}", styles["Heading2"]))
        story.append(Paragraph(f"Exported: {now_stamp()}", styles["Normal"]))
        story.append(Spacer(1, 10))

        data = [list(df.columns.astype(str))] + df.astype(str).values.tolist()
        tbl = Table(data, repeatRows=1)

        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2a2a2a")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
        ]))

        story.append(tbl)
        story.append(Spacer(1, 14))
        if si < len(xls.sheet_names):
            story.append(PageBreak())

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), rightMargin=18, leftMargin=18, topMargin=18, bottomMargin=18)
    doc.build(story)
    return buf.getvalue()


def excel_to_word_docx(xlsx_bytes: bytes, max_rows: int = 60, max_cols: int = 12) -> bytes:
    # Excel -> Word (tables). Web-safe (truncates large sheets)
    from docx import Document

    xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    doc = Document()
    doc.add_heading("Excel to Word Export", level=1)
    doc.add_paragraph(f"Created: {now_stamp()}")

    for sheet in xls.sheet_names:
        df = xls.parse(sheet).copy()
        df = df.iloc[:max_rows, :max_cols]

        doc.add_heading(f"Sheet: {sheet}", level=2)

        # Create table with headers
        rows = df.shape[0] + 1
        cols = df.shape[1] if df.shape[1] > 0 else 1
        table = doc.add_table(rows=rows, cols=cols)
        table.style = "Table Grid"

        # headers
        for j, col in enumerate(df.columns.astype(str).tolist()):
            table.cell(0, j).text = col

        # data
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                table.cell(i + 1, j).text = "" if pd.isna(df.iat[i, j]) else str(df.iat[i, j])

        doc.add_paragraph("")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def word_to_excel_tables(docx_bytes: bytes) -> bytes:
    # Word -> Excel (extract tables only)
    from docx import Document

    doc = Document(io.BytesIO(docx_bytes))
    tables = doc.tables

    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        if not tables:
            # create an empty sheet with note
            pd.DataFrame([{"note": "No tables found in DOCX."}]).to_excel(writer, sheet_name="Info", index=False)
        else:
            for idx, t in enumerate(tables, start=1):
                rows = []
                for row in t.rows:
                    rows.append([cell.text.strip() for cell in row.cells])
                df = pd.DataFrame(rows)
                df.to_excel(writer, sheet_name=f"Table_{idx}"[:31], index=False, header=False)

    return excel_buf.getvalue()


def docx_to_text_html_md(docx_bytes: bytes) -> Tuple[str, str, str]:
    from docx import Document
    from markdownify import markdownify as mdify

    doc = Document(io.BytesIO(docx_bytes))
    paragraphs = [p.text for p in doc.paragraphs if p.text and p.text.strip()]
    text = "\n".join(paragraphs).strip()

    html_lines = []
    for p in paragraphs:
        p2 = (p.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
        html_lines.append(f"<p>{p2}</p>")
    html = "\n".join(html_lines).strip()

    md = mdify(html) if html else (text or "")
    return text or "", html or "", md or ""


def xlsx_to_outputs_bundle(xlsx_bytes: bytes) -> Dict[str, bytes]:
    # Excel -> CSV/JSON/HTML/MD per sheet + combined JSON + PDF + Word + manifest
    from markdownify import markdownify as mdify

    xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    files: Dict[str, bytes] = {}
    combined = {"sheets": []}

    for sheet in xls.sheet_names:
        df = xls.parse(sheet)
        safe = safe_filename(sheet) or "Sheet"

        files[f"excel/sheets/{safe}.csv"] = df.to_csv(index=False).encode("utf-8")
        records = df_to_json_records(df)
        files[f"excel/sheets/{safe}.json"] = json.dumps({"sheet": sheet, "rows": records}, ensure_ascii=False, indent=2).encode("utf-8")

        html = df.to_html(index=False)
        files[f"excel/sheets/{safe}.html"] = html.encode("utf-8")
        files[f"excel/sheets/{safe}.md"] = mdify(html).encode("utf-8")

        combined["sheets"].append({"sheet": sheet, "rows": records})

    files["excel/combined.json"] = json.dumps(combined, ensure_ascii=False, indent=2).encode("utf-8")
    files["excel/output.pdf"] = excel_to_pdf_bytes(xlsx_bytes)
    files["excel/output.docx"] = excel_to_word_docx(xlsx_bytes)

    files["excel/manifest.json"] = json.dumps(
        {"type": "excel_bundle", "sheet_count": len(xls.sheet_names), "created": now_stamp()},
        ensure_ascii=False, indent=2
    ).encode("utf-8")

    return files


def pptx_to_text_json_images(pptx_bytes: bytes) -> Dict[str, bytes]:
    from pptx import Presentation

    prs = Presentation(io.BytesIO(pptx_bytes))
    files: Dict[str, bytes] = {}
    all_text = []
    slides_json = []
    img_count = 0

    for si, slide in enumerate(prs.slides, start=1):
        slide_text_parts = []
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text:
                t = shape.text.strip()
                if t:
                    slide_text_parts.append(t)
            # extract embedded images
            try:
                if shape.shape_type == 13:  # PICTURE
                    image = shape.image
                    img_bytes = image.blob
                    ext = (image.ext or "bin").lower()
                    img_count += 1
                    files[f"ppt/images/slide_{si:03d}_img_{img_count:03d}.{ext}"] = img_bytes
            except Exception:
                pass

        slide_text = "\n".join(slide_text_parts).strip()
        all_text.append(f"--- Slide {si} ---\n{slide_text}".strip())
        slides_json.append({"slide": si, "text": slide_text})

    files["ppt/slides.txt"] = ("\n\n".join(all_text).strip() + "\n").encode("utf-8")
    files["ppt/slides.json"] = json.dumps({"slides": slides_json}, ensure_ascii=False, indent=2).encode("utf-8")
    files["ppt/slides.pdf"] = text_to_pdf_bytes("PPT Text Export", files["ppt/slides.txt"].decode("utf-8", errors="ignore"))

    files["ppt/manifest.json"] = json.dumps(
        {"type": "ppt_bundle", "slide_count": len(prs.slides), "extracted_images": img_count, "created": now_stamp()},
        ensure_ascii=False, indent=2
    ).encode("utf-8")

    return files


# ============================================================
# Optional Searchable PDF (OCR layer) – requires ocrmypdf + OS deps
# ============================================================
def ocrmypdf_available() -> bool:
    try:
        import ocrmypdf  # noqa
        return True
    except Exception:
        return False


def make_searchable_pdf_from_pdf(pdf_bytes: bytes, lang: str = "eng") -> bytes:
    import ocrmypdf

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as f_in:
        f_in.write(pdf_bytes)
        in_path = f_in.name
    out_path = in_path.replace(".pdf", "_ocr.pdf")

    try:
        ocrmypdf.ocr(
            in_path,
            out_path,
            language=lang,
            skip_text=True,
            force_ocr=True,
            optimize=1,
            deskew=True,
        )
        with open(out_path, "rb") as f:
            return f.read()
    finally:
        for p in [in_path, out_path]:
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass


# ============================================================
# Image -> PDF with readable text (OCR Text PDF)
#   - Creates a PDF containing OCR text as selectable text (no image background).
#   - For true "searchable PDF over image", use ocrmypdf optional path.
# ============================================================
def image_to_ocr_text_pdf(img_bytes: bytes, lang: str = "eng") -> bytes:
    img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    text = ocr_image_to_text(img, lang=lang)
    title = "Image OCR to PDF (Selectable Text)"
    return text_to_pdf_bytes(title, text)


# ============================================================
# Session state: history & last outputs
# ============================================================
if "history" not in st.session_state:
    st.session_state.history = []  # list of dicts: time/task/input/outputs

if "last_outputs" not in st.session_state:
    st.session_state.last_outputs = {}  # filename -> bytes


def add_history(entry: Dict[str, Any]):
    st.session_state.history.insert(0, entry)
    st.session_state.history = st.session_state.history[:30]


# ============================================================
# Sidebar: Advanced processing options (renamed professionally)
# ============================================================
with st.sidebar:
    st.markdown("## ⚙️ Processing Options")
    st.caption("Adjust only if results need improvement.")

    clean_mode = st.selectbox("Text cleanup", ["Clean (Recommended)", "Raw"], index=0)
    prefer_text_layer_tables = st.toggle("Prefer PDF text-layer for tables", value=True)

    st.markdown("---")
    st.markdown("### OCR")
    ocr_lang = st.selectbox("Language", ["eng"], index=0)
    max_pages = st.slider("Max pages", 1, 60, 12)
    ocr_dpi = st.slider("OCR quality (DPI)", 200, 400, 260, step=10)

    st.markdown("---")
    st.markdown("### Table Detection")
    min_conf = st.slider("Confidence threshold", 0, 100, 50)

    st.markdown("---")
    st.markdown("### History")
    if st.session_state.history:
        with st.expander("Recent conversions", expanded=False):
            for h in st.session_state.history[:10]:
                st.markdown(
                    f"- **{h['time']}** • `{h['input']}` • **{h['task']}**"
                )
    else:
        st.caption("No conversions yet.")

normalizer = normalize_cell_text_clean if clean_mode.startswith("Clean") else normalize_cell_text_raw


# ============================================================
# Guided workflow UI
# ============================================================
st.markdown(
    """
    <div class="card">
      <span class="step">1) Upload</span>
      <span class="step">2) Choose</span>
      <span class="step">3) Convert</span>
      <span class="step">4) Download</span>
      <div class="muted" style="margin-top:8px;">
        The converter auto-detects your file type and shows only valid targets.
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# Step 1: Upload
uploaded = st.file_uploader(
    "Upload a file",
    type=["pdf", "png", "jpg", "jpeg", "webp", "docx", "xlsx", "xlsm", "pptx"],
    help="Supported: PDF, Images, Word (.docx), Excel (.xlsx/.xlsm), PowerPoint (.pptx)",
)

if not uploaded:
    st.info("Upload a file to see available conversions.")
    st.stop()

file_bytes = uploaded.read()
filename = uploaded.name
base_name = safe_filename(os.path.splitext(filename)[0])
kind = detect_type(filename)

# Top summary cards
c1, c2, c3 = st.columns([2.0, 1.0, 1.0], gap="medium")
with c1:
    st.markdown(
        f"<div class='card'><h4 style='margin:0 0 6px 0;'>Detected</h4>"
        f"<span class='badge'>{kind}</span> <span class='badge'>{filename}</span>"
        f"<div class='muted'>Ready for conversion • {now_stamp()}</div></div>",
        unsafe_allow_html=True
    )
with c2:
    st.markdown(
        f"<div class='card'><h4 style='margin:0 0 6px 0;'>Quality</h4>"
        f"<div class='kpi'>Cleanup: <b>{clean_mode}</b><br/>OCR DPI: <b>{ocr_dpi}</b></div></div>",
        unsafe_allow_html=True
    )
with c3:
    st.markdown(
        f"<div class='card'><h4 style='margin:0 0 6px 0;'>Tip</h4>"
        f"<div class='kpi'>Use <b>ZIP</b> for batch outputs.<br/>Use <b>Hybrid OCR</b> for scanned PDFs.</div></div>",
        unsafe_allow_html=True
    )

st.markdown('<hr class="divider">', unsafe_allow_html=True)

# Preview + Choice + Output
left, right = st.columns([1.0, 1.1], gap="large")

with left:
    st.markdown("<div class='card'><h4 style='margin:0 0 10px 0;'>Preview</h4>", unsafe_allow_html=True)
    if kind == "IMAGE":
        try:
            im = Image.open(io.BytesIO(file_bytes)).convert("RGB")
            st.image(im, use_container_width=True)
        except Exception:
            st.caption("Preview not available for this image.")
    else:
        st.caption("Preview is shown for images. For PDFs, use conversions to view results.")
    st.markdown("</div>", unsafe_allow_html=True)

# Define conversions (auto-filtered by detected type)
# Each conversion returns a dict: {filename: bytes}
CONVERSIONS: Dict[str, List[Dict[str, Any]]] = {
    "PDF": [
        {"id": "pdf_to_text_bundle", "label": "Convert to Text (Hybrid OCR)", "targets": ["TXT", "HTML", "MD", "ZIP"]},
        {"id": "pdf_to_word", "label": "Convert to Word (Editable Text)", "targets": ["DOCX"]},
        {"id": "pdf_tables_to_data", "label": "Extract Tables to Spreadsheet", "targets": ["XLSX", "CSV", "JSON", "ZIP"]},
        {"id": "pdf_to_images", "label": "Convert Pages to Images", "targets": ["ZIP(PNG)"]},
        {"id": "pdf_metadata", "label": "Extract PDF Metadata", "targets": ["JSON"]},
        {"id": "pdf_searchable_optional", "label": "Make Searchable PDF (OCR Layer) (Optional)", "targets": ["PDF(Searchable)"]},
    ],
    "IMAGE": [
        {"id": "img_to_text", "label": "Convert to Text (OCR)", "targets": ["TXT"]},
        {"id": "img_tables_to_data", "label": "Extract Tables to Spreadsheet", "targets": ["XLSX", "CSV", "JSON", "ZIP"]},
        {"id": "img_to_pdf", "label": "Convert to PDF (Image only)", "targets": ["PDF"]},
        {"id": "img_to_pdf_ocr_text", "label": "Convert to PDF (Readable Text via OCR)", "targets": ["PDF(Text)"]},
    ],
    "EXCEL": [
        {"id": "xlsx_to_bundle", "label": "Convert Excel to Other Formats (Bundle)", "targets": ["ZIP", "PDF", "DOCX", "CSV", "JSON", "HTML", "MD"]},
        {"id": "xlsx_to_pdf", "label": "Convert Excel to PDF", "targets": ["PDF"]},
        {"id": "xlsx_to_word", "label": "Convert Excel to Word", "targets": ["DOCX"]},
    ],
    "WORD": [
        {"id": "docx_to_text_bundle", "label": "Convert Word to Text / Web Formats (Bundle)", "targets": ["ZIP", "TXT", "HTML", "MD", "PDF(Text)"]},
        {"id": "docx_to_excel", "label": "Convert Word Tables to Excel", "targets": ["XLSX"]},
    ],
    "PPT": [
        {"id": "pptx_to_bundle", "label": "Convert PowerPoint to Text / Data (Bundle)", "targets": ["ZIP", "TXT", "JSON", "PDF(Text)", "Images(embedded)"]},
    ],
}

available = CONVERSIONS.get(kind, [])

with right:
    st.markdown("<div class='card'><h4 style='margin:0 0 10px 0;'>2) Choose conversion</h4>", unsafe_allow_html=True)

    if not available:
        st.error("This file type is not supported.")
        st.stop()

    # Conversion dropdown
    conv_id = st.selectbox(
        "What do you want to do?",
        options=[c["id"] for c in available],
        format_func=lambda cid: next(x["label"] for x in available if x["id"] == cid),
    )
    chosen = next(x for x in available if x["id"] == conv_id)

    # Target dropdown (for professional feel)
    target = st.selectbox("Convert to", options=chosen["targets"])

    st.caption("3) Click Convert — downloads will appear below.")
    run = st.button("Convert", type="primary")

    st.markdown("</div>", unsafe_allow_html=True)

if not run:
    st.stop()

# ============================================================
# Conversion execution (logic untouched; UI orchestrates)
# ============================================================
outputs: Dict[str, bytes] = {}
task_label = next(x["label"] for x in available if x["id"] == conv_id)

with st.spinner("Converting..."):
    try:
        # ---------------- PDF ----------------
        if conv_id == "pdf_to_text_bundle":
            pages = pdf_hybrid_text_extract(file_bytes, max_pages=max_pages, lang=ocr_lang, dpi=ocr_dpi)
            txt = "\n\n".join([f"--- Page {i+1} ---\n{p}".strip() for i, p in enumerate(pages)]).strip() + "\n"
            html = "\n".join(
                [f"<h2>Page {i+1}</h2>\n<pre>{(p or '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')}</pre>"
                 for i, p in enumerate(pages)]
            )
            try:
                from markdownify import markdownify as mdify
                md = mdify(html)
            except Exception:
                md = txt

            if target == "TXT":
                outputs[f"{base_name}.txt"] = txt.encode("utf-8")
            elif target == "HTML":
                outputs[f"{base_name}.html"] = html.encode("utf-8")
            elif target == "MD":
                outputs[f"{base_name}.md"] = (md or "").encode("utf-8")
            else:  # ZIP
                files = {
                    "text/output.txt": txt.encode("utf-8"),
                    "text/output.html": html.encode("utf-8"),
                    "text/output.md": (md or "").encode("utf-8"),
                    "text/manifest.json": json.dumps({"type": "pdf_hybrid_text", "pages": len(pages), "created": now_stamp()}, indent=2).encode("utf-8"),
                }
                outputs[f"{base_name}_text_{now_file_stamp()}.zip"] = build_zip(files)

        elif conv_id == "pdf_to_word":
            from docx import Document
            pages = pdf_hybrid_text_extract(file_bytes, max_pages=max_pages, lang=ocr_lang, dpi=ocr_dpi)

            doc = Document()
            doc.add_heading("PDF to Word (Hybrid OCR)", level=1)
            doc.add_paragraph(f"Source: {filename}")
            doc.add_paragraph(f"Created: {now_stamp()}")
            doc.add_paragraph("")

            for i, p in enumerate(pages, start=1):
                doc.add_heading(f"Page {i}", level=2)
                for line in (p or "").splitlines():
                    doc.add_paragraph(line)

            buf = io.BytesIO()
            doc.save(buf)
            outputs[f"{base_name}.docx"] = buf.getvalue()

        elif conv_id == "pdf_tables_to_data":
            tables_dfs: List[pd.DataFrame] = []

            if prefer_text_layer_tables:
                try:
                    tables_dfs = extract_tables_pdf_textlayer(file_bytes, max_pages=max_pages)
                except Exception:
                    tables_dfs = []

            if not tables_dfs:
                from img2table.ocr import TesseractOCR
                from img2table.document import PDF as Img2TablePDF

                ocr = TesseractOCR(lang=ocr_lang)
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as f:
                    f.write(file_bytes)
                    path = f.name
                try:
                    doc = Img2TablePDF(path)
                    tables_obj = doc.extract_tables(
                        ocr=ocr,
                        borderless_tables=True,
                        implicit_rows=True,
                        min_confidence=min_conf,
                    )
                finally:
                    try:
                        os.remove(path)
                    except Exception:
                        pass

                for t in flatten_img2table_tables(tables_obj):
                    df = table_to_df_safe(t)
                    if df is not None:
                        tables_dfs.append(df)

            if not tables_dfs:
                raise RuntimeError("No tables could be extracted from this PDF.")

            root = f"{base_name}_tables_{now_file_stamp()}"
            bundle = build_tables_bundle(tables_dfs, normalizer, base_root=root)

            if target == "XLSX":
                outputs[f"{root}.xlsx"] = bundle[f"{root}.xlsx"]
            elif target == "CSV":
                outputs[f"{root}_combined.csv"] = bundle[f"{root}/tables/combined.csv"]
            elif target == "JSON":
                outputs[f"{root}_combined.json"] = bundle[f"{root}/tables/combined.json"]
            else:  # ZIP
                outputs[f"{root}.zip"] = build_zip(bundle)

        elif conv_id == "pdf_to_images":
            zip_bytes, count = pdf_to_images_zip(file_bytes, max_pages=max_pages, dpi=220)
            outputs[f"{base_name}_pages_{now_file_stamp()}.zip"] = zip_bytes

        elif conv_id == "pdf_metadata":
            md = pdf_metadata_to_dict(file_bytes)
            outputs[f"{base_name}_metadata.json"] = json.dumps(md, ensure_ascii=False, indent=2).encode("utf-8")

        elif conv_id == "pdf_searchable_optional":
            if not ocrmypdf_available():
                raise RuntimeError("Searchable PDF not available. Install ocrmypdf + system packages ghostscript & qpdf.")
            out_pdf = make_searchable_pdf_from_pdf(file_bytes, lang=ocr_lang)
            outputs[f"{base_name}_searchable.pdf"] = out_pdf

        # ---------------- IMAGE ----------------
        elif conv_id == "img_to_text":
            im = Image.open(io.BytesIO(file_bytes)).convert("RGB")
            text = ocr_image_to_text(im, lang=ocr_lang)
            outputs[f"{base_name}.txt"] = (text or "").encode("utf-8")

        elif conv_id == "img_tables_to_data":
            from img2table.ocr import TesseractOCR
            from img2table.document import Image as Img2TableImage

            ocr = TesseractOCR(lang=ocr_lang)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as f:
                f.write(file_bytes)
                path = f.name
            try:
                doc = Img2TableImage(path)
                tables_obj = doc.extract_tables(
                    ocr=ocr,
                    borderless_tables=True,
                    implicit_rows=True,
                    min_confidence=min_conf,
                )
            finally:
                try:
                    os.remove(path)
                except Exception:
                    pass

            tables_dfs = []
            for t in flatten_img2table_tables(tables_obj):
                df = table_to_df_safe(t)
                if df is not None:
                    tables_dfs.append(df)

            if not tables_dfs:
                raise RuntimeError("No tables could be extracted from this image.")

            root = f"{base_name}_tables_{now_file_stamp()}"
            bundle = build_tables_bundle(tables_dfs, normalizer, base_root=root)

            if target == "XLSX":
                outputs[f"{root}.xlsx"] = bundle[f"{root}.xlsx"]
            elif target == "CSV":
                outputs[f"{root}_combined.csv"] = bundle[f"{root}/tables/combined.csv"]
            elif target == "JSON":
                outputs[f"{root}_combined.json"] = bundle[f"{root}/tables/combined.json"]
            else:
                outputs[f"{root}.zip"] = build_zip(bundle)

        elif conv_id == "img_to_pdf":
            img = Image.open(io.BytesIO(file_bytes)).convert("RGB")
            buf = io.BytesIO()
            img.save(buf, format="PDF")
            outputs[f"{base_name}.pdf"] = buf.getvalue()

        elif conv_id == "img_to_pdf_ocr_text":
            outputs[f"{base_name}_ocr_text.pdf"] = image_to_ocr_text_pdf(file_bytes, lang=ocr_lang)

        # ---------------- EXCEL ----------------
        elif conv_id == "xlsx_to_bundle":
            bundle = xlsx_to_outputs_bundle(file_bytes)

            # Choose a single output or bundle
            if target == "PDF":
                outputs[f"{base_name}.pdf"] = bundle["excel/output.pdf"]
            elif target == "DOCX":
                outputs[f"{base_name}.docx"] = bundle["excel/output.docx"]
            elif target == "CSV":
                # Provide combined zip-like CSV output (per sheet needs zip; here give zip)
                outputs[f"{base_name}_excel_csv_{now_file_stamp()}.zip"] = build_zip({k: v for k, v in bundle.items() if k.endswith(".csv") or k.endswith("manifest.json")})
            elif target == "JSON":
                outputs[f"{base_name}_excel.json"] = bundle["excel/combined.json"]
            elif target == "HTML":
                outputs[f"{base_name}_excel_html_{now_file_stamp()}.zip"] = build_zip({k: v for k, v in bundle.items() if k.endswith(".html") or k.endswith("manifest.json")})
            elif target == "MD":
                outputs[f"{base_name}_excel_md_{now_file_stamp()}.zip"] = build_zip({k: v for k, v in bundle.items() if k.endswith(".md") or k.endswith("manifest.json")})
            else:  # ZIP
                outputs[f"{base_name}_excel_bundle_{now_file_stamp()}.zip"] = build_zip(bundle)

        elif conv_id == "xlsx_to_pdf":
            outputs[f"{base_name}.pdf"] = excel_to_pdf_bytes(file_bytes)

        elif conv_id == "xlsx_to_word":
            outputs[f"{base_name}.docx"] = excel_to_word_docx(file_bytes)

        # ---------------- WORD ----------------
        elif conv_id == "docx_to_text_bundle":
            text, html, md = docx_to_text_html_md(file_bytes)
            pdf_bytes = text_to_pdf_bytes(f"Word Text Export: {base_name}", text)

            if target == "TXT":
                outputs[f"{base_name}.txt"] = (text or "").encode("utf-8")
            elif target == "HTML":
                outputs[f"{base_name}.html"] = (html or "").encode("utf-8")
            elif target == "MD":
                outputs[f"{base_name}.md"] = (md or "").encode("utf-8")
            elif target == "PDF(Text)":
                outputs[f"{base_name}.pdf"] = pdf_bytes
            else:  # ZIP
                files = {
                    "word/output.txt": (text or "").encode("utf-8"),
                    "word/output.html": (html or "").encode("utf-8"),
                    "word/output.md": (md or "").encode("utf-8"),
                    "word/output.pdf": pdf_bytes,
                    "word/manifest.json": json.dumps({"type": "word_bundle", "created": now_stamp()}, indent=2).encode("utf-8"),
                }
                outputs[f"{base_name}_word_bundle_{now_file_stamp()}.zip"] = build_zip(files)

        elif conv_id == "docx_to_excel":
            outputs[f"{base_name}_tables.xlsx"] = word_to_excel_tables(file_bytes)

        # ---------------- PPT ----------------
        elif conv_id == "pptx_to_bundle":
            bundle = pptx_to_text_json_images(file_bytes)
            outputs[f"{base_name}_ppt_bundle_{now_file_stamp()}.zip"] = build_zip(bundle)

        else:
            raise RuntimeError("Conversion not implemented.")
    except Exception as e:
        st.error(f"Conversion failed: {e}")
        st.stop()


# Save outputs to session + history
st.session_state.last_outputs = outputs
add_history({
    "time": now_stamp(),
    "input": filename,
    "task": task_label,
    "outputs": list(outputs.keys())
})

# ============================================================
# Step 4: Downloads (professional)
# ============================================================
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown("### ✅ Done. Download your files")

if not outputs:
    st.warning("No outputs produced.")
    st.stop()

# Show a compact summary + download buttons
sum_left, sum_right = st.columns([1.2, 0.8], gap="large")
with sum_left:
    st.markdown("<div class='card'><h4 style='margin:0 0 6px 0;'>Summary</h4>", unsafe_allow_html=True)
    st.markdown(f"- **Input:** `{filename}`")
    st.markdown(f"- **Action:** **{task_label}**")
    st.markdown(f"- **Generated:** {len(outputs)} file(s)")
    st.markdown("</div>", unsafe_allow_html=True)

with sum_right:
    st.markdown("<div class='card'><h4 style='margin:0 0 10px 0;'>Downloads</h4>", unsafe_allow_html=True)
    for out_name, out_bytes in outputs.items():
        lname = out_name.lower()
        mime = "application/octet-stream"
        if lname.endswith(".zip"):
            mime = "application/zip"
        elif lname.endswith(".pdf"):
            mime = "application/pdf"
        elif lname.endswith(".docx"):
            mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        elif lname.endswith(".pptx"):
            mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        elif lname.endswith(".xlsx"):
            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        elif lname.endswith(".csv"):
            mime = "text/csv"
        elif lname.endswith(".json"):
            mime = "application/json"
        elif lname.endswith(".txt"):
            mime = "text/plain"
        elif lname.endswith(".html"):
            mime = "text/html"
        elif lname.endswith(".md"):
            mime = "text/markdown"

        st.download_button(
            label=f"Download {out_name}",
            data=out_bytes,
            file_name=out_name,
            mime=mime,
            key=f"dl_{out_name}_{now_file_stamp()}",
            use_container_width=True
        )
    st.markdown("</div>", unsafe_allow_html=True)

# Optional: quick “Last outputs” panel
with st.expander("View output list", expanded=False):
    for out_name, out_bytes in outputs.items():
        st.write(f"- `{out_name}` • {len(out_bytes)} bytes")
