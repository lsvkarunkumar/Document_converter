import os
import re
import io
import json
import zipfile
import tempfile
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple

import streamlit as st
import pandas as pd
from PIL import Image


# ============================================================
# Page config + simple "professional" styling
# ============================================================
st.set_page_config(page_title="Pro Document Converter", layout="wide")

st.markdown(
    """
    <style>
      .block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
      .small-muted { color: rgba(255,255,255,0.65); font-size: 0.9rem; }
      .card {
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 14px;
        padding: 14px 16px;
        background: rgba(255,255,255,0.02);
      }
      .card h4 { margin: 0 0 6px 0; }
      .hr {
        height: 1px; background: rgba(255,255,255,0.08);
        border: none; margin: 12px 0;
      }
      .pill {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 999px;
        border: 1px solid rgba(255,255,255,0.14);
        background: rgba(255,255,255,0.03);
        font-size: 0.85rem;
        margin-right: 6px;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🧰 Pro Document Converter")
st.caption("Web-based multi-format converter: PDF/Image/Office → data/doc formats + ZIP bundles. Includes Excel → PDF and conversion history.")


# ============================================================
# Utilities
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
# Text cleanup (RAW vs CLEAN)
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
# PDF capabilities
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
# Tables extraction (PDF text-layer first, fallback OCR)
# ============================================================
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

    # Fix column names (None/empty/duplicates)
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

    # CSV + JSON per table and combined
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


# ============================================================
# Office conversions: DOCX, XLSX, PPTX
# ============================================================
def docx_to_text_html_md(docx_bytes: bytes) -> Tuple[str, str, str]:
    from docx import Document
    from markdownify import markdownify as mdify

    doc = Document(io.BytesIO(docx_bytes))
    paragraphs = [p.text for p in doc.paragraphs if p.text and p.text.strip()]
    text = "\n".join(paragraphs).strip()

    # Minimal HTML
    html_lines = []
    for p in paragraphs:
        # basic escaping
        p2 = (p.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
        html_lines.append(f"<p>{p2}</p>")
    html = "\n".join(html_lines).strip()

    md = mdify(html) if html else (text or "")
    return text or "", html or "", md or ""


def text_to_pdf_bytes(title: str, text: str) -> bytes:
    # Web-safe text-only PDF
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

        # wrap long lines
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
    """
    Excel -> PDF (table-style).
    - One section per sheet
    - Truncates very large sheets (web-safe)
    """
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet

    styles = getSampleStyleSheet()
    story = []

    xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    for si, sheet in enumerate(xls.sheet_names, start=1):
        df = xls.parse(sheet)
        df = df.copy()

        # truncate
        df = df.iloc[:max_rows, :max_cols]

        # title
        story.append(Paragraph(f"Sheet: {sheet}", styles["Heading2"]))
        story.append(Paragraph(f"Exported: {now_stamp()}", styles["Normal"]))
        story.append(Spacer(1, 10))

        # build table
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
    # landscape is better for tables
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), rightMargin=18, leftMargin=18, topMargin=18, bottomMargin=18)
    doc.build(story)
    return buf.getvalue()


def xlsx_to_outputs(xlsx_bytes: bytes) -> Dict[str, bytes]:
    """
    XLSX -> CSV/JSON/HTML/Markdown + Excel->PDF + ZIP
    Exports per sheet and combined JSON.
    """
    from markdownify import markdownify as mdify

    xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    files: Dict[str, bytes] = {}
    combined = {"sheets": []}

    # Sheet exports
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

    # NEW: Excel -> PDF
    files["excel/output.pdf"] = excel_to_pdf_bytes(xlsx_bytes)

    files["excel/manifest.json"] = json.dumps(
        {"type": "excel_export", "sheet_count": len(xls.sheet_names), "created": now_stamp()},
        ensure_ascii=False,
        indent=2,
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

            # extract embedded images (not full slide render)
            try:
                if shape.shape_type == 13:  # PICTURE
                    image = shape.image
                    img_bytes = image.blob
                    ext = (image.ext or "bin").lower()
                    img_count += 1
                    files[f"pptx/images/slide_{si:03d}_img_{img_count:03d}.{ext}"] = img_bytes
            except Exception:
                pass

        slide_text = "\n".join(slide_text_parts).strip()
        all_text.append(f"--- Slide {si} ---\n{slide_text}".strip())
        slides_json.append({"slide": si, "text": slide_text})

    files["pptx/slides.txt"] = ("\n\n".join(all_text).strip() + "\n").encode("utf-8")
    files["pptx/slides.json"] = json.dumps({"slides": slides_json}, ensure_ascii=False, indent=2).encode("utf-8")
    files["pptx/manifest.json"] = json.dumps(
        {"type": "pptx_export", "slide_count": len(prs.slides), "extracted_images": img_count, "created": now_stamp()},
        ensure_ascii=False,
        indent=2,
    ).encode("utf-8")
    return files


# ============================================================
# Optional: Searchable PDF (OCR Layer)
# ============================================================
def ocrmypdf_available() -> bool:
    try:
        import ocrmypdf  # noqa
        return True
    except Exception:
        return False


def make_searchable_pdf(pdf_bytes: bytes, lang: str = "eng") -> bytes:
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
# Session state: history + last outputs
# ============================================================
if "history" not in st.session_state:
    st.session_state.history = []  # list of dicts

if "last_outputs" not in st.session_state:
    st.session_state.last_outputs = {}  # filename -> bytes


def add_history(entry: Dict[str, Any]):
    st.session_state.history.insert(0, entry)
    st.session_state.history = st.session_state.history[:30]  # keep last 30


# ============================================================
# Sidebar (clean)
# ============================================================
with st.sidebar:
    st.markdown("### ⚙️ Settings")

    clean_mode = st.selectbox("Table text cleanup", ["Clean (recommended)", "Raw"], index=0)
    prefer_text_layer_tables = st.checkbox("PDF tables: prefer text layer first", value=True)

    ocr_lang = st.selectbox("OCR language", ["eng"], index=0)
    max_pages = st.slider("Max pages (PDF)", 1, 60, 12)
    ocr_dpi = st.slider("OCR DPI", 200, 400, 260, step=10)
    min_conf = st.slider("Table OCR min confidence", 0, 100, 50)

    show_preview = st.checkbox("Show preview (OCR/PDF render)", value=False)

    st.markdown('<hr class="hr">', unsafe_allow_html=True)
    st.markdown("### 🕘 History")
    if st.session_state.history:
        with st.expander("View recent conversions", expanded=False):
            for h in st.session_state.history[:10]:
                st.markdown(
                    f"- **{h['time']}** • `{h['input_name']}` • **{h['task']}** • outputs: {', '.join(h['outputs'])}"
                )
    else:
        st.caption("No conversions yet.")


# ============================================================
# Upload + auto task list by input type
# ============================================================
uploaded = st.file_uploader(
    "Upload file",
    type=["pdf", "png", "jpg", "jpeg", "webp", "docx", "xlsx", "xlsm", "pptx"],
)

if not uploaded:
    st.info("Upload a file to begin.")
    st.stop()

file_bytes = uploaded.read()
filename = uploaded.name
base_name = safe_filename(os.path.splitext(filename)[0])
normalizer = normalize_cell_text_clean if clean_mode.startswith("Clean") else normalize_cell_text_raw

# Determine valid tasks for this input type
tasks_pdf = [
    "PDF → Tables → Export (Excel/CSV/JSON/ZIP)",
    "PDF → Text (Hybrid) → TXT/HTML/Markdown/ZIP",
    "PDF → Word (Hybrid) → DOCX",
    "PDF → Images (PNG) → ZIP",
    "PDF → Metadata → JSON",
    "PDF → Smart Search PDF (OCR layer) (optional)",
]
tasks_img = [
    "Image → OCR Text → TXT",
    "Image → Tables → Export (Excel/CSV/JSON/ZIP)",
    "Image → PDF",
]
tasks_docx = [
    "DOCX → TXT/HTML/Markdown/PDF(text-only)/ZIP",
]
tasks_xlsx = [
    "XLSX → CSV/JSON/HTML/Markdown/PDF/ZIP",
]
tasks_pptx = [
    "PPTX → TXT/JSON/Images/PDF(text-only)/ZIP",
]

if is_pdf(filename):
    valid_tasks = tasks_pdf
    file_badge = "PDF"
elif is_image(filename):
    valid_tasks = tasks_img
    file_badge = "IMAGE"
elif is_docx(filename):
    valid_tasks = tasks_docx
    file_badge = "DOCX"
elif is_xlsx(filename):
    valid_tasks = tasks_xlsx
    file_badge = "XLSX"
elif is_pptx(filename):
    valid_tasks = tasks_pptx
    file_badge = "PPTX"
else:
    valid_tasks = []
    file_badge = "FILE"

# Header cards
colA, colB, colC = st.columns([2, 1, 1])
with colA:
    st.markdown(f"<div class='card'><h4>Input</h4><div class='pill'>{file_badge}</div> <span class='pill'>{filename}</span><div class='small-muted'>Loaded into memory • ready for conversion</div></div>", unsafe_allow_html=True)
with colB:
    st.markdown(f"<div class='card'><h4>Mode</h4><div class='small-muted'>Cleanup: <b>{clean_mode}</b><br>OCR lang: <b>{ocr_lang}</b></div></div>", unsafe_allow_html=True)
with colC:
    st.markdown(f"<div class='card'><h4>Pro Tips</h4><div class='small-muted'>Use ZIP for batch exports.<br>For scanned PDFs, Hybrid Text is best.</div></div>", unsafe_allow_html=True)

st.markdown('<hr class="hr">', unsafe_allow_html=True)

# Tabs layout
tab_convert, tab_outputs, tab_history = st.tabs(["🔁 Convert", "📦 Last Outputs", "🕘 History"])

with tab_convert:
    left, right = st.columns([1, 1], gap="large")

    with left:
        st.subheader("Preview")
        if is_image(filename):
            try:
                im = Image.open(io.BytesIO(file_bytes)).convert("RGB")
                st.image(im, use_container_width=True)
            except Exception:
                st.info("Image preview not available.")
        else:
            st.info("Preview is available for OCR/PDF render tasks when 'Show preview' is enabled.")

    with right:
        st.subheader("Convert")
        task = st.selectbox("Choose conversion", valid_tasks, index=0)
        run = st.button("Run conversion", type="primary")

        if not run:
            st.stop()

        outputs: Dict[str, bytes] = {}
        output_names: List[str] = []

        # -------------------------
        # PDF tasks
        # -------------------------
        if task == "PDF → Metadata → JSON":
            md = pdf_metadata_to_dict(file_bytes)
            b = json.dumps(md, ensure_ascii=False, indent=2).encode("utf-8")
            outputs[f"{base_name}_metadata.json"] = b
            st.json(md)

        elif task == "PDF → Images (PNG) → ZIP":
            zip_bytes, count = pdf_to_images_zip(file_bytes, max_pages=max_pages, dpi=220)
            outputs[f"{base_name}_pages_{now_file_stamp()}.zip"] = zip_bytes
            st.success(f"Rendered {count} pages into PNG files (zipped).")

        elif task == "PDF → Text (Hybrid) → TXT/HTML/Markdown/ZIP":
            pages = pdf_hybrid_text_extract(file_bytes, max_pages=max_pages, lang=ocr_lang, dpi=ocr_dpi)

            if show_preview:
                try:
                    imgs = pdf_render_pages_to_images(file_bytes, dpi=ocr_dpi, max_pages=1)
                    if imgs:
                        disp = to_displayable_image(imgs[0])
                        if disp is not None:
                            st.image(disp, caption="First page rendered", use_container_width=True)
                except Exception:
                    pass

            txt = "\n\n".join([f"--- Page {i+1} ---\n{p}".strip() for i, p in enumerate(pages)]).strip() + "\n"
            html = "\n".join(
                [f"<h2>Page {i+1}</h2>\n<pre>{(p or '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')}</pre>" for i, p in enumerate(pages)]
            )
            try:
                from markdownify import markdownify as mdify
                md = mdify(html)
            except Exception:
                md = txt

            files = {
                "text/output.txt": txt.encode("utf-8"),
                "text/output.html": html.encode("utf-8"),
                "text/output.md": (md or "").encode("utf-8"),
                "text/manifest.json": json.dumps({"type": "pdf_hybrid_text", "pages": len(pages), "created": now_stamp()}, indent=2).encode("utf-8"),
            }
            z = build_zip(files)
            outputs[f"{base_name}_text_{now_file_stamp()}.zip"] = z
            outputs[f"{base_name}.txt"] = files["text/output.txt"]
            outputs[f"{base_name}.html"] = files["text/output.html"]
            outputs[f"{base_name}.md"] = files["text/output.md"]
            st.text_area("Preview (first 1500 chars)", txt[:1500], height=260)

        elif task == "PDF → Word (Hybrid) → DOCX":
            from docx import Document

            pages = pdf_hybrid_text_extract(file_bytes, max_pages=max_pages, lang=ocr_lang, dpi=ocr_dpi)

            doc = Document()
            doc.add_heading("Hybrid Extracted Text (PDF)", level=1)
            doc.add_paragraph(f"Source: {filename}")
            doc.add_paragraph(f"Created: {now_stamp()}")
            doc.add_paragraph("")

            for i, p in enumerate(pages, start=1):
                doc.add_heading(f"Page {i}", level=2)
                for line in (p or "").splitlines():
                    doc.add_paragraph(line)

            out = io.BytesIO()
            doc.save(out)
            outputs[f"{base_name}_hybrid.docx"] = out.getvalue()
            st.success("DOCX created (editable text).")

        elif task == "PDF → Tables → Export (Excel/CSV/JSON/ZIP)":
            tables_dfs: List[pd.DataFrame] = []

            if prefer_text_layer_tables:
                with st.spinner("Trying PDF text-layer table extraction..."):
                    try:
                        tables_dfs = extract_tables_pdf_textlayer(file_bytes, max_pages=max_pages)
                    except Exception:
                        tables_dfs = []

            if not tables_dfs:
                with st.spinner("Falling back to OCR table extraction (img2table + Tesseract)..."):
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
                st.error("No tables extracted.")
                st.stop()

            st.success(f"Extracted {len(tables_dfs)} table(s). Preview Table 1:")
            st.dataframe(tables_dfs[0].applymap(normalizer), use_container_width=True)

            root = f"{base_name}_tables_{now_file_stamp()}"
            files = build_tables_bundle(tables_dfs, normalizer, base_root=root)
            z = build_zip(files)
            outputs[f"{root}.zip"] = z
            outputs[f"{root}.xlsx"] = files[f"{root}.xlsx"]

        elif task == "PDF → Smart Search PDF (OCR layer) (optional)":
            if not ocrmypdf_available():
                st.warning("Not available. Add `ocrmypdf` to requirements.txt and `ghostscript`, `qpdf` to packages.txt.")
                st.stop()

            with st.spinner("Creating searchable PDF (OCR layer)..."):
                out_pdf = make_searchable_pdf(file_bytes, lang=ocr_lang)

            outputs[f"{base_name}_searchable.pdf"] = out_pdf
            st.success("Searchable PDF created.")

        # -------------------------
        # Image tasks
        # -------------------------
        elif task == "Image → OCR Text → TXT":
            im = Image.open(io.BytesIO(file_bytes)).convert("RGB")
            with st.spinner("Running OCR on image..."):
                text = ocr_image_to_text(im, lang=ocr_lang)

            outputs[f"{base_name}.txt"] = (text or "").encode("utf-8")
            st.text_area("OCR Text", text or "(No text extracted)", height=280)

        elif task == "Image → PDF":
            im = Image.open(io.BytesIO(file_bytes)).convert("RGB")
            buf = io.BytesIO()
            im.save(buf, format="PDF")
            outputs[f"{base_name}.pdf"] = buf.getvalue()
            st.success("Converted image to PDF.")

        elif task == "Image → Tables → Export (Excel/CSV/JSON/ZIP)":
            from img2table.ocr import TesseractOCR
            from img2table.document import Image as Img2TableImage

            with st.spinner("Extracting tables from image (img2table + Tesseract)..."):
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
                st.error("No tables extracted from image.")
                st.stop()

            st.success(f"Extracted {len(tables_dfs)} table(s). Preview Table 1:")
            st.dataframe(tables_dfs[0].applymap(normalizer), use_container_width=True)

            root = f"{base_name}_tables_{now_file_stamp()}"
            files = build_tables_bundle(tables_dfs, normalizer, base_root=root)
            z = build_zip(files)
            outputs[f"{root}.zip"] = z

        # -------------------------
        # DOCX task
        # -------------------------
        elif task == "DOCX → TXT/HTML/Markdown/PDF(text-only)/ZIP":
            text, html, md = docx_to_text_html_md(file_bytes)
            pdf_bytes = text_to_pdf_bytes(f"DOCX Export: {base_name}", text)

            files = {
                "docx/output.txt": (text or "").encode("utf-8"),
                "docx/output.html": (html or "").encode("utf-8"),
                "docx/output.md": (md or "").encode("utf-8"),
                "docx/output.pdf": pdf_bytes,
                "docx/manifest.json": json.dumps({"type": "docx_export", "created": now_stamp()}, indent=2).encode("utf-8"),
            }
            z = build_zip(files)
            outputs[f"{base_name}_docx_{now_file_stamp()}.zip"] = z
            st.text_area("Preview (text)", (text or "")[:1500], height=260)

        # -------------------------
        # XLSX task (NEW includes Excel -> PDF)
        # -------------------------
        elif task == "XLSX → CSV/JSON/HTML/Markdown/PDF/ZIP":
            with st.spinner("Exporting Excel sheets + generating PDF..."):
                files = xlsx_to_outputs(file_bytes)
                z = build_zip(files)

            outputs[f"{base_name}_excel_{now_file_stamp()}.zip"] = z
            outputs[f"{base_name}.pdf"] = files["excel/output.pdf"]
            st.success("Excel exported: CSV/JSON/HTML/MD per sheet + combined JSON + PDF.")

        # -------------------------
        # PPTX task
        # -------------------------
        elif task == "PPTX → TXT/JSON/Images/PDF(text-only)/ZIP":
            with st.spinner("Extracting PPTX text + embedded images..."):
                files = pptx_to_text_json_images(file_bytes)

            slides_text = files.get("pptx/slides.txt", b"").decode("utf-8", errors="ignore")
            files["pptx/slides.pdf"] = text_to_pdf_bytes(f"PPTX Export: {base_name}", slides_text)

            z = build_zip(files)
            outputs[f"{base_name}_pptx_{now_file_stamp()}.zip"] = z
            st.success("PPTX exported: TXT + JSON + embedded images + text-only PDF.")

        else:
            st.error("Task not supported for this file type.")
            st.stop()

        # Save outputs into session_state for "Last Outputs"
        st.session_state.last_outputs = outputs

        output_names = list(outputs.keys())

        # Add history
        add_history({
            "time": now_stamp(),
            "task": task,
            "input_name": filename,
            "outputs": output_names,
        })

        # Show download buttons
        st.markdown('<hr class="hr">', unsafe_allow_html=True)
        st.subheader("Downloads")
        for out_name, out_bytes in outputs.items():
            mime = "application/octet-stream"
            lname = out_name.lower()
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
            )


with tab_outputs:
    st.subheader("📦 Last Outputs")
    if not st.session_state.last_outputs:
        st.info("Run a conversion to see outputs here.")
    else:
        st.markdown("<div class='small-muted'>These are stored in the current session only.</div>", unsafe_allow_html=True)
        for out_name, out_bytes in st.session_state.last_outputs.items():
            st.write(f"- `{out_name}` • {len(out_bytes)} bytes")


with tab_history:
    st.subheader("🕘 Conversion History")
    if not st.session_state.history:
        st.info("No history yet.")
    else:
        df = pd.DataFrame(st.session_state.history)
        st.dataframe(df, use_container_width=True)
