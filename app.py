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
# Page config + "Smallpdf-ish" clean UI
# ============================================================
st.set_page_config(page_title="DocFlow Converter", layout="wide")

st.markdown(
    """
    <style>
      .block-container { max-width: 1220px; padding-top: 1.0rem; padding-bottom: 2rem; }

      .topbar {
        display:flex; align-items:center; justify-content:space-between;
        padding: 12px 14px; border-radius: 14px;
        border: 1px solid rgba(0,0,0,0.08);
        background: linear-gradient(180deg, rgba(255,255,255,0.95), rgba(255,255,255,0.86));
      }
      .brand { font-size: 18px; font-weight: 850; letter-spacing: .2px; }
      .sub { color: rgba(0,0,0,0.62); font-size: 13px; margin-top: 2px; }
      .tag {
        font-size: 12px; padding: 4px 10px; border-radius: 999px;
        border: 1px solid rgba(0,0,0,0.10);
        background: rgba(255,255,255,0.92);
      }
      .card {
        border: 1px solid rgba(0,0,0,0.08);
        border-radius: 14px;
        padding: 14px 14px;
        background: rgba(255,255,255,0.92);
      }
      .muted { color: rgba(0,0,0,0.62); font-size: 13px; }
      .divider { height: 1px; background: rgba(0,0,0,0.08); margin: 14px 0; }

      .step {
        display: inline-flex; align-items: center; gap: 8px;
        font-size: 13px; padding: 6px 10px; border-radius: 999px;
        border: 1px solid rgba(0,0,0,0.08);
        background: rgba(255,255,255,0.88);
      }
      .step b {
        font-size: 12px; padding: 3px 8px; border-radius: 999px;
        background: rgba(0,0,0,0.06);
      }

      .pill {
        display:inline-block; padding: 4px 10px;
        border-radius: 999px;
        border: 1px solid rgba(0,0,0,0.10);
        background: rgba(0,0,0,0.04);
        font-size: 12px;
        margin-right: 6px;
      }

      .disabled {
        opacity: 0.45;
        filter: grayscale(0.2);
      }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="topbar">
      <div>
        <div class="brand">DocFlow Converter</div>
        <div class="sub">Upload → Choose → Convert → Download</div>
      </div>
      <div class="tag">Web-only • Streamlit</div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.write("")


# ============================================================
# Helpers
# ============================================================
def now_stamp() -> str:
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


def infer_type(filename: str) -> str:
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".pdf":
        return "PDF"
    if ext in [".png", ".jpg", ".jpeg", ".webp"]:
        return "IMAGE"
    if ext in [".xlsx", ".xlsm"]:
        return "EXCEL"
    if ext == ".docx":
        return "WORD"
    if ext == ".pptx":
        return "PPT"
    return "UNKNOWN"


def mime_for(name: str) -> str:
    lname = name.lower()
    if lname.endswith(".zip"):
        return "application/zip"
    if lname.endswith(".pdf"):
        return "application/pdf"
    if lname.endswith(".docx"):
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    if lname.endswith(".pptx"):
        return "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    if lname.endswith(".xlsx"):
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    if lname.endswith(".csv"):
        return "text/csv"
    if lname.endswith(".json"):
        return "application/json"
    if lname.endswith(".txt"):
        return "text/plain"
    if lname.endswith(".png"):
        return "image/png"
    if lname.endswith(".html"):
        return "text/html"
    if lname.endswith(".md"):
        return "text/markdown"
    return "application/octet-stream"


# ============================================================
# OCR / PDF helpers
# ============================================================
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


def pdf_hybrid_text_extract(pdf_bytes: bytes, max_pages: int, lang: str, dpi: int) -> List[str]:
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
        files[f"pages/page_{i:03d}.png"] = buf.getvalue()
    return build_zip(files), len(images)


def pdf_metadata_to_json(pdf_bytes: bytes) -> bytes:
    from pypdf import PdfReader
    r = PdfReader(io.BytesIO(pdf_bytes))
    md = r.metadata or {}
    out = {"page_count": len(r.pages)}
    for k, v in md.items():
        out[str(k)] = str(v) if v is not None else None
    return json.dumps(out, ensure_ascii=False, indent=2).encode("utf-8")


# ============================================================
# Tables extraction (text-layer first + OCR fallback)
# ============================================================
def normalize_cell_text_clean(val):
    if val is None:
        return val
    s = str(val).replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"\n+", " ", s)
    s = s.replace("\u00a0", " ")
    s = re.sub(r"[ \t]+", " ", s).strip()
    s = re.sub(r"^\|\s*", "", s)

    def join_spaced_letters(m):
        return m.group(0).replace(" ", "")

    s = re.sub(r"(?:\b[A-Za-z]\b(?:\s+|$)){4,}", join_spaced_letters, s)
    s = re.sub(r"(?:\b\d\b\s+){3,}\b\d\b", lambda m: m.group(0).replace(" ", ""), s)
    for _ in range(2):
        s = re.sub(r"\b([A-Za-z])\s+([A-Za-z]{2,})\b", r"\1\2", s)

    s = re.sub(r"\s*([,/:\.\-\+])\s*", r"\1", s)
    s = re.sub(r"\b(GB(?:/T)?)\s*([0-9])", r"\1 \2", s, flags=re.IGNORECASE)
    return s


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

    df2 = df2.where(pd.notnull(df2), None)
    return df2.to_dict(orient="records")


def build_tables_bundle(tables: List[pd.DataFrame], base: str) -> Dict[str, bytes]:
    from openpyxl.styles import Alignment

    cleaned = [df.applymap(normalize_cell_text_clean) for df in tables]
    files: Dict[str, bytes] = {}

    # Excel
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        for i, df in enumerate(cleaned, start=1):
            df.to_excel(writer, sheet_name=f"Table_{i}"[:31], index=False)
        wb = writer.book
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=False, vertical="top")
    files[f"{base}.xlsx"] = excel_buf.getvalue()

    combined_json = {"tables": []}
    for i, df in enumerate(cleaned, start=1):
        files[f"{base}/tables/table_{i}.csv"] = df.to_csv(index=False).encode("utf-8")
        one = {"table_index": i, "rows": df_to_json_records(df)}
        files[f"{base}/tables/table_{i}.json"] = json.dumps(one, ensure_ascii=False, indent=2).encode("utf-8")
        combined_json["tables"].append(one)

    files[f"{base}/tables/combined.json"] = json.dumps(combined_json, ensure_ascii=False, indent=2).encode("utf-8")
    files[f"{base}/manifest.json"] = json.dumps({"type": "tables_export", "table_count": len(cleaned)}, indent=2).encode("utf-8")
    return files


# ============================================================
# Office conversions
# ============================================================
def excel_to_pdf_bytes(xlsx_bytes: bytes, max_rows: int = 90, max_cols: int = 14) -> bytes:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet

    styles = getSampleStyleSheet()
    story = []

    xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    for si, sheet in enumerate(xls.sheet_names, start=1):
        df = xls.parse(sheet).iloc[:max_rows, :max_cols]
        story.append(Paragraph(f"{sheet}", styles["Heading2"]))
        story.append(Spacer(1, 8))
        data = [list(df.columns.astype(str))] + df.astype(str).values.tolist()
        tbl = Table(data, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f2937")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
        ]))
        story.append(tbl)
        if si < len(xls.sheet_names):
            story.append(PageBreak())

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=18, rightMargin=18, topMargin=18, bottomMargin=18)
    doc.build(story)
    return buf.getvalue()


def excel_to_word_docx_bytes(xlsx_bytes: bytes, max_rows: int = 120, max_cols: int = 14) -> bytes:
    from docx import Document
    from docx.shared import Pt

    xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)

    for si, sheet in enumerate(xls.sheet_names, start=1):
        df = xls.parse(sheet).iloc[:max_rows, :max_cols]
        doc.add_heading(sheet, level=2)

        table = doc.add_table(rows=1, cols=len(df.columns))
        hdr = table.rows[0].cells
        for j, c in enumerate(df.columns.astype(str).tolist()):
            hdr[j].text = c

        for _, row in df.astype(str).iterrows():
            cells = table.add_row().cells
            for j, val in enumerate(row.tolist()):
                cells[j].text = val

        if si < len(xls.sheet_names):
            doc.add_paragraph("")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def word_to_excel_tables(docx_bytes: bytes) -> Optional[bytes]:
    from docx import Document

    doc = Document(io.BytesIO(docx_bytes))
    if not doc.tables:
        return None

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for i, t in enumerate(doc.tables, start=1):
            rows = []
            for r in t.rows:
                rows.append([c.text.strip() for c in r.cells])
            df = pd.DataFrame(rows)
            df.to_excel(writer, sheet_name=f"Table_{i}"[:31], index=False, header=False)
    return out.getvalue()


def docx_to_plain_text(docx_bytes: bytes) -> str:
    from docx import Document
    doc = Document(io.BytesIO(docx_bytes))
    parts = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
    return "\n".join(parts).strip()


# ============================================================
# Optional searchable PDF (OCR layer)
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


def make_searchable_pdf_from_image(img_bytes: bytes, lang: str = "eng") -> bytes:
    img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    tmp_pdf = io.BytesIO()
    img.save(tmp_pdf, format="PDF")
    return make_searchable_pdf_from_pdf(tmp_pdf.getvalue(), lang=lang)


# ============================================================
# Session state
# ============================================================
if "outputs" not in st.session_state:
    st.session_state.outputs = {}  # name -> bytes

if "last_task_label" not in st.session_state:
    st.session_state.last_task_label = None

if "history" not in st.session_state:
    st.session_state.history = []


def push_history(task_label: str, fname: str):
    st.session_state.history.insert(0, {
        "time": datetime.utcnow().strftime("%H:%M:%S"),
        "task": task_label,
        "file": fname,
    })
    st.session_state.history = st.session_state.history[:20]


# ============================================================
# Sidebar - clean
# ============================================================
with st.sidebar:
    st.markdown("### Settings")
    with st.expander("Advanced options", expanded=False):
        ocr_lang = st.selectbox("OCR language", ["eng"], index=0)
        max_pages = st.slider("Max pages (PDF)", 1, 80, 12)
        ocr_dpi = st.slider("OCR quality (DPI)", 200, 400, 260, step=10)
        min_conf = st.slider("Table confidence", 0, 100, 50)

    st.markdown("### History")
    if st.session_state.history:
        for h in st.session_state.history[:8]:
            st.caption(f"• {h['time']} — {h['task']}")
    else:
        st.caption("No conversions yet.")


# ============================================================
# Permanent layout: Upload | Detected | Choose | Convert | Download
# ============================================================
col_left, col_right = st.columns([1.1, 1.0])

# ---------------- Left: Upload + detection + conversion choose
with col_left:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="step"><b>1</b> Upload</div>', unsafe_allow_html=True)
    st.write("")

    uploaded = st.file_uploader(
        "Upload file",
        type=["pdf", "png", "jpg", "jpeg", "webp", "docx", "xlsx", "xlsm", "pptx"],
        label_visibility="visible",
    )

    st.write("")
    st.markdown('<div class="step"><b>2</b> Detected</div>', unsafe_allow_html=True)
    st.write("")

    if uploaded:
        filename = uploaded.name
        file_bytes = uploaded.read()
        ftype = infer_type(filename)
        base = safe_filename(os.path.splitext(filename)[0])

        st.markdown(
            f"""
            <span class="pill"><b>{ftype}</b></span>
            <span class="pill">{filename}</span>
            """,
            unsafe_allow_html=True,
        )
        st.caption("File loaded and ready.")

    else:
        filename = None
        file_bytes = None
        ftype = "—"
        base = "output"
        st.markdown('<div class="muted">No file uploaded yet.</div>', unsafe_allow_html=True)

    st.write("")
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    st.markdown('<div class="step"><b>3</b> Choose conversion</div>', unsafe_allow_html=True)
    st.write("")

    TASKS_BY_TYPE = {
        "PDF": [
            ("Extract Tables → Excel/CSV/JSON (ZIP)", "pdf_tables"),
            ("Extract Text (OCR/Hybrid) → TXT", "pdf_text_txt"),
            ("Convert → Word (DOCX)", "pdf_to_docx"),
            ("Create Searchable PDF (OCR layer)", "pdf_searchable"),
            ("Pages → PNG (ZIP)", "pdf_pages_png"),
            ("Metadata → JSON", "pdf_meta_json"),
        ],
        "IMAGE": [
            ("OCR Image → TXT", "img_text_txt"),
            ("Extract Tables → Excel/CSV/JSON (ZIP)", "img_tables"),
            ("Convert → PDF", "img_to_pdf"),
            ("Convert → Searchable PDF (OCR layer)", "img_searchable_pdf"),
        ],
        "EXCEL": [
            ("Convert → PDF", "xlsx_to_pdf"),
            ("Convert → Word (DOCX)", "xlsx_to_docx"),
        ],
        "WORD": [
            ("Convert → TXT", "docx_to_txt"),
            ("Extract Tables → Excel (XLSX)", "docx_tables_to_xlsx"),
        ],
        "PPT": [
            ("Extract Text → TXT/JSON (ZIP)", "pptx_text_bundle"),
            ("Extract Embedded Images (ZIP)", "pptx_images_zip"),
        ],
    }

    task_options = TASKS_BY_TYPE.get(ftype, [])
    task_labels = [t[0] for t in task_options] if task_options else ["Upload a file to see conversions"]
    task_disabled = not bool(uploaded and task_options)

    task_label = st.selectbox(
        "Conversion",
        task_labels,
        index=0,
        disabled=task_disabled,
    )

    task_key = None
    if not task_disabled:
        task_key = dict(task_options)[task_label]

    st.markdown("</div>", unsafe_allow_html=True)

# ---------------- Right: Convert + Download (always visible)
with col_right:
    st.markdown('<div class="card">', unsafe_allow_html=True)

    st.markdown('<div class="step"><b>4</b> Convert</div>', unsafe_allow_html=True)
    st.write("")

    convert_disabled = not (uploaded and task_key)
    convert_btn = st.button("Convert", type="primary", disabled=convert_disabled)

    st.write("")
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    st.markdown('<div class="step"><b>5</b> Download</div>', unsafe_allow_html=True)
    st.write("")

    outputs_ready = bool(st.session_state.outputs)
    download_area_class = "" if outputs_ready else "disabled"
    st.markdown(f'<div class="{download_area_class}">', unsafe_allow_html=True)

    if not outputs_ready:
        st.caption("Downloads will appear here after conversion.")
        # Fake disabled buttons (visual only)
        st.button("Download output", disabled=True)
        st.button("Download ZIP", disabled=True)
    else:
        st.caption("Ready to download.")
        for out_name, out_bytes in st.session_state.outputs.items():
            st.download_button(
                label=f"Download {out_name}",
                data=out_bytes,
                file_name=out_name,
                mime=mime_for(out_name),
                key=f"dl_{out_name}_{now_stamp()}",
            )

        # Always offer "ZIP all"
        zname = f"{safe_filename(os.path.splitext(filename)[0])}_all_{now_stamp()}.zip" if filename else f"bundle_{now_stamp()}.zip"
        st.download_button(
            label=f"Download ALL as ZIP",
            data=build_zip(st.session_state.outputs),
            file_name=zname,
            mime="application/zip",
            key=f"dl_zip_{now_stamp()}",
        )

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# Conversion execution
# ============================================================
if convert_btn and uploaded and task_key and file_bytes and filename:
    st.session_state.outputs = {}
    base = safe_filename(os.path.splitext(filename)[0])

    try:
        outputs: Dict[str, bytes] = {}

        # PDF
        if task_key == "pdf_text_txt":
            pages = pdf_hybrid_text_extract(file_bytes, max_pages=max_pages, lang=ocr_lang, dpi=ocr_dpi)
            txt = "\n\n".join([p.strip() for p in pages if p is not None]).strip()
            outputs[f"{base}.txt"] = (txt + "\n").encode("utf-8")

        elif task_key == "pdf_to_docx":
            from docx import Document
            pages = pdf_hybrid_text_extract(file_bytes, max_pages=max_pages, lang=ocr_lang, dpi=ocr_dpi)
            doc = Document()
            for p in pages:
                p = (p or "").strip()
                if p:
                    for line in p.splitlines():
                        if line.strip():
                            doc.add_paragraph(line.strip())
                    doc.add_paragraph("")
            out = io.BytesIO()
            doc.save(out)
            outputs[f"{base}.docx"] = out.getvalue()

        elif task_key == "pdf_tables":
            tables = []
            try:
                tables = extract_tables_pdf_textlayer(file_bytes, max_pages=max_pages)
            except Exception:
                tables = []

            if not tables:
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
                        tables.append(df)

            if not tables:
                raise RuntimeError("No tables found in this PDF.")

            root = f"{base}_tables_{now_stamp()}"
            bundle = build_tables_bundle(tables, base=root)
            outputs[f"{root}.xlsx"] = bundle[f"{root}.xlsx"]
            outputs[f"{root}.zip"] = build_zip(bundle)

        elif task_key == "pdf_pages_png":
            z, _ = pdf_to_images_zip(file_bytes, max_pages=max_pages, dpi=220)
            outputs[f"{base}_pages_{now_stamp()}.zip"] = z

        elif task_key == "pdf_meta_json":
            outputs[f"{base}_metadata.json"] = pdf_metadata_to_json(file_bytes)

        elif task_key == "pdf_searchable":
            if not ocrmypdf_available():
                raise RuntimeError("Searchable PDF needs ocrmypdf + ghostscript + qpdf.")
            outputs[f"{base}_searchable.pdf"] = make_searchable_pdf_from_pdf(file_bytes, lang=ocr_lang)

        # IMAGE
        elif task_key == "img_text_txt":
            im = Image.open(io.BytesIO(file_bytes)).convert("RGB")
            txt = ocr_image_to_text(im, lang=ocr_lang)
            outputs[f"{base}.txt"] = (txt + "\n").encode("utf-8")

        elif task_key == "img_to_pdf":
            im = Image.open(io.BytesIO(file_bytes)).convert("RGB")
            buf = io.BytesIO()
            im.save(buf, format="PDF")
            outputs[f"{base}.pdf"] = buf.getvalue()

        elif task_key == "img_searchable_pdf":
            if not ocrmypdf_available():
                raise RuntimeError("Searchable PDF needs ocrmypdf + ghostscript + qpdf.")
            outputs[f"{base}_searchable.pdf"] = make_searchable_pdf_from_image(file_bytes, lang=ocr_lang)

        elif task_key == "img_tables":
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

            tables = []
            for t in flatten_img2table_tables(tables_obj):
                df = table_to_df_safe(t)
                if df is not None:
                    tables.append(df)

            if not tables:
                raise RuntimeError("No tables found in this image.")

            root = f"{base}_tables_{now_stamp()}"
            bundle = build_tables_bundle(tables, base=root)
            outputs[f"{root}.xlsx"] = bundle[f"{root}.xlsx"]
            outputs[f"{root}.zip"] = build_zip(bundle)

        # EXCEL
        elif task_key == "xlsx_to_pdf":
            outputs[f"{base}.pdf"] = excel_to_pdf_bytes(file_bytes)

        elif task_key == "xlsx_to_docx":
            outputs[f"{base}.docx"] = excel_to_word_docx_bytes(file_bytes)

        # WORD
        elif task_key == "docx_to_txt":
            txt = docx_to_plain_text(file_bytes)
            outputs[f"{base}.txt"] = (txt + "\n").encode("utf-8")

        elif task_key == "docx_tables_to_xlsx":
            xlsx = word_to_excel_tables(file_bytes)
            if xlsx is None:
                raise RuntimeError("No tables found in this Word document.")
            outputs[f"{base}_tables.xlsx"] = xlsx

        # PPT
        elif task_key == "pptx_text_bundle":
            from pptx import Presentation
            prs = Presentation(io.BytesIO(file_bytes))
            slides = []
            all_text = []
            for i, slide in enumerate(prs.slides, start=1):
                parts = []
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text:
                        t = shape.text.strip()
                        if t:
                            parts.append(t)
                stext = "\n".join(parts).strip()
                slides.append({"slide": i, "text": stext})
                all_text.append(f"Slide {i}\n{stext}".strip())

            files = {
                "slides.txt": ("\n\n".join(all_text).strip() + "\n").encode("utf-8"),
                "slides.json": json.dumps({"slides": slides}, ensure_ascii=False, indent=2).encode("utf-8"),
            }
            outputs[f"{base}_ppt_export_{now_stamp()}.zip"] = build_zip(files)

        elif task_key == "pptx_images_zip":
            from pptx import Presentation
            prs = Presentation(io.BytesIO(file_bytes))
            files = {}
            count = 0
            for si, slide in enumerate(prs.slides, start=1):
                for shape in slide.shapes:
                    try:
                        if shape.shape_type == 13:
                            image = shape.image
                            ext = (image.ext or "bin").lower()
                            count += 1
                            files[f"images/slide_{si:03d}_{count:03d}.{ext}"] = image.blob
                    except Exception:
                        pass
            if not files:
                raise RuntimeError("No embedded images found in this PPTX.")
            outputs[f"{base}_images_{now_stamp()}.zip"] = build_zip(files)

        else:
            raise RuntimeError("Conversion not implemented.")

        st.session_state.outputs = outputs
        st.session_state.last_task_label = task_label
        push_history(task_label, filename)

        # small success toast
        st.success("Conversion completed. Downloads are ready on the right panel.")

    except Exception as e:
        st.session_state.outputs = {}
        st.error(str(e))
