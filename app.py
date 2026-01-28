import os
import re
import io
import json
import zipfile
import base64
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple

import streamlit as st
import pandas as pd
import numpy as np
from PIL import Image

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(page_title="DocFlow Converter", layout="wide")

# ============================================================
# PREMIUM CSS (website-like)
# ============================================================
st.markdown(
    """
    <style>
      html, body, [class*="stApp"]{
        background: radial-gradient(1100px 800px at 10% 0%, rgba(99,102,241,0.18), transparent 55%),
                    radial-gradient(900px 700px at 90% 10%, rgba(16,185,129,0.14), transparent 55%),
                    linear-gradient(180deg, #f7f8fb, #f3f4f6);
        color: #0f172a;
      }
      .block-container{ max-width: 1260px; padding-top: 0.9rem; padding-bottom: 1.3rem; }
      #MainMenu, footer { visibility: hidden; }
      header { visibility: hidden; }

      .nav{
        display:flex; align-items:center; justify-content:space-between;
        padding: 12px 14px;
        border-radius: 18px;
        border: 1px solid rgba(15,23,42,0.08);
        background: rgba(255,255,255,0.72);
        backdrop-filter: blur(10px);
        box-shadow: 0 10px 26px rgba(15,23,42,0.06);
        margin-bottom: 12px;
      }
      .navL{ display:flex; align-items:center; gap:12px; }
      .logo{
        width: 40px; height: 40px; border-radius: 14px;
        background: linear-gradient(135deg, rgba(99,102,241,0.95), rgba(16,185,129,0.92));
        box-shadow: 0 12px 22px rgba(99,102,241,0.18);
      }
      .brand{ font-size: 16px; font-weight: 900; margin:0; }
      .subtitle{ margin-top:2px; font-size: 12px; color: rgba(15,23,42,0.62); }

      .chips{ display:flex; gap:8px; flex-wrap:wrap; justify-content:flex-end; }
      .chip{
        font-size: 12px; padding: 7px 10px;
        border-radius: 999px;
        border: 1px solid rgba(15,23,42,0.10);
        background: rgba(255,255,255,0.70);
        color: rgba(15,23,42,0.72);
      }

      .bar{
        border-radius: 18px;
        border: 1px solid rgba(15,23,42,0.08);
        background: rgba(255,255,255,0.72);
        backdrop-filter: blur(10px);
        box-shadow: 0 12px 28px rgba(15,23,42,0.06);
        padding: 12px 12px;
      }

      .sectionTitle{
        font-size: 12px;
        color: rgba(15,23,42,0.65);
        font-weight: 900;
        letter-spacing: .2px;
        margin: 0 0 8px 0;
      }

      .filecard{
        border-radius: 16px;
        border: 1px solid rgba(15,23,42,0.10);
        background: rgba(15,23,42,0.02);
        padding: 10px 10px;
        margin-top: 10px;
      }
      .filetop{ display:flex; justify-content:space-between; gap:10px; }
      .fname{ font-size: 13px; font-weight: 900; }
      .fmeta{ font-size: 12px; color: rgba(15,23,42,0.62); margin-top:2px; }
      .pill{
        font-size: 12px; padding: 6px 10px;
        border-radius: 999px;
        border: 1px solid rgba(15,23,42,0.10);
        background: rgba(255,255,255,0.65);
        color: rgba(15,23,42,0.72);
        height: fit-content;
      }

      .formats{
        display:flex;
        flex-direction:column;
        gap:10px;
        margin-top: 4px;
      }
      .fmt{
        border-radius: 16px;
        border: 1px solid rgba(15,23,42,0.10);
        padding: 10px 10px;
        background: rgba(255,255,255,0.65);
        box-shadow: 0 10px 20px rgba(15,23,42,0.04);
      }
      .fmtDisabled{
        opacity: .40;
        filter: grayscale(.2);
      }
      .fmtTitle{
        font-size: 13px;
        font-weight: 900;
        margin:0;
      }
      .fmtSub{
        font-size: 12px;
        color: rgba(15,23,42,0.62);
        margin-top: 4px;
        margin-bottom: 0;
      }

      .stButton button, .stDownloadButton button{
        border-radius: 14px !important;
        font-weight: 900 !important;
        padding: 10px 14px !important;
      }
      div[data-testid="stFileUploader"] section{
        border-radius: 16px !important;
        border: 1px dashed rgba(15,23,42,0.22) !important;
        background: rgba(255,255,255,0.70) !important;
      }
      hr{
        border: none;
        height: 1px;
        background: rgba(15,23,42,0.08);
        margin: 10px 0;
      }
      .muted{ color: rgba(15,23,42,0.62); font-size: 12px; }

      /* Compact expander */
      details summary{
        font-weight: 900;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="nav">
      <div class="navL">
        <div class="logo"></div>
        <div>
          <div class="brand">DocFlow Converter</div>
          <div class="subtitle">Upload → Auto Detect → Choose Output → Download</div>
        </div>
      </div>
      <div class="chips">
        <div class="chip">Website-style UI</div>
        <div class="chip">One-click convert</div>
        <div class="chip">Auto-download + fallback</div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# HELPERS
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
    if ext in [".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".bmp"]:
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
    return "application/octet-stream"


def auto_download_bytes(file_name: str, data: bytes, mime: str):
    """
    Attempts to auto-trigger a browser download via HTML/JS.
    Some browsers may block it; we always provide fallback download button too.
    """
    b64 = base64.b64encode(data).decode()
    html = f"""
    <html>
      <body>
        <a id="dl" download="{file_name}" href="data:{mime};base64,{b64}"></a>
        <script>
          const a = document.getElementById('dl');
          a.click();
        </script>
      </body>
    </html>
    """
    st.components.v1.html(html, height=0)


# ============================================================
# OCR/PDF (cloud-safe: EasyOCR + PyMuPDF + pdfplumber + pdf2docx)
# ============================================================
@st.cache_resource(show_spinner=False)
def _easyocr_reader(lang_code: str):
    import easyocr
    return easyocr.Reader([lang_code], gpu=False)


def ocr_image_to_text(pil_img: Image.Image, lang_ui: str = "eng") -> str:
    reader = _easyocr_reader("en")
    arr = np.array(pil_img.convert("RGB"))
    try:
        lines = reader.readtext(arr, detail=0, paragraph=True)
    except TypeError:
        lines = reader.readtext(arr, detail=0)
    return "\n".join([t.strip() for t in lines if t and str(t).strip()]).strip()


def pdf_textlayer_extract(pdf_bytes: bytes, max_pages: int) -> List[str]:
    import pdfplumber
    out = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for p in pdf.pages[:max_pages]:
            out.append(p.extract_text() or "")
    return out


def pdf_render_pages_to_images(pdf_bytes: bytes, dpi: int, max_pages: int) -> List[Image.Image]:
    import fitz
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    imgs = []
    for i in range(min(max_pages, doc.page_count)):
        pix = doc.load_page(i).get_pixmap(dpi=int(dpi))
        imgs.append(Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB"))
    doc.close()
    return imgs


def pdf_hybrid_text_extract(pdf_bytes: bytes, max_pages: int, dpi: int) -> List[str]:
    layer = pdf_textlayer_extract(pdf_bytes, max_pages=max_pages)
    needs_ocr = []
    for t in layer:
        t2 = re.sub(r"\s+", "", t or "")
        needs_ocr.append(len(t2) < 40)
    if not any(needs_ocr):
        return layer

    imgs = pdf_render_pages_to_images(pdf_bytes, dpi=dpi, max_pages=max_pages)
    out = []
    for i, base in enumerate(layer):
        if i < len(imgs) and needs_ocr[i]:
            try:
                txt = ocr_image_to_text(imgs[i], lang_ui="eng")
                out.append(txt if txt else base)
            except Exception:
                out.append(base)
        else:
            out.append(base)
    return out


def pdf_metadata_to_json(pdf_bytes: bytes) -> bytes:
    from pypdf import PdfReader
    r = PdfReader(io.BytesIO(pdf_bytes))
    md = r.metadata or {}
    out = {"page_count": len(r.pages)}
    for k, v in md.items():
        out[str(k)] = str(v) if v is not None else None
    return json.dumps(out, ensure_ascii=False, indent=2).encode("utf-8")


def pdf_to_images_zip(pdf_bytes: bytes, max_pages: int, dpi: int = 220) -> Tuple[bytes, int]:
    imgs = pdf_render_pages_to_images(pdf_bytes, dpi=dpi, max_pages=max_pages)
    files = {}
    for i, im in enumerate(imgs, start=1):
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        files[f"pages/page_{i:03d}.png"] = buf.getvalue()
    return build_zip(files), len(imgs)


def pdf_to_docx_high_fidelity(pdf_bytes: bytes) -> bytes:
    from pdf2docx import Converter
    tmp_id = datetime.utcnow().strftime("%Y%m%d_%H%M%S_%f")
    in_path = f"/tmp/in_{tmp_id}.pdf"
    out_path = f"/tmp/out_{tmp_id}.docx"
    with open(in_path, "wb") as f:
        f.write(pdf_bytes)
    try:
        cv = Converter(in_path)
        cv.convert(out_path, start=0, end=None)
        cv.close()
        with open(out_path, "rb") as f:
            return f.read()
    finally:
        for p in (in_path, out_path):
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass


# TABLES (text-layer)
def normalize_cell_text_clean(val):
    if val is None:
        return val
    s = str(val).replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"\n+", " ", s)
    s = s.replace("\u00a0", " ")
    s = re.sub(r"[ \t]+", " ", s).strip()
    return s


def extract_tables_pdf_textlayer(pdf_bytes: bytes, max_pages: int) -> List[pd.DataFrame]:
    import pdfplumber
    dfs: List[pd.DataFrame] = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages[:max_pages]:
            tbls = page.extract_tables()
            for t in tbls or []:
                if t:
                    df = pd.DataFrame(t)
                    df = df.replace("", np.nan).dropna(axis=0, how="all").dropna(axis=1, how="all").fillna("")
                    if not df.empty:
                        dfs.append(df)
    return dfs


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

    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        for i, df in enumerate(cleaned, start=1):
            df.to_excel(writer, sheet_name=f"Table_{i}"[:31], index=False, header=False)
        wb = writer.book
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=False, vertical="top")
    files[f"{base}.xlsx"] = excel_buf.getvalue()

    combined_json = {"tables": []}
    for i, df in enumerate(cleaned, start=1):
        files[f"{base}/tables/table_{i}.csv"] = df.to_csv(index=False, header=False).encode("utf-8")
        one = {"table_index": i, "rows": df_to_json_records(df)}
        files[f"{base}/tables/table_{i}.json"] = json.dumps(one, ensure_ascii=False, indent=2).encode("utf-8")
        combined_json["tables"].append(one)

    files[f"{base}/tables/combined.json"] = json.dumps(combined_json, ensure_ascii=False, indent=2).encode("utf-8")
    files[f"{base}/manifest.json"] = json.dumps({"type": "tables_export", "table_count": len(cleaned)}, indent=2).encode("utf-8")
    return files


# ============================================================
# BASIC Office conversions (keep minimal here)
# ============================================================
def docx_to_plain_text(docx_bytes: bytes) -> str:
    from docx import Document
    doc = Document(io.BytesIO(docx_bytes))
    parts = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
    return "\n".join(parts).strip()


def excel_to_pdf_bytes(xlsx_bytes: bytes, max_rows: int = 80, max_cols: int = 14) -> bytes:
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


# ============================================================
# STATE
# ============================================================
if "outputs" not in st.session_state:
    st.session_state.outputs = {}  # name->bytes

if "history" not in st.session_state:
    st.session_state.history = []  # list of dicts

if "pending_autodl" not in st.session_state:
    st.session_state.pending_autodl = None  # {"name":..., "bytes":..., "mime":...}


def push_history(task_label: str, fname: str):
    st.session_state.history.insert(0, {
        "time": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC"),
        "file": fname,
        "task": task_label,
    })
    st.session_state.history = st.session_state.history[:80]


# ============================================================
# SETTINGS (kept light)
# ============================================================
with st.sidebar:
    st.markdown("### Settings")
    with st.expander("Advanced", expanded=False):
        max_pages = st.slider("Max PDF pages", 1, 80, 12)
        ocr_dpi = st.slider("OCR DPI", 200, 400, 260, step=10)
    st.caption("Defaults are fine for most files.")


# ============================================================
# FORMAT MODEL (ALL outputs always shown; available highlighted)
# ============================================================
# Each option is an "output format" representation.
# We'll decide availability by input type.
FORMAT_CATALOG = [
    # title, subtitle, id
    ("PDF → Editable Word (DOCX)", "Best for editable documents", "pdf_docx"),
    ("PDF Tables → Excel (XLSX + ZIP)", "Extract schedules/tables to Excel", "pdf_tables"),
    ("PDF Text → TXT", "OCR/hybrid plain text", "pdf_txt"),
    ("PDF Pages → PNG (ZIP)", "Export pages as images", "pdf_pngzip"),
    ("PDF Metadata → JSON", "Technical PDF info", "pdf_meta"),

    ("Image OCR → TXT", "Scan images into text", "img_txt"),
    ("Image → PDF", "Wrap image into PDF", "img_pdf"),

    ("Word → TXT", "Extract paragraphs into text", "docx_txt"),

    ("Excel → PDF", "Printable PDF preview", "xlsx_pdf"),

    # You can expand later:
    # ("PPT → Text (ZIP)", "Extract slide text", "ppt_txt"),
    # ("PPT → Images (ZIP)", "Extract embedded images", "ppt_imgzip"),
]


def available_for_input(ftype: str) -> Dict[str, bool]:
    avail = {k: False for _, _, k in FORMAT_CATALOG}
    if ftype == "PDF":
        for k in ["pdf_docx", "pdf_tables", "pdf_txt", "pdf_pngzip", "pdf_meta"]:
            avail[k] = True
    elif ftype == "IMAGE":
        for k in ["img_txt", "img_pdf"]:
            avail[k] = True
    elif ftype == "WORD":
        for k in ["docx_txt"]:
            avail[k] = True
    elif ftype == "EXCEL":
        for k in ["xlsx_pdf"]:
            avail[k] = True
    return avail


# ============================================================
# TOP BAR LAYOUT (website converter bar)
# ============================================================
col_upload, col_formats, col_download = st.columns([1.05, 1.35, 0.95], gap="large")

with col_upload:
    st.markdown('<div class="bar">', unsafe_allow_html=True)
    st.markdown('<div class="sectionTitle">UPLOAD & DETECT</div>', unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "Upload",
        type=["pdf", "png", "jpg", "jpeg", "webp", "tif", "tiff", "bmp", "docx", "xlsx", "xlsm", "pptx"],
        label_visibility="collapsed",
    )

    if uploaded:
        filename = uploaded.name
        file_bytes = uploaded.read()
        ftype = infer_type(filename)
        base = safe_filename(os.path.splitext(filename)[0])

        st.markdown(
            f"""
            <div class="filecard">
              <div class="filetop">
                <div>
                  <div class="fname">{filename}</div>
                  <div class="fmeta">{ftype} • {len(file_bytes)/1024:.1f} KB</div>
                </div>
                <div class="pill">{ftype}</div>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        filename, file_bytes, ftype, base = None, None, "—", "output"
        st.markdown('<div class="muted">Upload a file to unlock conversions.</div>', unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)


with col_formats:
    st.markdown('<div class="bar">', unsafe_allow_html=True)
    st.markdown('<div class="sectionTitle">OUTPUT FORMATS (ALL) • AVAILABLE HIGHLIGHTED</div>', unsafe_allow_html=True)

    avail_map = available_for_input(ftype) if filename else {k: False for _, _, k in FORMAT_CATALOG}

    st.markdown('<div class="formats">', unsafe_allow_html=True)

    # Render format cards + buttons
    for title, sub, key in FORMAT_CATALOG:
        is_avail = avail_map.get(key, False)
        cls = "fmt" if is_avail else "fmt fmtDisabled"
        st.markdown(f'<div class="{cls}">', unsafe_allow_html=True)
        st.markdown(f'<div class="fmtTitle">{title}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="fmtSub">{sub}</div>', unsafe_allow_html=True)

        btn_label = "Convert & Download" if is_avail else "Not available"
        clicked = st.button(btn_label, key=f"run_{key}", disabled=not (filename and is_avail), use_container_width=True)

        st.markdown("</div>", unsafe_allow_html=True)

        # One click conversion trigger (run immediately)
        if clicked and filename and file_bytes:
            st.session_state.outputs = {}
            try:
                with st.spinner("Converting…"):
                    outputs: Dict[str, bytes] = {}

                    # PDF
                    if key == "pdf_docx":
                        docx = pdf_to_docx_high_fidelity(file_bytes)
                        out_name = f"{base}.docx"
                        outputs[out_name] = docx

                    elif key == "pdf_tables":
                        tables = extract_tables_pdf_textlayer(file_bytes, max_pages=max_pages)
                        if not tables:
                            raise RuntimeError("No tables detected from PDF text-layer. If this PDF is scanned, we can add OCR-table extraction next.")
                        root = f"{base}_tables_{now_stamp()}"
                        bundle = build_tables_bundle(tables, base=root)
                        # give a single ZIP (best website UX)
                        zip_bytes = build_zip(bundle)
                        out_name = f"{root}.zip"
                        outputs[out_name] = zip_bytes

                    elif key == "pdf_txt":
                        pages = pdf_hybrid_text_extract(file_bytes, max_pages=max_pages, dpi=ocr_dpi)
                        txt = "\n\n".join([p.strip() for p in pages if p is not None]).strip()
                        out_name = f"{base}.txt"
                        outputs[out_name] = (txt + "\n").encode("utf-8")

                    elif key == "pdf_pngzip":
                        z, _ = pdf_to_images_zip(file_bytes, max_pages=max_pages, dpi=220)
                        out_name = f"{base}_pages_{now_stamp()}.zip"
                        outputs[out_name] = z

                    elif key == "pdf_meta":
                        out_name = f"{base}_metadata.json"
                        outputs[out_name] = pdf_metadata_to_json(file_bytes)

                    # IMAGE
                    elif key == "img_txt":
                        im = Image.open(io.BytesIO(file_bytes)).convert("RGB")
                        txt = ocr_image_to_text(im, lang_ui="eng")
                        out_name = f"{base}.txt"
                        outputs[out_name] = (txt + "\n").encode("utf-8")

                    elif key == "img_pdf":
                        im = Image.open(io.BytesIO(file_bytes)).convert("RGB")
                        buf = io.BytesIO()
                        im.save(buf, format="PDF")
                        out_name = f"{base}.pdf"
                        outputs[out_name] = buf.getvalue()

                    # WORD
                    elif key == "docx_txt":
                        txt = docx_to_plain_text(file_bytes)
                        out_name = f"{base}.txt"
                        outputs[out_name] = (txt + "\n").encode("utf-8")

                    # EXCEL
                    elif key == "xlsx_pdf":
                        out_name = f"{base}.pdf"
                        outputs[out_name] = excel_to_pdf_bytes(file_bytes)

                    else:
                        raise RuntimeError("This conversion is not implemented yet.")

                st.session_state.outputs = outputs
                push_history(title, filename)

                # AUTO DOWNLOAD (first file)
                first_name = list(outputs.keys())[0]
                first_bytes = outputs[first_name]
                st.session_state.pending_autodl = {
                    "name": first_name,
                    "bytes": first_bytes,
                    "mime": mime_for(first_name),
                }

                st.success("Converted. Download should start automatically (if browser allows).")

                # force rerun so download triggers in the download panel
                st.rerun()

            except Exception as e:
                st.session_state.outputs = {}
                st.error(str(e))

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


with col_download:
    st.markdown('<div class="bar">', unsafe_allow_html=True)
    st.markdown('<div class="sectionTitle">DOWNLOAD</div>', unsafe_allow_html=True)

    # Trigger pending auto download
    if st.session_state.pending_autodl:
        p = st.session_state.pending_autodl
        # Try to auto-download
        auto_download_bytes(p["name"], p["bytes"], p["mime"])
        # Clear the flag so it doesn't repeatedly download
        st.session_state.pending_autodl = None

    if not st.session_state.outputs:
        st.markdown(
            """
            <div class="filecard">
              <div class="fname">No output yet</div>
              <div class="fmeta">Choose an available conversion and click “Convert & Download”.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown('<div class="muted">If auto-download was blocked, use the buttons below.</div>', unsafe_allow_html=True)
        for out_name, out_bytes in st.session_state.outputs.items():
            st.download_button(
                label=f"Download {out_name}",
                data=out_bytes,
                file_name=out_name,
                mime=mime_for(out_name),
                use_container_width=True,
                key=f"dl_{out_name}_{now_stamp()}",
            )

        # Optional: always offer "ZIP all"
        if len(st.session_state.outputs) > 1:
            zip_name = f"{base}_bundle_{now_stamp()}.zip"
            st.download_button(
                label="Download ALL as ZIP",
                data=build_zip(st.session_state.outputs),
                file_name=zip_name,
                mime="application/zip",
                use_container_width=True,
                key=f"dl_zip_{now_stamp()}",
            )

        if st.button("Clear output", use_container_width=True):
            st.session_state.outputs = {}
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# HISTORY (BOTTOM, DOES NOT INTERACT)
# ============================================================
st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
with st.expander("Recent conversions", expanded=False):
    if not st.session_state.history:
        st.caption("No history yet.")
    else:
        dfh = pd.DataFrame(st.session_state.history)[["time", "file", "task"]]
        st.dataframe(dfh, use_container_width=True, hide_index=True)
        st.caption("History is informational only.")
