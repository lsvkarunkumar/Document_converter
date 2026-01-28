import os
import re
import io
import json
import zipfile
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple

import streamlit as st
import pandas as pd
import numpy as np
from PIL import Image


# ============================================================
# PAGE
# ============================================================
st.set_page_config(page_title="DocFlow Converter", layout="wide")

# ============================================================
# STYLES (website feel)
# ============================================================
st.markdown(
    """
    <style>
      html, body, [class*="stApp"] {
        background: radial-gradient(1100px 800px at 10% 0%, rgba(99,102,241,0.18), transparent 55%),
                    radial-gradient(900px 700px at 90% 10%, rgba(16,185,129,0.14), transparent 55%),
                    linear-gradient(180deg, #f7f8fb, #f3f4f6);
        color: #0f172a;
      }
      .block-container { max-width: 1220px; padding-top: 1rem; padding-bottom: 1.6rem; }
      #MainMenu, footer { visibility: hidden; }
      header { visibility: hidden; }

      /* Top nav */
      .nav {
        display:flex; align-items:center; justify-content:space-between;
        padding: 14px 16px;
        border-radius: 18px;
        border: 1px solid rgba(15,23,42,0.08);
        background: rgba(255,255,255,0.72);
        backdrop-filter: blur(10px);
        box-shadow: 0 10px 26px rgba(15,23,42,0.06);
        margin-bottom: 14px;
      }
      .navL { display:flex; align-items:center; gap:12px; }
      .logo {
        width: 42px; height: 42px; border-radius: 14px;
        background: linear-gradient(135deg, rgba(99,102,241,0.95), rgba(16,185,129,0.92));
        box-shadow: 0 12px 22px rgba(99,102,241,0.18);
      }
      .brand { font-size: 16px; font-weight: 900; margin:0; }
      .subtitle { margin-top:2px; font-size: 12px; color: rgba(15,23,42,0.62); }
      .navR { display:flex; gap:8px; flex-wrap:wrap; justify-content:flex-end; }
      .chip {
        font-size: 12px; padding: 7px 10px;
        border-radius: 999px;
        border: 1px solid rgba(15,23,42,0.10);
        background: rgba(255,255,255,0.70);
        color: rgba(15,23,42,0.72);
      }

      /* Panels */
      .panel {
        border-radius: 18px;
        border: 1px solid rgba(15,23,42,0.08);
        background: rgba(255,255,255,0.72);
        backdrop-filter: blur(10px);
        box-shadow: 0 12px 28px rgba(15,23,42,0.06);
        padding: 14px 14px;
      }
      .phead { display:flex; align-items:center; justify-content:space-between; margin-bottom: 10px; }
      .ptitle { font-size: 13px; font-weight: 900; margin:0; color: rgba(15,23,42,0.90); }
      .psub { font-size: 12px; color: rgba(15,23,42,0.58); margin:0; }

      /* File card */
      .filecard {
        border-radius: 16px;
        border: 1px solid rgba(15,23,42,0.08);
        background: rgba(15,23,42,0.02);
        padding: 12px 12px;
        margin-top: 10px;
      }
      .filetop { display:flex; justify-content:space-between; gap:12px; align-items:flex-start; }
      .fname { font-weight: 900; font-size: 13px; }
      .fmeta { font-size: 12px; color: rgba(15,23,42,0.62); }
      .badgeRow { display:flex; gap:8px; flex-wrap:wrap; margin-top: 8px; }
      .badge {
        font-size: 12px; padding: 6px 10px;
        border-radius: 999px;
        border: 1px solid rgba(15,23,42,0.10);
        background: rgba(255,255,255,0.55);
        color: rgba(15,23,42,0.72);
      }

      /* Conversion tiles */
      .tiles { display:grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap:10px; margin-top: 10px; }
      .tile {
        border-radius: 16px;
        border: 1px solid rgba(15,23,42,0.10);
        background: rgba(255,255,255,0.70);
        padding: 12px 12px;
        box-shadow: 0 10px 20px rgba(15,23,42,0.04);
      }
      .tile h4 { margin:0; font-size: 13px; font-weight: 900; }
      .tile p { margin:6px 0 0 0; font-size: 12px; color: rgba(15,23,42,0.60); }

      /* Empty state */
      .empty {
        border-radius: 16px;
        border: 1px dashed rgba(15,23,42,0.18);
        background: rgba(255,255,255,0.60);
        padding: 14px 14px;
      }
      .empty h4 { margin:0; font-size: 13px; font-weight: 900; }
      .empty p { margin:6px 0 0 0; font-size: 12px; color: rgba(15,23,42,0.62); }

      /* Widget tweaks */
      div[data-testid="stFileUploader"] section {
        border-radius: 16px !important;
        border: 1px dashed rgba(15,23,42,0.22) !important;
        background: rgba(255,255,255,0.70) !important;
      }
      .stButton button, .stDownloadButton button {
        border-radius: 14px !important;
        font-weight: 900 !important;
        padding: 10px 14px !important;
      }
      hr { border: none; height: 1px; background: rgba(15,23,42,0.08); margin: 12px 0; }
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
          <div class="subtitle">Fast, clean file conversions — web-only</div>
        </div>
      </div>
      <div class="navR">
        <div class="chip">OCR Tables → Excel</div>
        <div class="chip">PDF → Editable Word</div>
        <div class="chip">ZIP Bundles</div>
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


# ============================================================
# OCR / PDF (cloud-safe)
# ============================================================
@st.cache_resource(show_spinner=False)
def _easyocr_reader(lang_code: str):
    import easyocr
    return easyocr.Reader([lang_code], gpu=False)


def _ui_lang_to_easyocr(ui_lang: str) -> str:
    return "en" if ui_lang == "eng" else "en"


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


def ocr_image_to_text(pil_img: Image.Image, lang_ui: str = "eng") -> str:
    reader = _easyocr_reader(_ui_lang_to_easyocr(lang_ui))
    arr = np.array(pil_img.convert("RGB"))
    try:
        lines = reader.readtext(arr, detail=0, paragraph=True)
    except TypeError:
        lines = reader.readtext(arr, detail=0)
    return "\n".join([t.strip() for t in lines if t and str(t).strip()]).strip()


def pdf_hybrid_text_extract(pdf_bytes: bytes, max_pages: int, lang: str, dpi: int) -> List[str]:
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
                txt = ocr_image_to_text(imgs[i], lang_ui=lang)
                out.append(txt if txt else base)
            except Exception:
                out.append(base)
        else:
            out.append(base)
    return out


def pdf_to_images_zip(pdf_bytes: bytes, max_pages: int, dpi: int = 220) -> Tuple[bytes, int]:
    imgs = pdf_render_pages_to_images(pdf_bytes, dpi=dpi, max_pages=max_pages)
    files = {}
    for i, im in enumerate(imgs, start=1):
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        files[f"pages/page_{i:03d}.png"] = buf.getvalue()
    return build_zip(files), len(imgs)


def pdf_metadata_to_json(pdf_bytes: bytes) -> bytes:
    from pypdf import PdfReader
    r = PdfReader(io.BytesIO(pdf_bytes))
    md = r.metadata or {}
    out = {"page_count": len(r.pages)}
    for k, v in md.items():
        out[str(k)] = str(v) if v is not None else None
    return json.dumps(out, ensure_ascii=False, indent=2).encode("utf-8")


# ============================================================
# TABLES
# ============================================================
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

    for sheet in xls.sheet_names:
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
            pd.DataFrame(rows).to_excel(writer, sheet_name=f"Table_{i}"[:31], index=False, header=False)
    return out.getvalue()


def docx_to_plain_text(docx_bytes: bytes) -> str:
    from docx import Document
    doc = Document(io.BytesIO(docx_bytes))
    parts = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
    return "\n".join(parts).strip()


# ============================================================
# STATE
# ============================================================
if "outputs" not in st.session_state:
    st.session_state.outputs = {}

if "history" not in st.session_state:
    st.session_state.history = []


def push_history(task_label: str, task_key: str, fname: str):
    st.session_state.history.insert(0, {
        "time": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC"),
        "file": fname,
        "task": task_label,
        "task_key": task_key,
    })
    st.session_state.history = st.session_state.history[:80]


# ============================================================
# Sidebar (minimal)
# ============================================================
with st.sidebar:
    st.markdown("### Settings")
    with st.expander("Advanced", expanded=False):
        ocr_lang = st.selectbox("OCR language", ["eng"], index=0)
        max_pages = st.slider("Max PDF pages", 1, 80, 12)
        ocr_dpi = st.slider("OCR DPI", 200, 400, 260, step=10)
    st.caption("Keep defaults unless the scan is unclear.")

# ============================================================
# Tasks
# ============================================================
TASKS_BY_TYPE = {
    "PDF": [
        ("PDF → Editable Word (DOCX)", "pdf_to_docx"),
        ("PDF Tables → Excel/CSV/JSON (ZIP)", "pdf_tables"),
        ("PDF Text (Hybrid OCR) → TXT", "pdf_text_txt"),
        ("PDF Pages → PNG (ZIP)", "pdf_pages_png"),
        ("PDF Metadata → JSON", "pdf_meta_json"),
    ],
    "IMAGE": [
        ("Image OCR → TXT", "img_text_txt"),
        ("Image Table → Excel/CSV/JSON (ZIP)", "img_tables"),  # placeholder - can add later
        ("Image → PDF", "img_to_pdf"),
    ],
    "EXCEL": [
        ("Excel → PDF", "xlsx_to_pdf"),
        ("Excel → Word (DOCX)", "xlsx_to_docx"),
    ],
    "WORD": [
        ("Word → TXT", "docx_to_txt"),
        ("Word Tables → Excel (XLSX)", "docx_tables_to_xlsx"),
    ],
}

# ============================================================
# Layout
# ============================================================
colL, colR = st.columns([1.15, 1.0], gap="large")

with colL:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown(
        '<div class="phead"><div><div class="ptitle">Upload</div><div class="psub">Drag & drop your file</div></div></div>',
        unsafe_allow_html=True
    )

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
                <div class="badge">{ftype}</div>
              </div>
              <div class="badgeRow">
                <div class="badge">High-quality Word conversion</div>
                <div class="badge">Table extraction → Excel</div>
                <div class="badge">ZIP downloads</div>
              </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        st.markdown("<hr/>", unsafe_allow_html=True)
        st.markdown(
            '<div class="phead"><div><div class="ptitle">Choose conversion</div><div class="psub">Recommended options for your file</div></div></div>',
            unsafe_allow_html=True
        )

        opts = TASKS_BY_TYPE.get(ftype, [])
        if not opts:
            st.info("No conversions available for this file type.")
            task_label, task_key = None, None
        else:
            # tiles (primary), dropdown (fallback)
            tiles = opts[:4]
            tile_cols = st.columns(2)
            selected = st.session_state.get("selected_task_key", None)

            for i, (lab, key) in enumerate(tiles):
                with tile_cols[i % 2]:
                    st.markdown('<div class="tile">', unsafe_allow_html=True)
                    st.markdown(f"<h4>{lab}</h4>", unsafe_allow_html=True)
                    st.markdown("<p>One-click conversion with clean outputs.</p>", unsafe_allow_html=True)
                    if st.button("Select", key=f"pick_{key}", use_container_width=True):
                        st.session_state.selected_task_key = key
                    st.markdown("</div>", unsafe_allow_html=True)

            st.markdown("<div class='spacer'></div>", unsafe_allow_html=True)

            # dropdown fallback shows full list
            label_to_key = dict(opts)
            keys = [k for _, k in opts]
            default_idx = 0
            if st.session_state.get("selected_task_key") in keys:
                default_idx = keys.index(st.session_state.get("selected_task_key"))

            task_label = st.selectbox("All conversions", [l for l, _ in opts], index=default_idx)
            task_key = label_to_key.get(task_label)

    else:
        filename, file_bytes, ftype, base = None, None, "—", "output"
        task_label, task_key = None, None

        st.markdown(
            """
            <div class="empty">
              <h4>Start by uploading a file</h4>
              <p>PDF, images, Word, Excel — then choose conversion and download instantly.</p>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("</div>", unsafe_allow_html=True)

with colR:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown(
        '<div class="phead"><div><div class="ptitle">Output</div><div class="psub">Convert & download</div></div></div>',
        unsafe_allow_html=True
    )

    convert_disabled = not (uploaded and task_key)
    c1, c2 = st.columns([0.7, 0.3])
    with c1:
        convert_btn = st.button("Convert", type="primary", disabled=convert_disabled, use_container_width=True)
    with c2:
        clear_btn = st.button("Clear", disabled=not bool(st.session_state.outputs), use_container_width=True)
        if clear_btn:
            st.session_state.outputs = {}
            st.rerun()

    st.markdown("<hr/>", unsafe_allow_html=True)

    if not st.session_state.outputs:
        st.markdown(
            """
            <div class="empty">
              <h4>Nothing here yet</h4>
              <p>Select a conversion on the left and click <b>Convert</b>. Outputs will appear here.</p>
            </div>
            """,
            unsafe_allow_html=True
        )
    else:
        for out_name, out_bytes in st.session_state.outputs.items():
            st.download_button(
                label=f"Download {out_name}",
                data=out_bytes,
                file_name=out_name,
                mime=mime_for(out_name),
                use_container_width=True,
                key=f"dl_{out_name}_{now_stamp()}",
            )

        zname = f"{safe_filename(os.path.splitext(filename)[0])}_bundle_{now_stamp()}.zip" if filename else f"bundle_{now_stamp()}.zip"
        st.download_button(
            label="Download ALL as ZIP",
            data=build_zip(st.session_state.outputs),
            file_name=zname,
            mime="application/zip",
            use_container_width=True,
            key=f"dl_zip_{now_stamp()}",
        )

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# Execute conversion
# ============================================================
if convert_btn and uploaded and task_key and file_bytes and filename:
    st.session_state.outputs = {}
    base = safe_filename(os.path.splitext(filename)[0])

    try:
        outputs: Dict[str, bytes] = {}

        if task_key == "pdf_to_docx":
            outputs[f"{base}.docx"] = pdf_to_docx_high_fidelity(file_bytes)

        elif task_key == "pdf_tables":
            tables = extract_tables_pdf_textlayer(file_bytes, max_pages=max_pages)
            if not tables:
                raise RuntimeError("No tables found in this PDF (text-layer). If scanned, we can add OCR-table mode next.")
            root = f"{base}_tables_{now_stamp()}"
            bundle = build_tables_bundle(tables, base=root)
            outputs[f"{root}.xlsx"] = bundle[f"{root}.xlsx"]
            outputs[f"{root}.zip"] = build_zip(bundle)

        elif task_key == "pdf_text_txt":
            pages = pdf_hybrid_text_extract(file_bytes, max_pages=max_pages, lang=ocr_lang, dpi=ocr_dpi)
            txt = "\n\n".join([p.strip() for p in pages if p is not None]).strip()
            outputs[f"{base}.txt"] = (txt + "\n").encode("utf-8")

        elif task_key == "pdf_pages_png":
            z, _ = pdf_to_images_zip(file_bytes, max_pages=max_pages, dpi=220)
            outputs[f"{base}_pages_{now_stamp()}.zip"] = z

        elif task_key == "pdf_meta_json":
            outputs[f"{base}_metadata.json"] = pdf_metadata_to_json(file_bytes)

        elif task_key == "img_text_txt":
            im = Image.open(io.BytesIO(file_bytes)).convert("RGB")
            txt = ocr_image_to_text(im, lang_ui=ocr_lang)
            outputs[f"{base}.txt"] = (txt + "\n").encode("utf-8")

        elif task_key == "img_to_pdf":
            im = Image.open(io.BytesIO(file_bytes)).convert("RGB")
            buf = io.BytesIO()
            im.save(buf, format="PDF")
            outputs[f"{base}.pdf"] = buf.getvalue()

        elif task_key == "xlsx_to_pdf":
            outputs[f"{base}.pdf"] = excel_to_pdf_bytes(file_bytes)

        elif task_key == "xlsx_to_docx":
            outputs[f"{base}.docx"] = excel_to_word_docx_bytes(file_bytes)

        elif task_key == "docx_to_txt":
            txt = docx_to_plain_text(file_bytes)
            outputs[f"{base}.txt"] = (txt + "\n").encode("utf-8")

        elif task_key == "docx_tables_to_xlsx":
            xlsx = word_to_excel_tables(file_bytes)
            if xlsx is None:
                raise RuntimeError("No tables found in this Word document.")
            outputs[f"{base}_tables.xlsx"] = xlsx

        else:
            raise RuntimeError("Conversion not implemented.")

        st.session_state.outputs = outputs
        push_history(task_label=task_label, task_key=task_key, fname=filename)
        st.success("Done. Downloads are ready.")

    except Exception as e:
        st.session_state.outputs = {}
        st.error(str(e))


# ============================================================
# History (bottom, clean, optional)
# ============================================================
with st.expander("Recent conversions", expanded=False):
    if not st.session_state.history:
        st.caption("No history yet.")
    else:
        dfh = pd.DataFrame(st.session_state.history)[["time", "file", "task"]]
        st.dataframe(dfh, use_container_width=True, hide_index=True)
        st.caption("History is informational only.")
