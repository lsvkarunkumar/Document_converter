import os
import re
import io
import json
import zipfile
from datetime import datetime
from typing import List, Dict, Any, Tuple

import streamlit as st
import pandas as pd
import numpy as np
from PIL import Image

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(page_title="DocFlow Converter (MVP)", layout="wide")

# ============================================================
# CLEAN MVP CSS (premium but minimal)
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
      .block-container{ max-width: 1240px; padding-top: 0.9rem; padding-bottom: 1.4rem; }
      #MainMenu, footer { visibility: hidden; }
      header { visibility: hidden; }

      .top{
        display:flex; align-items:center; justify-content:space-between;
        padding: 12px 14px;
        border-radius: 18px;
        border: 1px solid rgba(15,23,42,0.08);
        background: rgba(255,255,255,0.72);
        backdrop-filter: blur(10px);
        box-shadow: 0 10px 26px rgba(15,23,42,0.06);
        margin-bottom: 12px;
      }
      .left{ display:flex; align-items:center; gap:12px; }
      .logo{
        width: 40px; height: 40px; border-radius: 14px;
        background: linear-gradient(135deg, rgba(99,102,241,0.95), rgba(16,185,129,0.92));
        box-shadow: 0 12px 22px rgba(99,102,241,0.18);
      }
      .brand{ font-size: 16px; font-weight: 900; margin:0; }
      .sub{ margin-top:2px; font-size: 12px; color: rgba(15,23,42,0.62); }

      .chipRow{ display:flex; gap:8px; flex-wrap:wrap; justify-content:flex-end; }
      .chip{
        font-size: 12px; padding: 7px 10px;
        border-radius: 999px;
        border: 1px solid rgba(15,23,42,0.10);
        background: rgba(255,255,255,0.70);
        color: rgba(15,23,42,0.72);
      }

      .panel{
        border-radius: 18px;
        border: 1px solid rgba(15,23,42,0.08);
        background: rgba(255,255,255,0.72);
        backdrop-filter: blur(10px);
        box-shadow: 0 12px 28px rgba(15,23,42,0.06);
        padding: 12px 12px;
      }
      .title{
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

      .muted{ color: rgba(15,23,42,0.62); font-size: 12px; }
      hr{ border: none; height: 1px; background: rgba(15,23,42,0.08); margin: 10px 0; }

      div[data-testid="stFileUploader"] section{
        border-radius: 16px !important;
        border: 1px dashed rgba(15,23,42,0.22) !important;
        background: rgba(255,255,255,0.70) !important;
      }
      .stButton button, .stDownloadButton button{
        border-radius: 14px !important;
        font-weight: 900 !important;
        padding: 10px 14px !important;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="top">
      <div class="left">
        <div class="logo"></div>
        <div>
          <div class="brand">DocFlow Converter</div>
          <div class="sub">Phase-1 MVP • Reliable conversions • FAST & HQ DOCX modes</div>
        </div>
      </div>
      <div class="chipRow">
        <div class="chip">FAST DOCX</div>
        <div class="chip">HQ DOCX (page-capped)</div>
        <div class="chip">Tables → Excel ZIP</div>
        <div class="chip">OCR Text</div>
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
def _easyocr_reader():
    import easyocr
    return easyocr.Reader(["en"], gpu=False)


def ocr_image_to_text(pil_img: Image.Image) -> str:
    reader = _easyocr_reader()
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
                txt = ocr_image_to_text(imgs[i])
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


# ============================================================
# DOCX CONVERSION (FAST + HQ page-capped)
# ============================================================
def pdf_to_docx_fast_text(pdf_bytes: bytes, max_pages: int, dpi: int) -> bytes:
    """
    FAST DOCX: extract text (hybrid) and write to docx.
    Formatting is basic but speed is excellent.
    """
    from docx import Document
    pages = pdf_hybrid_text_extract(pdf_bytes, max_pages=max_pages, dpi=dpi)

    doc = Document()
    for p in pages:
        p = (p or "").strip()
        if not p:
            continue
        for line in p.splitlines():
            line = line.strip()
            if line:
                doc.add_paragraph(line)
        doc.add_paragraph("")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def pdf_to_docx_high_fidelity(pdf_bytes: bytes, max_pages: int = 20) -> bytes:
    """
    High fidelity conversion using pdf2docx (can be slow).
    We cap pages to avoid endless runs.
    """
    from pdf2docx import Converter
    import fitz

    tmp_id = datetime.utcnow().strftime("%Y%m%d_%H%M%S_%f")
    in_path = f"/tmp/in_{tmp_id}.pdf"
    out_path = f"/tmp/out_{tmp_id}.docx"

    with open(in_path, "wb") as f:
        f.write(pdf_bytes)

    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        page_count = doc.page_count
        doc.close()

        end = min(max_pages, page_count) - 1
        if end < 0:
            raise RuntimeError("PDF has no pages.")

        cv = Converter(in_path)
        cv.convert(out_path, start=0, end=end)
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
# TABLES (digital PDF)
# ============================================================
def normalize_cell_text(val):
    if val is None:
        return val
    s = str(val).replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    s = s.replace("\u00a0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def extract_tables_pdf_textlayer(pdf_bytes: bytes, max_pages: int) -> List[pd.DataFrame]:
    import pdfplumber
    dfs: List[pd.DataFrame] = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages[:max_pages]:
            tbls = page.extract_tables() or []
            for t in tbls:
                if not t:
                    continue
                df = pd.DataFrame(t)
                df = df.replace("", np.nan).dropna(axis=0, how="all").dropna(axis=1, how="all").fillna("")
                if not df.empty:
                    dfs.append(df.applymap(normalize_cell_text))
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


def build_tables_zip(tables: List[pd.DataFrame], base: str) -> Tuple[str, bytes]:
    """
    ZIP includes XLSX + CSV/JSON + manifest
    """
    from openpyxl.styles import Alignment

    files: Dict[str, bytes] = {}

    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        for i, df in enumerate(tables, start=1):
            df.to_excel(writer, sheet_name=f"Table_{i}"[:31], index=False, header=False)
        wb = writer.book
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=False, vertical="top")
    files[f"{base}.xlsx"] = excel_buf.getvalue()

    combined = {"tables": []}
    for i, df in enumerate(tables, start=1):
        files[f"{base}/tables/table_{i}.csv"] = df.to_csv(index=False, header=False).encode("utf-8")
        one = {"table_index": i, "rows": df_to_json_records(df)}
        files[f"{base}/tables/table_{i}.json"] = json.dumps(one, ensure_ascii=False, indent=2).encode("utf-8")
        combined["tables"].append(one)

    files[f"{base}/tables/combined.json"] = json.dumps(combined, ensure_ascii=False, indent=2).encode("utf-8")
    files[f"{base}/manifest.json"] = json.dumps({"type": "tables_export", "table_count": len(tables)}, indent=2).encode("utf-8")

    zip_bytes = build_zip(files)
    return f"{base}.zip", zip_bytes


# ============================================================
# EXCEL -> PDF (MVP)
# ============================================================
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
# SESSION STATE
# ============================================================
if "outputs" not in st.session_state:
    st.session_state.outputs = {}  # name -> bytes

if "history" not in st.session_state:
    st.session_state.history = []  # list dict


def push_history(task_label: str, fname: str):
    st.session_state.history.insert(0, {
        "time": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC"),
        "file": fname,
        "task": task_label,
    })
    st.session_state.history = st.session_state.history[:60]


# ============================================================
# SIDEBAR (minimal but important for speed controls)
# ============================================================
with st.sidebar:
    st.markdown("### Settings")
    max_pages = st.slider("Max PDF pages", 1, 200, 25)
    ocr_dpi = st.slider("OCR quality (DPI)", 200, 400, 260, step=10)
    st.caption("Tip: HQ DOCX is page-capped to avoid endless runs.")


# ============================================================
# TASKS
# ============================================================
TASKS_BY_TYPE = {
    "PDF": [
        ("PDF → Word (FAST, Text-based) (DOCX)", "pdf_docx_fast"),
        ("PDF → Word (HQ Layout, First N pages) (DOCX)", "pdf_docx_hq"),
        ("PDF Tables → Excel Bundle (ZIP)", "pdf_tables_zip"),
        ("PDF → Text (Hybrid OCR) (TXT)", "pdf_txt"),
        ("PDF → Pages (PNG ZIP)", "pdf_pngzip"),
        ("PDF → Metadata (JSON)", "pdf_meta"),
    ],
    "IMAGE": [
        ("Image → Text (OCR) (TXT)", "img_txt"),
        ("Image → PDF", "img_pdf"),
    ],
    "EXCEL": [
        ("Excel → PDF", "xlsx_pdf"),
    ],
    "WORD": [
        ("Word → Text (TXT)", "docx_txt"),
    ],
}

# ============================================================
# MAIN LAYOUT
# ============================================================
left, right = st.columns([1.08, 0.92], gap="large")

with left:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="title">UPLOAD</div>', unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "Upload file",
        type=["pdf", "png", "jpg", "jpeg", "webp", "tif", "tiff", "bmp", "docx", "xlsx", "xlsm"],
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

        st.markdown("<hr/>", unsafe_allow_html=True)
        st.markdown('<div class="title">AVAILABLE CONVERSIONS</div>', unsafe_allow_html=True)

        options = TASKS_BY_TYPE.get(ftype, [])
        if not options:
            st.info("No conversions available for this file type in MVP.")
            task_label, task_key = None, None
        else:
            labels = [x[0] for x in options]
            lab_to_key = {l: k for l, k in options}
            task_label = st.selectbox("Conversion", labels, label_visibility="collapsed")
            task_key = lab_to_key[task_label]

            # user guidance (short)
            if task_key == "pdf_docx_hq":
                st.markdown(f"<div class='muted'>HQ DOCX converts only first <b>{max_pages}</b> pages to keep it fast.</div>", unsafe_allow_html=True)
            elif task_key == "pdf_docx_fast":
                st.markdown(f"<div class='muted'>FAST DOCX uses text/OCR extraction. Very quick, basic formatting.</div>", unsafe_allow_html=True)
            else:
                st.markdown("<div class='muted'>Click Convert once. Downloads appear on the right.</div>", unsafe_allow_html=True)

    else:
        filename, file_bytes, ftype, base = None, None, "—", "output"
        task_label, task_key = None, None
        st.markdown('<div class="muted">Upload a file to see available conversions.</div>', unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

with right:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="title">CONVERT & DOWNLOAD</div>', unsafe_allow_html=True)

    convert_disabled = not (uploaded and task_key)
    c1, c2 = st.columns([0.68, 0.32])
    with c1:
        convert_btn = st.button("Convert", type="primary", disabled=convert_disabled, use_container_width=True)
    with c2:
        clear_btn = st.button("Clear", disabled=not bool(st.session_state.outputs), use_container_width=True)
        if clear_btn:
            st.session_state.outputs = {}
            st.rerun()

    st.markdown("<hr/>", unsafe_allow_html=True)

    # Execute conversion
    if convert_btn and uploaded and task_key and file_bytes and filename:
        st.session_state.outputs = {}
        try:
            with st.spinner("Converting…"):
                outputs: Dict[str, bytes] = {}

                # PDF -> DOCX
                if task_key == "pdf_docx_fast":
                    outputs[f"{base}_FAST.docx"] = pdf_to_docx_fast_text(
                        file_bytes, max_pages=max_pages, dpi=ocr_dpi
                    )

                elif task_key == "pdf_docx_hq":
                    outputs[f"{base}_HQ_first{max_pages}.docx"] = pdf_to_docx_high_fidelity(
                        file_bytes, max_pages=max_pages
                    )

                # PDF Tables -> ZIP
                elif task_key == "pdf_tables_zip":
                    tables = extract_tables_pdf_textlayer(file_bytes, max_pages=max_pages)
                    if not tables:
                        raise RuntimeError(
                            "No tables found (text-layer). If this PDF is scanned, OCR-table extraction is next upgrade."
                        )
                    zip_name, zip_bytes = build_tables_zip(
                        tables, base=f"{base}_tables_{now_stamp()}"
                    )
                    outputs[zip_name] = zip_bytes

                # PDF -> TXT
                elif task_key == "pdf_txt":
                    pages = pdf_hybrid_text_extract(file_bytes, max_pages=max_pages, dpi=ocr_dpi)
                    txt = "\n\n".join([p.strip() for p in pages if p is not None]).strip()
                    outputs[f"{base}.txt"] = (txt + "\n").encode("utf-8")

                # PDF pages -> PNG ZIP
                elif task_key == "pdf_pngzip":
                    z, _ = pdf_to_images_zip(file_bytes, max_pages=max_pages, dpi=220)
                    outputs[f"{base}_pages_{now_stamp()}.zip"] = z

                # PDF metadata
                elif task_key == "pdf_meta":
                    outputs[f"{base}_metadata.json"] = pdf_metadata_to_json(file_bytes)

                # IMAGE -> TXT OCR
                elif task_key == "img_txt":
                    im = Image.open(io.BytesIO(file_bytes)).convert("RGB")
                    txt = ocr_image_to_text(im)
                    outputs[f"{base}.txt"] = (txt + "\n").encode("utf-8")

                # IMAGE -> PDF
                elif task_key == "img_pdf":
                    im = Image.open(io.BytesIO(file_bytes)).convert("RGB")
                    buf = io.BytesIO()
                    im.save(buf, format="PDF")
                    outputs[f"{base}.pdf"] = buf.getvalue()

                # EXCEL -> PDF
                elif task_key == "xlsx_pdf":
                    outputs[f"{base}.pdf"] = excel_to_pdf_bytes(file_bytes)

                # WORD -> TXT
                elif task_key == "docx_txt":
                    from docx import Document
                    doc = Document(io.BytesIO(file_bytes))
                    parts = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
                    outputs[f"{base}.txt"] = ("\n".join(parts).strip() + "\n").encode("utf-8")

                else:
                    raise RuntimeError("Conversion not implemented in MVP.")

            st.session_state.outputs = outputs
            push_history(task_label, filename)
            st.success("Done. Download below.")

        except Exception as e:
            st.session_state.outputs = {}
            st.error(str(e))

    # Download area
    if not st.session_state.outputs:
        st.markdown('<div class="muted">No outputs yet. Convert to see downloads here.</div>', unsafe_allow_html=True)
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

        if st.button("Clear output", use_container_width=True):
            st.session_state.outputs = {}
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# HISTORY (BOTTOM, informational only)
# ============================================================
with st.expander("Recent conversions", expanded=False):
    if not st.session_state.history:
        st.caption("No history yet.")
    else:
        dfh = pd.DataFrame(st.session_state.history)[["time", "file", "task"]]
        st.dataframe(dfh, use_container_width=True, hide_index=True)
        st.caption("History does not affect conversions.")
