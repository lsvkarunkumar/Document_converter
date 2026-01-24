import os
import re
import json
import zipfile
import tempfile
from io import BytesIO
from typing import List, Optional, Dict, Any

import streamlit as st
import pandas as pd
from PIL import Image


# =============================
# App config
# =============================
st.set_page_config(page_title="Document Converter", layout="wide")
st.title("📎 Document Converter (PDF/Image → Excel/Word/PPT/PDF)")
st.caption("Now includes: Excel + CSV + JSON + ZIP bundle outputs for extracted tables.")


# =============================
# Helpers: file types
# =============================
def is_pdf(name: str) -> bool:
    return name.lower().endswith(".pdf")


def is_image(name: str) -> bool:
    return name.lower().endswith((".png", ".jpg", ".jpeg", ".webp"))


# =============================
# Display helper (safe)
# =============================
def to_displayable_image(obj):
    """
    Convert various image-like objects to something Streamlit can display reliably.
    Returns either a PIL Image or a numpy array; or None if not possible.
    """
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


# =============================
# img2table output normalization
# =============================
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


# =============================
# Text cleanup: RAW vs CLEAN
# =============================
def _collapse_spaces(s: str) -> str:
    return re.sub(r"[ \t]+", " ", s).strip()


def normalize_cell_text_raw(val):
    if val is None:
        return val
    s = str(val)
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s


def normalize_cell_text_clean(val):
    """
    Strong cleaning for OCR/table text:
    - joins spaced letters/digits: "G B" -> "GB", "1 1 9 6" -> "1196"
    - fixes split-words like "s tandard" -> "standard"
    - converts newlines to spaces for Excel-friendly output
    - cleans spaces around punctuation
    """
    if val is None:
        return val
    s = str(val)

    # normalize line breaks -> spaces
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"\n+", "\n", s).replace("\n", " ")
    s = s.replace("\u00a0", " ")
    s = _collapse_spaces(s)

    # remove leading junk
    s = re.sub(r"^\|\s*", "", s)

    # join spaced letters (>=4 single-letter tokens)
    def join_spaced_letters(m):
        return m.group(0).replace(" ", "")

    s = re.sub(r"(?:\b[A-Za-z]\b(?:\s+|$)){4,}", join_spaced_letters, s)

    # join spaced digits (>=4 digits)
    s = re.sub(r"(?:\b\d\b\s+){3,}\b\d\b", lambda m: m.group(0).replace(" ", ""), s)

    # fix split-words like "s tandard" -> "standard"
    for _ in range(2):
        s = re.sub(r"\b([A-Za-z])\s+([A-Za-z]{2,})\b", r"\1\2", s)

    # clean punctuation spacing
    s = re.sub(r"\s*([,/:\.\-\+])\s*", r"\1", s)

    # GB code spacing: "GB/T1196" -> "GB/T 1196"
    s = re.sub(r"\b(GB(?:/T)?)\s*([0-9])", r"\1 \2", s, flags=re.IGNORECASE)

    return _collapse_spaces(s)


# =============================
# PDF text-layer table extraction (cleaner for text PDFs)
# =============================
def extract_tables_pdf_textlayer(pdf_bytes: bytes, max_pages: int = 20) -> List[pd.DataFrame]:
    import pdfplumber
    dfs: List[pd.DataFrame] = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages[:max_pages]:
            tbls = page.extract_tables()
            if not tbls:
                continue
            for t in tbls:
                if t:
                    dfs.append(pd.DataFrame(t))
    return dfs


# =============================
# Export helpers: CSV/JSON/ZIP
# =============================
def df_to_json_records(df: pd.DataFrame) -> List[Dict[str, Any]]:
    """
    Convert DF to list-of-dicts in a stable way, forcing string keys.
    Handles duplicate/None headers by auto-naming columns.
    """
    df2 = df.copy()

    # Fix empty/None column names
    cols = []
    for i, c in enumerate(df2.columns):
        name = str(c).strip() if c is not None else ""
        if not name or name.lower() in {"nan", "none"}:
            name = f"col_{i+1}"
        cols.append(name)
    df2.columns = cols

    # If duplicate column names exist, make them unique
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

    # Replace NaN with None for JSON
    df2 = df2.where(pd.notnull(df2), None)

    return df2.to_dict(orient="records")


def build_tables_bundle(
    tables: List[pd.DataFrame],
    normalizer,
    base_name: str = "tables",
) -> Dict[str, bytes]:
    """
    Returns a dict of filename -> bytes for:
    - Excel (multi-sheet)
    - CSV per table + combined CSV
    - JSON per table + combined JSON
    - manifest.json
    """
    from openpyxl.styles import Alignment

    files: Dict[str, bytes] = {}
    cleaned = [df.applymap(normalizer) for df in tables]

    # ---- Excel (multi-sheet)
    excel_buf = BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        for i, df in enumerate(cleaned, start=1):
            df.to_excel(writer, sheet_name=f"Table_{i}"[:31], index=False)

        wb = writer.book
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=False, vertical="top")

    files[f"{base_name}.xlsx"] = excel_buf.getvalue()

    # ---- CSV per table + combined
    combined_csv_parts = []
    for i, df in enumerate(cleaned, start=1):
        csv_bytes = df.to_csv(index=False).encode("utf-8")
        files[f"{base_name}_table_{i}.csv"] = csv_bytes

        combined_csv_parts.append(f"# --- Table {i} ---\n")
        combined_csv_parts.append(df.to_csv(index=False))

    files[f"{base_name}_combined.csv"] = "".join(combined_csv_parts).encode("utf-8")

    # ---- JSON per table + combined
    combined_json = {"tables": []}
    for i, df in enumerate(cleaned, start=1):
        records = df_to_json_records(df)
        table_json = {
            "table_index": i,
            "rows": records,
        }
        files[f"{base_name}_table_{i}.json"] = json.dumps(table_json, ensure_ascii=False, indent=2).encode("utf-8")
        combined_json["tables"].append(table_json)

    files[f"{base_name}_combined.json"] = json.dumps(combined_json, ensure_ascii=False, indent=2).encode("utf-8")

    # ---- Manifest
    manifest = {
        "bundle_type": "tables_export",
        "table_count": len(cleaned),
        "files": sorted(list(files.keys())),
        "notes": "Excel contains one sheet per table. CSV/JSON have per-table and combined exports.",
    }
    files[f"{base_name}_manifest.json"] = json.dumps(manifest, ensure_ascii=False, indent=2).encode("utf-8")

    return files


def build_zip(files: Dict[str, bytes], zip_name: str = "bundle.zip") -> bytes:
    buf = BytesIO()
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as z:
        for fname, data in files.items():
            z.writestr(fname, data)
    return buf.getvalue()


# =============================
# Sidebar
# =============================
with st.sidebar:
    st.header("Controls")

    task = st.selectbox(
        "Task",
        [
            "PDF/Image → Tables → Export (Excel/CSV/JSON/ZIP)",
            "PDF → Word (.docx) (text-based PDF)",
            "Scanned PDF (image) → Word (.docx) (OCR editable text)",
            "PDF → PPT (.pptx) (text-based)",
            "Image → PDF",
            "Extract Text (OCR) → TXT",
        ],
        index=0,
    )

    output_mode = st.selectbox(
        "Table output mode",
        ["Clean (recommended)", "Raw (as extracted)"],
        index=0,
    )

    prefer_text_layer = st.checkbox(
        "For PDF tables: prefer PDF text layer (cleaner) when possible",
        value=True,
    )

    show_ocr_preview = st.checkbox(
        "Show OCR preview image (first page)",
        value=False,
        help="OFF avoids Streamlit image display issues and is faster.",
    )

    ocr_lang = st.selectbox("OCR language", ["eng"], index=0)
    max_pages = st.slider("Max pages", 1, 50, 10)
    min_conf = st.slider("Min OCR confidence (OCR tables)", 0, 100, 50)

uploaded = st.file_uploader("Upload PDF or Image", type=["pdf", "png", "jpg", "jpeg", "webp"])
run = st.button("Run", type="primary")

if not uploaded:
    st.info("Upload a file to begin.")
    st.stop()

file_bytes = uploaded.read()
filename = uploaded.name

left, right = st.columns([1, 1], gap="large")

with left:
    st.subheader("Input preview")
    if is_image(filename):
        try:
            im = Image.open(BytesIO(file_bytes)).convert("RGB")
            left.image(im, caption="Uploaded image", use_container_width=True)
        except Exception:
            left.warning("Could not preview image.")
    else:
        left.info("PDF uploaded. Preview shown only if OCR preview is enabled.")

if not run:
    st.stop()

with right:
    st.subheader("Output")


# =============================
# Task: Image -> PDF
# =============================
if task == "Image → PDF":
    if not is_image(filename):
        st.error("Please upload an image for Image → PDF.")
        st.stop()

    img = Image.open(BytesIO(file_bytes)).convert("RGB")
    out = BytesIO()
    img.save(out, format="PDF")

    st.success("Converted image to PDF.")
    st.download_button("Download PDF", out.getvalue(), "output.pdf", "application/pdf")
    st.stop()


# =============================
# Task: OCR -> TXT
# =============================
if task == "Extract Text (OCR) → TXT":
    import pytesseract

    if is_pdf(filename):
        from pdf2image import convert_from_bytes

        with st.spinner("Rendering PDF pages for OCR..."):
            images = convert_from_bytes(file_bytes, dpi=240)[:max_pages]

        if show_ocr_preview and images:
            disp = to_displayable_image(images[0])
            if disp is not None:
                left.image(disp, caption="First page rendered (OCR)", use_container_width=True)
            else:
                left.warning("Could not preview rendered page (OCR will still work).")

        parts = []
        with st.spinner("Running OCR..."):
            for i, im in enumerate(images, start=1):
                txt = pytesseract.image_to_string(im, lang=ocr_lang).strip()
                if txt:
                    parts.append(f"--- Page {i} ---\n{txt}")

        text = "\n\n".join(parts).strip() or "(No text extracted)"
        st.text_area("Extracted OCR text", text, height=350)
        st.download_button("Download TXT", text.encode("utf-8"), "output.txt", "text/plain")
        st.stop()

    else:
        import pytesseract
        img = Image.open(BytesIO(file_bytes)).convert("RGB")
        with st.spinner("Running OCR on image..."):
            text = pytesseract.image_to_string(img, lang=ocr_lang).strip() or "(No text extracted)"
        st.text_area("Extracted OCR text", text, height=350)
        st.download_button("Download TXT", text.encode("utf-8"), "output.txt", "text/plain")
        st.stop()


# =============================
# Task: Tables -> Export Excel/CSV/JSON/ZIP
# =============================
if task == "PDF/Image → Tables → Export (Excel/CSV/JSON/ZIP)":
    normalizer = normalize_cell_text_clean if output_mode.startswith("Clean") else normalize_cell_text_raw
    tables_dfs: List[pd.DataFrame] = []

    # 1) Try PDF text-layer tables first (for PDFs only)
    if is_pdf(filename) and prefer_text_layer:
        with st.spinner("Trying PDF text-layer table extraction (pdfplumber)..."):
            try:
                tables_dfs = extract_tables_pdf_textlayer(file_bytes, max_pages=max_pages)
            except Exception:
                tables_dfs = []

    # 2) Fallback to OCR-based img2table
    if not tables_dfs:
        with st.spinner("Falling back to OCR-based table extraction (img2table + Tesseract)..."):
            from img2table.ocr import TesseractOCR
            from img2table.document import PDF as Img2TablePDF
            from img2table.document import Image as Img2TableImage

            ocr = TesseractOCR(lang=ocr_lang)

            def extract_tables_pdf(pdf_bytes: bytes):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as f:
                    f.write(pdf_bytes)
                    path = f.name
                try:
                    doc = Img2TablePDF(path)
                    return doc.extract_tables(
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

            def extract_tables_img(img_bytes: bytes):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as f:
                    f.write(img_bytes)
                    path = f.name
                try:
                    doc = Img2TableImage(path)
                    return doc.extract_tables(
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

            tables_obj = extract_tables_pdf(file_bytes) if is_pdf(filename) else extract_tables_img(file_bytes)
            tables = flatten_img2table_tables(tables_obj)

            for t in tables:
                df = table_to_df_safe(t)
                if df is not None:
                    tables_dfs.append(df)

    if not tables_dfs:
        st.error("No tables could be extracted. Try a clearer scan or adjust OCR confidence.")
        st.stop()

    # Preview first table (cleaned/raw)
    preview_df = tables_dfs[0].applymap(normalizer)
    st.success(f"Extracted {len(tables_dfs)} table(s). Previewing Table 1:")
    st.dataframe(preview_df, use_container_width=True)

    # Build bundle files
    base = "tables_export"
    files = build_tables_bundle(tables_dfs, normalizer=normalizer, base_name=base)

    # Individual downloads
    st.download_button(
        "Download Excel (.xlsx)",
        files[f"{base}.xlsx"],
        f"{base}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.download_button(
        "Download Combined CSV",
        files[f"{base}_combined.csv"],
        f"{base}_combined.csv",
        "text/csv",
    )
    st.download_button(
        "Download Combined JSON",
        files[f"{base}_combined.json"],
        f"{base}_combined.json",
        "application/json",
    )

    # ZIP download (everything)
    zip_bytes = build_zip(files, zip_name=f"{base}.zip")
    st.download_button(
        "Download ZIP bundle (Excel + all CSV + all JSON + manifest)",
        zip_bytes,
        f"{base}.zip",
        "application/zip",
    )

    st.stop()


# =============================
# Task: PDF -> Word (text-based)
# =============================
if task == "PDF → Word (.docx) (text-based PDF)":
    if not is_pdf(filename):
        st.error("Please upload a PDF for PDF → Word.")
        st.stop()

    from pdf2docx import Converter

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as f:
        f.write(file_bytes)
        pdf_path = f.name

    docx_path = pdf_path.replace(".pdf", ".docx")

    try:
        with st.spinner("Converting PDF → DOCX (text-based)..."):
            cv = Converter(pdf_path)
            cv.convert(docx_path, start=0, end=None)
            cv.close()

        with open(docx_path, "rb") as f:
            docx_bytes = f.read()

        st.success("Converted PDF to Word (text-based).")
        st.download_button(
            "Download DOCX",
            docx_bytes,
            "output.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    finally:
        for p in [pdf_path, docx_path]:
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass

    st.stop()


# =============================
# Task: Scanned PDF -> Word (OCR editable text)
# =============================
if task == "Scanned PDF (image) → Word (.docx) (OCR editable text)":
    if not is_pdf(filename):
        st.error("Please upload a PDF for OCR → Word.")
        st.stop()

    import pytesseract
    from pdf2image import convert_from_bytes
    from docx import Document

    with st.spinner("Rendering PDF pages to images..."):
        images = convert_from_bytes(file_bytes, dpi=260)[:max_pages]

    if show_ocr_preview and images:
        disp = to_displayable_image(images[0])
        if disp is not None:
            left.image(disp, caption="First page rendered (OCR)", use_container_width=True)
        else:
            left.warning("Could not preview rendered page (but OCR will still work).")

    doc = Document()
    doc.add_heading("OCR Extracted Text", level=1)
    doc.add_paragraph(f"Source: {uploaded.name}")
    doc.add_paragraph("")

    with st.spinner("Running OCR and building DOCX..."):
        for i, im in enumerate(images, start=1):
            doc.add_heading(f"Page {i}", level=2)
            txt = pytesseract.image_to_string(im, lang=ocr_lang).strip()
            txt = txt if txt else "(No text extracted)"
            for line in txt.splitlines():
                doc.add_paragraph(line)

    out = BytesIO()
    doc.save(out)

    st.success("Created DOCX with editable OCR text.")
    st.download_button(
        "Download DOCX (OCR)",
        out.getvalue(),
        "output_ocr.docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    st.stop()


# =============================
# Task: PDF -> PPT (text-based)
# =============================
if task == "PDF → PPT (.pptx) (text-based)":
    if not is_pdf(filename):
        st.error("Please upload a PDF for PDF → PPT.")
        st.stop()

    import pdfplumber
    from pptx import Presentation
    from pptx.util import Pt

    def clean_line(s: str) -> str:
        s = (s or "").replace("\x00", " ")
        return re.sub(r"[ \t]+", " ", s).strip()

    with st.spinner("Extracting text from PDF..."):
        parts = []
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            for i, page in enumerate(pdf.pages[:max_pages], start=1):
                txt = clean_line(page.extract_text() or "")
                if txt:
                    parts.append((i, txt))

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "PDF → PPT"
    slide.placeholders[1].text = "Generated from extracted PDF text"

    for page_num, page_text in parts[:max_pages]:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = f"Page {page_num}"
        tf = slide.placeholders[1].text_frame
        tf.clear()

        lines = [ln.strip() for ln in page_text.splitlines() if ln.strip()]
        lines = lines[:20] if lines else ["(No text extracted)"]

        first = True
        for ln in lines:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.text = ln[:180]
            p.level = 0
            p.font.size = Pt(14)

    out = BytesIO()
    prs.save(out)

    st.success("Created PPTX.")
    st.download_button(
        "Download PPTX",
        out.getvalue(),
        "output.pptx",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
    st.stop()


st.error("Unknown task selected.")
