import os
import re
import tempfile
from io import BytesIO
from typing import List, Optional, Tuple

import streamlit as st
import pandas as pd
from PIL import Image


# -----------------------------
# Streamlit config
# -----------------------------
st.set_page_config(page_title="Document Converter", layout="wide")
st.title("📎 Document Converter (PDF/Image → Excel/Word/PPT/PDF)")
st.caption("Upload a file → choose task → Run → preview → download.")


# -----------------------------
# Helpers: file types
# -----------------------------
def is_pdf(name: str) -> bool:
    return name.lower().endswith(".pdf")


def is_image(name: str) -> bool:
    return name.lower().endswith((".png", ".jpg", ".jpeg", ".webp"))


# -----------------------------
# Helpers: img2table output normalization
# -----------------------------
def flatten_img2table_tables(tables_obj) -> List:
    """
    img2table may return:
      - list[Table]
      - dict[page -> list[Table]]
    Return a flat list[Table].
    """
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
    """
    Convert a table-like object to a DataFrame safely.
    """
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


# -----------------------------
# TEXT CLEANING (important)
# -----------------------------
def _collapse_spaces(s: str) -> str:
    return re.sub(r"[ \t]+", " ", s).strip()


def normalize_cell_text_raw(val):
    """Minimal cleaning only."""
    if val is None:
        return val
    s = str(val)
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s


def normalize_cell_text_clean(val):
    """
    Strong cleaning for OCR/table text:
    - Convert newlines to spaces
    - Fix spaced letters: 'G B' -> 'GB', 'C o d e' -> 'Code'
    - Fix spaced digits: '1 1 9 6' -> '1196'
    - Fix split-words like 's tandard' -> 'standard'
    - Remove weird separators and extra spaces around punctuation
    """
    if val is None:
        return val
    s = str(val)

    # Normalize line breaks to spaces
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"\n+", "\n", s)
    s = s.replace("\n", " ")

    # Remove non-breaking spaces etc.
    s = s.replace("\u00a0", " ")

    # Collapse whitespace
    s = _collapse_spaces(s)

    # Remove leading junk like "| " if present
    s = re.sub(r"^\|\s*", "", s)

    # Join spaced single letters (>=4 letters)
    # e.g., "C o d e f o r" -> "Codefor" (we will add spaces back later using word rules)
    def join_spaced_letters(match):
        return match.group(0).replace(" ", "")

    s = re.sub(r"(?:\b[A-Za-z]\b(?:\s+|$)){4,}", join_spaced_letters, s)

    # Join spaced digits sequences (>=4 digits)
    s = re.sub(r"(?:\b\d\b\s+){3,}\b\d\b", lambda m: m.group(0).replace(" ", ""), s)

    # Now fix common OCR “split word” pattern: one letter + space + rest
    # e.g., "s tandard" -> "standard", "p ollutants" -> "pollutants"
    # Apply repeatedly because it may appear multiple times.
    for _ in range(2):
        s = re.sub(r"\b([A-Za-z])\s+([A-Za-z]{2,})\b", r"\1\2", s)

    # Clean spaces around punctuation: commas, periods, slashes, hyphens, plus, colon
    s = re.sub(r"\s*([,/:\.\-\+])\s*", r"\1", s)

    # Special handling for GB codes: add a space between prefix and number if missing
    # e.g., "GB/T1196" -> "GB/T 1196", "GB25465" -> "GB 25465"
    s = re.sub(r"\b(GB(?:/T)?)\s*([0-9])", r"\1 \2", s, flags=re.IGNORECASE)

    # Put spaces between words if we accidentally glued "Codefor" type strings:
    # A simple heuristic: insert space before lowercase->uppercase transitions (rare here),
    # and between long glued segments "Codeforconstruction" is hard; better to keep as is.
    # Most OCR outputs already have spaces preserved; we mainly fix broken spaces.

    s = _collapse_spaces(s)
    return s


# -----------------------------
# PDF Text-layer table extraction (better for text-based PDFs)
# -----------------------------
def extract_tables_pdf_textlayer(pdf_bytes: bytes, max_pages: int = 20) -> List[pd.DataFrame]:
    """
    Use pdfplumber to extract tables from text-based PDFs.
    Returns list of DataFrames.
    This is often MUCH cleaner than OCR for docs like standards tables.
    """
    import pdfplumber

    dfs: List[pd.DataFrame] = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages[:max_pages]:
            # Extract tables (list of list rows)
            tbls = page.extract_tables()
            if not tbls:
                continue
            for t in tbls:
                if not t:
                    continue
                df = pd.DataFrame(t)
                # Sometimes first row is header; we keep raw to avoid wrong assumptions
                dfs.append(df)
    return dfs


# -----------------------------
# Sidebar controls
# -----------------------------
with st.sidebar:
    st.header("Controls")

    task = st.selectbox(
        "Task",
        [
            "PDF/Image → Tables → Excel (.xlsx)",
            "PDF → Word (.docx)",
            "PDF → PPT (.pptx)",
            "Image → PDF",
            "Extract Text (OCR) → TXT",
        ],
        index=0,
    )

    output_mode = st.selectbox(
        "Output mode (Tables → Excel)",
        ["Clean (recommended)", "Raw (as extracted)"],
        index=0,
    )

    prefer_text_layer = st.checkbox(
        "For PDF tables: prefer PDF text layer (cleaner) when possible",
        value=True,
        help="If checked, app first tries pdfplumber tables for PDFs (no OCR). If none found, it falls back to OCR/img2table.",
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
            st.image(Image.open(BytesIO(file_bytes)), use_container_width=True)
        except Exception:
            st.warning("Could not preview image.")
    else:
        st.info("PDF uploaded. Preview shown only for OCR tasks.")

if not run:
    st.stop()

with right:
    st.subheader("Output")


# -----------------------------
# Task: Image -> PDF
# -----------------------------
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


# -----------------------------
# Task: Extract Text (OCR) -> TXT
# -----------------------------
if task == "Extract Text (OCR) → TXT":
    import pytesseract

    if is_pdf(filename):
        from pdf2image import convert_from_bytes

        with st.spinner("Rendering PDF pages for OCR..."):
            images = convert_from_bytes(file_bytes, dpi=240)[:max_pages]

        if images:
            left.image(images[0], caption="First page rendered", use_container_width=True)

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
        img = Image.open(BytesIO(file_bytes)).convert("RGB")
        with st.spinner("Running OCR on image..."):
            text = pytesseract.image_to_string(img, lang=ocr_lang).strip() or "(No text extracted)"

        st.text_area("Extracted OCR text", text, height=350)
        st.download_button("Download TXT", text.encode("utf-8"), "output.txt", "text/plain")
        st.stop()


# -----------------------------
# Task: PDF/Image -> Tables -> Excel
# -----------------------------
if task == "PDF/Image → Tables → Excel (.xlsx)":
    from openpyxl.styles import Alignment

    normalizer = normalize_cell_text_clean if output_mode.startswith("Clean") else normalize_cell_text_raw

    tables_dfs: List[pd.DataFrame] = []

    # 1) Try PDF text layer extraction (cleaner) if enabled and input is PDF
    if is_pdf(filename) and prefer_text_layer:
        with st.spinner("Trying PDF text-layer table extraction (pdfplumber)..."):
            try:
                tables_dfs = extract_tables_pdf_textlayer(file_bytes, max_pages=max_pages)
            except Exception:
                tables_dfs = []

    # If no tables found, fallback to OCR-based img2table
    if not tables_dfs:
        with st.spinner("Falling back to OCR-based table extraction (img2table + Tesseract)..."):
            # Heavy imports only when needed
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

            if not tables:
                st.error("No tables detected. Try clearer input or adjust OCR confidence.")
                st.stop()

            for t in tables:
                df = table_to_df_safe(t)
                if df is not None:
                    tables_dfs.append(df)

    if not tables_dfs:
        st.error("No tables could be extracted.")
        st.stop()

    # Clean/Raw normalize all DFs for export + preview
    cleaned_dfs = [df.applymap(normalizer) for df in tables_dfs]

    st.success(f"Extracted {len(cleaned_dfs)} table(s). Previewing Table 1:")
    st.dataframe(cleaned_dfs[0], use_container_width=True)

    # Export all tables to Excel (multi-sheet), wrap OFF
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for i, df in enumerate(cleaned_dfs, start=1):
            sheet = f"Table_{i}"[:31]
            df.to_excel(writer, sheet_name=sheet, index=False)

        wb = writer.book
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=False, vertical="top")

    st.download_button(
        "Download Excel (.xlsx)",
        out.getvalue(),
        "tables.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.download_button(
        "Download Table 1 (CSV)",
        cleaned_dfs[0].to_csv(index=False).encode("utf-8"),
        "table_1.csv",
        "text/csv",
    )
    st.stop()


# -----------------------------
# Task: PDF -> Word
# -----------------------------
if task == "PDF → Word (.docx)":
    if not is_pdf(filename):
        st.error("Please upload a PDF for PDF → Word.")
        st.stop()

    from pdf2docx import Converter

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as f:
        f.write(file_bytes)
        pdf_path = f.name

    docx_path = pdf_path.replace(".pdf", ".docx")

    try:
        with st.spinner("Converting PDF → DOCX (may take time)..."):
            cv = Converter(pdf_path)
            cv.convert(docx_path, start=0, end=None)
            cv.close()

        with open(docx_path, "rb") as f:
            docx_bytes = f.read()

        st.success("Converted PDF to Word.")
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


# -----------------------------
# Task: PDF -> PPT
# -----------------------------
if task == "PDF → PPT (.pptx)":
    if not is_pdf(filename):
        st.error("Please upload a PDF for PDF → PPT.")
        st.stop()

    import pdfplumber
    from pptx import Presentation
    from pptx.util import Pt

    def clean_line(s: str) -> str:
        s = (s or "").replace("\x00", " ")
        s = re.sub(r"[ \t]+", " ", s)
        return s.strip()

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
            if first:
                p = tf.paragraphs[0]
                first = False
            else:
                p = tf.add_paragraph()
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
