import os
import re
import tempfile
from io import BytesIO
from typing import List, Optional

import streamlit as st
import pandas as pd
from PIL import Image


st.set_page_config(page_title="Document Converter", layout="wide")
st.title("📎 Document Converter (PDF/Image → Excel/Word/PPT/PDF)")
st.caption("Text PDFs convert directly. Scanned/image PDFs need OCR to become editable text.")


def is_pdf(name: str) -> bool:
    return name.lower().endswith(".pdf")


def is_image(name: str) -> bool:
    return name.lower().endswith((".png", ".jpg", ".jpeg", ".webp"))


def pil_to_png_bytes(im: Image.Image) -> bytes:
    """Convert PIL image to PNG bytes for Streamlit display (avoids TypeError)."""
    buf = BytesIO()
    im.save(buf, format="PNG")
    return buf.getvalue()


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

    # join spaced letters / digits
    s = re.sub(r"(?:\b[A-Za-z]\b(?:\s+|$)){4,}", join_spaced_letters, s)
    s = re.sub(r"(?:\b\d\b\s+){3,}\b\d\b", lambda m: m.group(0).replace(" ", ""), s)

    # fix split words like "s tandard" -> "standard"
    for _ in range(2):
        s = re.sub(r"\b([A-Za-z])\s+([A-Za-z]{2,})\b", r"\1\2", s)

    # cleanup punctuation spacing
    s = re.sub(r"\s*([,/:\.\-\+])\s*", r"\1", s)

    # GB codes: add space between prefix and number
    s = re.sub(r"\b(GB(?:/T)?)\s*([0-9])", r"\1 \2", s, flags=re.IGNORECASE)

    return _collapse_spaces(s)


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


with st.sidebar:
    st.header("Controls")

    task = st.selectbox(
        "Task",
        [
            "PDF/Image → Tables → Excel (.xlsx)",
            "PDF → Word (.docx) (text-based PDF)",
            "Scanned PDF (image) → Word (.docx) (OCR editable text)",
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
            im = Image.open(BytesIO(file_bytes))
            left.image(pil_to_png_bytes(im), caption="Uploaded image", use_container_width=True)
        except Exception:
            st.warning("Could not preview image.")
    else:
        st.info("PDF uploaded. Preview shown for OCR operations.")

if not run:
    st.stop()

with right:
    st.subheader("Output")


# -----------------------------
# Image -> PDF
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
# OCR -> TXT
# -----------------------------
if task == "Extract Text (OCR) → TXT":
    import pytesseract

    if is_pdf(filename):
        from pdf2image import convert_from_bytes

        with st.spinner("Rendering PDF pages for OCR..."):
            images = convert_from_bytes(file_bytes, dpi=240)[:max_pages]

        if images:
            left.image(pil_to_png_bytes(images[0]), caption="First page rendered (OCR)", use_container_width=True)

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
# PDF/Image -> Tables -> Excel
# -----------------------------
if task == "PDF/Image → Tables → Excel (.xlsx)":
    from openpyxl.styles import Alignment

    normalizer = normalize_cell_text_clean if output_mode.startswith("Clean") else normalize_cell_text_raw

    tables_dfs: List[pd.DataFrame] = []

    if is_pdf(filename) and prefer_text_layer:
        with st.spinner("Trying PDF text-layer table extraction (pdfplumber)..."):
            try:
                tables_dfs = extract_tables_pdf_textlayer(file_bytes, max_pages=max_pages)
            except Exception:
                tables_dfs = []

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
        st.error("No tables could be extracted.")
        st.stop()

    cleaned_dfs = [df.applymap(normalizer) for df in tables_dfs]

    st.success(f"Extracted {len(cleaned_dfs)} table(s). Previewing Table 1:")
    st.dataframe(cleaned_dfs[0], use_container_width=True)

    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for i, df in enumerate(cleaned_dfs, start=1):
            df.to_excel(writer, sheet_name=f"Table_{i}"[:31], index=False)

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
    st.stop()


# -----------------------------
# PDF -> Word (text-based)
# -----------------------------
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


# -----------------------------
# Scanned PDF -> Word (OCR editable text)
# -----------------------------
if task == "Scanned PDF (image) → Word (.docx) (OCR editable text)":
    if not is_pdf(filename):
        st.error("Please upload a PDF for OCR → Word.")
        st.stop()

    import pytesseract
    from pdf2image import convert_from_bytes
    from docx import Document

    with st.spinner("Rendering PDF pages to images..."):
        images = convert_from_bytes(file_bytes, dpi=260)[:max_pages]

    if images:
        left.image(pil_to_png_bytes(images[0]), caption="First page rendered (OCR)", use_container_width=True)

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


# -----------------------------
# PDF -> PPT (text layer)
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
