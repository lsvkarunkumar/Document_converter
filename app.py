import os
import re
import tempfile
from io import BytesIO
from typing import List

import streamlit as st
import pandas as pd
from PIL import Image

st.set_page_config(page_title="Doc Utility Hub", layout="wide")
st.title("📎 Doc Utility Hub (PDF/Image → Excel/Word/PPT/PDF)")
st.caption("Upload a file → choose conversion → preview extracted content → download result.")

# -----------------------------
# Helpers (lightweight)
# -----------------------------
def clean_text(s: str) -> str:
    s = (s or "").replace("\x00", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()

def split_pages(text: str) -> List[str]:
    if not text:
        return []
    chunks = re.split(r"--- Page \d+ ---", text)
    chunks = [clean_text(c) for c in chunks if clean_text(c)]
    return chunks

def image_to_pdf_bytes(image_bytes: bytes) -> bytes:
    img = Image.open(BytesIO(image_bytes)).convert("RGB")
    out = BytesIO()
    img.save(out, format="PDF")
    return out.getvalue()

def is_pdf(filename: str) -> bool:
    return filename.lower().endswith(".pdf")

def is_image(filename: str) -> bool:
    return any(filename.lower().endswith(ext) for ext in [".png", ".jpg", ".jpeg", ".webp"])

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

    ocr_lang = st.selectbox("OCR language", ["eng"], index=0)
    max_pages = st.slider("Max pages (PDF OCR/Text)", 1, 50, 20)
    min_conf = st.slider("Min OCR confidence (tables)", 0, 100, 50)

    st.markdown("---")
    st.write("Tips")
    st.write("- Table extraction works best when tables are clear & straight.")
    st.write("- PDF→Word works best for text-based PDFs.")
    st.write("- For scanned PDFs, use OCR text or PPT from OCR text.")

uploaded = st.file_uploader("Upload PDF/Image", type=["pdf", "png", "jpg", "jpeg", "webp"])
if not uploaded:
    st.stop()

file_bytes = uploaded.read()
filename = uploaded.name

c1, c2 = st.columns([1, 1], gap="large")

with c1:
    st.subheader("Input preview")
    if is_image(filename):
        try:
            st.image(Image.open(BytesIO(file_bytes)), use_container_width=True)
        except Exception:
            st.info("Image preview not available.")
    else:
        st.info("PDF uploaded. (Preview shown when OCR renders pages.)")

run = c2.button("Run", type="primary")
if not run:
    st.stop()

# -----------------------------
# Task implementations (lazy imports)
# -----------------------------
if task == "Image → PDF":
    if not is_image(filename):
        st.error("Please upload an image for Image → PDF.")
        st.stop()

    pdf_out = image_to_pdf_bytes(file_bytes)
    c2.success("Converted image to PDF.")
    c2.download_button("Download PDF", pdf_out, "output.pdf", "application/pdf")

elif task == "Extract Text (OCR) → TXT":
    # Lazy imports
    import pytesseract
    from pdf2image import convert_from_bytes

    def ocr_image_bytes(img_bytes: bytes) -> str:
        img = Image.open(BytesIO(img_bytes)).convert("RGB")
        return clean_text(pytesseract.image_to_string(img, lang=ocr_lang))

    def ocr_pdf_bytes(pdf_bytes: bytes) -> str:
        images = convert_from_bytes(pdf_bytes, dpi=200)
        images = images[:max_pages]
        parts = []
        for idx, im in enumerate(images, start=1):
            txt = clean_text(pytesseract.image_to_string(im, lang=ocr_lang))
            if txt:
                parts.append(f"--- Page {idx} ---\n{txt}")
        return clean_text("\n\n".join(parts)), images

    if is_pdf(filename):
        with st.spinner("OCR on PDF pages..."):
            text, imgs = ocr_pdf_bytes(file_bytes)
        if imgs:
            c1.image(imgs[0], caption="First page rendered for OCR", use_container_width=True)
        c2.text_area("Extracted OCR text", text, height=350)
        c2.download_button("Download TXT", (text or "").encode("utf-8"), "output.txt", "text/plain")
    else:
        with st.spinner("OCR on image..."):
            text = ocr_image_bytes(file_bytes)
        c2.text_area("Extracted OCR text", text, height=350)
        c2.download_button("Download TXT", (text or "").encode("utf-8"), "output.txt", "text/plain")

elif task == "PDF/Image → Tables → Excel (.xlsx)":
    # Lazy imports
    from img2table.ocr import TesseractOCR
    from img2table.document import PDF as Img2TablePDF
    from img2table.document import Image as Img2TableImage

    ocr = TesseractOCR(lang=ocr_lang)

    def extract_tables_from_pdf(pdf_bytes: bytes):
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

    def extract_tables_from_image(img_bytes: bytes):
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

    with st.spinner("Extracting tables..."):
        tables = extract_tables_from_pdf(file_bytes) if is_pdf(filename) else extract_tables_from_image(file_bytes)

    if not tables:
        c2.error("No tables detected.")
        st.stop()

    c2.success(f"Found {len(tables)} table(s).")

    # Preview first table
    first_df = tables[0].df.copy()
    c2.dataframe(first_df, use_container_width=True)

    # Multi-sheet Excel
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for i, t in enumerate(tables, start=1):
            t.df.to_excel(writer, sheet_name=f"Table_{i}"[:31], index=False)

    c2.download_button(
        "Download Excel (.xlsx)",
        out.getvalue(),
        "tables.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

elif task == "PDF → Word (.docx)":
    if not is_pdf(filename):
        st.error("Please upload a PDF for PDF → Word.")
        st.stop()

    # Lazy imports
    from pdf2docx import Converter
    import pdfplumber

    def looks_scanned(pdf_bytes: bytes) -> bool:
        try:
            with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
                total = 0
                for page in pdf.pages[:2]:
                    txt = (page.extract_text() or "").strip()
                    total += len(txt)
            return total < 50
        except Exception:
            return True

    scanned = looks_scanned(file_bytes)
    if scanned:
        c2.warning("This PDF looks scanned/image-based. PDF→Word may be poor. Consider OCR→TXT instead.")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as fpdf:
        fpdf.write(file_bytes)
        pdf_path = fpdf.name

    docx_path = pdf_path.replace(".pdf", ".docx")
    try:
        with st.spinner("Converting PDF → DOCX..."):
            cv = Converter(pdf_path)
            cv.convert(docx_path, start=0, end=None)
            cv.close()

        with open(docx_path, "rb") as fdoc:
            docx_bytes = fdoc.read()

        c2.success("Converted PDF to Word.")
        c2.download_button(
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

elif task == "PDF → PPT (.pptx)":
    if not is_pdf(filename):
        st.error("Please upload a PDF for PDF → PPT.")
        st.stop()

    # Lazy imports
    import pdfplumber
    import pytesseract
    from pdf2image import convert_from_bytes
    from pptx import Presentation
    from pptx.util import Pt

    def looks_scanned(pdf_bytes: bytes) -> bool:
        try:
            with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
                total = 0
                for page in pdf.pages[:2]:
                    txt = (page.extract_text() or "").strip()
                    total += len(txt)
            return total < 50
        except Exception:
            return True

    def extract_text_textpdf(pdf_bytes: bytes) -> str:
        parts = []
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            for i, page in enumerate(pdf.pages[:max_pages], start=1):
                txt = clean_text(page.extract_text() or "")
                if txt:
                    parts.append(f"--- Page {i} ---\n{txt}")
        return clean_text("\n\n".join(parts))

    def extract_text_ocr(pdf_bytes: bytes):
        images = convert_from_bytes(pdf_bytes, dpi=200)[:max_pages]
        parts = []
        for i, im in enumerate(images, start=1):
            txt = clean_text(pytesseract.image_to_string(im, lang=ocr_lang))
            if txt:
                parts.append(f"--- Page {i} ---\n{txt}")
        return clean_text("\n\n".join(parts)), images

    scanned = looks_scanned(file_bytes)

    with st.spinner("Extracting text..."):
        if scanned:
            text, imgs = extract_text_ocr(file_bytes)
            if imgs:
                c1.image(imgs[0], caption="First page rendered for OCR", use_container_width=True)
        else:
            text = extract_text_textpdf(file_bytes)

    pages = split_pages(text)
    if not pages:
        pages = [text] if text else ["(No text extracted)"]

    with st.spinner("Building PPTX..."):
        prs = Presentation()
        # Title slide
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "PDF → PPT"
        slide.placeholders[1].text = "Generated from extracted text/OCR"

        for idx, page_text in enumerate(pages[:max_pages], start=1):
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = f"Page {idx}"
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
        pptx_bytes = out.getvalue()

    c2.success("Created PPT from extracted text.")
    c2.download_button(
        "Download PPTX",
        pptx_bytes,
        "output.pptx",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
    c2.text_area("Text used (preview)", text[:8000], height=250)

else:
    st.error("Unknown task.")
