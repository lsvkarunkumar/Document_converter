import os
import re
import tempfile
from io import BytesIO
from typing import List, Tuple

import pandas as pd
import streamlit as st
from PIL import Image

import pytesseract
from pdf2image import convert_from_bytes
import pdfplumber

from img2table.document import Image as Img2TableImage
from img2table.document import PDF as Img2TablePDF
from img2table.ocr import TesseractOCR

from pdf2docx import Converter
from pptx import Presentation
from pptx.util import Inches, Pt
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4


# -----------------------------
# Helpers
# -----------------------------
def clean_text(s: str) -> str:
    s = s.replace("\x00", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def is_probably_scanned_pdf(pdf_bytes: bytes, max_pages_check: int = 2) -> bool:
    """Heuristic: if pdfplumber extracts almost no text from first pages, treat as scanned."""
    try:
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            pages = min(len(pdf.pages), max_pages_check)
            total = 0
            for i in range(pages):
                txt = pdf.pages[i].extract_text() or ""
                total += len(txt.strip())
        return total < 50
    except Exception:
        return True


def ocr_image_bytes(image_bytes: bytes, lang: str = "eng") -> str:
    img = Image.open(BytesIO(image_bytes)).convert("RGB")
    return clean_text(pytesseract.image_to_string(img, lang=lang))


def ocr_pdf_bytes(pdf_bytes: bytes, lang: str = "eng", max_pages: int = 20) -> Tuple[str, List[Image.Image]]:
    """Convert PDF pages to images and OCR them. Returns (text, page_images)."""
    images = convert_from_bytes(pdf_bytes, dpi=200)
    images = images[:max_pages]
    parts = []
    for idx, im in enumerate(images, start=1):
        txt = pytesseract.image_to_string(im, lang=lang)
        txt = clean_text(txt)
        if txt:
            parts.append(f"--- Page {idx} ---\n{txt}")
    return clean_text("\n\n".join(parts)), images


def pdf_text_bytes(pdf_bytes: bytes, max_pages: int = 50) -> str:
    """Extract text from text-based PDF using pdfplumber."""
    parts = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages[:max_pages], start=1):
            txt = page.extract_text() or ""
            txt = clean_text(txt)
            if txt:
                parts.append(f"--- Page {i} ---\n{txt}")
    return clean_text("\n\n".join(parts))


def tables_to_excel_bytes(tables: List, sheet_prefix: str = "Table") -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for i, t in enumerate(tables, start=1):
            df = t.df.copy()
            sheet = f"{sheet_prefix}_{i}"[:31]
            df.to_excel(writer, sheet_name=sheet, index=False)
    return out.getvalue()


def image_to_pdf_bytes(image_bytes: bytes) -> bytes:
    img = Image.open(BytesIO(image_bytes)).convert("RGB")
    out = BytesIO()
    img.save(out, format="PDF")
    return out.getvalue()


def text_to_pdf_bytes(text: str) -> bytes:
    """Simple text->PDF for readable output (optional helper)."""
    out = BytesIO()
    c = canvas.Canvas(out, pagesize=A4)
    width, height = A4
    margin = 40
    y = height - margin
    lines = (text or "").splitlines() if text else ["(No text)"]

    c.setFont("Helvetica", 10)
    for line in lines:
        if y < margin:
            c.showPage()
            c.setFont("Helvetica", 10)
            y = height - margin
        c.drawString(margin, y, line[:140])  # simple clipping
        y -= 12

    c.save()
    return out.getvalue()


def pdf_to_docx_bytes(pdf_bytes: bytes) -> bytes:
    """Convert PDF to DOCX using pdf2docx (best for text-based PDFs)."""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as fpdf:
        fpdf.write(pdf_bytes)
        pdf_path = fpdf.name

    docx_path = pdf_path.replace(".pdf", ".docx")
    try:
        cv = Converter(pdf_path)
        # Convert all pages
        cv.convert(docx_path, start=0, end=None)
        cv.close()
        with open(docx_path, "rb") as fdocx:
            return fdocx.read()
    finally:
        for p in [pdf_path, docx_path]:
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass


def text_to_pptx_bytes(title: str, pages: List[str]) -> bytes:
    prs = Presentation()
    # Title slide
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    slide.placeholders[1].text = "Generated from PDF text/OCR"

    # Content slides
    for idx, page_text in enumerate(pages, start=1):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = f"Page {idx}"
        tf = slide.placeholders[1].text_frame
        tf.clear()

        # Add bullets (limit to keep slides readable)
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
    return out.getvalue()


def split_into_page_blocks(text: str) -> List[str]:
    """Split combined text that has '--- Page N ---' markers into list of page blocks."""
    if not text:
        return []
    chunks = re.split(r"--- Page \d+ ---", text)
    chunks = [clean_text(c) for c in chunks if clean_text(c)]
    return chunks


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="Doc Utility Hub", layout="wide")
st.title("📎 Doc Utility Hub (PDF/Image → Excel/Word/PPT/PDF)")
st.caption("Upload a file → choose a task → preview extracted content → download converted output.")

with st.sidebar:
    st.header("Input & Output")
    task = st.selectbox(
        "Choose conversion task",
        [
            "PDF → Word (.docx)",
            "PDF → PPT (.pptx)",
            "PDF/Image → Tables → Excel (.xlsx)",
            "Image → PDF",
            "Extract Text (OCR) → TXT",
        ],
        index=2,
    )

    ocr_lang = st.selectbox("OCR language", ["eng"], index=0)
    max_pages = st.slider("Max PDF pages to process (OCR/text)", 1, 50, 20)
    st.markdown("---")
    st.write("Notes:")
    st.write("- Table extraction uses OCR (good for scanned docs).")
    st.write("- PDF→Word is best on text-based PDFs.")
    st.write("- PDF→PPT uses extracted text/OCR to build slides.")

uploaded = st.file_uploader(
    "Upload PDF or Image",
    type=["pdf", "png", "jpg", "jpeg", "webp"],
)

if not uploaded:
    st.stop()

file_bytes = uploaded.read()
name = uploaded.name.lower()

left, right = st.columns([1, 1], gap="large")

with left:
    st.subheader("Preview")
    if name.endswith(".pdf"):
        st.info("PDF uploaded. Preview images will show only if OCR path is used.")
    else:
        try:
            img = Image.open(BytesIO(file_bytes))
            st.image(img, use_container_width=True)
        except Exception:
            st.warning("Could not preview image.")

with right:
    st.subheader("Output")
    run = st.button("Run conversion", type="primary")

if not run:
    st.stop()

# -----------------------------
# Run tasks
# -----------------------------
if task == "Image → PDF":
    if not any(name.endswith(ext) for ext in [".png", ".jpg", ".jpeg", ".webp"]):
        st.error("Please upload an image file for Image → PDF.")
        st.stop()

    pdf_out = image_to_pdf_bytes(file_bytes)
    st.success("Converted image to PDF.")
    st.download_button(
        "Download PDF",
        data=pdf_out,
        file_name="output.pdf",
        mime="application/pdf",
    )

elif task == "Extract Text (OCR) → TXT":
    if name.endswith(".pdf"):
        with st.spinner("OCR on PDF pages..."):
            text, imgs = ocr_pdf_bytes(file_bytes, lang=ocr_lang, max_pages=max_pages)
        with left:
            if imgs:
                st.image(imgs[0], caption="First page (rendered for OCR)", use_container_width=True)
        with right:
            st.text_area("Extracted text (OCR)", text, height=350)
            st.download_button(
                "Download TXT",
                data=(text or "").encode("utf-8"),
                file_name="output.txt",
                mime="text/plain",
            )
    else:
        with st.spinner("OCR on image..."):
            text = ocr_image_bytes(file_bytes, lang=ocr_lang)
        with right:
            st.text_area("Extracted text (OCR)", text, height=350)
            st.download_button(
                "Download TXT",
                data=(text or "").encode("utf-8"),
                file_name="output.txt",
                mime="text/plain",
            )

elif task == "PDF/Image → Tables → Excel (.xlsx)":
    # Use img2table with Tesseract OCR (stable)
    ocr = TesseractOCR(lang=ocr_lang)

    with st.spinner("Detecting & extracting tables..."):
        if name.endswith(".pdf"):
            # save to temp pdf (most reliable)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as f:
                f.write(file_bytes)
                pdf_path = f.name
            try:
                doc = Img2TablePDF(pdf_path)
                tables = doc.extract_tables(
                    ocr=ocr,
                    borderless_tables=True,
                    implicit_rows=True,
                    min_confidence=50,
                )
            finally:
                try:
                    os.remove(pdf_path)
                except Exception:
                    pass
        else:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as f:
                f.write(file_bytes)
                img_path = f.name
            try:
                doc = Img2TableImage(img_path)
                tables = doc.extract_tables(
                    ocr=ocr,
                    borderless_tables=True,
                    implicit_rows=True,
                    min_confidence=50,
                )
            finally:
                try:
                    os.remove(img_path)
                except Exception:
                    pass

    if not tables:
        st.error("No tables detected. Try a clearer scan/photo, or a different document.")
        st.stop()

    # Preview first table
    first_df = tables[0].df
    with right:
        st.success(f"Found {len(tables)} table(s).")
        st.dataframe(first_df, use_container_width=True)

        xlsx_bytes = tables_to_excel_bytes(tables)
        st.download_button(
            "Download Excel (.xlsx)",
            data=xlsx_bytes,
            file_name="tables.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.download_button(
            "Download First Table (.csv)",
            data=first_df.to_csv(index=False).encode("utf-8"),
            file_name="table_1.csv",
            mime="text/csv",
        )

elif task == "PDF → Word (.docx)":
    if not name.endswith(".pdf"):
        st.error("Please upload a PDF file for PDF → Word.")
        st.stop()

    scanned = is_probably_scanned_pdf(file_bytes, max_pages_check=2)

    if scanned:
        st.warning("This PDF looks scanned (image-based). PDF→Word works best on text-based PDFs.")
        st.info("Tip: Use 'Extract Text (OCR) → TXT' for scanned PDFs, or convert OCR text to PDF/Word in a later step.")

    with st.spinner("Converting PDF to DOCX..."):
        try:
            docx_bytes = pdf_to_docx_bytes(file_bytes)
        except Exception as e:
            st.error(f"Conversion failed: {e}")
            st.stop()

    st.success("Converted PDF to Word.")
    st.download_button(
        "Download DOCX",
        data=docx_bytes,
        file_name="output.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    # Also show extracted text preview (quick)
    try:
        txt = pdf_text_bytes(file_bytes, max_pages=min(max_pages, 20))
        if txt:
            st.text_area("Extracted text preview (text-based)", txt[:8000], height=250)
    except Exception:
        pass

elif task == "PDF → PPT (.pptx)":
    if not name.endswith(".pdf"):
        st.error("Please upload a PDF file for PDF → PPT.")
        st.stop()

    scanned = is_probably_scanned_pdf(file_bytes, max_pages_check=2)

    with st.spinner("Extracting text for slides..."):
        if scanned:
            text, imgs = ocr_pdf_bytes(file_bytes, lang=ocr_lang, max_pages=max_pages)
            with left:
                if imgs:
                    st.image(imgs[0], caption="First page (rendered for OCR)", use_container_width=True)
        else:
            text = pdf_text_bytes(file_bytes, max_pages=max_pages)

    pages = split_into_page_blocks(text)
    if not pages:
        pages = [text] if text else ["(No text extracted)"]

    with st.spinner("Building PPTX..."):
        pptx_bytes = text_to_pptx_bytes("PDF → PPT", pages[:min(len(pages), max_pages)])

    st.success("Created PPT from extracted text.")
    st.download_button(
        "Download PPTX",
        data=pptx_bytes,
        file_name="output.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )

    st.text_area("Text used for PPT (preview)", text[:8000], height=250)

else:
    st.error("Unknown task selected.")
