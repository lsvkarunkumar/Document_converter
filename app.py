# app.py
# Document Converter (Web-based Streamlit, GitHub + Streamlit Cloud)
# Focus: Image table -> Excel, Scanned/Image-PDF table -> Excel, PDF -> Editable Word (high fidelity)
#
# ✅ Pure pip-only (no system installs)
# ✅ Multi-file upload, dynamic options, multi-output selection, ZIP download
#
# ---------------------------
# requirements.txt (suggested)
# ---------------------------
# streamlit==1.41.1
# pymupdf==1.24.14
# pdfplumber==0.11.4
# pdf2docx==0.5.8
# python-docx==1.1.2
# pillow==10.4.0
# pandas==2.2.2
# openpyxl==3.1.5
# opencv-python-headless==4.10.0.84
# numpy==2.0.2
# easyocr==1.7.2
#
# Notes:
# - easyocr pulls torch; Streamlit Cloud can run it but cold-start may be slower.
# - If easyocr is not installed/loads fail, OCR-based conversions will be disabled with a clear UI message.

from __future__ import annotations

import io
import os
import re
import zipfile
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import streamlit as st

from PIL import Image

import pandas as pd
import numpy as np

# PDF tools
import fitz  # PyMuPDF
import pdfplumber

# DOCX handling
from docx import Document

# High-fidelity PDF->DOCX (text PDFs)
try:
    from pdf2docx import Converter as PDF2DOCX_Converter
    PDF2DOCX_OK = True
except Exception:
    PDF2DOCX_OK = False

# CV / OCR (for images & scanned PDFs)
try:
    import cv2  # opencv-python-headless
    CV2_OK = True
except Exception:
    CV2_OK = False

try:
    import easyocr
    EASY_OCR_OK = True
except Exception:
    EASY_OCR_OK = False


# ----------------------------
# Data structures
# ----------------------------
@dataclass
class DetectedFile:
    kind: str              # pdf_text | pdf_scanned | image | docx | pptx | xlsx | csv | txt | unknown
    ext: str
    mime: str
    details: Dict[str, str]


@dataclass
class OutputArtifact:
    filename: str
    mime: str
    data: bytes
    log: str


# ----------------------------
# UI labels and conversion IDs
# ----------------------------
CONVERSIONS = {
    # PDF text-based
    "pdf_to_docx_hifi": "PDF → Editable Word (High-Fidelity)",
    "pdf_to_xlsx_tables": "PDF → Excel (Extract Tables)",
    "pdf_to_csv_tables": "PDF → CSV (Extract Tables)",
    "pdf_to_txt": "PDF → Text (Extract)",
    "pdf_to_images": "PDF → Images (PNG)",

    # Scanned / image PDF
    "scanned_pdf_table_to_xlsx": "Scanned PDF → Excel (Extract Tables)",
    "scanned_pdf_table_to_csv": "Scanned PDF → CSV (Extract Tables)",
    "scanned_pdf_ocr_to_docx": "Scanned PDF → Word (Editable OCR)",
    "scanned_pdf_ocr_to_txt": "Scanned PDF → Text (OCR)",

    # Image
    "image_table_to_xlsx": "Image → Excel (Extract Table)",
    "image_table_to_csv": "Image → CSV (Extract Table)",
    "image_ocr_to_docx": "Image → Word (Editable OCR)",
    "image_ocr_to_txt": "Image → Text (OCR)",
    "image_to_pdf": "Image → PDF",
}

FIDELITY_BADGE = {
    "pdf_to_docx_hifi": "✅ High-Fidelity",
    "image_table_to_xlsx": "⚠️ Best-Effort OCR",
    "scanned_pdf_table_to_xlsx": "⚠️ Best-Effort OCR",
    "scanned_pdf_ocr_to_docx": "⚠️ OCR Best-Effort",
    "image_ocr_to_docx": "⚠️ OCR Best-Effort",
}


# ----------------------------
# Helpers: detection
# ----------------------------
def guess_mime(ext: str) -> str:
    ext = ext.lower()
    return {
        ".pdf": "application/pdf",
        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".csv": "text/csv",
        ".txt": "text/plain",
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".webp": "image/webp",
        ".tif": "image/tiff",
        ".tiff": "image/tiff",
        ".bmp": "image/bmp",
    }.get(ext, "application/octet-stream")


def is_image_ext(ext: str) -> bool:
    return ext.lower() in {".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".bmp"}


def detect_file(file_name: str, file_bytes: bytes) -> DetectedFile:
    ext = os.path.splitext(file_name)[1].lower()
    mime = guess_mime(ext)

    if ext == ".pdf":
        # Determine if text-based or scanned:
        # Heuristic: sample first few pages -> if extracted text length is tiny, assume scanned
        text_chars = 0
        page_count = 0
        try:
            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                page_count = len(pdf.pages)
                sample_pages = pdf.pages[: min(3, page_count)]
                for p in sample_pages:
                    t = p.extract_text() or ""
                    text_chars += len(t.strip())
        except Exception:
            # If pdfplumber fails, fallback to PyMuPDF
            try:
                doc = fitz.open(stream=file_bytes, filetype="pdf")
                page_count = doc.page_count
                for i in range(min(3, page_count)):
                    text_chars += len((doc.load_page(i).get_text() or "").strip())
                doc.close()
            except Exception:
                pass

        # If no real text, likely scanned
        kind = "pdf_text" if text_chars >= 60 else "pdf_scanned"
        return DetectedFile(kind=kind, ext=ext, mime=mime, details={"pages": str(page_count), "text_chars_sample": str(text_chars)})

    if ext == ".docx":
        return DetectedFile(kind="docx", ext=ext, mime=mime, details={})

    if is_image_ext(ext):
        return DetectedFile(kind="image", ext=ext, mime=mime, details={})

    if ext == ".xlsx":
        return DetectedFile(kind="xlsx", ext=ext, mime=mime, details={})

    if ext == ".csv":
        return DetectedFile(kind="csv", ext=ext, mime=mime, details={})

    if ext == ".txt":
        return DetectedFile(kind="txt", ext=ext, mime=mime, details={})

    if ext == ".pptx":
        return DetectedFile(kind="pptx", ext=ext, mime=mime, details={})

    return DetectedFile(kind="unknown", ext=ext, mime=mime, details={})


def available_conversions(d: DetectedFile) -> List[str]:
    if d.kind == "pdf_text":
        return ["pdf_to_docx_hifi", "pdf_to_xlsx_tables", "pdf_to_csv_tables", "pdf_to_txt", "pdf_to_images"]
    if d.kind == "pdf_scanned":
        return ["scanned_pdf_table_to_xlsx", "scanned_pdf_table_to_csv", "scanned_pdf_ocr_to_docx", "scanned_pdf_ocr_to_txt", "pdf_to_images"]
    if d.kind == "image":
        return ["image_table_to_xlsx", "image_table_to_csv", "image_ocr_to_docx", "image_ocr_to_txt", "image_to_pdf"]
    if d.kind == "docx":
        # Keep minimal (pure-pip). You can add docx->txt easily.
        return ["docx_to_txt", "docx_to_pdf_best_effort"]
    if d.kind in ("xlsx", "csv", "txt", "pptx"):
        return []
    return []


# ----------------------------
# OCR & Table extraction (image/scanned PDF)
# ----------------------------
@st.cache_resource(show_spinner=False)
def get_easyocr_reader(lang_list: Tuple[str, ...]):
    # easyocr.Reader loads models; caching avoids repeated loads
    return easyocr.Reader(list(lang_list), gpu=False)


def pil_to_cv(img: Image.Image) -> np.ndarray:
    arr = np.array(img.convert("RGB"))
    return arr[:, :, ::-1]  # RGB->BGR


def cv_to_pil(img_bgr: np.ndarray) -> Image.Image:
    rgb = img_bgr[:, :, ::-1]
    return Image.fromarray(rgb)


def preprocess_for_table(img_bgr: np.ndarray, enhance: bool, deskew: bool) -> np.ndarray:
    if not CV2_OK:
        return img_bgr

    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)

    if enhance:
        # CLAHE improves local contrast
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        gray = clahe.apply(gray)
        gray = cv2.GaussianBlur(gray, (3, 3), 0)

    # Binarize
    bw = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 25, 10
    )

    if deskew:
        # Find angle based on text/lines
        coords = np.column_stack(np.where(bw > 0))
        if coords.size > 0:
            angle = cv2.minAreaRect(coords)[-1]
            if angle < -45:
                angle = -(90 + angle)
            else:
                angle = -angle
            (h, w) = bw.shape[:2]
            M = cv2.getRotationMatrix2D((w // 2, h // 2), angle, 1.0)
            bw = cv2.warpAffine(bw, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)

    return bw


def find_table_cells_bordered(bw: np.ndarray) -> List[Tuple[int, int, int, int]]:
    """
    Detect bordered table cell rectangles using line morphology.
    Returns list of bounding boxes (x, y, w, h) sorted top-to-bottom, left-to-right.
    """
    if not CV2_OK:
        return []

    h, w = bw.shape[:2]

    # Detect horizontal and vertical lines
    hor_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(10, w // 40), 1))
    ver_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(10, h // 40)))

    horizontal = cv2.erode(bw, hor_kernel, iterations=1)
    horizontal = cv2.dilate(horizontal, hor_kernel, iterations=2)

    vertical = cv2.erode(bw, ver_kernel, iterations=1)
    vertical = cv2.dilate(vertical, ver_kernel, iterations=2)

    grid = cv2.addWeighted(horizontal, 0.5, vertical, 0.5, 0.0)
    grid = cv2.dilate(grid, cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3)), iterations=1)

    # Find contours of cells (approx)
    contours, _ = cv2.findContours(grid, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

    boxes = []
    for c in contours:
        x, y, ww, hh = cv2.boundingRect(c)
        # Filter small noise; tune thresholds
        if ww < 30 or hh < 18:
            continue
        if ww > w * 0.98 and hh > h * 0.98:
            continue
        boxes.append((x, y, ww, hh))

    # De-duplicate overlapping boxes by size preference
    boxes = sorted(boxes, key=lambda b: (b[1], b[0], -(b[2] * b[3])))

    # Heuristic: keep mid-sized boxes likely cells
    # Remove very large boxes that capture the full table
    filtered = []
    for b in boxes:
        x, y, ww, hh = b
        area = ww * hh
        if area > (w * h) * 0.6:
            continue
        filtered.append(b)

    # Sort by y then x
    filtered = sorted(filtered, key=lambda b: (b[1], b[0]))
    return filtered


def ocr_image_regions(img_bgr: np.ndarray, regions: List[Tuple[int, int, int, int]], reader, conf_min: float) -> List[str]:
    """
    OCR each region and return list of strings.
    """
    texts = []
    for (x, y, w, h) in regions:
        crop = img_bgr[y:y + h, x:x + w]
        # easyocr expects RGB
        crop_rgb = crop[:, :, ::-1]
        results = reader.readtext(crop_rgb)
        # Pick best line(s) above threshold
        good = [r for r in results if len(r) >= 3 and float(r[2]) >= conf_min]
        if not good:
            texts.append("")
        else:
            good = sorted(good, key=lambda r: -float(r[2]))
            # Join multiple fragments by x-position
            good_sorted = sorted(good, key=lambda r: r[0][0][0])  # left-most x
            texts.append(" ".join([g[1] for g in good_sorted]).strip())
    return texts


def boxes_to_grid(regions: List[Tuple[int, int, int, int]], texts: List[str]) -> pd.DataFrame:
    """
    Convert detected cell boxes + OCR text into a 2D grid by clustering rows and columns.
    This is heuristic but works well for typical bordered tables.
    """
    if not regions:
        return pd.DataFrame()

    # Compute centers
    centers = np.array([(x + w / 2, y + h / 2) for (x, y, w, h) in regions], dtype=float)
    xs = centers[:, 0]
    ys = centers[:, 1]

    # Cluster rows by y (simple gap threshold)
    order = np.argsort(ys)
    row_ids = np.zeros(len(regions), dtype=int)
    row = 0
    prev_y = None
    for idx in order:
        y = ys[idx]
        if prev_y is None:
            row_ids[idx] = row
            prev_y = y
            continue
        if abs(y - prev_y) > 18:  # row gap threshold (tune)
            row += 1
        row_ids[idx] = row
        prev_y = y

    # Cluster cols by x (simple gap threshold)
    order_x = np.argsort(xs)
    col_ids = np.zeros(len(regions), dtype=int)
    col = 0
    prev_x = None
    for idx in order_x:
        x = xs[idx]
        if prev_x is None:
            col_ids[idx] = col
            prev_x = x
            continue
        if abs(x - prev_x) > 28:  # col gap threshold (tune)
            col += 1
        col_ids[idx] = col
        prev_x = x

    n_rows = int(row_ids.max()) + 1
    n_cols = int(col_ids.max()) + 1

    grid = [["" for _ in range(n_cols)] for _ in range(n_rows)]
    # Fill with nearest; if collisions, concatenate
    for i, t in enumerate(texts):
        r = int(row_ids[i]); c = int(col_ids[i])
        if grid[r][c]:
            grid[r][c] = (grid[r][c] + " " + t).strip()
        else:
            grid[r][c] = t

    # Trim empty trailing columns
    df = pd.DataFrame(grid)
    # Drop all-empty rows/cols
    df = df.replace("", np.nan)
    df = df.dropna(axis=0, how="all").dropna(axis=1, how="all").fillna("")
    return df


def extract_table_from_image(
    img: Image.Image,
    mode: str,
    enhance: bool,
    deskew: bool,
    ocr_langs: Tuple[str, ...],
    conf_min: float
) -> Tuple[pd.DataFrame, str]:
    """
    mode: 'bordered' or 'borderless'
    """
    if not (CV2_OK and EASY_OCR_OK):
        missing = []
        if not CV2_OK:
            missing.append("opencv-python-headless")
        if not EASY_OCR_OK:
            missing.append("easyocr")
        return pd.DataFrame(), f"OCR/table extraction unavailable. Missing: {', '.join(missing)}"

    img_bgr = pil_to_cv(img)
    bw = preprocess_for_table(img_bgr, enhance=enhance, deskew=deskew)

    reader = get_easyocr_reader(ocr_langs)

    if mode == "bordered":
        regions = find_table_cells_bordered(bw)
        if len(regions) < 4:
            return pd.DataFrame(), "Could not detect a bordered table grid (too few cells found). Try Borderless mode or enable Enhance/Deskew."
        texts = ocr_image_regions(img_bgr, regions, reader, conf_min=conf_min)
        df = boxes_to_grid(regions, texts)
        if df.empty:
            return df, "Detected cells but resulted in an empty table after grid building. Try different settings."
        return df, f"Extracted table: {df.shape[0]} rows × {df.shape[1]} cols (bordered mode)."

    # Borderless mode: OCR whole image and split by whitespace blocks (best-effort)
    # This is a fallback; not “same format” guaranteed.
    img_rgb = np.array(img.convert("RGB"))
    results = reader.readtext(img_rgb)
    good = [r for r in results if len(r) >= 3 and float(r[2]) >= conf_min]
    if not good:
        return pd.DataFrame(), "No OCR text found above confidence threshold."
    # Sort by y then x
    good = sorted(good, key=lambda r: (r[0][0][1], r[0][0][0]))
    lines: List[str] = []
    current = []
    last_y = None
    for box, txt, conf in good:
        y = box[0][1]
        if last_y is None or abs(y - last_y) < 14:
            current.append(txt)
        else:
            lines.append(" ".join(current))
            current = [txt]
        last_y = y
    if current:
        lines.append(" ".join(current))

    # Split lines into columns by 2+ spaces
    rows = [re.split(r"\s{2,}", ln.strip()) for ln in lines if ln.strip()]
    max_cols = max((len(r) for r in rows), default=0)
    padded = [r + [""] * (max_cols - len(r)) for r in rows]
    df = pd.DataFrame(padded).replace("", np.nan).dropna(axis=0, how="all").dropna(axis=1, how="all").fillna("")
    if df.empty:
        return df, "OCR extracted text but could not build a table grid."
    return df, f"Extracted table: {df.shape[0]} rows × {df.shape[1]} cols (borderless best-effort)."


# ----------------------------
# PDF conversions
# ----------------------------
def pdf_to_images(file_bytes: bytes) -> List[OutputArtifact]:
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    out = []
    for i in range(doc.page_count):
        page = doc.load_page(i)
        pix = page.get_pixmap(dpi=200)
        png_bytes = pix.tobytes("png")
        out.append(OutputArtifact(
            filename=f"page_{i+1:03d}.png",
            mime="image/png",
            data=png_bytes,
            log=f"Rendered page {i+1} to PNG (200 dpi)."
        ))
    doc.close()
    return out


def pdf_text_to_txt(file_bytes: bytes) -> OutputArtifact:
    texts = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for i, p in enumerate(pdf.pages):
            t = p.extract_text() or ""
            texts.append(t)
    full = "\n\n".join(texts).strip()
    return OutputArtifact(
        filename="document.txt",
        mime="text/plain",
        data=full.encode("utf-8", errors="ignore"),
        log=f"Extracted text from {len(texts)} pages."
    )


def pdf_text_to_docx_hifi(file_bytes: bytes) -> Tuple[Optional[OutputArtifact], str]:
    if not PDF2DOCX_OK:
        return None, "pdf2docx is not available. Add 'pdf2docx' to requirements.txt."

    # pdf2docx requires file path; use temp in memory via BytesIO is not supported directly
    # We'll write to a temp file in /tmp (Streamlit Cloud allows)
    tmp_id = datetime.utcnow().strftime("%Y%m%d_%H%M%S_%f")
    in_path = f"/tmp/in_{tmp_id}.pdf"
    out_path = f"/tmp/out_{tmp_id}.docx"
    try:
        with open(in_path, "wb") as f:
            f.write(file_bytes)

        cvt = PDF2DOCX_Converter(in_path)
        # Convert all pages
        cvt.convert(out_path, start=0, end=None)
        cvt.close()

        with open(out_path, "rb") as f:
            docx_bytes = f.read()

        return OutputArtifact(
            filename="document_editable.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            data=docx_bytes,
            log="Converted using pdf2docx (high-fidelity attempt)."
        ), "OK"
    except Exception as e:
        return None, f"PDF→DOCX conversion failed: {e}"
    finally:
        for p in (in_path, out_path):
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass


def pdf_tables_to_frames(file_bytes: bytes) -> List[pd.DataFrame]:
    """
    Extract tables from a text-based PDF using pdfplumber.
    Returns list of dataframes (each table).
    """
    tables = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            # extract_table returns list of rows, best-effort
            table = page.extract_table()
            if table and any(any(cell for cell in row) for row in table):
                df = pd.DataFrame(table)
                # Clean fully empty rows/cols
                df = df.replace("", np.nan).dropna(axis=0, how="all").dropna(axis=1, how="all").fillna("")
                if not df.empty:
                    tables.append(df)
    return tables


def frames_to_xlsx_bytes(frames: List[pd.DataFrame], sheet_prefix: str = "Table") -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        if not frames:
            pd.DataFrame([["No tables found"]]).to_excel(writer, index=False, header=False, sheet_name="Result")
        else:
            for i, df in enumerate(frames, start=1):
                name = f"{sheet_prefix}_{i}"
                # Excel sheet name max 31 chars
                name = name[:31]
                df.to_excel(writer, index=False, header=False, sheet_name=name)
    return bio.getvalue()


def frames_to_csv_bytes(frames: List[pd.DataFrame]) -> bytes:
    # If multiple tables, join with separators
    out = []
    for i, df in enumerate(frames, start=1):
        out.append(f"# --- TABLE {i} ---")
        out.append(df.to_csv(index=False, header=False))
    return "\n".join(out).encode("utf-8", errors="ignore")


# ----------------------------
# OCR conversions to TXT / DOCX
# ----------------------------
def ocr_image_to_txt(img: Image.Image, ocr_langs: Tuple[str, ...], conf_min: float) -> Tuple[str, str]:
    if not EASY_OCR_OK:
        return "", "easyocr not available. Add 'easyocr' to requirements.txt."
    reader = get_easyocr_reader(ocr_langs)
    arr = np.array(img.convert("RGB"))
    results = reader.readtext(arr)
    good = [r for r in results if len(r) >= 3 and float(r[2]) >= conf_min]
    if not good:
        return "", "No OCR text found above confidence threshold."
    # Sort by y then x and join lines
    good = sorted(good, key=lambda r: (r[0][0][1], r[0][0][0]))
    lines = [r[1] for r in good]
    return "\n".join(lines).strip(), f"OCR lines: {len(lines)}"


def text_to_docx_bytes(text: str, title: str = "OCR Result") -> bytes:
    doc = Document()
    doc.add_heading(title, level=1)
    for para in text.split("\n"):
        if para.strip():
            doc.add_paragraph(para.strip())
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


# ----------------------------
# ZIP packaging
# ----------------------------
def build_zip(outputs_by_file: Dict[str, List[OutputArtifact]]) -> bytes:
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, "w", compression=zipfile.ZIP_DEFLATED) as z:
        report_lines = []
        for src_name, outs in outputs_by_file.items():
            safe_src = re.sub(r"[^a-zA-Z0-9._-]+", "_", src_name).strip("_")[:80] or "file"
            base_dir = f"converted/{safe_src}"
            for art in outs:
                safe_out = re.sub(r"[^a-zA-Z0-9._-]+", "_", art.filename).strip("_")
                z.writestr(f"{base_dir}/{safe_out}", art.data)
                report_lines.append(f"[{src_name}] -> {art.filename} :: {art.log}")
        z.writestr("conversion_report.txt", "\n".join(report_lines) or "No outputs produced.")
    return bio.getvalue()


# ----------------------------
# Core conversion router
# ----------------------------
def convert_one(
    file_name: str,
    file_bytes: bytes,
    detected: DetectedFile,
    selected: List[str],
    opts: Dict
) -> List[OutputArtifact]:
    outs: List[OutputArtifact] = []

    # Common options
    ocr_langs = tuple(opts.get("ocr_langs", ("en",)))
    conf_min = float(opts.get("ocr_conf", 0.45))
    table_mode = opts.get("table_mode", "bordered")
    enhance = bool(opts.get("enhance", True))
    deskew = bool(opts.get("deskew", True))

    if "pdf_to_images" in selected:
        outs.extend(pdf_to_images(file_bytes))

    if detected.kind == "pdf_text":
        if "pdf_to_txt" in selected:
            outs.append(pdf_text_to_txt(file_bytes))

        if "pdf_to_docx_hifi" in selected:
            art, msg = pdf_text_to_docx_hifi(file_bytes)
            if art:
                outs.append(art)
            else:
                outs.append(OutputArtifact(
                    filename="pdf_to_docx_failed.txt",
                    mime="text/plain",
                    data=msg.encode("utf-8", errors="ignore"),
                    log=msg
                ))

        if "pdf_to_xlsx_tables" in selected or "pdf_to_csv_tables" in selected:
            frames = pdf_tables_to_frames(file_bytes)
            if "pdf_to_xlsx_tables" in selected:
                xlsx = frames_to_xlsx_bytes(frames, sheet_prefix="PDF_Table")
                outs.append(OutputArtifact(
                    filename="tables.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    data=xlsx,
                    log=f"Extracted {len(frames)} tables via pdfplumber."
                ))
            if "pdf_to_csv_tables" in selected:
                csvb = frames_to_csv_bytes(frames)
                outs.append(OutputArtifact(
                    filename="tables.csv",
                    mime="text/csv",
                    data=csvb,
                    log=f"Extracted {len(frames)} tables via pdfplumber."
                ))

    elif detected.kind == "pdf_scanned":
        # Render pages to images first
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        page_imgs: List[Image.Image] = []
        for i in range(doc.page_count):
            pix = doc.load_page(i).get_pixmap(dpi=250)
            page_imgs.append(Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB"))
        doc.close()

        # Table to Excel/CSV
        if "scanned_pdf_table_to_xlsx" in selected or "scanned_pdf_table_to_csv" in selected:
            frames = []
            logs = []
            for i, im in enumerate(page_imgs, start=1):
                df, lg = extract_table_from_image(
                    im,
                    mode=table_mode,
                    enhance=enhance,
                    deskew=deskew,
                    ocr_langs=ocr_langs,
                    conf_min=conf_min
                )
                logs.append(f"Page {i}: {lg}")
                if not df.empty:
                    df = df.copy()
                    # add optional page marker
                    frames.append(df)
            if "scanned_pdf_table_to_xlsx" in selected:
                xlsx = frames_to_xlsx_bytes(frames, sheet_prefix="Page")
                outs.append(OutputArtifact(
                    filename="scanned_pdf_tables.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    data=xlsx,
                    log="; ".join(logs)[:1400]
                ))
            if "scanned_pdf_table_to_csv" in selected:
                csvb = frames_to_csv_bytes(frames)
                outs.append(OutputArtifact(
                    filename="scanned_pdf_tables.csv",
                    mime="text/csv",
                    data=csvb,
                    log="; ".join(logs)[:1400]
                ))

        # OCR to TXT/DOCX
        if "scanned_pdf_ocr_to_txt" in selected or "scanned_pdf_ocr_to_docx" in selected:
            all_txt = []
            logs = []
            for i, im in enumerate(page_imgs, start=1):
                t, lg = ocr_image_to_txt(im, ocr_langs=ocr_langs, conf_min=conf_min)
                logs.append(f"Page {i}: {lg}")
                if t.strip():
                    all_txt.append(f"--- PAGE {i} ---\n{t}")
            merged = "\n\n".join(all_txt).strip()
            if "scanned_pdf_ocr_to_txt" in selected:
                outs.append(OutputArtifact(
                    filename="scanned_pdf_ocr.txt",
                    mime="text/plain",
                    data=merged.encode("utf-8", errors="ignore"),
                    log="; ".join(logs)[:1400]
                ))
            if "scanned_pdf_ocr_to_docx" in selected:
                docx_bytes = text_to_docx_bytes(merged, title="Scanned PDF OCR")
                outs.append(OutputArtifact(
                    filename="scanned_pdf_ocr.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    data=docx_bytes,
                    log="; ".join(logs)[:1400]
                ))

    elif detected.kind == "image":
        img = Image.open(io.BytesIO(file_bytes)).convert("RGB")

        if "image_to_pdf" in selected:
            bio = io.BytesIO()
            img.save(bio, format="PDF")
            outs.append(OutputArtifact(
                filename="image.pdf",
                mime="application/pdf",
                data=bio.getvalue(),
                log="Saved image as PDF."
            ))

        if "image_table_to_xlsx" in selected or "image_table_to_csv" in selected:
            df, lg = extract_table_from_image(
                img,
                mode=table_mode,
                enhance=enhance,
                deskew=deskew,
                ocr_langs=ocr_langs,
                conf_min=conf_min
            )
            if "image_table_to_xlsx" in selected:
                xlsx = frames_to_xlsx_bytes([df] if not df.empty else [], sheet_prefix="Image_Table")
                outs.append(OutputArtifact(
                    filename="image_table.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    data=xlsx,
                    log=lg
                ))
            if "image_table_to_csv" in selected:
                csvb = (df.to_csv(index=False, header=False).encode("utf-8", errors="ignore") if not df.empty else b"")
                outs.append(OutputArtifact(
                    filename="image_table.csv",
                    mime="text/csv",
                    data=csvb or b"No table found",
                    log=lg
                ))

        if "image_ocr_to_txt" in selected or "image_ocr_to_docx" in selected:
            t, lg = ocr_image_to_txt(img, ocr_langs=ocr_langs, conf_min=conf_min)
            if "image_ocr_to_txt" in selected:
                outs.append(OutputArtifact(
                    filename="image_ocr.txt",
                    mime="text/plain",
                    data=t.encode("utf-8", errors="ignore"),
                    log=lg
                ))
            if "image_ocr_to_docx" in selected:
                docx_bytes = text_to_docx_bytes(t, title="Image OCR")
                outs.append(OutputArtifact(
                    filename="image_ocr.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    data=docx_bytes,
                    log=lg
                ))

    else:
        outs.append(OutputArtifact(
            filename="not_supported.txt",
            mime="text/plain",
            data=f"Input type '{detected.kind}' not supported yet in this build.".encode("utf-8"),
            log="No conversion performed."
        ))

    return outs


# ----------------------------
# Streamlit App
# ----------------------------
st.set_page_config(page_title="Document Converter", layout="wide")
st.title("Document Converter")
st.caption("Upload files → auto-detect → choose outputs → convert → download (ZIP).")

# Capability banner
cap = []
cap.append("✅ PDF→Editable Word (High-Fidelity)" if PDF2DOCX_OK else "❌ PDF→DOCX needs pdf2docx")
cap.append("✅ OCR/Table extraction" if (CV2_OK and EASY_OCR_OK) else "❌ OCR/Table extraction needs easyocr + opencv")
st.info(" | ".join(cap))

uploaded = st.file_uploader(
    "Upload files",
    type=None,
    accept_multiple_files=True
)

if "files" not in st.session_state:
    st.session_state.files = []
if "detected" not in st.session_state:
    st.session_state.detected = {}
if "selections" not in st.session_state:
    st.session_state.selections = {}
if "results" not in st.session_state:
    st.session_state.results = {}

if uploaded:
    st.session_state.files = uploaded
    # detect cache
    for uf in uploaded:
        if uf.name not in st.session_state.detected:
            b = uf.getvalue()
            st.session_state.detected[uf.name] = detect_file(uf.name, b)
        if uf.name not in st.session_state.selections:
            d = st.session_state.detected[uf.name]
            st.session_state.selections[uf.name] = []

# Layout
left, right = st.columns([1.1, 1.0], gap="large")

with left:
    st.subheader("Files")
    if not uploaded:
        st.write("Upload one or more files to begin.")
    else:
        rows = []
        for i, uf in enumerate(st.session_state.files):
            d = st.session_state.detected[uf.name]
            rows.append({
                "File": uf.name,
                "Detected": d.kind,
                "Ext": d.ext,
                "Size (KB)": round(len(uf.getvalue()) / 1024, 1),
                "Notes": f"pages={d.details.get('pages','')}" if d.kind.startswith("pdf") else ""
            })
        df_files = pd.DataFrame(rows)
        st.dataframe(df_files, use_container_width=True, hide_index=True)

        selected_file = st.selectbox(
            "Select a file to configure",
            [uf.name for uf in st.session_state.files],
            index=0
        )

with right:
    st.subheader("Conversion Options")
    if not uploaded:
        st.write("Waiting for upload…")
    else:
        d = st.session_state.detected[selected_file]
        st.markdown(f"**Detected type:** `{d.kind}`")

        # Global options
        with st.expander("OCR / Table options (applies to OCR-based conversions)", expanded=True):
            if not (CV2_OK and EASY_OCR_OK):
                st.warning("OCR/table extraction disabled (missing easyocr or opencv). Add them to requirements.txt.")
            col1, col2 = st.columns(2)
            with col1:
                ocr_lang = st.selectbox("OCR language", ["en"], index=0)  # extend later
                table_mode = st.selectbox("Table mode", ["bordered", "borderless"], index=0,
                                          help="Bordered = has visible grid lines. Borderless = best-effort OCR.")
                conf_min = st.slider("OCR confidence threshold", 0.10, 0.90, 0.45, 0.05)
            with col2:
                enhance = st.checkbox("Enhance image", value=True, help="Improves contrast/clarity for OCR.")
                deskew = st.checkbox("Deskew (straighten)", value=True, help="Helps when photo is tilted.")
                preview_extract = st.checkbox("Preview table extraction (before converting)", value=True)

        opts = {
            "ocr_langs": (ocr_lang,),
            "table_mode": table_mode,
            "ocr_conf": conf_min,
            "enhance": enhance,
            "deskew": deskew,
        }

        # List conversions
        conv_ids = available_conversions(d)
        if not conv_ids:
            st.warning("No conversions available for this file type yet.")
        else:
            st.markdown("**Choose outputs:**")
            current = set(st.session_state.selections.get(selected_file, []))
            new_selected = set()
            for cid in conv_ids:
                label = CONVERSIONS.get(cid, cid)
                badge = FIDELITY_BADGE.get(cid, "")
                show = f"{label}  —  {badge}" if badge else label

                # disable OCR actions if missing
                needs_ocr = cid in {
                    "image_table_to_xlsx", "image_table_to_csv", "image_ocr_to_docx", "image_ocr_to_txt",
                    "scanned_pdf_table_to_xlsx", "scanned_pdf_table_to_csv", "scanned_pdf_ocr_to_docx", "scanned_pdf_ocr_to_txt",
                }
                disabled = needs_ocr and not (CV2_OK and EASY_OCR_OK)

                checked = st.checkbox(show, value=(cid in current), disabled=disabled)
                if checked:
                    new_selected.add(cid)

            st.session_state.selections[selected_file] = sorted(new_selected)

        # Apply-to-all
        st.divider()
        apply_all = st.checkbox("Apply these outputs to all uploaded files (where compatible)", value=False)
        if apply_all and uploaded:
            chosen = st.session_state.selections[selected_file]
            for uf in st.session_state.files:
                dd = st.session_state.detected[uf.name]
                allowed = set(available_conversions(dd))
                st.session_state.selections[uf.name] = [c for c in chosen if c in allowed]

        # Preview (table extraction) for the selected file
        if preview_extract and uploaded:
            if d.kind == "image":
                uf = next(u for u in st.session_state.files if u.name == selected_file)
                img = Image.open(io.BytesIO(uf.getvalue())).convert("RGB")
                st.image(img, caption="Input image (preview)", use_container_width=True)
                if ("image_table_to_xlsx" in st.session_state.selections[selected_file]) or ("image_table_to_csv" in st.session_state.selections[selected_file]):
                    df_prev, lg = extract_table_from_image(
                        img,
                        mode=table_mode,
                        enhance=enhance,
                        deskew=deskew,
                        ocr_langs=(ocr_lang,),
                        conf_min=conf_min
                    )
                    st.write(lg)
                    if not df_prev.empty:
                        st.dataframe(df_prev, use_container_width=True, hide_index=True)
            elif d.kind == "pdf_scanned":
                uf = next(u for u in st.session_state.files if u.name == selected_file)
                b = uf.getvalue()
                doc = fitz.open(stream=b, filetype="pdf")
                if doc.page_count > 0:
                    pix = doc.load_page(0).get_pixmap(dpi=220)
                    img0 = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
                    doc.close()
                    st.image(img0, caption="Page 1 preview", use_container_width=True)
                    if ("scanned_pdf_table_to_xlsx" in st.session_state.selections[selected_file]) or ("scanned_pdf_table_to_csv" in st.session_state.selections[selected_file]):
                        df_prev, lg = extract_table_from_image(
                            img0,
                            mode=table_mode,
                            enhance=enhance,
                            deskew=deskew,
                            ocr_langs=(ocr_lang,),
                            conf_min=conf_min
                        )
                        st.write(lg)
                        if not df_prev.empty:
                            st.dataframe(df_prev, use_container_width=True, hide_index=True)

# Convert button (full-width bottom)
st.divider()
convert_clicked = st.button("Convert", type="primary", use_container_width=True, disabled=not uploaded)

if convert_clicked and uploaded:
    outputs_by_file: Dict[str, List[OutputArtifact]] = {}
    progress = st.progress(0)
    status = st.empty()

    for idx, uf in enumerate(st.session_state.files, start=1):
        name = uf.name
        b = uf.getvalue()
        d = st.session_state.detected[name]
        selected = st.session_state.selections.get(name, [])

        status.write(f"Converting **{name}** ({d.kind}) …")
        if not selected:
            outputs_by_file[name] = [OutputArtifact(
                filename="skipped.txt",
                mime="text/plain",
                data=b"No outputs selected for this file.",
                log="Skipped"
            )]
        else:
            try:
                outs = convert_one(name, b, d, selected, opts)
                outputs_by_file[name] = outs
            except Exception as e:
                outputs_by_file[name] = [OutputArtifact(
                    filename="error.txt",
                    mime="text/plain",
                    data=str(e).encode("utf-8", errors="ignore"),
                    log=f"Failed: {e}"
                )]

        progress.progress(int((idx / len(st.session_state.files)) * 100))

    st.session_state.results = outputs_by_file
    status.write("✅ Done.")

# Downloads
if st.session_state.results:
    st.subheader("Downloads")

    # If only one file and one output, show direct download
    total_outputs = sum(len(v) for v in st.session_state.results.values())
    if len(st.session_state.results) == 1 and total_outputs == 1:
        only_file = next(iter(st.session_state.results))
        only_art = st.session_state.results[only_file][0]
        st.download_button(
            label=f"Download {only_art.filename}",
            data=only_art.data,
            file_name=only_art.filename,
            mime=only_art.mime,
            use_container_width=True
        )
    else:
        zip_bytes = build_zip(st.session_state.results)
        st.download_button(
            label="Download ZIP (all outputs)",
            data=zip_bytes,
            file_name="converted_outputs.zip",
            mime="application/zip",
            use_container_width=True
        )

    # Also show per-file download buttons
    with st.expander("Per-file outputs", expanded=False):
        for src, outs in st.session_state.results.items():
            st.markdown(f"**{src}**")
            for art in outs:
                st.download_button(
                    label=f"Download {art.filename}",
                    data=art.data,
                    file_name=art.filename,
                    mime=art.mime,
                    key=f"{src}__{art.filename}__dl"
                )
            st.caption("Logs: " + " | ".join([o.log for o in outs])[:1200])
