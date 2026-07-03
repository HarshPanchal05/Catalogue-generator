import hashlib
import importlib
import os
import re
import threading
import urllib.parse
import urllib.request
from io import BytesIO
from pathlib import Path
from tkinter import BooleanVar, END, StringVar, Tk, Toplevel, filedialog, messagebox
from tkinter import ttk

import pandas as pd
try:
    _pdf_module = importlib.import_module("pypdf")
except ModuleNotFoundError:
    try:
        _pdf_module = importlib.import_module("PyPDF2")
    except ModuleNotFoundError as e:
        raise ModuleNotFoundError(
            "Missing PDF dependency. Install it with: py -m pip install pypdf"
        ) from e
PdfReader = _pdf_module.PdfReader
PdfWriter = _pdf_module.PdfWriter
try:
    from pypdf.generic import ContentStream
except ImportError:
    ContentStream = None
from PIL import Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


APP_TITLE = "Catalog Generator"

PAGE_WIDTH, PAGE_HEIGHT = A4

LEFT_MARGIN = 0
RIGHT_MARGIN = 0
BOTTOM_MARGIN = 0
X_GAP = 0
Y_GAP = 2

HEADER_TOP_MARGIN = 10
HEADER_MAX_HEIGHT = 78
GRID_TOP = PAGE_HEIGHT - 108

CODE_FONT_SIZE = 9
NAME_FONT_SIZE = 8.7
LINE_HEIGHT = 10.5
TEXT_BLOCK_HEIGHT = 38

PDF_IMAGE_DPI = 220
PDF_JPEG_QUALITY = 88
WATERMARK_OPACITY = 0.32
WATERMARK_SCALE = 0.58
PDF_WATERMARK_DPI = 90
PDF_WATERMARK_QUALITY = 65


def get_layout(grid_columns=3, grid_rows=5):
    grid_columns = int(grid_columns)
    grid_rows = int(grid_rows)
    card_w = (
        PAGE_WIDTH
        - LEFT_MARGIN
        - RIGHT_MARGIN
        - ((grid_columns - 1) * X_GAP)
    ) / grid_columns
    card_h = (
        GRID_TOP
        - BOTTOM_MARGIN
        - ((grid_rows - 1) * Y_GAP)
    ) / grid_rows

    return {
        "columns": grid_columns,
        "rows": grid_rows,
        "products_per_page": grid_columns * grid_rows,
        "card_w": card_w,
        "card_h": card_h,
    }


def app_base_dir():
    return Path(getattr(__import__("sys"), "_MEIPASS", Path(__file__).parent))


def register_font(font_path=None):
    candidates = []

    if font_path:
        candidates.append(Path(font_path))

    candidates.extend(
        [
            Path.cwd() / "CanvaSans-Bold.ttf",
            app_base_dir() / "CanvaSans-Bold.ttf",
            Path(__file__).parent / "CanvaSans-Bold.ttf",
        ]
    )

    for candidate in candidates:
        if candidate.is_file():
            try:
                pdfmetrics.registerFont(TTFont("CatalogFont", str(candidate)))
                return "CatalogFont"
            except Exception:
                pass

    return "Helvetica-Bold"


def clean_filename(value):
    value = str(value).strip()
    value = re.sub(r"[^\w\- .]", "_", value)
    return value[:80] or "image"


def split_code_and_name(product_name):
    product_name = str(product_name).strip()
    match = re.match(r"^([A-Za-z0-9\-\/]+)\s+(.*)", product_name)

    if not match:
        return "", product_name

    first = match.group(1)
    remaining = match.group(2).strip()

    if any(c.isdigit() for c in first):
        return first, remaining

    return "", product_name


def excel_column_letter_to_index(value):
    value = str(value).strip().upper()

    if not re.fullmatch(r"[A-Z]+", value):
        return None

    index = 0

    for char in value:
        index = index * 26 + (ord(char) - ord("A") + 1)

    return index - 1


def resolve_excel_column(df, value, label):
    value = str(value).strip()

    if value in df.columns:
        return value

    for column in df.columns:
        if str(column).strip().lower() == value.lower():
            return column

    index = excel_column_letter_to_index(value)

    if index is not None and 0 <= index < len(df.columns):
        return df.columns[index]

    raise ValueError(
        f"{label} column not found: {value}. "
        "Use the Excel header name or a column letter like B or T."
    )


def wrap_to_width(text, font_name, font_size, max_width, max_lines):
    words = str(text).split()

    if not words:
        return []

    lines = []
    current = ""

    for word in words:
        trial = word if not current else f"{current} {word}"

        if pdfmetrics.stringWidth(trial, font_name, font_size) <= max_width:
            current = trial
            continue

        if current:
            lines.append(current)
            current = word
        else:
            lines.append(word)
            current = ""

        if len(lines) == max_lines:
            break

    if current and len(lines) < max_lines:
        lines.append(current)

    if len(lines) == max_lines:
        used_words = " ".join(lines).split()

        if len(used_words) < len(words):
            last = lines[-1]

            while (
                pdfmetrics.stringWidth(last + "...", font_name, font_size)
                > max_width
                and " " in last
            ):
                last = last.rsplit(" ", 1)[0]

            lines[-1] = last + "..."

    return lines[:max_lines]


def draw_centered_fit(pdf, font_name, text, x_center, y, max_width, font_size):
    size = font_size

    while (
        size > 7
        and pdfmetrics.stringWidth(text, font_name, size) > max_width
    ):
        size -= 0.25

    pdf.setFont(font_name, size)
    pdf.drawCentredString(x_center, y, text)


def is_url(value):
    parsed = urllib.parse.urlparse(str(value).strip())
    return parsed.scheme.lower() in {"http", "https"}


def image_cache_target(value, cache_dir, product_name):
    digest = hashlib.sha1(str(value).encode("utf-8")).hexdigest()[:16]
    parsed = urllib.parse.urlparse(str(value))
    suffix = Path(parsed.path).suffix

    if suffix.lower() not in {".jpg", ".jpeg", ".png", ".webp", ".bmp"}:
        suffix = ".jpg"

    return Path(cache_dir) / f"{clean_filename(product_name)}-{digest}{suffix}"


def local_or_download_image(value, cache_dir, product_name, download_missing=True):
    value = str(value).strip()

    if not value:
        return None, "Image URL/path is empty."

    if is_url(value):
        target = image_cache_target(value, cache_dir, product_name)

        if not target.exists() and download_missing:
            try:
                target.parent.mkdir(parents=True, exist_ok=True)
                request = urllib.request.Request(
                    value,
                    headers={"User-Agent": "Mozilla/5.0 CatalogGenerator/1.0"},
                )

                with urllib.request.urlopen(request, timeout=30) as response:
                    status = getattr(response, "status", None)
                    content_type = response.headers.get("Content-Type", "")
                    data = response.read()

                if status and status >= 400:
                    return None, f"Download failed with HTTP status {status}."

                if not data:
                    return None, "Download returned an empty file."

                target.write_bytes(data)

                if content_type and "image" not in content_type.lower():
                    return (
                        str(target),
                        f"Downloaded, but server content type is {content_type}.",
                    )

            except Exception as e:
                return None, f"Download failed: {type(e).__name__}: {e}"

        if target.exists():
            return str(target), ""

        return None, "Image was not downloaded and no cached file was found."

    path = Path(value)

    if path.is_file():
        return str(path), ""

    return None, f"Local image file not found: {value}"


def apply_watermark(img, watermark_path):
    if not watermark_path or not Path(watermark_path).is_file():
        return img

    base = img.convert("RGBA")

    with Image.open(watermark_path) as watermark_source:
        watermark = watermark_source.convert("RGBA")

    max_w = max(1, int(base.width * WATERMARK_SCALE))
    max_h = max(1, int(base.height * WATERMARK_SCALE))
    watermark.thumbnail((max_w, max_h), Image.LANCZOS)

    alpha = watermark.getchannel("A")
    alpha = alpha.point(lambda value: int(value * WATERMARK_OPACITY))
    watermark.putalpha(alpha)

    x = (base.width - watermark.width) // 2
    y = (base.height - watermark.height) // 2
    base.alpha_composite(watermark, (x, y))

    return base.convert("RGB")


def optimized_image_for_pdf(
    image_path,
    optimize_dir,
    draw_w_points,
    draw_h_points,
    label,
    watermark_path=None,
):
    image_path = Path(image_path)
    optimize_dir = Path(optimize_dir)
    optimize_dir.mkdir(parents=True, exist_ok=True)

    watermark_identity = ""

    if watermark_path and Path(watermark_path).is_file():
        watermark_stat = Path(watermark_path).stat()
        watermark_identity = (
            f"|wm:{Path(watermark_path).resolve()}|"
            f"{watermark_stat.st_mtime_ns}|{watermark_stat.st_size}"
        )

    source_stat = image_path.stat()
    digest_text = (
        f"{image_path.resolve()}|{source_stat.st_mtime_ns}|"
        f"{source_stat.st_size}|{draw_w_points:.2f}|{draw_h_points:.2f}|"
        f"{PDF_IMAGE_DPI}|{PDF_JPEG_QUALITY}{watermark_identity}"
    )
    digest = hashlib.sha1(digest_text.encode("utf-8")).hexdigest()[:16]
    target = optimize_dir / f"{clean_filename(label)}-{digest}.jpg"

    if target.exists():
        return str(target)

    max_w = max(1, int(draw_w_points / 72 * PDF_IMAGE_DPI))
    max_h = max(1, int(draw_h_points / 72 * PDF_IMAGE_DPI))

    with Image.open(image_path) as source:
        img = source.copy()

    if img.mode in ("RGBA", "LA") or (
        img.mode == "P" and "transparency" in img.info
    ):
        background = Image.new("RGB", img.size, "white")
        alpha = img.convert("RGBA").split()[-1]
        background.paste(img.convert("RGBA"), mask=alpha)
        img = background
    else:
        img = img.convert("RGB")

    img = apply_watermark(img, watermark_path)
    img.thumbnail((max_w, max_h), Image.LANCZOS)
    img.save(
        target,
        "JPEG",
        quality=PDF_JPEG_QUALITY,
        optimize=True,
        progressive=True,
        subsampling=0,
    )

    return str(target)


def collect_rows(df, product_column, image_column):
    product_column = resolve_excel_column(df, product_column, "Product")
    image_column = resolve_excel_column(df, image_column, "Image URL")
    rows = []

    for _, row in df.iterrows():
        product_name = str(row.get(product_column, "")).strip()
        image_value = str(row.get(image_column, "")).strip()

        if product_name and image_value and image_value.lower() != "nan":
            rows.append((product_name, image_value))

    return rows, product_column, image_column


def download_images_for_excel(
    excel_path,
    product_column,
    image_column,
    download_folder,
    sheet_name=0,
    progress_callback=None,
):
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    rows, product_column, image_column = collect_rows(df, product_column, image_column)

    if not rows:
        raise ValueError("No rows found with both product name and image URL.")

    download_folder = Path(download_folder)
    download_folder.mkdir(parents=True, exist_ok=True)
    downloaded = 0
    existing = 0
    skipped = 0
    manifest_rows = []

    for i, (product_name, image_value) in enumerate(rows):
        local_path = None

        try:
            if is_url(image_value):
                target = image_cache_target(image_value, download_folder, product_name)

                if target.exists():
                    existing += 1
                    local_path = str(target)
                else:
                    local_path, reason = local_or_download_image(
                        image_value,
                        download_folder,
                        product_name,
                        download_missing=True,
                    )

                    if local_path:
                        downloaded += 1
                    else:
                        skipped += 1
            else:
                local_path, reason = local_or_download_image(
                    image_value,
                    download_folder,
                    product_name,
                    download_missing=False,
                )

                if local_path:
                    existing += 1
                else:
                    skipped += 1

        except Exception as e:
            reason = f"{type(e).__name__}: {e}"
            skipped += 1

        manifest_rows.append(
            {
                "product_name": product_name,
                "image_source": image_value,
                "local_image_path": local_path or "",
                "status": "ready" if local_path else "skipped",
                "reason": reason if not local_path else "",
            }
        )

        if progress_callback:
            progress_callback(i + 1, len(rows), downloaded + existing, skipped)

    manifest_path = download_folder / "image_download_manifest.csv"
    pd.DataFrame(manifest_rows).to_csv(manifest_path, index=False)

    return downloaded, existing, skipped, str(manifest_path)


def draw_header(pdf, header_path, optimize_dir=None):
    if not header_path:
        return

    if not Path(header_path).is_file():
        return

    with Image.open(header_path) as header:
        header_h = min(
            HEADER_MAX_HEIGHT,
            PAGE_WIDTH * header.height / header.width
        )

    draw_path = header_path

    if optimize_dir:
        draw_path = optimized_image_for_pdf(
            header_path,
            optimize_dir,
            PAGE_WIDTH,
            header_h,
            "header",
        )

    pdf.drawImage(
        draw_path,
        0,
        PAGE_HEIGHT - HEADER_TOP_MARGIN - header_h,
        width=PAGE_WIDTH,
        height=header_h,
        preserveAspectRatio=False,
        mask="auto",
    )


def draw_watermark_on_slot(pdf, watermark_path, x, y, card_w, card_h):
    image_area_x = x + 9
    image_area_y = y + TEXT_BLOCK_HEIGHT + 4
    image_area_w = card_w - 18
    image_area_h = card_h - TEXT_BLOCK_HEIGHT - 12
    draw_watermark_on_rect(
        pdf,
        watermark_path,
        image_area_x,
        image_area_y,
        image_area_w,
        image_area_h,
    )


def draw_watermark_on_rect(pdf, watermark_path, x, y, width, height):
    with Image.open(watermark_path) as watermark:
        wm_w, wm_h = watermark.size

    max_w = width * WATERMARK_SCALE
    max_h = height * WATERMARK_SCALE
    scale = min(max_w / wm_w, max_h / wm_h)
    draw_w = wm_w * scale
    draw_h = wm_h * scale
    draw_x = x + (width - draw_w) / 2
    draw_y = y + (height - draw_h) / 2

    try:
        pdf.saveState()
        pdf.drawImage(
            watermark_path,
            draw_x,
            draw_y,
            width=draw_w,
            height=draw_h,
            preserveAspectRatio=True,
            mask="auto",
        )
        pdf.restoreState()
    except Exception:
        pdf.drawImage(
            watermark_path,
            draw_x,
            draw_y,
            width=draw_w,
            height=draw_h,
            preserveAspectRatio=True,
            mask="auto",
        )


def optimized_watermark_for_existing_pdf(
    watermark_path,
    output_pdf_path,
    draw_w_points,
    draw_h_points,
):
    watermark_path = Path(watermark_path)
    cache_dir = Path(output_pdf_path).with_name(
        Path(output_pdf_path).stem + "_watermark_cache"
    )
    cache_dir.mkdir(parents=True, exist_ok=True)

    source_stat = watermark_path.stat()
    digest_text = (
        f"{watermark_path.resolve()}|{source_stat.st_mtime_ns}|"
        f"{source_stat.st_size}|{draw_w_points:.2f}|{draw_h_points:.2f}|"
        f"{PDF_WATERMARK_DPI}|{PDF_WATERMARK_QUALITY}|{WATERMARK_OPACITY}"
    )
    digest = hashlib.sha1(digest_text.encode("utf-8")).hexdigest()[:16]
    target = cache_dir / f"watermark-{digest}.png"

    if target.exists():
        return str(target)

    max_w = max(1, int(draw_w_points / 72 * PDF_WATERMARK_DPI))
    max_h = max(1, int(draw_h_points / 72 * PDF_WATERMARK_DPI))

    with Image.open(watermark_path) as source:
        watermark = source.convert("RGBA")

    watermark.thumbnail((max_w, max_h), Image.LANCZOS)
    alpha = watermark.getchannel("A")
    alpha = alpha.point(lambda value: int(value * WATERMARK_OPACITY))
    watermark.putalpha(alpha)

    watermark.save(target, "PNG", optimize=True, compress_level=9)
    return str(target)


def _is_catalog_product_image(width, height):
    if width <= 0 or height <= 0:
        return False

    aspect = width / height

    # Catalog headers are very wide and short; product images are counted.
    if aspect > 4.5 and width > 1200:
        return False

    # Skip tiny decorative or repeated overlay images.
    if width < 60 or height < 60:
        return False

    return True


def _transformed_rect_from_bounds(ctm, x0, y0, x1, y1):
    corners = [(x0, y0), (x1, y0), (x0, y1), (x1, y1)]
    xs = []
    ys = []

    for px, py in corners:
        xs.append(ctm[0] * px + ctm[2] * py + ctm[4])
        ys.append(ctm[1] * px + ctm[3] * py + ctm[5])

    left = min(xs)
    bottom = min(ys)

    return {
        "x": left,
        "y": bottom,
        "width": max(xs) - left,
        "height": max(ys) - bottom,
    }


def _multiply_pdf_matrix(left, right):
    return [
        left[0] * right[0] + left[2] * right[1],
        left[1] * right[0] + left[3] * right[1],
        left[0] * right[2] + left[2] * right[3],
        left[1] * right[2] + left[3] * right[3],
        left[0] * right[4] + left[2] * right[5] + left[4],
        left[1] * right[4] + left[3] * right[5] + left[5],
    ]


def enumerate_product_images_on_page(page):
    if ContentStream is None:
        return []

    try:
        contents = page.get_contents()
    except Exception:
        return []

    if contents is None:
        return []

    try:
        operations = ContentStream(contents, page).operations
    except Exception:
        return []

    resources = page.get("/Resources") or {}
    xobjects = resources.get("/XObject") or {}

    try:
        xobjects = xobjects.get_object()
    except Exception:
        pass

    stack = []
    ctm = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0]
    draws = []

    for operands, operator in operations:
        if operator == b"q":
            stack.append(list(ctm))
        elif operator == b"Q":
            ctm = stack.pop() if stack else [1.0, 0.0, 0.0, 1.0, 0.0, 0.0]
        elif operator == b"cm":
            ctm = _multiply_pdf_matrix([float(value) for value in operands], ctm)
        elif operator == b"Do":
            xobject_ref = xobjects.get(operands[0])
            if xobject_ref is None:
                continue

            try:
                xobject = xobject_ref.get_object()
            except Exception:
                continue

            subtype = xobject.get("/Subtype")
            if subtype == "/Image":
                native_w = float(xobject.get("/Width", 0) or 0)
                native_h = float(xobject.get("/Height", 0) or 0)
                if not _is_catalog_product_image(native_w, native_h):
                    continue
                rect = _transformed_rect_from_bounds(ctm, 0, 0, 1, 1)
            elif subtype == "/Form":
                bbox = xobject.get("/BBox", [0, 0, 0, 0])
                native_w = float(bbox[2]) - float(bbox[0])
                native_h = float(bbox[3]) - float(bbox[1])
                if not _is_catalog_product_image(native_w, native_h):
                    continue
                rect = _transformed_rect_from_bounds(
                    ctm,
                    float(bbox[0]),
                    float(bbox[1]),
                    float(bbox[2]),
                    float(bbox[3]),
                )
            else:
                continue

            if rect["width"] <= 0 or rect["height"] <= 0:
                continue

            draws.append(rect)

    return draws


def count_products_in_existing_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    product_images = 0

    for page in reader.pages:
        product_images += len(enumerate_product_images_on_page(page))

    return product_images


def add_watermark_to_existing_pdf(
    input_pdf_path,
    watermark_path,
    output_pdf_path,
    grid_columns=3,
    grid_rows=5,
    total_products=None,
    progress_callback=None,
):
    if not Path(input_pdf_path).is_file():
        raise ValueError("Please choose an existing PDF file.")

    if not Path(watermark_path).is_file():
        raise ValueError("Please choose a watermark image.")

    layout = get_layout(grid_columns, grid_rows)
    image_area_w = layout["card_w"] - 18
    image_area_h = layout["card_h"] - TEXT_BLOCK_HEIGHT - 12
    max_watermark_w = image_area_w * WATERMARK_SCALE
    max_watermark_h = image_area_h * WATERMARK_SCALE
    watermark_path = optimized_watermark_for_existing_pdf(
        watermark_path,
        output_pdf_path,
        max_watermark_w,
        max_watermark_h,
    )
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()
    watermark_limit = int(total_products) if total_products else None
    watermarked_count = 0
    total_pages = len(reader.pages)

    for page_index, page in enumerate(reader.pages):
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)
        product_draws = enumerate_product_images_on_page(page)
        slots_on_page = 0

        if product_draws and (watermark_limit is None or watermarked_count < watermark_limit):
            packet = BytesIO()
            overlay = canvas.Canvas(packet, pagesize=(page_width, page_height))

            for draw in product_draws:
                if watermark_limit is not None and watermarked_count >= watermark_limit:
                    break

                draw_watermark_on_rect(
                    overlay,
                    watermark_path,
                    draw["x"],
                    draw["y"],
                    draw["width"],
                    draw["height"],
                )
                watermarked_count += 1
                slots_on_page += 1

            if slots_on_page > 0:
                overlay.save()
                packet.seek(0)
                overlay_pdf = PdfReader(packet)

                if overlay_pdf.pages:
                    overlay_page = overlay_pdf.pages[0]

                    try:
                        page.merge_page(overlay_page)
                    except Exception:
                        try:
                            page.mergePage(overlay_page)
                        except Exception as e:
                            raise RuntimeError(
                                f"Could not apply watermark on page {page_index + 1}: {e}"
                            ) from e

        writer.add_page(page)

        if progress_callback:
            progress_callback(page_index + 1, total_pages)

    output_pdf_path = str(output_pdf_path)
    Path(output_pdf_path).parent.mkdir(parents=True, exist_ok=True)

    with open(output_pdf_path, "wb") as output_file:
        writer.write(output_file)

    return output_pdf_path


def draw_product(
    pdf,
    font_name,
    x,
    y,
    product_name,
    image_path,
    card_w,
    card_h,
    optimize_dir=None,
    watermark_path=None,
):
    try:
        image_area_x = x + 9
        image_area_y = y + TEXT_BLOCK_HEIGHT + 4
        image_area_w = card_w - 18
        image_area_h = card_h - TEXT_BLOCK_HEIGHT - 12

        with Image.open(image_path) as source:
            img_w, img_h = source.size

        scale = min(image_area_w / img_w, image_area_h / img_h)
        draw_w = img_w * scale
        draw_h = img_h * scale

        img_x = image_area_x + (image_area_w - draw_w) / 2
        img_y = image_area_y + (image_area_h - draw_h) / 2
        draw_path = image_path

        if optimize_dir:
            draw_path = optimized_image_for_pdf(
                image_path,
                optimize_dir,
                draw_w,
                draw_h,
                product_name,
                watermark_path,
            )

        pdf.drawImage(
            ImageReader(draw_path),
            img_x,
            img_y,
            width=draw_w,
            height=draw_h,
            mask="auto",
        )

    except Exception as e:
        return False, f"Image could not be opened/optimized/drawn: {type(e).__name__}: {e}"

    code, name = split_code_and_name(product_name)

    pdf.setFillColor(colors.black)
    max_text_width = min(card_w - 18, 110)
    text_x = x + card_w / 2

    if code:
        draw_centered_fit(
            pdf,
            font_name,
            code,
            text_x,
            y + 29,
            max_text_width,
            CODE_FONT_SIZE,
        )

        text_y = y + 17
        max_name_lines = 2
    else:
        text_y = y + 29
        max_name_lines = 3

    lines = wrap_to_width(
        name,
        font_name,
        NAME_FONT_SIZE,
        max_text_width,
        max_name_lines,
    )

    for line in lines:
        draw_centered_fit(
            pdf,
            font_name,
            line,
            text_x,
            text_y,
            max_text_width,
            NAME_FONT_SIZE,
        )
        text_y -= LINE_HEIGHT

    return True, ""


def generate_catalog(
    excel_path,
    product_column,
    image_column,
    output_path,
    header_path=None,
    sheet_name=0,
    font_path=None,
    image_download_folder=None,
    grid_columns=3,
    grid_rows=5,
    watermark_path=None,
    progress_callback=None,
):
    font_name = register_font(font_path)
    layout = get_layout(grid_columns, grid_rows)
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    rows, product_column, image_column = collect_rows(df, product_column, image_column)

    if not rows:
        raise ValueError("No rows found with both product name and image URL.")

    output_path = str(output_path)
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    pdf = canvas.Canvas(output_path, pagesize=A4)
    cache_dir = Path(image_download_folder) if image_download_folder else Path(output_path).with_suffix("").with_name(Path(output_path).stem + "_images")
    cache_dir.mkdir(parents=True, exist_ok=True)
    optimize_dir = cache_dir / "_pdf_optimized"
    generated = 0
    skipped = 0
    skipped_rows = []

    for i, (product_name, image_value) in enumerate(rows):
        if i % layout["products_per_page"] == 0:
            if i != 0:
                pdf.showPage()
            draw_header(pdf, header_path, optimize_dir)

        position = i % layout["products_per_page"]
        col = position % layout["columns"]
        row_no = position // layout["columns"]

        x = LEFT_MARGIN + col * (layout["card_w"] + X_GAP)
        y = GRID_TOP - layout["card_h"] - row_no * (layout["card_h"] + Y_GAP)

        try:
            image_path, image_reason = local_or_download_image(
                image_value,
                cache_dir,
                product_name,
                download_missing=True,
            )

            if not image_path:
                skipped += 1
                skipped_rows.append(
                    {
                        "row_number": i + 2,
                        "product_name": product_name,
                        "image_source": image_value,
                        "reason": image_reason or "Image path could not be resolved.",
                    }
                )
                continue

            ok, draw_reason = draw_product(
                pdf,
                font_name,
                x,
                y,
                product_name,
                image_path,
                layout["card_w"],
                layout["card_h"],
                optimize_dir,
                watermark_path,
            )

            if ok:
                generated += 1
            else:
                skipped += 1
                skipped_rows.append(
                    {
                        "row_number": i + 2,
                        "product_name": product_name,
                        "image_source": image_value,
                        "local_image_path": image_path,
                        "reason": draw_reason or "Could not draw image in PDF.",
                    }
                )

        except Exception as e:
            skipped += 1
            skipped_rows.append(
                {
                    "row_number": i + 2,
                    "product_name": product_name,
                    "image_source": image_value,
                    "reason": f"{type(e).__name__}: {e}",
                }
            )

        if progress_callback:
            progress_callback(i + 1, len(rows), generated, skipped)

    pdf.save()

    skip_report_path = ""

    if skipped_rows:
        skip_report_path = str(Path(output_path).with_name(Path(output_path).stem + "_skipped_images.csv"))
        pd.DataFrame(skipped_rows).to_csv(skip_report_path, index=False)

    return generated, skipped, output_path, skip_report_path, skipped_rows[:10]


class CatalogGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("800x620")
        self.root.minsize(760, 580)

        self.excel_path = StringVar()
        self.header_path = StringVar()
        self.output_folder = StringVar()
        self.output_filename = StringVar(value="Catalog.pdf")
        self.existing_pdf_path = StringVar()
        self.watermarked_pdf_filename = StringVar(value="Catalog_watermarked.pdf")
        self.existing_pdf_product_count = StringVar()
        self.download_folder = StringVar()
        self.grid_columns = StringVar(value="3")
        self.grid_rows = StringVar(value="5")
        self.use_watermark = BooleanVar(value=False)
        self.watermark_path = StringVar()
        self.product_column = StringVar()
        self.image_column = StringVar()
        self.status = StringVar(value="Choose an Excel file to begin.")

        self.columns = []

        self.build_ui()

    def build_ui(self):
        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill="both", expand=True)

        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Excel file").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=self.excel_path).grid(row=0, column=1, sticky="ew", pady=6, padx=8)
        ttk.Button(frame, text="Browse", command=self.pick_excel).grid(row=0, column=2, pady=6)

        ttk.Label(frame, text="Product name column").grid(row=1, column=0, sticky="w", pady=6)
        self.product_combo = ttk.Combobox(frame, textvariable=self.product_column)
        self.product_combo.grid(row=1, column=1, sticky="ew", pady=6, padx=8)

        ttk.Label(frame, text="Image URL column").grid(row=2, column=0, sticky="w", pady=6)
        self.image_combo = ttk.Combobox(frame, textvariable=self.image_column)
        self.image_combo.grid(row=2, column=1, sticky="ew", pady=6, padx=8)

        ttk.Label(frame, text="Header image").grid(row=3, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=self.header_path).grid(row=3, column=1, sticky="ew", pady=6, padx=8)
        ttk.Button(frame, text="Browse", command=self.pick_header).grid(row=3, column=2, pady=6)

        ttk.Label(frame, text="Product grid").grid(row=4, column=0, sticky="w", pady=6)
        grid_frame = ttk.Frame(frame)
        grid_frame.grid(row=4, column=1, sticky="w", pady=6, padx=8)
        ttk.Label(grid_frame, text="Columns").pack(side="left")
        ttk.Spinbox(
            grid_frame,
            from_=1,
            to=8,
            width=5,
            textvariable=self.grid_columns,
        ).pack(side="left", padx=(6, 12))
        ttk.Label(grid_frame, text="Rows").pack(side="left")
        ttk.Spinbox(
            grid_frame,
            from_=1,
            to=10,
            width=5,
            textvariable=self.grid_rows,
        ).pack(side="left", padx=(6, 0))

        ttk.Label(frame, text="Watermark image").grid(row=5, column=0, sticky="w", pady=6)
        watermark_frame = ttk.Frame(frame)
        watermark_frame.grid(row=5, column=1, sticky="ew", pady=6, padx=8)
        watermark_frame.columnconfigure(1, weight=1)
        ttk.Checkbutton(watermark_frame, variable=self.use_watermark).grid(row=0, column=0, sticky="w")
        ttk.Entry(watermark_frame, textvariable=self.watermark_path).grid(row=0, column=1, sticky="ew", padx=(6, 0))
        ttk.Button(frame, text="Browse", command=self.pick_watermark).grid(row=5, column=2, pady=6)

        ttk.Label(frame, text="Image download folder").grid(row=6, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=self.download_folder).grid(row=6, column=1, sticky="ew", pady=6, padx=8)
        ttk.Button(frame, text="Browse", command=self.pick_download_folder).grid(row=6, column=2, pady=6)

        ttk.Label(frame, text="Generated PDF folder").grid(row=7, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=self.output_folder).grid(row=7, column=1, sticky="ew", pady=6, padx=8)
        ttk.Button(frame, text="Browse", command=self.pick_output_folder).grid(row=7, column=2, pady=6)

        ttk.Label(frame, text="PDF file name").grid(row=8, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=self.output_filename).grid(row=8, column=1, sticky="ew", pady=6, padx=8)

        self.progress = ttk.Progressbar(frame, mode="determinate")
        self.progress.grid(row=9, column=0, columnspan=3, sticky="ew", pady=(16, 6))

        action_frame = ttk.Frame(frame)
        action_frame.grid(row=10, column=0, columnspan=3, pady=10)
        ttk.Button(action_frame, text="Download Images First", command=self.start_download).pack(side="left", padx=6)
        ttk.Button(action_frame, text="Generate PDF", command=self.start_generation).pack(
            side="left", padx=6
        )

        ttk.Label(frame, textvariable=self.status).grid(row=11, column=0, columnspan=3, sticky="w", pady=(4, 8))

        self.log = ttk.Treeview(frame, columns=("message",), show="headings", height=8)
        self.log.heading("message", text="Activity")
        self.log.grid(row=12, column=0, columnspan=3, sticky="nsew")
        frame.rowconfigure(12, weight=1)

        watermark_existing = ttk.LabelFrame(frame, text="Add watermark to existing PDF", padding=10)
        watermark_existing.grid(row=13, column=0, columnspan=3, sticky="ew", pady=(12, 0))
        watermark_existing.columnconfigure(1, weight=1)

        ttk.Label(watermark_existing, text="Existing PDF").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(watermark_existing, textvariable=self.existing_pdf_path).grid(
            row=0, column=1, sticky="ew", padx=8, pady=4
        )
        ttk.Button(watermark_existing, text="Browse", command=self.pick_existing_pdf).grid(
            row=0, column=2, pady=4
        )

        ttk.Label(watermark_existing, text="Total products auto").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(watermark_existing, textvariable=self.existing_pdf_product_count, width=12).grid(
            row=1, column=1, sticky="w", padx=8, pady=4
        )

        ttk.Label(watermark_existing, text="Watermarked file name").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(watermark_existing, textvariable=self.watermarked_pdf_filename).grid(
            row=2, column=1, sticky="ew", padx=8, pady=4
        )
        ttk.Button(
            watermark_existing,
            text="Create Watermarked Copy",
            command=self.start_existing_pdf_watermark,
        ).grid(row=2, column=2, pady=4)

    def log_message(self, message):
        self.log.insert("", END, values=(message,))
        self.log.yview_moveto(1)

    def pick_excel(self):
        path = filedialog.askopenfilename(
            title="Choose Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )

        if not path:
            return

        self.excel_path.set(path)

        if not self.output_folder.get():
            self.output_folder.set(str(Path(path).parent))
        self.output_filename.set(Path(path).stem + "_Catalog.pdf")

        default_download_folder = Path(path).with_name(Path(path).stem + "_images")
        self.download_folder.set(str(default_download_folder))

        self.load_columns()

    def load_columns(self):
        try:
            df = pd.read_excel(self.excel_path.get(), nrows=0)
            self.columns = [str(col) for col in df.columns]

            self.product_combo["values"] = self.columns
            self.image_combo["values"] = self.columns

            for col in self.columns:
                lowered = col.lower()
                if not self.product_column.get() and "product" in lowered and "name" in lowered:
                    self.product_column.set(col)
                if not self.image_column.get() and ("image" in lowered or "url" in lowered):
                    self.image_column.set(col)

            self.status.set(f"Loaded {len(self.columns)} columns.")
            self.log_message("Columns loaded from Excel.")

        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Could not read Excel columns:\n{e}")

    def pick_header(self):
        path = filedialog.askopenfilename(
            title="Choose header image",
            filetypes=[("Image files", "*.png *.jpg *.jpeg"), ("All files", "*.*")],
        )

        if path:
            self.header_path.set(path)

    def pick_watermark(self):
        path = filedialog.askopenfilename(
            title="Choose watermark image",
            filetypes=[("Image files", "*.png *.jpg *.jpeg"), ("All files", "*.*")],
        )

        if path:
            self.watermark_path.set(path)
            self.use_watermark.set(True)

    def pick_output_folder(self):
        path = filedialog.askdirectory(title="Choose generated PDF folder")

        if path:
            self.output_folder.set(path)

    def pick_existing_pdf(self):
        path = filedialog.askopenfilename(
            title="Choose existing catalog PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )

        if path:
            self.existing_pdf_path.set(path)
            source = Path(path)
            if not self.output_folder.get():
                self.output_folder.set(str(source.parent))
            self.watermarked_pdf_filename.set(source.stem + "_watermarked.pdf")
            try:
                product_count = count_products_in_existing_pdf(path)
                if product_count:
                    self.existing_pdf_product_count.set(str(product_count))
                    self.log_message(
                        f"Auto-detected {product_count} products in existing PDF."
                    )
                else:
                    self.existing_pdf_product_count.set("")
                    self.log_message(
                        "Could not auto-detect products. You can enter the count manually."
                    )
            except Exception as e:
                self.existing_pdf_product_count.set("")
                self.log_message(f"Product count detection failed: {e}")

    def pick_download_folder(self):
        path = filedialog.askdirectory(title="Choose image download folder")

        if path:
            self.download_folder.set(path)

    def selected_grid(self):
        try:
            columns = int(str(self.grid_columns.get()).strip())
            rows = int(str(self.grid_rows.get()).strip())
        except ValueError:
            raise ValueError("Grid columns and rows must be whole numbers.")

        if columns < 1 or rows < 1:
            raise ValueError("Grid rows and columns must be greater than zero.")

        if columns > 8 or rows > 10:
            raise ValueError("Please use up to 8 columns and 10 rows.")

        return columns, rows

    def selected_output_path(self):
        folder = Path(str(self.output_folder.get()).strip())
        filename = str(self.output_filename.get()).strip()

        if not filename:
            raise ValueError("Please enter a PDF file name.")

        if Path(filename).suffix.lower() != ".pdf":
            filename += ".pdf"

        return str(folder / filename)

    def selected_watermarked_output_path(self):
        folder = Path(str(self.output_folder.get()).strip())
        filename = str(self.watermarked_pdf_filename.get()).strip()

        if not filename:
            raise ValueError("Please enter a watermarked PDF file name.")

        if Path(filename).suffix.lower() != ".pdf":
            filename += ".pdf"

        return str(folder / filename)

    def selected_existing_pdf_product_count(self):
        value = str(self.existing_pdf_product_count.get()).strip()

        if not value:
            return None

        try:
            count = int(value)
        except ValueError:
            raise ValueError("Total products must be a whole number.")

        if count < 1:
            raise ValueError("Total products must be greater than zero.")

        return count

    def open_pdf(self, output_path):
        os.startfile(output_path)

    def open_output_folder(self, output_path):
        os.startfile(str(Path(output_path).parent))

    def show_pdf_ready_popup(self, output_path):
        popup = Toplevel(self.root)
        popup.title(APP_TITLE)
        popup.resizable(False, False)
        popup.transient(self.root)
        popup.grab_set()

        frame = ttk.Frame(popup, padding=18)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="PDF created successfully.").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 12)
        )
        ttk.Label(frame, text=str(output_path)).grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(0, 14)
        )

        ttk.Button(frame, text="Open PDF", command=lambda: self.open_pdf(output_path)).grid(
            row=2, column=0, padx=4
        )
        ttk.Button(
            frame,
            text="Open Folder",
            command=lambda: self.open_output_folder(output_path),
        ).grid(row=2, column=1, padx=4)
        ttk.Button(frame, text="Close", command=popup.destroy).grid(row=2, column=2, padx=4)

        popup.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - popup.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - popup.winfo_height()) // 2
        popup.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def validate(self):
        if not self.excel_path.get():
            raise ValueError("Please choose an Excel file.")
        if not self.product_column.get():
            raise ValueError("Please enter/select the product name column.")
        if not self.image_column.get():
            raise ValueError("Please enter/select the image URL column.")
        if not self.output_folder.get():
            raise ValueError("Please choose the generated PDF folder.")
        self.selected_output_path()
        self.selected_grid()
        if self.use_watermark.get() and not self.watermark_path.get():
            raise ValueError("Please choose a watermark image or turn off watermark.")

    def validate_download(self):
        if not self.excel_path.get():
            raise ValueError("Please choose an Excel file.")
        if not self.product_column.get():
            raise ValueError("Please enter/select the product name column.")
        if not self.image_column.get():
            raise ValueError("Please enter/select the image URL column.")
        if not self.download_folder.get():
            raise ValueError("Please choose an image download folder.")

    def validate_existing_pdf_watermark(self):
        if not self.existing_pdf_path.get():
            raise ValueError("Please choose the existing PDF.")
        if not self.watermark_path.get():
            raise ValueError("Please choose a watermark image.")
        if not self.output_folder.get():
            raise ValueError("Please choose the generated PDF folder.")
        self.selected_grid()
        self.selected_existing_pdf_product_count()
        self.selected_watermarked_output_path()

    def start_download(self):
        try:
            self.validate_download()
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))
            return

        self.progress["value"] = 0
        self.status.set("Downloading images...")
        self.log_message("Started image download.")

        thread = threading.Thread(target=self.download_worker, daemon=True)
        thread.start()

    def start_generation(self):
        try:
            self.validate()
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))
            return

        self.progress["value"] = 0
        self.status.set("Generating PDF...")
        self.log_message("Started PDF generation.")

        thread = threading.Thread(target=self.generate_worker, daemon=True)
        thread.start()

    def start_existing_pdf_watermark(self):
        try:
            self.validate_existing_pdf_watermark()
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))
            return

        self.progress["value"] = 0
        self.status.set("Adding watermark to existing PDF...")
        self.log_message("Started existing PDF watermarking.")

        thread = threading.Thread(target=self.existing_pdf_watermark_worker, daemon=True)
        thread.start()

    def update_progress(self, done, total, generated, skipped):
        def apply_update():
            self.progress["maximum"] = total
            self.progress["value"] = done
            self.status.set(
                f"Processed {done}/{total}. Added {generated}, skipped {skipped}."
            )

        self.root.after(0, apply_update)

    def update_download_progress(self, done, total, completed, skipped):
        def apply_update():
            self.progress["maximum"] = total
            self.progress["value"] = done
            self.status.set(
                f"Downloaded/ready {completed}/{total}. Skipped {skipped}."
            )

        self.root.after(0, apply_update)

    def update_pdf_watermark_progress(self, done, total):
        def apply_update():
            self.progress["maximum"] = total
            self.progress["value"] = done
            self.status.set(f"Watermarked pages {done}/{total}.")

        self.root.after(0, apply_update)

    def download_worker(self):
        try:
            downloaded, existing, skipped, manifest_path = download_images_for_excel(
                excel_path=self.excel_path.get(),
                product_column=self.product_column.get(),
                image_column=self.image_column.get(),
                download_folder=self.download_folder.get(),
                progress_callback=self.update_download_progress,
            )

            def done():
                self.status.set(
                    f"Images ready. New {downloaded}, existing {existing}, skipped {skipped}."
                )
                self.log_message(f"Image manifest saved: {manifest_path}")
                messagebox.showinfo(
                    APP_TITLE,
                    "Images are ready.\n"
                    f"New downloads: {downloaded}\n"
                    f"Already existed: {existing}\n"
                    f"Skipped: {skipped}\n\n"
                    "Now click Generate PDF.",
                )

            self.root.after(0, done)

        except Exception as e:
            error_message = str(e)

            def failed():
                self.status.set("Image download failed.")
                self.log_message(error_message)
                messagebox.showerror(APP_TITLE, error_message)

            self.root.after(0, failed)

    def generate_worker(self):
        try:
            grid_columns, grid_rows = self.selected_grid()
            watermark_path = self.watermark_path.get() if self.use_watermark.get() else None
            generated, skipped, output_path, skip_report_path, skip_preview = generate_catalog(
                excel_path=self.excel_path.get(),
                product_column=self.product_column.get(),
                image_column=self.image_column.get(),
                output_path=self.selected_output_path(),
                header_path=self.header_path.get() or None,
                image_download_folder=self.download_folder.get() or None,
                grid_columns=grid_columns,
                grid_rows=grid_rows,
                watermark_path=watermark_path,
                progress_callback=self.update_progress,
            )

            def done():
                self.status.set(f"Done. Added {generated}, skipped {skipped}.")
                self.log_message(f"Saved PDF: {output_path}")
                if skip_report_path:
                    self.log_message(f"Skipped image report: {skip_report_path}")
                for item in skip_preview:
                    self.log_message(
                        f"Skipped row {item.get('row_number')}: "
                        f"{item.get('product_name')} - {item.get('reason')}"
                    )
                self.show_pdf_ready_popup(output_path)

            self.root.after(0, done)

        except Exception as e:
            error_message = str(e)

            def failed():
                self.status.set("Failed.")
                self.log_message(error_message)
                messagebox.showerror(APP_TITLE, error_message)

            self.root.after(0, failed)

    def existing_pdf_watermark_worker(self):
        try:
            grid_columns, grid_rows = self.selected_grid()
            output_path = add_watermark_to_existing_pdf(
                input_pdf_path=self.existing_pdf_path.get(),
                watermark_path=self.watermark_path.get(),
                output_pdf_path=self.selected_watermarked_output_path(),
                grid_columns=grid_columns,
                grid_rows=grid_rows,
                total_products=self.selected_existing_pdf_product_count(),
                progress_callback=self.update_pdf_watermark_progress,
            )

            def done():
                self.status.set("Watermarked PDF created.")
                self.log_message(f"Saved watermarked PDF: {output_path}")
                self.show_pdf_ready_popup(output_path)

            self.root.after(0, done)

        except Exception as e:
            error_message = str(e)

            def failed():
                self.status.set("Watermarking failed.")
                self.log_message(error_message)
                messagebox.showerror(APP_TITLE, error_message)

            self.root.after(0, failed)


def main():
    root = Tk()
    CatalogGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
