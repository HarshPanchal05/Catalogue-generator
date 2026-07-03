import os
import re

import pandas as pd
from PIL import Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

# ==========================
# BASE DIRECTORY
# ==========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_FILE = os.path.join(BASE_DIR, "Picanol.xlsx")
HEADER_FILE = os.path.join(BASE_DIR, "header_cropped.png")
OUTPUT_FILE = os.path.join(BASE_DIR, "Picanol_Catalog.pdf")
IMAGE_FOLDER = os.path.join(BASE_DIR, "images")
FONT_FILE = os.path.join(BASE_DIR, "CanvaSans-Bold.ttf")

# ==========================
# LOAD FONT
# ==========================
try:
    pdfmetrics.registerFont(TTFont("CanvaSans", FONT_FILE))
    FONT_NAME = "CanvaSans"
    print("Canva Sans loaded.")
except Exception as e:
    print("Font Error:", e)
    FONT_NAME = "Helvetica-Bold"

# ==========================
# PAGE SETTINGS
# ==========================
PAGE_WIDTH, PAGE_HEIGHT = A4

COLS = 3
ROWS = 5
PRODUCTS_PER_PAGE = COLS * ROWS

LEFT_MARGIN = 0
RIGHT_MARGIN = 0
BOTTOM_MARGIN = 0

X_GAP = 0
Y_GAP = 2

HEADER_TOP_MARGIN = 10
HEADER_MAX_HEIGHT = 78
GRID_TOP = PAGE_HEIGHT - 108

CARD_W = (
    PAGE_WIDTH
    - LEFT_MARGIN
    - RIGHT_MARGIN
    - ((COLS - 1) * X_GAP)
) / COLS

CARD_H = (
    GRID_TOP
    - BOTTOM_MARGIN
    - ((ROWS - 1) * Y_GAP)
) / ROWS

CARD_BORDER_COLOR = colors.HexColor("#9a72ff")
DRAW_CARD_BORDER = False

CODE_FONT_SIZE = 9
NAME_FONT_SIZE = 8.7
LINE_HEIGHT = 10.5
TEXT_BLOCK_HEIGHT = 38

# ==========================
# READ EXCEL
# ==========================
df = pd.read_excel(EXCEL_FILE)
df = df[["Product Name"]]

df["Product Name"] = (
    df["Product Name"]
    .fillna("")
    .astype(str)
    .str.strip()
)

# ==========================
# READ LOCAL IMAGES
# ==========================
image_files = {
    os.path.splitext(f)[0].strip().lower(): f
    for f in os.listdir(IMAGE_FOLDER)
    if f.lower().endswith((".jpg", ".jpeg", ".png"))
}

valid_rows = []

for _, row in df.iterrows():

    product_name = row["Product Name"]

    if not product_name:
        continue

    filename = "".join(
        c for c in product_name
        if c.isalnum() or c in (" ", "-", "_")
    ).strip().lower()

    if filename in image_files:
        image_path = os.path.join(IMAGE_FOLDER, image_files[filename])

        if os.path.isfile(image_path):
            valid_rows.append((product_name, image_path))

print()
print(f"Image files found      : {len(image_files)}")
print(f"Products to generate   : {len(valid_rows)}")
print(f"Products skipped       : {len(df) - len(valid_rows)}")
print()

# ==========================
# CREATE PDF
# ==========================
pdf = canvas.Canvas(OUTPUT_FILE, pagesize=A4)

# ==========================
# HEADER
# ==========================
def draw_header():

    try:
        with Image.open(HEADER_FILE) as header:
            header_h = min(
                HEADER_MAX_HEIGHT,
                PAGE_WIDTH * header.height / header.width
            )

        pdf.drawImage(
            HEADER_FILE,
            0,
            PAGE_HEIGHT - HEADER_TOP_MARGIN - header_h,
            width=PAGE_WIDTH,
            height=header_h,
            preserveAspectRatio=False,
            mask="auto"
        )

    except Exception as e:
        print("Header Error:", e)

# ==========================
# TEXT HELPERS
# ==========================
def split_code_and_name(product_name):

    product_name = str(product_name).strip()

    match = re.match(
        r"^([A-Za-z0-9\-\/]+)\s+(.*)",
        product_name
    )

    if not match:
        return "", product_name

    first = match.group(1)
    remaining = match.group(2).strip()

    if any(c.isdigit() for c in first):
        return first, remaining

    return "", product_name


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


def draw_centered_fit(text, x_center, y, max_width, font_size):

    size = font_size

    while (
        size > 7
        and pdfmetrics.stringWidth(text, FONT_NAME, size) > max_width
    ):
        size -= 0.25

    pdf.setFont(FONT_NAME, size)
    pdf.drawCentredString(x_center, y, text)

# ==========================
# PRODUCT CARD
# ==========================
def draw_product(x, y, product_name, image_path):

    if DRAW_CARD_BORDER:
        pdf.setStrokeColor(CARD_BORDER_COLOR)
        pdf.setLineWidth(0.5)

        pdf.rect(
            x,
            y,
            CARD_W,
            CARD_H,
            stroke=1,
            fill=0
        )

    # ----------------------
    # IMAGE
    # ----------------------
    try:
        image_area_x = x + 9
        image_area_y = y + TEXT_BLOCK_HEIGHT + 4
        image_area_w = CARD_W - 18
        image_area_h = CARD_H - TEXT_BLOCK_HEIGHT - 12

        with Image.open(image_path) as source:
            img_w, img_h = source.size

        scale = min(image_area_w / img_w, image_area_h / img_h)
        draw_w = img_w * scale
        draw_h = img_h * scale

        img_x = image_area_x + (image_area_w - draw_w) / 2
        img_y = image_area_y + (image_area_h - draw_h) / 2

        pdf.drawImage(
            ImageReader(image_path),
            img_x,
            img_y,
            width=draw_w,
            height=draw_h,
            mask="auto"
        )

    except Exception as e:
        print("Image Error:", product_name, e)
        return

    # ----------------------
    # PRODUCT TEXT
    # ----------------------
    code, name = split_code_and_name(product_name)

    pdf.setFillColor(colors.black)
    max_text_width = min(CARD_W - 18, 110)
    text_x = x + CARD_W / 2

    if code:
        draw_centered_fit(
            code,
            text_x,
            y + 29,
            max_text_width,
            CODE_FONT_SIZE
        )

        text_y = y + 17
        max_name_lines = 2
    else:
        text_y = y + 29
        max_name_lines = 3

    lines = wrap_to_width(
        name,
        FONT_NAME,
        NAME_FONT_SIZE,
        max_text_width,
        max_name_lines
    )

    for line in lines:
        draw_centered_fit(
            line,
            text_x,
            text_y,
            max_text_width,
            NAME_FONT_SIZE
        )

        text_y -= LINE_HEIGHT

# ==========================
# GENERATE PDF
# ==========================
total = len(valid_rows)

for i, (product_name, image_path) in enumerate(valid_rows):

    print(f"Creating {i + 1}/{total}")

    if i % PRODUCTS_PER_PAGE == 0:

        if i != 0:
            pdf.showPage()

        draw_header()

    position = i % PRODUCTS_PER_PAGE

    col = position % COLS
    row_no = position // COLS

    x = LEFT_MARGIN + col * (CARD_W + X_GAP)
    y = GRID_TOP - CARD_H - row_no * (CARD_H + Y_GAP)

    draw_product(
        x,
        y,
        product_name,
        image_path
    )

pdf.save()

print()
print("DONE")
print("Saved as:")
print(OUTPUT_FILE)
