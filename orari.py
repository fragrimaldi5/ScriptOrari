import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# ================= CONFIG =================

INPUT_DIR = "input_excel"
OUTPUT_DIR = "output_img"
LOGO_PATH = "logo_studentingegneria.png"

GIORNI = ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì"]

SCALE = 3
LOGO_RATIO = 0.10

BASE_CELL_WIDTH = int(460 * SCALE)
BASE_CELL_HEIGHT = int(260 * SCALE)

BASE_PADDING = int(10 * SCALE)
BASE_GAP = int(10 * SCALE)
MARGIN = int(30 * SCALE)

FONT_PATH = "C:/Windows/Fonts/arial.ttf"

FONT_HEADER_BASE = int(56 * SCALE)
FONT_ORARIO_BASE = int(72 * SCALE)
FONT_ESAME_BASE = int(84 * SCALE)
FONT_AULA_BASE = int(60 * SCALE)

PASTEL_COLORS = [
    (235, 242, 250),
    (238, 250, 235),
    (250, 245, 235),
    (245, 235, 250),
    (250, 235, 240),
    (235, 250, 245),
]

# ==========================================

os.makedirs(OUTPUT_DIR, exist_ok=True)


def pulisci_testo(cella):
    if pd.isna(cella):
        return [], [], None

    testo = str(cella)

    gruppo = None
    if "(A-H)" in testo:
        gruppo = "AH"
    elif "(I-Z)" in testo:
        gruppo = "IZ"

    esami = []
    aule = []

    for r in testo.splitlines():
        r = r.strip()
        if not r:
            continue
        if "aula" in r.lower():
            aule.append(r)
        else:
            esami.append(r)

    return esami, aule, gruppo


def genera_report(excel_path, filtro=None):
    df = pd.read_excel(excel_path, header=None)
    orari = df.iloc[3:, 0]
    dati = df.iloc[3:, 1:6]

    cols = len(GIORNI) + 1
    rows = len(orari) + 1

    table_width = cols * BASE_CELL_WIDTH + (cols - 1) * BASE_GAP
    side = table_width + MARGIN * 2

    table_area_h = int(side * (1 - LOGO_RATIO)) - MARGIN
    base_table_h = rows * BASE_CELL_HEIGHT + (rows - 1) * BASE_GAP
    scale_y = min(1.0, table_area_h / base_table_h)

    CELL_WIDTH = BASE_CELL_WIDTH
    CELL_HEIGHT = int(BASE_CELL_HEIGHT * scale_y)
    CELL_PADDING = int(BASE_PADDING * scale_y)
    CELL_GAP = int(BASE_GAP * scale_y)

    FONT_HEADER = int(FONT_HEADER_BASE * scale_y)
    FONT_ORARIO = int(FONT_ORARIO_BASE * scale_y)
    FONT_ESAME_MAX = int(FONT_ESAME_BASE * scale_y)
    FONT_AULA = int(FONT_AULA_BASE * scale_y)

    font_header = ImageFont.truetype(FONT_PATH, FONT_HEADER)
    font_orario = ImageFont.truetype(FONT_PATH, FONT_ORARIO)
    font_aula = ImageFont.truetype(FONT_PATH, FONT_AULA)

    img = Image.new("RGB", (side, side), "white")
    draw = ImageDraw.Draw(img)

    table_h = rows * CELL_HEIGHT + (rows - 1) * CELL_GAP
    x0 = (side - table_width) // 2
    y0 = MARGIN + (table_area_h - table_h) // 2

    exam_colors = {}
    color_i = 0

    def wrap_text(text, max_w, font):
        words = text.split()
        lines, cur = [], ""
        for w in words:
            t = cur + (" " if cur else "") + w
            if draw.textlength(t, font=font) <= max_w:
                cur = t
            else:
                lines.append(cur)
                cur = w
        if cur:
            lines.append(cur)
        return lines

    def fit_exam(text):
        max_w = CELL_WIDTH - 2 * CELL_PADDING
        size = FONT_ESAME_MAX
        while size > 28:
            f = ImageFont.truetype(FONT_PATH, size)
            lines = wrap_text(text, max_w, f)[:3]
            if len(lines) * (f.size + 6) <= CELL_HEIGHT * 0.55:
                return f, lines
            size -= 2
        f = ImageFont.truetype(FONT_PATH, size)
        return f, wrap_text(text, max_w, f)[:3]

    # ===== HEADER =====
    for c in range(cols):
        x = x0 + c * (CELL_WIDTH + CELL_GAP)
        draw.rectangle([x, y0, x + CELL_WIDTH, y0 + CELL_HEIGHT], outline="black", width=3)
        if c > 0:
            g = GIORNI[c - 1]
            w = draw.textlength(g, font=font_header)
            draw.text((x + (CELL_WIDTH - w) / 2,
                       y0 + CELL_HEIGHT / 2 - FONT_HEADER / 2),
                      g, font=font_header, fill="black")

    # ===== TABELLA =====
    for r, orario in enumerate(orari):
        y = y0 + (r + 1) * (CELL_HEIGHT + CELL_GAP)

        draw.rectangle([x0, y, x0 + CELL_WIDTH, y + CELL_HEIGHT], outline="black", width=3)
        w = draw.textlength(str(orario), font=font_orario)
        draw.text((x0 + (CELL_WIDTH - w) / 2,
                   y + CELL_HEIGHT / 2 - FONT_ORARIO / 2),
                  str(orario), font=font_orario, fill="black")

        for c in range(len(GIORNI)):
            x = x0 + (c + 1) * (CELL_WIDTH + CELL_GAP)
            val = dati.iloc[r, c]

            esami, aule, gruppo = pulisci_testo(val)

            if filtro and gruppo != filtro:
                draw.rectangle([x, y, x + CELL_WIDTH, y + CELL_HEIGHT], outline="black", width=3)
                continue

            if not esami:
                draw.rectangle([x, y, x + CELL_WIDTH, y + CELL_HEIGHT], outline="black", width=3)
                continue

            key = " | ".join(esami)
            if key not in exam_colors:
                exam_colors[key] = PASTEL_COLORS[color_i % len(PASTEL_COLORS)]
                color_i += 1

            draw.rectangle(
                [x, y, x + CELL_WIDTH, y + CELL_HEIGHT],
                fill=exam_colors[key],
                outline="black",
                width=3
            )

            blocks = []

            for i, es in enumerate(esami):
                font_esame, exam_lines = fit_exam(es)
                blocks.append((exam_lines, font_esame, 6))

                if i < len(aule):
                    aula_lines = wrap_text(
                        aule[i],
                        CELL_WIDTH - 2 * CELL_PADDING,
                        font_aula
                    )[:2]
                    blocks.append((aula_lines, font_aula, 4))

            total_h = sum(len(lines) * (font.size + gap) for lines, font, gap in blocks)
            ty = y + max(CELL_PADDING, (CELL_HEIGHT - total_h) / 2)

            for lines, font, gap in blocks:
                for line in lines:
                    w = draw.textlength(line, font=font)
                    draw.text(
                        (x + (CELL_WIDTH - w) / 2, ty),
                        line,
                        font=font,
                        fill="black"
                    )
                    ty += font.size + gap

    # ===== LOGO =====
    if os.path.exists(LOGO_PATH):
        logo = Image.open(LOGO_PATH).convert("RGBA")
        logo_h = int(side * LOGO_RATIO)
        scale = logo_h / logo.height
        logo = logo.resize((int(logo.width * scale), logo_h), Image.LANCZOS)
        img.paste(
            logo,
            ((side - logo.width) // 2, side - logo.height - MARGIN),
            logo
        )

    base = os.path.splitext(os.path.basename(excel_path))[0]
    suffix = f"_{filtro.lower()}" if filtro else ""
    img.save(os.path.join(OUTPUT_DIR, base + suffix + ".png"))


# ===== BATCH =====
for f in os.listdir(INPUT_DIR):
    if not f.lower().endswith((".xls", ".xlsx")):
        continue

    path = os.path.join(INPUT_DIR, f)
    df = pd.read_excel(path, header=None)
    testi = df.astype(str).values.flatten()

    has_ah = any("(A-H)" in t for t in testi)
    has_iz = any("(I-Z)" in t for t in testi)

    base = os.path.splitext(f)[0]

    if has_ah and has_iz:
        for flt in ("AH", "IZ"):
            out = os.path.join(OUTPUT_DIR, f"{base}_{flt.lower()}.png")
            if os.path.exists(out):
                print(f"⏭️  {out} già esiste")
                continue
            genera_report(path, flt)
    else:
        out = os.path.join(OUTPUT_DIR, f"{base}.png")
        if not os.path.exists(out):
            genera_report(path)

print("✅ Completato.")
