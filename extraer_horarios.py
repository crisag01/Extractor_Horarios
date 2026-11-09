"""
extraer_horarios.py
-------------------
Extrae datos de imágenes de horarios universitarios y genera Horarios.xlsx.
100% local, sin IA.

Dependencias:
    pip install pytesseract pillow openpyxl numpy
    + Tesseract OCR instalado:
      Windows: https://github.com/UB-Mannheim/tesseract/wiki

Uso:
    python extraer_horarios.py
"""

import re
from pathlib import Path

import numpy as np
from PIL import Image
import pytesseract
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Configuración ─────────────────────────────────────────────────────────────
IMAGE_FOLDER = r"screenshots"
OUTPUT_FILE  = r"Horarios.xlsx"

# Descomenta si Tesseract no está en el PATH:
#pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

SUPPORTED_EXT = {".png", ".jpg", ".jpeg", ".webp", ".bmp"}

COL_NAMES = [
    "Carrera", "Acronimo", "Sistema", "Semestre",
    "Profesor", "CVE", "Asignatura", "Grupo", "Tipo",
    "Lun", "Mar", "Mie", "Jue", "Vie", "Sab", "Salon", "Cupo",
]

FIELD_MAP = [
    "Profesor", "CVE", "Asignatura", "Grupo", "Tipo",
    "Lun", "Mar", "Mie", "Jue", "Vie", "Sab", "Salon", "Cupo",
]

TIME_FIELDS  = {"Lun", "Mar", "Mie", "Jue", "Vie", "Sab"}
DIGIT_FIELDS = {"CVE", "Cupo"}
LEFT_COLS    = {"Carrera", "Profesor", "Asignatura"}

# ── Preprocesado OCR ──────────────────────────────────────────────────────────

def preprocess(pil_img):
    """Escala 3x y binariza: texto negro sobre fondo blanco."""
    big = pil_img.resize((pil_img.width * 3, pil_img.height * 3), Image.LANCZOS)
    g = np.array(big.convert("L"))
    thresh = np.where(g < 180, 0, 255).astype(np.uint8)
    return Image.fromarray(thresh)


def ocr(pil_img):
    proc = preprocess(pil_img)
    txt = pytesseract.image_to_string(proc, config="--psm 6 --oem 3")
    return " ".join(txt.split())


def clean_time(txt):
    times = re.findall(r'\d{1,2}:\d{2}', txt)
    return "\n".join(times) if times else ""


def only_digits(txt):
    return re.sub(r'\D', '', txt)

# ── Detección de estructura ───────────────────────────────────────────────────

def detect_columns(arr):
    """Detecta los segmentos x de cada columna a partir del encabezado azul."""
    blue = (arr[:, :, 0] < 80) & (arr[:, :, 1] < 80) & (arr[:, :, 2] > 100)
    in_blue = False
    segments = []
    start = None
    for c in range(arr.shape[1]):
        has = blue[:, c].any()
        if has and not in_blue:
            start = c
            in_blue = True
        elif not has and in_blue:
            segments.append((start, c))
            in_blue = False
    if in_blue:
        segments.append((start, arr.shape[1]))
    return segments


def detect_row_borders(arr, col_segs):
    """
    Detecta las líneas horizontales oscuras que separan las filas de datos.
    Devuelve lista de y donde hay un separador (una línea por borde).
    """
    x1 = col_segs[0][0]
    x2 = col_segs[-1][1]
    width = x2 - x1

    # Fin del encabezado: última fila con MUCHOS píxeles azules (>100).
    # Ignoramos píxeles azules sueltos que aparecen más abajo.
    blue = (arr[:, :, 0] < 80) & (arr[:, :, 1] < 80) & (arr[:, :, 2] > 100)
    blue_count = blue.sum(axis=1)
    header_rows = np.where(blue_count > 100)[0]
    if len(header_rows) == 0:
        return []
    search_start = int(header_rows.max()) + 1

    borders = []
    prev_border = -10
    for y in range(search_start, arr.shape[0]):
        row = arr[y, x1:x2]
        n_dark   = ((row[:, 0] < 60) & (row[:, 1] < 60) & (row[:, 2] < 60)).sum()
        n_yellow = ((row[:, 0] > 200) & (row[:, 1] > 200) & (row[:, 2] < 200)).sum()
        # Línea separadora: mayoría oscura, nada amarillo
        if n_dark > width * 0.6 and n_yellow < 50:
            if y - prev_border > 5:   # no contar líneas dobles
                borders.append(y)
                prev_border = y

    return borders


def detect_yellow_separator(arr):
    """Línea dorada delgada encima del título (separa carrera del resto)."""
    gold = (arr[:, :, 0] > 180) & (arr[:, :, 1] > 140) & (arr[:, :, 2] < 80)
    rows = np.where(gold.sum(axis=1) > 50)[0]
    return int(rows.min()) if len(rows) else 110

# ── Parseo de una imagen ──────────────────────────────────────────────────────

def parse_image(path: Path):
    img = Image.open(path)
    arr = np.array(img)

    # 1) Carrera
    sep_y = detect_yellow_separator(arr)
    carrera = ocr(img.crop((0, 0, img.width, sep_y))).strip()

    # 2) Semestre
    full_text = pytesseract.image_to_string(img, config="--psm 3")
    m = re.search(r'Semestre[:\s]+(\S+)', full_text, re.IGNORECASE)
    semestre = m.group(1) if m else ""

    # 3) Sistema
    sistema = "SUA" if "SISTEMA DE UNIVERSIDAD ABIERTA" in carrera.upper() else "ESC"

    # 4) Columnas
    col_segs = detect_columns(arr)
    if len(col_segs) < 3:
        print(f"  ⚠  No se detectaron columnas en {path.name}")
        return []

    # 5) Bordes de filas (líneas oscuras entre datos)
    borders = detect_row_borders(arr, col_segs)
    if len(borders) < 2:
        print(f"  ⚠  No se detectaron filas de datos en {path.name}")
        return []

    # Cada par consecutivo de bordes define una fila de datos
    data_rows = []
    for ri in range(len(borders) - 1):
        ry1 = borders[ri] + 2      # dentro de la fila (evitar la línea del borde)
        ry2 = borders[ri + 1] - 2

        if ry2 - ry1 < 10:
            continue  # franja demasiado pequeña, saltar

        row_data = {
            "Carrera":  carrera,
            "Acronimo": "",
            "Sistema":  sistema,
            "Semestre": semestre,
        }

        for fi, (x1, x2) in enumerate(col_segs):
            if fi >= len(FIELD_MAP):
                break
            field = FIELD_MAP[fi]
            raw = ocr(img.crop((x1, ry1, x2, ry2)))

            if field in TIME_FIELDS:
                row_data[field] = clean_time(raw)
            elif field in DIGIT_FIELDS:
                row_data[field] = only_digits(raw)
            else:
                row_data[field] = raw

        data_rows.append(row_data)

    return data_rows

# ── Excel ─────────────────────────────────────────────────────────────────────

HDR_FILL  = PatternFill("solid", fgColor="00008B")
HDR_FONT  = Font(bold=True, color="FFFFFF", name="Arial", size=10)
CELL_FONT = Font(name="Arial", size=10)
CENTER    = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT      = Alignment(horizontal="left",   vertical="center", wrap_text=True)
THIN      = Side(style="thin", color="000000")
BORDER    = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

COL_WIDTHS = {
    "Carrera": 28, "Acronimo": 10, "Sistema": 10, "Semestre": 12,
    "Profesor": 26, "CVE": 8, "Asignatura": 34, "Grupo": 8,
    "Tipo": 14, "Lun": 13, "Mar": 13, "Mie": 13,
    "Jue": 13, "Vie": 13, "Sab": 13, "Salon": 10, "Cupo": 8,
}


def save_excel(rows, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Horarios"

    for ci, name in enumerate(COL_NAMES, 1):
        c = ws.cell(row=1, column=ci, value=name)
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[1].height = 22

    for ri, row in enumerate(rows, 2):
        for ci, name in enumerate(COL_NAMES, 1):
            val = row.get(name, "")
            c = ws.cell(row=ri, column=ci, value=val)
            c.font = CELL_FONT
            c.alignment = LEFT if name in LEFT_COLS else CENTER
            c.border = BORDER

    for ci, name in enumerate(COL_NAMES, 1):
        ws.column_dimensions[get_column_letter(ci)].width = COL_WIDTHS.get(name, 14)

    ws.freeze_panes = "A2"
    wb.save(output_path)

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    folder = Path(IMAGE_FOLDER)
    if not folder.exists():
        print(f"ERROR: La carpeta no existe: {folder}")
        return

    images = sorted(p for p in folder.iterdir()
                    if p.is_file() and p.suffix.lower() in SUPPORTED_EXT)

    if not images:
        print("No se encontraron imágenes en la carpeta.")
        return

    print(f"Carpeta: {folder}")
    print(f"Imágenes encontradas: {len(images)}\n")

    all_rows = []
    for i, img_path in enumerate(images, 1):
        print(f"  [{i}/{len(images)}] Procesando: {img_path.name} ...", end=" ", flush=True)
        try:
            rows = parse_image(img_path)
            all_rows.extend(rows)
            print(f"OK ({len(rows)} fila(s))")
        except Exception as e:
            print(f"ERROR: {e}")

    if not all_rows:
        print("\nNo se extrajeron datos.")
        return

    save_excel(all_rows, OUTPUT_FILE)
    print(f"\nArchivo guardado: {OUTPUT_FILE}")
    print(f"Total de filas: {len(all_rows)}")


if __name__ == "__main__":
    main()
