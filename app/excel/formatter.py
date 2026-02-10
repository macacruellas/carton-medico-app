"""
Funciones de formateo para celdas de Excel.
"""
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Font

from config import FONT_SIZE_BASE, FONT_SIZE_MIN


def escribir_en_una_linea(ws, cell_addr, texto, base_font_size=FONT_SIZE_BASE,
                          min_font_size=FONT_SIZE_MIN, horizontal="center"):
    """
    Escribe texto en UNA sola línea con centrado prolijo:
    - shrink_to_fit
    - ajusta tamaño de fuente según longitud
    - sin wrap
    """
    if not texto:
        return

    texto = str(texto).strip().replace("\n", " ")

    cell = ws[cell_addr]
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                break

    cell.value = texto

    cell.alignment = Alignment(
        wrap_text=False,
        shrink_to_fit=True,
        vertical="center",
        horizontal=horizontal,
    )

    # Tamaño de fuente según longitud
    largo = len(texto)
    size = base_font_size
    if largo > 25:
        size = 11
    if largo > 35:
        size = 10
    if largo > 45:
        size = 9
    if largo > 55:
        size = 8
    if size < min_font_size:
        size = min_font_size

    try:
        cell.font = Font(
            name=cell.font.name,
            bold=cell.font.bold,
            italic=cell.font.italic,
            size=size,
        )
    except Exception:
        cell.font = Font(size=size)


def normalizar_interrogatorio(txt):
    """
    Une saltos de línea "de corte" (wrap) del PDF para no desperdiciar espacio,
    pero respeta:
      - ítems que empiezan con '-' o '•'
      - línea inicial tipo 'Pte de ...'
    """
    if not txt:
        return txt

    lineas = [l.strip() for l in txt.splitlines()]
    lineas = [l for l in lineas if l != ""]

    out = []
    for l in lineas:
        es_item = l.startswith("-") or l.startswith("•")

        if not out:
            out.append(l)
            continue

        prev = out[-1].rstrip()

        if es_item:
            out.append(l)
            continue

        if prev.startswith("-"):
            out[-1] = prev + " " + l
            continue

        if (
            (prev.endswith(",") or not prev.endswith((".", ":", ";", "?", "!", ")")))
            and l[:1].islower()
        ):
            out[-1] = prev + " " + l
        else:
            out.append(l)

    return "\n".join(out)
