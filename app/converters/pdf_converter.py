"""
Conversión de XLSX a PDF usando LibreOffice + recorte de márgenes con PyMuPDF.
"""
import io
import os
import subprocess
import tempfile

import fitz  # PyMuPDF

from config import (
    SOFFICE_PATH,
    MARGEN_PAG1_LEFT, MARGEN_PAG1_RIGHT, MARGEN_PAG1_TOP, MARGEN_PAG1_BOTTOM,
    MARGEN_PAG2_LEFT, MARGEN_PAG2_RIGHT, MARGEN_PAG2_TOP, MARGEN_PAG2_BOTTOM,
)


def _mm_a_pt(mm):
    """Convierte milímetros a puntos PDF."""
    return mm * 72.0 / 25.4


def xlsx_a_pdf_con_libreoffice(xlsx_bytes):
    """
    Recibe un BytesIO con el XLSX,
    lo guarda en un archivo temporal,
    llama a LibreOffice en modo headless para convertir a PDF,
    recorta márgenes y devuelve otro BytesIO con el PDF.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        xlsx_path = os.path.join(tmpdir, "carton_temp.xlsx")
        pdf_path = os.path.join(tmpdir, "carton_temp.pdf")

        # Guardar XLSX en disco
        with open(xlsx_path, "wb") as f:
            f.write(xlsx_bytes.getvalue())

        # Llamar a LibreOffice en modo headless
        cmd = [
            SOFFICE_PATH,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", tmpdir,
            xlsx_path,
        ]

        resultado = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        if resultado.returncode != 0 or not os.path.exists(pdf_path):
            raise RuntimeError("No se pudo convertir el XLSX a PDF con LibreOffice.")

        # Recortar márgenes del PDF
        doc = fitz.open(pdf_path)

        for i, page in enumerate(doc):
            rect = page.rect

            if i == 0:  # Página 1 (Frente)
                nuevo = fitz.Rect(
                    rect.x0 + _mm_a_pt(MARGEN_PAG1_LEFT),
                    rect.y0 + _mm_a_pt(MARGEN_PAG1_TOP),
                    rect.x1 - _mm_a_pt(MARGEN_PAG1_RIGHT),
                    rect.y1 - _mm_a_pt(MARGEN_PAG1_BOTTOM),
                )
            elif i == 1:  # Página 2 (Dorso)
                nuevo = fitz.Rect(
                    rect.x0 + _mm_a_pt(MARGEN_PAG2_LEFT),
                    rect.y0 + _mm_a_pt(MARGEN_PAG2_TOP),
                    rect.x1 - _mm_a_pt(MARGEN_PAG2_RIGHT),
                    rect.y1 - _mm_a_pt(MARGEN_PAG2_BOTTOM),
                )
            else:  # Páginas extra (sin recorte)
                nuevo = rect

            page.set_cropbox(nuevo)

        pdf_bytes = io.BytesIO()
        doc.save(pdf_bytes)
        doc.close()
        pdf_bytes.seek(0)
        return pdf_bytes
