"""
Extracción de texto desde archivos PDF.
"""
import fitz  # PyMuPDF


def extraer_texto_pdf(file_storage):
    """Devuelve TODO el texto del PDF como string."""
    with fitz.open(stream=file_storage.read(), filetype="pdf") as doc:
        partes = []
        for page in doc:
            partes.append(page.get_text())
    return "\n".join(partes)
