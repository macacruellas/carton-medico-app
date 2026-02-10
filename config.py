"""
Configuración centralizada de la aplicación Cartón Médico.
"""
import os
import shutil

# --- Rutas ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "Cartón médico.xlsx")

# --- Servidor Flask ---
HOST = "0.0.0.0"
PORT = 5001

# --- LibreOffice ---
# Detecta automáticamente la ruta de soffice según la plataforma
def _detectar_soffice():
    """Busca el ejecutable de LibreOffice en el sistema."""
    # 1) Si está en el PATH
    path = shutil.which("soffice")
    if path:
        return path
    # 2) Rutas comunes por plataforma
    candidatos = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",  # Windows
        "/usr/bin/soffice",                                     # Linux
        "/usr/local/bin/soffice",                               # Linux alt
        "/Applications/LibreOffice.app/Contents/MacOS/soffice", # macOS
    ]
    for c in candidatos:
        if os.path.isfile(c):
            return c
    # 3) Fallback: confiar en que esté en el PATH
    return "soffice"

SOFFICE_PATH = os.environ.get("SOFFICE_PATH", _detectar_soffice())

# --- Márgenes de recorte del PDF (en mm) ---
# Página 1 (Frente)
MARGEN_PAG1_LEFT = 28
MARGEN_PAG1_RIGHT = 28
MARGEN_PAG1_TOP = 33
MARGEN_PAG1_BOTTOM = 33

# Página 2 (Dorso)
MARGEN_PAG2_LEFT = 47
MARGEN_PAG2_RIGHT = 47
MARGEN_PAG2_TOP = 17
MARGEN_PAG2_BOTTOM = 17

# --- Fuentes ---
FONT_SIZE_BASE = 12
FONT_SIZE_MIN = 8
