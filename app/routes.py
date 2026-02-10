"""
Rutas Flask de la aplicación.
"""
from flask import Blueprint, render_template, request, send_file

from app.parsers import extraer_texto_pdf, parsear_historia_clinica
from app.excel import completar_carton_medico
from app.converters import xlsx_a_pdf_con_libreoffice

bp = Blueprint("main", __name__)


@bp.route("/", methods=["GET"])
def home():
    return render_template("index.html", error=None)


@bp.route("/generar", methods=["POST"])
def generar():
    file = request.files.get("pdf")
    if not file or not file.filename.lower().endswith(".pdf"):
        return render_template("index.html", error="Subí un archivo PDF válido.")

    # 1) Leer texto del PDF
    texto = extraer_texto_pdf(file)

    # 2) Parsear datos relevantes
    datos = parsear_historia_clinica(texto)

    # 3) Completar la plantilla en XLSX (en memoria)
    try:
        xlsx_bytes = completar_carton_medico(datos)
    except FileNotFoundError as e:
        return render_template("index.html", error=str(e))

    # 4) Convertir ese XLSX a PDF con LibreOffice
    try:
        pdf_bytes = xlsx_a_pdf_con_libreoffice(xlsx_bytes)
    except Exception as e:
        return render_template(
            "index.html",
            error=f"Error al convertir a PDF con LibreOffice: {e}",
        )

    # 5) Enviar el PDF como descarga
    id_paciente = (datos.get("id") or "").strip()
    safe_id = "".join(ch for ch in id_paciente if ch.isalnum() or ch in ("-", "_"))

    if safe_id:
        nombre_archivo = f"Carton_medico_{safe_id}.pdf"
    else:
        nombre_archivo = "Carton_medico_sin_id.pdf"

    return send_file(
        pdf_bytes,
        as_attachment=True,
        download_name=nombre_archivo,
        mimetype="application/pdf",
    )
