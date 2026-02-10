import os
import io
import re
from flask import Flask, request, render_template_string, send_file
from openpyxl import load_workbook
from openpyxl.worksheet.page import PageMargins
import fitz  # PyMuPDF
import tempfile
import subprocess
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Font



from flask import Flask, request, render_template_string, send_file
from openpyxl import load_workbook
import fitz  # PyMuPDF

app = Flask(__name__)

# ========== HTML SENCILLO (después si querés le copiamos el CSS lindo del otro servidor) ==========
PAGE = """
<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>Cartón médico - Autocompletado</title>
</head>
<body style="font-family: system-ui, sans-serif; background:#0f172a; color:#e5e7eb;">
  <div style="max-width:800px;margin:40px auto;padding:24px;border-radius:16px;background:#111827;">
    <h1 style="margin-top:0;">Cartón médico – Autocompletar desde Historia Clínica (PDF)</h1>
    <p>Subir la historia clínica en PDF y el servidor va a rellenar automáticamente algunos datos en la plantilla <b>Cartón médico.xlsx</b>.</p>

    {% if error %}
      <div style="margin:12px 0;padding:12px;border-radius:8px;background:rgba(248,113,113,.15);border:1px solid rgba(248,113,113,.6);">
        <b>Error:</b> {{ error }}
      </div>
    {% endif %}

    <form method="post" action="/generar" enctype="multipart/form-data">
      <label>Historia clínica (PDF):
        <input type="file" name="pdf" accept="application/pdf" required>
      </label>
      <br><br>
      <button type="submit" style="padding:10px 16px;border-radius:10px;border:none;
              background:linear-gradient(180deg,#22d3ee,#38bdf8);color:#022c22;font-weight:600;cursor:pointer;">
        Generar cartón médico
      </button>
    </form>

    <p style="font-size:12px;color:#9ca3af;margin-top:20px;">
      El archivo generado será un <b>.pdf</b> que se debera guardar en la carpeta del paciente, listo para imprimir.
    </p>
  </div>
</body>
</html>
"""

# ========== PARSEAR TEXTO DEL PDF ==========

def extraer_texto_pdf(file_storage):
    """Devuelve TODO el texto del PDF como string."""
    with fitz.open(stream=file_storage.read(), filetype="pdf") as doc:
        partes = []
        for page in doc:
            partes.append(page.get_text())
    return "\n".join(partes)

def parsear_historia_clinica(texto):
    """
    Saca:
    - nombre_completo  
    - dni
    - edad
    - id
    Devuelve dict con esas keys (pueden ser None si no se encuentran).
    """
    data = {
        "nombre_completo": None,
        "edad": None,
        "peso": None,
        "id": None,
        "diagnostico": None,
        "histologia": None,
        "estad_t": None,
        "estad_n": None,
        "estad_m": None,
        "estad_estadio": None,
        "interrogatorio":None,


        # acá más adelante agregamos: diagnostico, estadio, etc.
    }

    data["_bloques_interrogatorio"] = []
    lineas = [l.strip() for l in texto.splitlines() if l.strip()]

    for i, linea in enumerate(lineas):
        # Nombre del paciente: en tu PDF es "Paciente" en una línea, y EN LA SIGUIENTE el nombre
        if linea.upper() == "PACIENTE" and i + 1 < len(lineas):
            data["nombre_completo"] = lineas[i + 1].strip()
        
        
          # === ID del paciente (robusto, evita confundir con años) ===
        if data.get("id") is None and "id" in linea.lower():

            # Vamos a mirar: linea actual, siguiente y la otra
            posibles = []
            ventanas = [linea]

            if i + 1 < len(lineas):
                ventanas.append(lineas[i+1])
            if i + 2 < len(lineas):
                ventanas.append(lineas[i+2])

            # Buscamos SOLO números grandes (evita 19, 25, 2025, etc.)
            for v in ventanas:
                nums = re.findall(r"\b(\d{5,8})\b", v)
                for n in nums:
                    # Filtrar números tipo 2025 (años)
                    if int(n) > 3000:   # Un ID nunca es menor a esto
                        posibles.append(int(n))

            # Si encontramos candidatos, agarramos el MAYOR (siempre es el ID)
            if posibles:
                data["id"] = str(max(posibles))

        # Edad: "Edad" en una línea y en la siguiente "39 años"
        if linea.upper() == "EDAD" and i + 1 < len(lineas):
            m_edad = re.search(r"(\d+)", lineas[i + 1])
            if m_edad:
                data["edad"] = int(m_edad.group(1))

        # Peso corporal: aparece como "Peso corporal" y en la línea siguiente el valor
        if "peso corporal" in linea.lower() and i + 1 < len(lineas):
            m_peso = re.search(r"(\d+(?:[.,]\d+)?)", lineas[i+1])
            if m_peso:
                data["peso"] = float(m_peso.group(1).replace(",", "."))
           # === ID del paciente ===
        # Buscamos el ID tomando la línea actual + la siguiente (por si el PDF corta el texto raro)
        
        # ========= DIAGNÓSTICO (desde Grupo) =========
        # En el PDF aparece así:
        # Grupo:
        # TUMOR MALIGNO DEL CUELLO DEL UTERO
        if "GRUPO" in linea.upper():
            if i + 1 < len(lineas):
                linea_diagnostico = lineas[i + 1].strip()

        # evitar texto irrelevante como "diagnostico"
                if linea_diagnostico.lower() != "diagnostico":
                    data["diagnostico"] = linea_diagnostico

        # ========= ESTADIFICACIÓN: T =========
        # Busca líneas tipo "T: T2b"
        if linea.strip().upper().startswith("T:"):
            m_t = re.search(r"T\s*:\s*([A-Za-z0-9]+)", linea)
            if m_t:
                data["estad_t"] = m_t.group(1)   # ej: "T2b"


        
        # ========= ESTADIFICACIÓN: N =========
        if linea.strip().upper().startswith("N:"):
            m_n = re.search(r"N\s*:\s*([A-Za-z0-9]+)", linea)
            if m_n:
                data["estad_n"] = m_n.group(1)   # ej: "N1"


        
        # ========= ESTADIFICACIÓN: M =========
        # Ejemplo: "M: M0"
        if linea.strip().upper().startswith("M:"):
            m_m = re.search(r"M\s*:\s*([A-Za-z0-9]+)", linea, re.IGNORECASE)
            if m_m:
                data["estad_m"] = m_m.group(1)   # ej: "M0"


        # ========= ESTADIFICACIÓN: ESTADIO =========
        # Ejemplo en el PDF: "Estadio: Stage IIIC1"
        if linea.upper().startswith("ESTADIO:"):
            m_e = re.search(r"Estadio:\s*(.+)", linea, re.IGNORECASE)
            if m_e:
                data["estad_estadio"] = m_e.group(1).strip()
        
         # ========= HISTOLOGÍA =========
        # Ejemplo en PDF: "Histologia: CARCINOMA ESCAMOSO"
        if linea.lower().startswith("histologia"):
            m_h = re.search(r"histologia\s*:\s*(.+)", linea, re.IGNORECASE)
            if m_h:
                data["histologia"] = m_h.group(1).strip()

    

        # ========= INTERROGATORIO (último bloque) =========
        if "INTERROGATORIO" in linea.upper():
            j = i + 1
            bloque = []

            # saltar posibles líneas vacías
            while j < len(lineas) and not lineas[j].strip():
                j += 1

            # saltar línea tipo "BRUNL   19/05/2025 18:35:52" si está
            if j < len(lineas) and re.search(r"\d{2}/\d{2}/\d{4}", lineas[j]):
                j += 1

            # ahora sí, juntar TODO el texto del interrogatorio
            while j < len(lineas):
                l2 = lineas[j].rstrip()

                # si está vacía, la guardamos como salto de línea y seguimos
                if not l2:
                    bloque.append("")
                    j += 1
                    continue

                # si aparece un título nuevo en MAYÚSCULAS, cortamos el bloque
                if (
                    l2.isupper()
                    and len(l2) <= 60
                    and "INTERROGATORIO" not in l2.upper()
                    and not l2.lstrip().startswith(("*", "-"))
                    ):
                    break

                bloque.append(l2)
                j += 1

            if bloque:
                data["_bloques_interrogatorio"].append("\n".join(bloque))

      # Elegimos SIEMPRE el último interrogatorio encontrado
    if data["_bloques_interrogatorio"]:
            data["interrogatorio"] = data["_bloques_interrogatorio"][-1]
    else:
            data["interrogatorio"] = None

     # limpiamos la clave interna
    data.pop("_bloques_interrogatorio", None)

   # ====== PRESCRIPCIÓN BRAQUI (dosis total / N° fx / dosis por fx) ======
    presc_braqui = parsear_prescripcion_braqui(texto)
    if presc_braqui:
        data["braqui_dosis_total"] = presc_braqui["dosis_total"]
        data["braqui_n_fracciones"] = presc_braqui["n_fracciones"]
        data["braqui_dosis_por_fraccion"] = presc_braqui["dosis_por_fraccion"]
    else:
        data["braqui_dosis_total"] = None
        data["braqui_n_fracciones"] = None
        data["braqui_dosis_por_fraccion"] = None



    return data


def parsear_prescripcion_braqui(texto):
    """
    Busca en el PDF la tabla 'Dosis por fracción / N° de Fracciones / ...'
    y devuelve SOLO la prescripción de BRAQUI (no la de RTE).

    Devuelve dict con:
      - dosis_por_fraccion
      - n_fracciones
      - fracciones_por_semana
      - dosis_total
      - dosis_total_con_externa
    o None si no se pudo encontrar.
    """
    lineas = [l.strip() for l in texto.splitlines()]

    # 1) Busco la zona que arranca en 'Dosis por fracción'
    idx_dosispor = next(
        (i for i, l in enumerate(lineas) if "dosis por fracción" in l.lower()),
        None
    )
    if idx_dosispor is None:
        return None

    # 2) Hasta 'Conducta Terapéutica' (fin de esa tabla)
    idx_end = next(
        (i for i, l in enumerate(lineas[idx_dosispor:], start=idx_dosispor)
         if "conducta terapéutica" in l.lower()),
        len(lineas)
    )

    # Me quedo con la parte intermedia
    relevantes = [l.strip() for l in lineas[idx_dosispor+1:idx_end] if l.strip()]

    grupos = []
    num_re = re.compile(r"^\d+(?:[.,]\d+)?$")  # 2  2.00  24,5

    k = 0
    while k < len(relevantes):
        if num_re.match(relevantes[k]):
            nums = []
            start = k
            while k < len(relevantes) and len(nums) < 5 and num_re.match(relevantes[k]):
                nums.append(float(relevantes[k].replace(",", ".")))
                k += 1

            # Si tengo exactamente 5 números → (Gy/fx, N fx, fx/sem, Dosis total, Dosis total+RTE)
            if len(nums) == 5:
                prev = " ".join(relevantes[max(0, start-6):start])
                nxt  = " ".join(relevantes[k:min(len(relevantes), k+6)])
                grupos.append({"values": nums, "prev": prev, "next": nxt})
        else:
            k += 1

    if not grupos:
        return None

    def ctx(g):
        return (g["prev"] + " " + g["next"]).lower()

    marcadores_braqui = ["bqt", "braqui", "uterovaginal",
                         "vaginal", "hr-ctv", "cervicovaginal"]

    # 1) Si algún grupo habla de anestesia → ese es braqui (tomamos el último por si hay varios)
    grupos_anestesia = [g for g in grupos if "anestesia" in ctx(g)]
    if grupos_anestesia:
        elegido = grupos_anestesia[-1]
    else:
        # 2) Si no, tomamos el ÚLTIMO grupo que mencione BQT / Braqui / etc.
        grupos_marcadores = [
            g for g in grupos
            if any(m in ctx(g) for m in marcadores_braqui)
        ]
        if grupos_marcadores:
            elegido = grupos_marcadores[-1]
        else:
            # 3) Recontra-fallback: el último grupo de todos
            elegido = grupos[-1]

    v = elegido["values"]
    return {
        "dosis_por_fraccion": v[0],
        "n_fracciones": v[1],
        "fracciones_por_semana": v[2],
        "dosis_total": v[3],
        "dosis_total_con_externa": v[4],
    }


def separar_apellido_nombre(nombre_completo):
    """
    Dado algo tipo 'LUDUEÑA CLAUDIA ESTEFANIA'
    devuelve ('LUDUEÑA', 'CLAUDIA ESTEFANIA')
    Si no se puede, todo va a Apellido.
    """
    if not nombre_completo:
        return None, None
    partes = nombre_completo.strip().split()
    if len(partes) == 1:
        return partes[0], ""
    apellido = partes[0]
    nombre = " ".join(partes[1:])
    return apellido, nombre
def escribir_en_una_linea(ws, cell_addr, texto, base_font_size=12, min_font_size=8, horizontal="center"):
    """
    UNA sola línea + centrado prolijo:
    - shrink_to_fit
    - ajusta tamaño de fuente si es largo
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

    # Alineación como el resto (centrado)
    cell.alignment = Alignment(
        wrap_text=False,
        shrink_to_fit=True,
        vertical="center",
        horizontal=horizontal  # "center" o "centerContinuous"
    )

    # Tamaño de fuente según longitud
    largo = len(texto)
    size = base_font_size
    if largo > 25: size = 11
    if largo > 35: size = 10
    if largo > 45: size = 9
    if largo > 55: size = 8
    if size < min_font_size:
        size = min_font_size

    try:
        cell.font = Font(
            name=cell.font.name,
            bold=cell.font.bold,
            italic=cell.font.italic,
            size=size
        )
    except:
        cell.font = Font(size=size)

# ========== LLENAR PLANTILLA CARTÓN MÉDICO ==========
def normalizar_interrogatorio(txt: str) -> str:
    """
    Une saltos de línea "de corte" (wrap) del PDF para no desperdiciar espacio,
    pero respeta:
      - ítems que empiezan con '-'
      - línea inicial tipo 'Pte de ...'
    """
    if not txt:
        return txt

    lineas = [l.strip() for l in txt.splitlines()]
    lineas = [l for l in lineas if l != ""]  # sacamos vacías

    out = []
    for l in lineas:
        es_item = l.startswith("-") or l.startswith("•")

        if not out:
            out.append(l)
            continue

        prev = out[-1].rstrip()

        # Regla 1: si la línea actual es un ítem ("-..."), siempre va en nueva línea
        if es_item:
            out.append(l)
            continue

        # Regla 2: si la anterior es un ítem, lo que sigue suele ser continuación -> unir
        # (pero con tu PDF generalmente las continuaciones vienen sin '-')
        if prev.startswith("-"):
            # si la continuación empieza en minúscula o con letra, la pegamos
            out[-1] = prev + " " + l
            continue

        # Regla 3: unir cortes típicos del PDF:
        # - anterior termina en coma o no termina en punto
        # - y la siguiente empieza en minúscula (continuación de frase)
        if (
            (prev.endswith(",") or (not prev.endswith((".", ":", ";", "?", "!", ")"))))
            and l[:1].islower()
        ):
            out[-1] = prev + " " + l
        else:
            out.append(l)

    return "\n".join(out)


def completar_carton_medico(datos):
    """
    Abre 'Cartón médico.xlsx' (mismo directorio), rellena algunos campos en la hoja 'Frente'
    y devuelve el contenido del xlsx como bytes.
    """
    base_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(base_dir, "Cartón médico.xlsx")

    if not os.path.exists(template_path):
        raise FileNotFoundError(f"No se encontró la plantilla Cartón médico.xlsx en {template_path}")

    wb = load_workbook(template_path)
    # === Márgenes en CERO para que el PDF salga sin bordes ===
    margins = PageMargins(
        left=0,   # margen izquierdo
        right=0,  # margen derecho
        top=0,    # margen superior
        bottom=0, # margen inferior
        header=0,
        footer=0
    )

    # Aplicar esos márgenes a todas las hojas (Frente y Dorso)
    for hoja in wb.worksheets:
        hoja.page_margins = margins
    ws_f = wb["Frente"]  # hoja Frente

    nombre_completo = datos.get("nombre_completo")
    hc = datos.get("hc")
    dni = datos.get("dni")
    edad = datos.get("edad")

    apellido, nombre = separar_apellido_nombre(nombre_completo)

    # celdas según la plantilla que miramos:
    # C9:F9  → Apellido
    # C10:F10 → Nombre
    # I9:J9  → Edad
    # (más cosas las vamos agregando luego)

    if apellido:
        escribir_en_una_linea(ws_f, "C9", apellido, horizontal="center")

    if nombre:
        escribir_en_una_linea(ws_f, "C10", nombre, horizontal="center")

    if edad is not None:
        ws_f["I9"] = edad
            # Peso corporal
        # Peso corporal
    peso = datos.get("peso")
    if peso is not None:
        ws_f["I10"] = f"{peso} kg"
     # ID del paciente en el cartón médico
    id_paciente = datos.get("id")
    if id_paciente:
        ws_f["C11"] = id_paciente  
     # DIAGNÓSTICO – GRUPO
    if datos.get("diagnostico"):
        ws_f["C14"] = datos["diagnostico"]

    # HISTOLOGÍA
    if datos.get("histologia"):
        ws_f["C16"] = datos["histologia"]

    
    # === ESTADIFICACIÓN: T ===
    if datos.get("estad_t"):
        ws_f["C15"] = f"T: {datos['estad_t']}"

    # === ESTADIFICACIÓN: N ===
    if datos.get("estad_n"):
        ws_f["D15"] = f"N: {datos["estad_n"]}"

    # === ESTADIFICACIÓN: N ===
    if datos.get("estad_m"):
        ws_f["E15"] = f"M: {datos["estad_m"]}"

    # === ESTADIFICACIÓN: ESTADIO ===
    if datos.get("estad_estadio"):
        ws_f["G15"] = f"ESTADIO: {datos["estad_estadio"]}"

    # === HISTOLOGÍA ===
    if datos.get("histologia"):
        ws_f["C16"] = datos["histologia"]

    # === INTERROGATORIO (C17) ===
    if datos.get("interrogatorio"):
        texto_inter = normalizar_interrogatorio(datos["interrogatorio"])
        cell = ws_f["C17"]   # C17 es donde va el texto largo

    if isinstance(cell, MergedCell):
        for merged_range in ws_f.merged_cells.ranges:
            if cell.coordinate in merged_range:
                top_left = ws_f.cell(
                    row=merged_range.min_row,
                    column=merged_range.min_col
                )
                top_left.value = texto_inter
                # 👉 ACÁ va la alineación
                top_left.alignment = Alignment(
                    wrap_text=True,
                    vertical="top",
                    horizontal="left"
                )
                break
    else:
        cell.value = texto_inter
        # 👉 ACÁ va la alineación
        cell.alignment = Alignment(
            wrap_text=True,
            vertical="top",
            horizontal="left"
        )


    # Podríamos guardar HC y DNI en algún lugar del frente o dorso.
    # Por ahora los dejamos solo impresos en la parte superior de Dorso si definimos dónde:
    ws_d = wb["Dorso"]
    

     # =======================
    # 2. PRESCRIPCIÓN BRAQUI
    # =======================
    def formatear_dosis(valor):
        """Devuelve 'X.Y Gy' con 1 decimal."""
        if valor is None:
            return None
        try:
            v = float(valor)
            return f"{v:.1f} Gy"
        except:
            return f"{valor} Gy"


    dosis_total = datos.get("braqui_dosis_total")
    n_fx        = datos.get("braqui_n_fracciones")
    dosis_fx    = datos.get("braqui_dosis_por_fraccion")

    # Tabla 2. Prescripción (columna 'Prescripción')
    # C36 → Dosis Total
    # C37 → N˚ de fracción
    # C38 → Dosis x fracción

    if dosis_total is not None:
          ws_f["C36"] = formatear_dosis(dosis_total)  # podés poner f"{dosis_total} Gy" si querés

    if n_fx is not None:
        ws_f["C37"] = int(n_fx) if float(n_fx).is_integer() else n_fx

    if dosis_fx is not None:
         ws_f["C38"] = formatear_dosis(dosis_fx)



    # Guardar a un buffer en memoria
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

   

# ========== RUTAS FLASK ==========

@app.route("/", methods=["GET"])
def home():
    return render_template_string(PAGE, error=None)

@app.route("/generar", methods=["POST"])
def generar():
    file = request.files.get("pdf")
    if not file or not file.filename.lower().endswith(".pdf"):
        return render_template_string(PAGE, error="Subí un archivo PDF válido.")

    # 1) Leer texto del PDF
    texto = extraer_texto_pdf(file)

    # 2) Parsear datos relevantes
    datos = parsear_historia_clinica(texto)

    # 3) Completar la plantilla en XLSX (en memoria)
    try:
        xlsx_bytes = completar_carton_medico(datos)
    except FileNotFoundError as e:
        return render_template_string(PAGE, error=str(e))

    # 4) Convertir ese XLSX a PDF con LibreOffice en modo headless
    try:
        pdf_bytes = xlsx_a_pdf_con_libreoffice(xlsx_bytes)
    except Exception as e:
        # Si algo falla, mostramos error en la página
        return render_template_string(
            PAGE,
            error=f"Error al convertir a PDF con LibreOffice: {e}"
        )

    # 5) Enviar el PDF como descarga (nombre dinámico)
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
        mimetype="application/pdf"
)


def xlsx_a_pdf_con_libreoffice(xlsx_bytes):
    """
    Recibe un BytesIO con el XLSX,
    lo guarda en un archivo temporal,
    llama a LibreOffice en modo headless para convertir a PDF
    y devuelve otro BytesIO con el PDF.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        xlsx_path = os.path.join(tmpdir, "carton_temp.xlsx")
        pdf_path = os.path.join(tmpdir, "carton_temp.pdf")

        # Guardar XLSX en disco
        with open(xlsx_path, "wb") as f:
            f.write(xlsx_bytes.getvalue())

        # Ruta al ejecutable de LibreOffice.
        # Si 'soffice' no está en el PATH, podés poner la ruta completa, por ejemplo:
        # soffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
        # Ruta del ejecutable de LibreOffice (ajustada para Windows)
        soffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"



        # Llamar a LibreOffice en modo headless
        cmd = [
            soffice_path,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", tmpdir,
            xlsx_path,
        ]

        resultado = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        if resultado.returncode != 0 or not os.path.exists(pdf_path):
            raise RuntimeError("No se pudo convertir el XLSX a PDF con LibreOffice.")

        # === Recortar márgenes del PDF usando fitz ===
        doc = fitz.open(pdf_path)

        # margen a recortar (en milímetros)
        margen_mm = 30  # podés subirlo a 8-10 si querés aún menos borde
        margen_pt = margen_mm * 72.0 / 25.4  # conversión mm -> puntos

                # === Recortar márgenes del PDF usando fitz, por página ===
        # Márgenes en mm para cada página (índice 0 = página 1, índice 1 = página 2)
        # Cambiá estos valores a gusto:
        margen_pag1_mm = 30   # página 1
        margen_pag2_mm = 20   # página 2
        margen_default_mm = 5  # por si hubiera más páginas

        def mm_a_pt(mm):
            return mm * 72.0 / 25.4

                # === Recortar márgenes del PDF usando fitz, por página ===

        # Configuración de márgenes por página (en mm)
        # PÁGINA 1 → recorte total (ya te quedó perfecto)
        margen_pag1_mm = 30  

        # PÁGINA 2 → recortar SOLO los costados
        margen_left_right_pag2_mm = 40  # ajustá este valor si querés más o menos recorte
        # No tocamos el margen superior ni inferior en esta página

        def mm_to_pt(mm):
            return mm * 72.0 / 25.4

                # === Recortar márgenes del PDF usando fitz, por página ===

        # ----- CONFIGURACIÓN DE MÁRGENES (en mm) -----
        # Página 1 (ya te quedaba bien):
        margen1_left  = 28
        margen1_right = 28
        margen1_top   = 33
        margen1_bottom= 33

        # Página 2 (AHORA TOTALMENTE PERSONALIZABLE):
        margen2_left   = 47   # modificá este
        margen2_right  = 47  # modificá este
        margen2_top    = 17   # modificá este
        margen2_bottom = 17   # modificá este

        # Conversor mm -> puntos PDF
        def mm_to_pt(mm):
            return mm * 72.0 / 25.4

        for i, page in enumerate(doc):
            rect = page.rect

            if i == 0:   # ====== PÁGINA 1 ======
                nuevo = fitz.Rect(
                    rect.x0 + mm_to_pt(margen1_left),
                    rect.y0 + mm_to_pt(margen1_top),
                    rect.x1 - mm_to_pt(margen1_right),
                    rect.y1 - mm_to_pt(margen1_bottom)
                )

            elif i == 1: # ====== PÁGINA 2 ======
                nuevo = fitz.Rect(
                    rect.x0 + mm_to_pt(margen2_left),
                    rect.y0 + mm_to_pt(margen2_top),
                    rect.x1 - mm_to_pt(margen2_right),
                    rect.y1 - mm_to_pt(margen2_bottom)
                )

            else:        # Páginas extra (por si acaso)
                nuevo = rect

            page.set_cropbox(nuevo)




        pdf_bytes = io.BytesIO()
        doc.save(pdf_bytes)
        doc.close()
        pdf_bytes.seek(0)
        return pdf_bytes



if __name__ == "__main__":
    print(">> Servidor Cartón médico en http://127.0.0.1:5001")
    app.run(host="0.0.0.0", port=5001, debug=False)
