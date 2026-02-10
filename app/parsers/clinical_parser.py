"""
Parseo de datos clínicos desde el texto extraído del PDF.
"""
import re


def parsear_historia_clinica(texto):
    """
    Extrae datos clínicos del texto del PDF:
    nombre_completo, dni, edad, id, diagnostico, histologia,
    estadificación (T, N, M, estadio), interrogatorio,
    y prescripción de braquiterapia.
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
        "interrogatorio": None,
    }

    bloques_interrogatorio = []
    lineas = [l.strip() for l in texto.splitlines() if l.strip()]

    for i, linea in enumerate(lineas):
        # --- Nombre del paciente ---
        if linea.upper() == "PACIENTE" and i + 1 < len(lineas):
            data["nombre_completo"] = lineas[i + 1].strip()

        # --- ID del paciente (robusto, evita confundir con años) ---
        if data.get("id") is None and "id" in linea.lower():
            posibles = []
            ventanas = [linea]
            if i + 1 < len(lineas):
                ventanas.append(lineas[i + 1])
            if i + 2 < len(lineas):
                ventanas.append(lineas[i + 2])

            for v in ventanas:
                nums = re.findall(r"\b(\d{5,8})\b", v)
                for n in nums:
                    if int(n) > 3000:
                        posibles.append(int(n))

            if posibles:
                data["id"] = str(max(posibles))

        # --- Edad ---
        if linea.upper() == "EDAD" and i + 1 < len(lineas):
            m_edad = re.search(r"(\d+)", lineas[i + 1])
            if m_edad:
                data["edad"] = int(m_edad.group(1))

        # --- Peso corporal ---
        if "peso corporal" in linea.lower() and i + 1 < len(lineas):
            m_peso = re.search(r"(\d+(?:[.,]\d+)?)", lineas[i + 1])
            if m_peso:
                data["peso"] = float(m_peso.group(1).replace(",", "."))

        # --- Diagnóstico (desde Grupo) ---
        if "GRUPO" in linea.upper():
            if i + 1 < len(lineas):
                linea_diagnostico = lineas[i + 1].strip()
                if linea_diagnostico.lower() != "diagnostico":
                    data["diagnostico"] = linea_diagnostico

        # --- Estadificación: T ---
        if linea.strip().upper().startswith("T:"):
            m_t = re.search(r"T\s*:\s*([A-Za-z0-9]+)", linea)
            if m_t:
                data["estad_t"] = m_t.group(1)

        # --- Estadificación: N ---
        if linea.strip().upper().startswith("N:"):
            m_n = re.search(r"N\s*:\s*([A-Za-z0-9]+)", linea)
            if m_n:
                data["estad_n"] = m_n.group(1)

        # --- Estadificación: M ---
        if linea.strip().upper().startswith("M:"):
            m_m = re.search(r"M\s*:\s*([A-Za-z0-9]+)", linea, re.IGNORECASE)
            if m_m:
                data["estad_m"] = m_m.group(1)

        # --- Estadificación: Estadio ---
        if linea.upper().startswith("ESTADIO:"):
            m_e = re.search(r"Estadio:\s*(.+)", linea, re.IGNORECASE)
            if m_e:
                data["estad_estadio"] = m_e.group(1).strip()

        # --- Histología ---
        if linea.lower().startswith("histologia"):
            m_h = re.search(r"histologia\s*:\s*(.+)", linea, re.IGNORECASE)
            if m_h:
                data["histologia"] = m_h.group(1).strip()

        # --- Interrogatorio (último bloque) ---
        if "INTERROGATORIO" in linea.upper():
            j = i + 1
            bloque = []

            while j < len(lineas) and not lineas[j].strip():
                j += 1

            if j < len(lineas) and re.search(r"\d{2}/\d{2}/\d{4}", lineas[j]):
                j += 1

            while j < len(lineas):
                l2 = lineas[j].rstrip()

                if not l2:
                    bloque.append("")
                    j += 1
                    continue

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
                bloques_interrogatorio.append("\n".join(bloque))

    # Elegimos SIEMPRE el último interrogatorio encontrado
    if bloques_interrogatorio:
        data["interrogatorio"] = bloques_interrogatorio[-1]

    # Prescripción braquiterapia
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

    idx_dosispor = next(
        (i for i, l in enumerate(lineas) if "dosis por fracción" in l.lower()),
        None,
    )
    if idx_dosispor is None:
        return None

    idx_end = next(
        (
            i
            for i, l in enumerate(lineas[idx_dosispor:], start=idx_dosispor)
            if "conducta terapéutica" in l.lower()
        ),
        len(lineas),
    )

    relevantes = [l.strip() for l in lineas[idx_dosispor + 1 : idx_end] if l.strip()]

    grupos = []
    num_re = re.compile(r"^\d+(?:[.,]\d+)?$")

    k = 0
    while k < len(relevantes):
        if num_re.match(relevantes[k]):
            nums = []
            start = k
            while k < len(relevantes) and len(nums) < 5 and num_re.match(relevantes[k]):
                nums.append(float(relevantes[k].replace(",", ".")))
                k += 1

            if len(nums) == 5:
                prev = " ".join(relevantes[max(0, start - 6) : start])
                nxt = " ".join(relevantes[k : min(len(relevantes), k + 6)])
                grupos.append({"values": nums, "prev": prev, "next": nxt})
        else:
            k += 1

    if not grupos:
        return None

    def ctx(g):
        return (g["prev"] + " " + g["next"]).lower()

    marcadores_braqui = [
        "bqt", "braqui", "uterovaginal",
        "vaginal", "hr-ctv", "cervicovaginal",
    ]

    grupos_anestesia = [g for g in grupos if "anestesia" in ctx(g)]
    if grupos_anestesia:
        elegido = grupos_anestesia[-1]
    else:
        grupos_marcadores = [
            g for g in grupos if any(m in ctx(g) for m in marcadores_braqui)
        ]
        if grupos_marcadores:
            elegido = grupos_marcadores[-1]
        else:
            elegido = grupos[-1]

    v = elegido["values"]
    return {
        "dosis_por_fraccion": v[0],
        "n_fracciones": v[1],
        "fracciones_por_semana": v[2],
        "dosis_total": v[3],
        "dosis_total_con_externa": v[4],
    }
