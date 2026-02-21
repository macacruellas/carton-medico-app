"""
Funciones auxiliares de uso general.
"""
from app.utils.nombres_comunes import NOMBRES_COMUNES


def separar_apellido_nombre(nombre_completo):
    """
    Dado algo tipo 'DE LA CRUZ JUAN CARLOS'
    devuelve ('DE LA CRUZ', 'JUAN CARLOS').

    Estrategia: recorre las palabras de derecha a izquierda.
    Mientras la palabra sea un nombre de pila conocido, la acumula como "nombre".
    El resto queda como "apellido".

    Ejemplos:
      'LUDUEÑA CLAUDIA ESTEFANIA'   → ('LUDUEÑA', 'CLAUDIA ESTEFANIA')
      'DE LA CRUZ JUAN CARLOS'      → ('DE LA CRUZ', 'JUAN CARLOS')
      'GARCIA LOPEZ MARIA'          → ('GARCIA LOPEZ', 'MARIA')
      'PEREZ JUAN'                   → ('PEREZ', 'JUAN')
    """
    if not nombre_completo:
        return None, None

    partes = nombre_completo.strip().upper().split()

    if len(partes) == 1:
        return partes[0], ""

    # Recorremos de derecha a izquierda buscando nombres de pila conocidos
    idx_primer_nombre = len(partes)
    for i in range(len(partes) - 1, 0, -1):
        if partes[i] in NOMBRES_COMUNES:
            idx_primer_nombre = i
        else:
            break

    # Si no encontró ningún nombre conocido, fallback: primera palabra = apellido
    if idx_primer_nombre == len(partes):
        return partes[0], " ".join(partes[1:])

    # Si el apellido termina en preposición/artículo (DE, DEL, LA, LOS, etc.)
    # la palabra siguiente es parte del apellido, no del nombre.
    # Ej: "DE LOS SANTOS FLORENCIA" → "DE LOS" + "SANTOS" es apellido
    articulos = {"DE", "DEL", "LA", "LAS", "LOS", "EL", "DI", "DA", "DO"}
    while (
        idx_primer_nombre < len(partes)
        and partes[idx_primer_nombre - 1] in articulos
    ):
        idx_primer_nombre += 1

    apellido = " ".join(partes[:idx_primer_nombre])
    nombre = " ".join(partes[idx_primer_nombre:])
    return apellido, nombre


def formatear_dosis(valor):
    """Devuelve 'X.Y Gy' con 1 decimal."""
    if valor is None:
        return None
    try:
        v = float(valor)
        return f"{v:.1f} Gy"
    except (ValueError, TypeError):
        return f"{valor} Gy"
