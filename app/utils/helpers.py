"""
Funciones auxiliares de uso general.
"""


def separar_apellido_nombre(nombre_completo):
    """
    Dado algo tipo 'LUDUEÑA CLAUDIA ESTEFANIA'
    devuelve ('LUDUEÑA', 'CLAUDIA ESTEFANIA').
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


def formatear_dosis(valor):
    """Devuelve 'X.Y Gy' con 1 decimal."""
    if valor is None:
        return None
    try:
        v = float(valor)
        return f"{v:.1f} Gy"
    except (ValueError, TypeError):
        return f"{valor} Gy"
