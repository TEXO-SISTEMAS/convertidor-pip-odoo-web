"""
Funciones auxiliares: manejo de fechas, redondeo, formato de strings.
"""

import json
from pathlib import Path
from datetime import datetime, timedelta


def _cargar_mappings_custom() -> list[dict]:
    """
    Carga los mapeos personalizados desde mappings.json.
    En Vercel el archivo se copia a /tmp/mappings.json al iniciar app.py,
    por eso buscamos primero en /tmp (Vercel) y después en la ruta del repo.
    """
    import os
    candidatas = []
    # Vercel: /tmp/mappings.json (donde app.py guarda los cambios)
    if os.environ.get("VERCEL"):
        candidatas.append(Path("/tmp/mappings.json"))
    # Ruta normal: junto al raíz del repo
    candidatas.append(Path(__file__).parent.parent / "mappings.json")

    for ruta in candidatas:
        if ruta.exists():
            try:
                with open(ruta, encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                continue
    return []


def _aplicar_mapping_custom(codigo_pip: str, nombre_pip: str) -> tuple[str, str] | None:
    """
    Busca si el producto coincide con algún mapeo personalizado.
    Retorna (codigo_odoo, nombre_odoo) si hay match, o None si no.
    """
    nombre_upper = (nombre_pip or "").upper().strip()
    codigo_upper = (codigo_pip or "").upper().strip()

    for m in _cargar_mappings_custom():
        tipo = m.get("tipo_match", "")
        valor = (m.get("valor_match") or "").upper().strip()
        cod_odoo = m.get("codigo_odoo", "")
        nom_odoo = m.get("nombre_odoo", "")

        if tipo == "codigo" and valor and codigo_upper == valor:
            return (cod_odoo, nom_odoo)
        elif tipo == "codigo_prefijo" and valor and codigo_upper.startswith(valor):
            return (cod_odoo, nom_odoo)
        elif tipo == "nombre" and valor and valor in nombre_upper:
            return (cod_odoo, nom_odoo)


def convertir_fecha_pip_a_odoo(fecha_str: str) -> str | None:
    """
    Convierte una fecha del formato PIP al formato requerido por Odoo.

    Args:
        fecha_str: Fecha en formato "DD/MM/YYYY", o "-" si no aplica.

    Returns:
        Fecha en formato "YYYY-MM-DD" (SIN HORA), o None si fecha_str es "-".

    Raises:
        ValueError: Si el formato de fecha no es válido.
    """
    if not fecha_str or fecha_str.strip() == "-":
        return None

    try:
        fecha = datetime.strptime(fecha_str.strip(), "%d/%m/%Y")
    except ValueError:
        raise ValueError(f"Formato de fecha inválido: '{fecha_str}'. Se esperaba DD/MM/YYYY.")

    # Odoo requiere YYYY-MM-DD sin hora
    return fecha.strftime("%Y-%m-%d")


def calcular_fecha_vencimiento(fecha_emision_str: str) -> str:
    """
    Calcula la fecha de vencimiento como 30 días después de la fecha de emisión.
    Se usa cuando el Excel PIP trae "-" en la columna Fecha de Vencimiento.

    Args:
        fecha_emision_str: Fecha de emisión en formato "DD/MM/YYYY".

    Returns:
        Fecha de vencimiento en formato "YYYY-MM-DD".

    Raises:
        ValueError: Si el formato de fecha de emisión no es válido.
    """
    if not fecha_emision_str or fecha_emision_str.strip() == "-":
        raise ValueError("La fecha de emisión no puede estar vacía o ser '-'.")

    try:
        fecha_emision = datetime.strptime(fecha_emision_str.strip(), "%d/%m/%Y")
    except ValueError:
        raise ValueError(f"Formato de fecha de emisión inválido: '{fecha_emision_str}'. Se esperaba DD/MM/YYYY.")

    # Regla de negocio: vencimiento = emisión + 30 días calendario
    fecha_vencimiento = fecha_emision + timedelta(days=30)

    return fecha_vencimiento.strftime("%Y-%m-%d")


def extraer_partes_numero_factura(numero_doc: str) -> tuple[str, str, str]:
    """
    Extrae Suc, Sec y Nro del número de documento.
    
    Args:
        numero_doc: Número de documento en formato "001-003-0002871"
    
    Returns:
        Tupla (suc, sec, nro)
        Ejemplo: ("001", "003", "0002871")
    """
    if not numero_doc or not isinstance(numero_doc, str):
        return ("", "", "")
    
    partes = numero_doc.strip().split("-")
    
    if len(partes) != 3:
        return ("", "", "")
    
    suc = partes[0].strip()
    sec = partes[1].strip()
    nro = partes[2].strip()
    
    return (suc, sec, nro)


def calcular_terminos_pago(fecha_emision_str: str, fecha_vencimiento_str: str) -> str:
    """
    Calcula los términos de pago según valores permitidos por Odoo.
    
    Args:
        fecha_emision_str: Fecha de emisión en formato DD/MM/YYYY
        fecha_vencimiento_str: Fecha de vencimiento en formato DD/MM/YYYY o "-"
    
    Returns:
        String con el término de pago exacto que acepta Odoo
        
    Valores posibles:
        - "Pago inmediato" (cuando vencimiento es "-")
        - "5 Días", "8 Dias", "15 días", "21 días", "30 días", 
          "45 días", "60 días", "75 días", "90 días", "135 días"
    """
    # Si vencimiento es "-" o vacío → Pago inmediato (contado)
    if not fecha_vencimiento_str or fecha_vencimiento_str.strip() == "-":
        return "Pago inmediato"
    
    try:
        fecha_emision = datetime.strptime(fecha_emision_str.strip(), "%d/%m/%Y")
        fecha_venc = datetime.strptime(fecha_vencimiento_str.strip(), "%d/%m/%Y")
        
        diferencia = (fecha_venc - fecha_emision).days
        
        # Mapear a los valores exactos permitidos por Odoo
        # Rangos ajustados para redondear al valor más cercano
        if diferencia == 0:
            return "Pago inmediato"
        elif diferencia <= 6:
            return "5 Días"
        elif diferencia <= 11:
            return "8 Dias"
        elif diferencia <= 18:
            return "15 días"
        elif diferencia <= 25:
            return "21 días"
        elif diferencia <= 37:
            return "30 días"
        elif diferencia <= 52:
            return "45 días"
        elif diferencia <= 67:
            return "60 días"
        elif diferencia <= 82:
            return "75 días"
        elif diferencia <= 112:
            return "90 días"
        else:
            return "135 días"
            
    except ValueError:
        # Si hay error al parsear, asumir pago inmediato
        return "Pago inmediato"


def mapear_producto_odoo(codigo_pip: str, nombre_pip: str) -> tuple[str, str]:
    """
    Mapea el código y nombre del producto PIP a código y etiqueta de Odoo.
    
    Args:
        codigo_pip: Código del producto en PIP (ej: "F1-L180", "AT-L03", "1-INT")
        nombre_pip: Nombre del producto en PIP
    
    Returns:
        Tupla (codigo_odoo, etiqueta_odoo)
    """
    # Verificar primero los mapeos personalizados
    custom = _aplicar_mapping_custom(codigo_pip, nombre_pip)
    if custom:
        return custom

    # Normalizar para búsqueda (mayúsculas, sin espacios extras)
    nombre_upper = (nombre_pip or "").upper().strip()
    codigo_upper = (codigo_pip or "").upper().strip()

    # GRUPO 1: Floresta X + (CUOTA, ENTREGA INICIAL, PAGO 50%, VALOR DEL TERRENO) o código F<N>- → Ventas Floresta
    # El prefijo F<N>- NO aplica si el nombre contiene palabras de recupero/servicios
    _kw_floresta = ("CUOTA", "ENTREGA INICIAL", "PAGO 50%", "VALOR DEL TERRENO")
    _kw_rec01 = ("LIMPIEZA", "RECUPERO", "CONSUMO", "RECOLECCION", "RECOLECCIÓN", "BASURA", "GASTOS COMUNES")
    _es_rec01 = any(k in nombre_upper for k in _kw_rec01)

    if (("FLORESTA 1" in nombre_upper and any(k in nombre_upper for k in _kw_floresta)) or (codigo_upper.startswith("F1-") and not _es_rec01)):
        return ("FLO1", "Venta Floresta 1")
    elif (("FLORESTA 2" in nombre_upper and any(k in nombre_upper for k in _kw_floresta)) or (codigo_upper.startswith("F2-") and not _es_rec01)):
        return ("flo2", "Venta Floresta 2")
    elif (("FLORESTA 3" in nombre_upper and any(k in nombre_upper for k in _kw_floresta)) or (codigo_upper.startswith("F3-") and not _es_rec01)):
        return ("flo3", "Venta Floresta 3")
    elif (("FLORESTA 4" in nombre_upper and any(k in nombre_upper for k in _kw_floresta)) or (codigo_upper.startswith("F4-") and not _es_rec01)):
        return ("flo4", "Venta Floresta 4")
    elif (("FLORESTA 5" in nombre_upper and any(k in nombre_upper for k in _kw_floresta)) or (codigo_upper.startswith("F5-") and not _es_rec01)):
        return ("flo5", "Venta Floresta 5")
    elif (("FLORESTA 6" in nombre_upper and any(k in nombre_upper for k in _kw_floresta)) or (codigo_upper.startswith("F6-") and not _es_rec01)):
        return ("flo6", "Venta Floresta 6")
    
    # GRUPO 3: Productos específicos (antes de rec01 para evitar conflictos)
    # Altos Terport
    if "ALTOS TERPORT" in nombre_upper or codigo_upper == "AT-L03":
        return ("venter", "Ventas Terport")
    
    # Intereses
    if "INTERESES" in nombre_upper or "INTERÉS" in nombre_upper or "INTERES" in nombre_upper or "1-INT" in codigo_upper:
        return ("inte01", "Intereses")
    
    # GRUPO 2: Recupero de Servicios Básicos
    palabras_rec01 = [
        "LIMPIEZA",
        "RECUPERO",
        "CONSUMO",
        "RECOLECCION",  # Sin tilde
        "RECOLECCIÓN",   # Con tilde
        "BASURA",
        "GASTOS COMUNES",
    ]
    
    # También código "14" es rec01
    if codigo_upper == "14":
        return ("rec01", nombre_pip or "")
    
    # Verificar si contiene alguna palabra clave
    for palabra in palabras_rec01:
        if palabra in nombre_upper:
            return ("rec01", nombre_pip or "")
    
    # Si no matchea nada, retornar el código original del PIP y el nombre como etiqueta
    return (codigo_pip or "", nombre_pip or "")
