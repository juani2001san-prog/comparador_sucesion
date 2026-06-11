"""
Funciones de normalización (puras y tolerantes a nulos).

Pensadas para aplicarse con ``Series.apply`` sobre columnas enteras sin romper
ante celdas vacías o con texto basura.
"""

from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime

import pandas as pd


# --------------------------------------------------------------------------- #
# Texto
# --------------------------------------------------------------------------- #

def quitar_acentos(texto: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(c)
    )


def quitar_espacios_dobles(texto: str) -> str:
    return re.sub(r"\s+", " ", str(texto)).strip()


def limpiar_texto(valor) -> str:
    """Mayúsculas, sin acentos, sin espacios dobles. Vacío si es nulo."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    return quitar_espacios_dobles(quitar_acentos(str(valor)).upper())


# --------------------------------------------------------------------------- #
# Fechas
# --------------------------------------------------------------------------- #

_EXCEL_EPOCH = pd.Timestamp("1899-12-30")


def parsear_fecha(valor):
    """Convierte cualquier representación a ``pd.Timestamp`` (sin hora) o ``NaT``."""
    if valor is None:
        return pd.NaT
    if isinstance(valor, float) and pd.isna(valor):
        return pd.NaT
    if isinstance(valor, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(valor).normalize()
    if isinstance(valor, (int, float)):
        try:
            return (_EXCEL_EPOCH + pd.to_timedelta(int(valor), unit="D")).normalize()
        except (ValueError, OverflowError):
            return pd.NaT
    texto = str(valor).strip()
    if not texto:
        return pd.NaT
    # Formato ISO (aaaa-mm-dd) -> NO usar dayfirst (evita swap día/mes y warning).
    iso = bool(re.match(r"^\d{4}-\d{1,2}-\d{1,2}", texto))
    fecha = pd.to_datetime(texto, dayfirst=not iso, errors="coerce")
    return pd.NaT if pd.isna(fecha) else pd.Timestamp(fecha).normalize()


# --------------------------------------------------------------------------- #
# Importes (formato argentino)
# --------------------------------------------------------------------------- #

def parsear_importe(valor) -> float:
    """Convierte un importe en formato argentino (``1.234.567,89``) a float.

    Soporta valores numéricos, negativos, paréntesis ``(1.234,56)`` y ``$``.
    Devuelve ``0.0`` ante vacíos o no parseables.
    """
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        if pd.isna(valor):
            return 0.0
        return round(float(valor), 2)

    texto = str(valor).strip()
    if texto == "" or texto.lower() in ("nan", "none", "-", "s/d"):
        return 0.0

    negativo = False
    if texto.startswith("(") and texto.endswith(")"):
        negativo = True
        texto = texto[1:-1]

    texto = texto.replace(" ", "").replace(" ", "").replace("$", "")
    texto = re.sub(r"[^0-9,.\-]", "", texto)
    if texto in ("", "-", ".", ","):
        return 0.0

    tiene_coma = "," in texto
    tiene_punto = "." in texto
    if tiene_coma and tiene_punto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")   # AR: 1.234,56
        else:
            texto = texto.replace(",", "")                     # US: 1,234.56
    elif tiene_coma:
        texto = texto.replace(",", ".")
    elif tiene_punto:
        partes = texto.split(".")
        if len(partes) > 2 or len(partes[1]) == 3:
            texto = texto.replace(".", "")                     # miles

    try:
        numero = float(texto)
    except ValueError:
        return 0.0
    if negativo:
        numero = -numero
    return round(numero, 2)


def redondear(valor, decimales: int = 2) -> float:
    try:
        if valor is None or (isinstance(valor, float) and pd.isna(valor)):
            return 0.0
        return round(float(valor), decimales)
    except (TypeError, ValueError):
        return 0.0


def formato_ar(valor) -> str:
    """Formatea al estilo argentino 1.234.567,89 (solo para mostrar)."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    try:
        texto = f"{float(valor):,.2f}"
    except (TypeError, ValueError):
        return str(valor)
    return texto.replace(",", "@").replace(".", ",").replace("@", ".")
