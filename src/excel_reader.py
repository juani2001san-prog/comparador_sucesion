"""Lectura de archivos Excel (listar hojas / leer una hoja a DataFrame)."""

from __future__ import annotations

import io

import pandas as pd


def _deduplicar(columnas) -> list[str]:
    """Renombra columnas repetidas agregando sufijos .1, .2, …

    Necesario porque al recortar espacios de los encabezados pueden quedar
    nombres idénticos (p. ej. 'DETALLE' y 'DETALLE '), y pandas no permite
    seleccionar por una etiqueta duplicada.
    """
    vistos: dict[str, int] = {}
    salida = []
    for col in columnas:
        if col in vistos:
            vistos[col] += 1
            salida.append(f"{col}.{vistos[col]}")
        else:
            vistos[col] = 0
            salida.append(col)
    return salida


def _asegurar_buffer(archivo) -> io.BytesIO:
    if isinstance(archivo, (bytes, bytearray)):
        return io.BytesIO(archivo)
    try:
        archivo.seek(0)
    except (AttributeError, ValueError):
        pass
    return archivo


def listar_hojas(archivo) -> list[str]:
    buffer = _asegurar_buffer(archivo)
    with pd.ExcelFile(buffer, engine="openpyxl") as xls:
        return list(xls.sheet_names)


def leer_hoja(archivo, hoja: str, fila_encabezado: int = 0) -> pd.DataFrame:
    """Lee una hoja a DataFrame.

    ``fila_encabezado`` es el índice (base 0) de la fila de encabezados.
    Se conservan los datos como objeto; la normalización la hace ``normalizar``.
    """
    buffer = _asegurar_buffer(archivo)
    df = pd.read_excel(
        buffer, sheet_name=hoja, header=fila_encabezado,
        engine="openpyxl", dtype=object,
    )
    df.columns = _deduplicar([str(c).strip() for c in df.columns])
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def vista_previa(df: pd.DataFrame, filas: int = 8) -> pd.DataFrame:
    return df.head(filas)
