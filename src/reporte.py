"""Resumen y exportación a Excel (XlsxWriter)."""

from __future__ import annotations

import io
import re

import pandas as pd

from . import normalizar as N
from .conciliar import (
    ESTADO_DIF_FECHA,
    ESTADO_DIF_IMPORTE,
    ESTADO_DUPLICADO,
    ESTADO_FALTA_CONTAB,
    ESTADO_FALTA_PLANILLA,
)


def resumen_df(resumen: dict) -> pd.DataFrame:
    filas = []
    for k, v in resumen.items():
        if k in ("Total contabilidad", "Total planilla", "Diferencia de totales",
                 "Total diferencia de importe"):
            v = N.formato_ar(v)
        filas.append({"Métrica": k, "Valor": v})
    return pd.DataFrame(filas)


def _hoja_segura(nombre: str) -> str:
    limpio = re.sub(r'[:\\/?*\[\]]', " ", str(nombre)).strip() or "Hoja"
    return limpio[:31]


def _subset(df: pd.DataFrame, estados: list[str]) -> pd.DataFrame:
    if not len(df):
        return df
    return df[df["estado"].isin(estados)].reset_index(drop=True)


def _escribir(writer, nombre, df, formatos):
    hoja = _hoja_segura(nombre)
    if df is None or not len(df):
        pd.DataFrame({"Info": ["Sin registros."]}).to_excel(writer, sheet_name=hoja, index=False)
        return
    df.to_excel(writer, sheet_name=hoja, index=False)
    ws = writer.sheets[hoja]
    for j, col in enumerate(df.columns):
        ws.write(0, j, str(col), formatos["enc"])
    for j, col in enumerate(df.columns):
        nombre_col = str(col).lower()
        max_largo = df[col].astype(str).str.len().max()
        max_largo = 10 if pd.isna(max_largo) else int(max_largo)
        ancho = max(12, min(45, max_largo + 2))
        if "fecha" in nombre_col:
            ws.set_column(j, j, max(ancho, 12), formatos["fecha"])
        elif any(k in nombre_col for k in ("importe", "total", "diferencia", "valor", "saldo")):
            ws.set_column(j, j, max(ancho, 14), formatos["num"])
        else:
            ws.set_column(j, j, ancho)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(df), len(df.columns) - 1)


def exportar_excel(df_resultado, df_contab, df_planilla, resumen_tabla) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        libro = writer.book
        formatos = {
            "enc": libro.add_format({
                "bold": True, "bg_color": "#1F4E78", "font_color": "white",
                "border": 1, "align": "center", "valign": "vcenter",
            }),
            "fecha": libro.add_format({"num_format": "dd/mm/yyyy"}),
            "num": libro.add_format({"num_format": "#,##0.00"}),
        }
        _escribir(writer, "Resumen", resumen_tabla, formatos)
        _escribir(writer, "Resultado", df_resultado, formatos)
        _escribir(writer, "Diferencias de importe", _subset(df_resultado, [ESTADO_DIF_IMPORTE]), formatos)
        _escribir(writer, "Falta en contabilidad", _subset(df_resultado, [ESTADO_FALTA_CONTAB]), formatos)
        _escribir(writer, "Falta en planilla", _subset(df_resultado, [ESTADO_FALTA_PLANILLA]), formatos)
        _escribir(writer, "Diferencias de fecha", _subset(df_resultado, [ESTADO_DIF_FECHA]), formatos)
        _escribir(writer, "Posibles duplicados", _subset(df_resultado, [ESTADO_DUPLICADO]), formatos)
        _escribir(writer, "Mov contabilidad", df_contab, formatos)
        _escribir(writer, "Mov planilla", df_planilla, formatos)
    buffer.seek(0)
    return buffer.getvalue()
