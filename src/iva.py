"""
Posición de IVA — cálculo a partir de "Mis Comprobantes" de AFIP
================================================================

Formula:

    IVA Débito Fiscal  (a) = IVA(emitidos facturas + ND) - IVA(emitidos NC)
    IVA Crédito Fiscal (b) = IVA(recibidos facturas + ND) - IVA(recibidos NC)
    POSICIÓN = a - b - saldo_SJ_anterior_técnico - saldo_SJ_anterior_LD

    > 0  → a pagar al fisco
    < 0  → saldo a favor (sigue al próximo período)

Los archivos de entrada son los Excel descargables del sitio de AFIP
("Mis Comprobantes" → Emitidos / Recibidos). El cálculo distingue facturas y
notas de débito (que SUMAN) de notas de crédito (que RESTAN) por el texto de
la columna "Tipo" (ej. ``"3 - Nota de Crédito A"``).
"""

from __future__ import annotations

import io
from typing import Optional

import pandas as pd


# Columnas estándar de AFIP. Algunos archivos varían el orden o tienen
# pequeñas diferencias de encabezado; las buscamos por nombre (insensible
# a mayúsculas y espacios extra).
_NOMBRES = {
    "fecha": ["Fecha"],
    "tipo": ["Tipo"],
    "punto_venta": ["Punto de Venta", "Punto Venta"],
    "numero": ["Número Desde", "Numero Desde", "Nro. Desde"],
    "denominacion": [
        "Denominación Receptor", "Denominacion Receptor",
        "Denominación Emisor", "Denominacion Emisor",
    ],
    "total_iva": ["Total IVA", "Total Iva"],
    "imp_total": ["Imp. Total", "Imp Total", "Importe Total"],
}

_ALICUOTAS_IVA = ["IVA 0%", "IVA 2,5%", "IVA 5%", "IVA 10,5%", "IVA 21%", "IVA 27%"]


def _buscar_columna(columnas: list[str], candidatos: list[str]) -> Optional[str]:
    """Devuelve el nombre real de columna que coincida con alguno de los candidatos."""
    norm = {c.lower().strip(): c for c in columnas}
    for cand in candidatos:
        if cand in columnas:
            return cand
        if cand.lower().strip() in norm:
            return norm[cand.lower().strip()]
    return None


def es_nota_credito(tipo: str) -> bool:
    """``True`` si el comprobante es una Nota de Crédito (cualquier letra/clase)."""
    if not isinstance(tipo, str):
        return False
    s = tipo.upper()
    return "NOTA DE CREDITO" in s or "NOTA DE CRÉDITO" in s or "NOTA CREDITO" in s


def cargar_archivo_afip(archivo, hoja: Optional[str] = None) -> pd.DataFrame:
    """
    Carga un Excel 'Mis Comprobantes' (AFIP) y devuelve un DataFrame normalizado.

    Columnas de salida:

    - ``fecha`` (Timestamp)
    - ``tipo`` (str — el texto original, ej. ``"1 - Factura A"``)
    - ``es_nc`` (bool)
    - ``punto_venta``, ``numero`` (int)
    - ``denominacion`` (str)
    - ``total_iva`` (float)
    - ``imp_total`` (float)
    - una columna por cada alícuota detectada (``IVA 21%``, etc.)
    """
    if isinstance(archivo, (bytes, bytearray)):
        archivo = io.BytesIO(archivo)

    xl = pd.ExcelFile(archivo)
    hoja = hoja or xl.sheet_names[0]
    raw = pd.read_excel(archivo, sheet_name=hoja, header=0)
    raw.columns = [str(c).strip() for c in raw.columns]

    cols = list(raw.columns)
    col_tipo = _buscar_columna(cols, _NOMBRES["tipo"])
    col_total = _buscar_columna(cols, _NOMBRES["total_iva"])
    if not col_tipo or not col_total:
        raise ValueError(
            "El Excel no parece ser de 'Mis Comprobantes' (AFIP): "
            f"no se encontró 'Tipo' o 'Total IVA'. Columnas: {cols}"
        )

    col_fecha = _buscar_columna(cols, _NOMBRES["fecha"])
    col_pv = _buscar_columna(cols, _NOMBRES["punto_venta"])
    col_num = _buscar_columna(cols, _NOMBRES["numero"])
    col_denom = _buscar_columna(cols, _NOMBRES["denominacion"])
    col_imp = _buscar_columna(cols, _NOMBRES["imp_total"])

    out = pd.DataFrame()
    out["fecha"] = (
        pd.to_datetime(raw[col_fecha], errors="coerce", dayfirst=True)
        if col_fecha else pd.NaT
    )
    out["tipo"] = raw[col_tipo].astype(str).str.strip()
    out["es_nc"] = out["tipo"].apply(es_nota_credito)
    out["punto_venta"] = pd.to_numeric(raw[col_pv], errors="coerce").fillna(0).astype(int) if col_pv else 0
    out["numero"] = pd.to_numeric(raw[col_num], errors="coerce").fillna(0).astype(int) if col_num else 0
    out["denominacion"] = raw[col_denom].astype(str).str.strip() if col_denom else ""
    out["total_iva"] = pd.to_numeric(raw[col_total], errors="coerce").fillna(0.0)
    out["imp_total"] = pd.to_numeric(raw[col_imp], errors="coerce").fillna(0.0) if col_imp else 0.0

    # IVA por alícuota (si el archivo no la trae, queda en 0).
    for ali in _ALICUOTAS_IVA:
        col_ali = _buscar_columna(cols, [ali])
        out[ali] = pd.to_numeric(raw[col_ali], errors="coerce").fillna(0.0) if col_ali else 0.0

    # Filtro líneas sin tipo válido (filas de totales del Excel).
    out = out[out["tipo"].str.contains("-", na=False)].reset_index(drop=True)
    return out


def filtrar_por_periodo(df: pd.DataFrame, periodo: Optional[str]) -> pd.DataFrame:
    """Filtra el DataFrame al período 'YYYY-MM' indicado. Si es None, no filtra."""
    if not periodo or df.empty:
        return df
    per = df["fecha"].dt.to_period("M").astype(str)
    return df[per == periodo].reset_index(drop=True)


def periodos_presentes(df: pd.DataFrame) -> list[str]:
    """Lista ordenada de períodos 'YYYY-MM' presentes en el df."""
    if df.empty or "fecha" not in df.columns:
        return []
    per = df["fecha"].dropna().dt.to_period("M").astype(str)
    return sorted(set(per))


def resumen_lado(df: pd.DataFrame) -> dict:
    """Suma de IVA en facturas+ND vs notas de crédito, y neto del lado."""
    fact_y_nd = df[~df["es_nc"]]
    notas_credito = df[df["es_nc"]]
    fc = round(float(fact_y_nd["total_iva"].sum()), 2)
    nc = round(float(notas_credito["total_iva"].sum()), 2)
    return {
        "fc": fc,
        "nc": nc,
        "neto": round(fc - nc, 2),
        "cant_fact_nd": int(len(fact_y_nd)),
        "cant_nc": int(len(notas_credito)),
        "imp_total_fact_nd": round(float(fact_y_nd["imp_total"].sum()), 2),
        "imp_total_nc": round(float(notas_credito["imp_total"].sum()), 2),
    }


def calcular_posicion(
    df_emitidos: pd.DataFrame,
    df_recibidos: pd.DataFrame,
    saldo_anterior_tecnico: float = 0.0,
    saldo_anterior_ld: float = 0.0,
) -> dict:
    """Cálculo completo de la posición IVA del período."""
    df_lado = resumen_lado(df_emitidos)
    cf_lado = resumen_lado(df_recibidos)
    saldo_tec = round(float(saldo_anterior_tecnico), 2)
    saldo_ld = round(float(saldo_anterior_ld), 2)
    pos = round(df_lado["neto"] - cf_lado["neto"] - saldo_tec - saldo_ld, 2)
    return {
        "debito_fiscal": df_lado,
        "credito_fiscal": cf_lado,
        "saldo_anterior_tecnico": saldo_tec,
        "saldo_anterior_ld": saldo_ld,
        "posicion": pos,
        "a_pagar": pos if pos > 0 else 0.0,
        "saldo_a_favor": abs(pos) if pos < 0 else 0.0,
    }


def _alicuotas_no_cero(df: pd.DataFrame) -> list[str]:
    return [a for a in _ALICUOTAS_IVA if a in df.columns and df[a].sum() != 0]


def exportar_excel(
    df_emitidos: pd.DataFrame,
    df_recibidos: pd.DataFrame,
    posicion: dict,
    periodo: Optional[str] = None,
) -> bytes:
    """Genera un Excel descargable con la posición y el detalle de comprobantes."""
    df_l = posicion["debito_fiscal"]
    cf_l = posicion["credito_fiscal"]
    titulo_periodo = f" — {periodo}" if periodo else ""

    pos_filas = [
        ("DÉBITO FISCAL" + titulo_periodo, None, None),
        ("Facturas + Notas de Débito emitidas",       df_l["fc"],   df_l["cant_fact_nd"]),
        ("(-) Notas de Crédito emitidas",            -df_l["nc"],   df_l["cant_nc"]),
        ("    IVA DF neto (a)",                       df_l["neto"], df_l["cant_fact_nd"] + df_l["cant_nc"]),
        ("", None, None),
        ("CRÉDITO FISCAL" + titulo_periodo, None, None),
        ("Facturas + Notas de Débito recibidas",      cf_l["fc"],   cf_l["cant_fact_nd"]),
        ("(-) Notas de Crédito recibidas",           -cf_l["nc"],   cf_l["cant_nc"]),
        ("    IVA CF neto (b)",                       cf_l["neto"], cf_l["cant_fact_nd"] + cf_l["cant_nc"]),
        ("", None, None),
        ("(-) Saldo SJ anterior técnico (c)",        -posicion["saldo_anterior_tecnico"], None),
        ("(-) Saldo SJ anterior LD / retenciones (d)", -posicion["saldo_anterior_ld"], None),
        ("", None, None),
        ("POSICIÓN = a - b - c - d",                  posicion["posicion"], None),
    ]
    pos_df = pd.DataFrame(pos_filas, columns=["Concepto", "Importe", "Cantidad"])

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        pos_df.to_excel(writer, index=False, sheet_name="Posición")
        df_emitidos.to_excel(writer, index=False, sheet_name="Emitidos")
        df_recibidos.to_excel(writer, index=False, sheet_name="Recibidos")

        wb = writer.book
        money = wb.add_format({"num_format": "#,##0.00"})
        bold = wb.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white"})
        bold_pos = wb.add_format({"bold": True, "num_format": "#,##0.00", "bg_color": "#FFF2CC"})

        ws = writer.sheets["Posición"]
        ws.set_column("A:A", 50)
        ws.set_column("B:B", 18, money)
        ws.set_column("C:C", 12)
        ws.write(0, 0, "Concepto", bold)
        ws.write(0, 1, "Importe", bold)
        ws.write(0, 2, "Cantidad", bold)
        # Última fila (POSICIÓN) en formato destacado.
        ws.write(len(pos_df), 0, pos_df.iloc[-1, 0], wb.add_format({"bold": True, "bg_color": "#FFF2CC"}))
        ws.write(len(pos_df), 1, pos_df.iloc[-1, 1], bold_pos)

        for hoja in ("Emitidos", "Recibidos"):
            ws = writer.sheets[hoja]
            ws.set_column("A:Z", 14, money)

    buf.seek(0)
    return buf.getvalue()
