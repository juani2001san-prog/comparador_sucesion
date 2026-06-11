"""
Conciliación movimiento a movimiento entre Contabilidad (sistema) y Planilla
(lo que envían). Como las cajas no tienen comprobante ni CUIT, el cruce se hace
por **importe + fecha**.

Pasada 1 - importe + fecha exactos  -> OK.
Pasada 2 - mismo importe, distinta fecha (dentro de la tolerancia => OK;
           si no => "Diferencia de fecha").
Resto    - faltantes en cada lado o posibles duplicados.
"""

from __future__ import annotations

import re
from collections import defaultdict, deque

import numpy as np
import pandas as pd

ESTADO_OK = "OK"
ESTADO_DIF_IMPORTE = "Diferencia de importe"
ESTADO_DIF_FECHA = "Diferencia de fecha"
ESTADO_FALTA_CONTAB = "Falta en contabilidad"   # está en planilla, no en el sistema
ESTADO_FALTA_PLANILLA = "Falta en planilla"     # está en el sistema, no en la planilla
ESTADO_DUPLICADO = "Posible duplicado"

ESTADOS = [
    ESTADO_OK, ESTADO_DIF_IMPORTE, ESTADO_DIF_FECHA, ESTADO_FALTA_CONTAB,
    ESTADO_FALTA_PLANILLA, ESTADO_DUPLICADO,
]

COLUMNAS_RESULTADO = [
    "estado", "cuenta",
    "fecha_contab", "fecha_planilla",
    "detalle_contab", "detalle_planilla",
    "importe_contab", "importe_planilla",
    "diferencia", "observacion",
]


def _clave_importe(importe: float, comparar_abs: bool) -> str:
    valor = abs(importe) if comparar_abs else importe
    return f"{valor:.2f}"


def _dias_entre(a, b) -> int | None:
    if pd.isna(a) or pd.isna(b):
        return None
    return abs((pd.Timestamp(a).normalize() - pd.Timestamp(b).normalize()).days)


def _signos_opuestos(a: float, b: float) -> bool:
    return a != 0 and b != 0 and (a > 0) != (b > 0)


def _fila_par(c: pd.Series, p: pd.Series, comparar_abs: bool, tolerancia: int) -> dict:
    imp_c, imp_p = float(c["importe"]), float(p["importe"])
    dif = round(imp_c - imp_p, 2)
    # Diferencia "efectiva" según se compare con o sin signo.
    val_c = abs(imp_c) if comparar_abs else imp_c
    val_p = abs(imp_p) if comparar_abs else imp_p
    dif_valor = round(val_c - val_p, 2)
    dias = _dias_entre(c["fecha"], p["fecha"])

    if abs(dif_valor) >= 0.01:
        estado = ESTADO_DIF_IMPORTE
        observacion = f"Importe distinto (dif {dif:+.2f})."
    elif dias is not None and dias > tolerancia:
        estado = ESTADO_DIF_FECHA
        observacion = f"Mismo importe, fecha difiere {dias} día(s)."
    else:
        estado = ESTADO_OK
        observacion = ""

    if comparar_abs and _signos_opuestos(imp_c, imp_p):
        observacion = (observacion + " Signo invertido.").strip()

    return {
        "estado": estado,
        "cuenta": c["cuenta"] or p["cuenta"],
        "fecha_contab": c["fecha"],
        "fecha_planilla": p["fecha"],
        "detalle_contab": c["detalle"],
        "detalle_planilla": p["detalle"],
        "importe_contab": round(float(c["importe"]), 2),
        "importe_planilla": round(float(p["importe"]), 2),
        "diferencia": dif,
        "observacion": observacion,
    }


def _fila_sola(fila: pd.Series, *, es_contab: bool, es_dup: bool) -> dict:
    importe = round(float(fila["importe"]), 2)
    if es_dup:
        estado = ESTADO_DUPLICADO
        observacion = "Aparece repetido (mismo importe y fecha)."
    elif es_contab:
        estado = ESTADO_FALTA_PLANILLA
        observacion = "Está en la contabilidad pero no en la planilla."
    else:
        estado = ESTADO_FALTA_CONTAB
        observacion = "Está en la planilla pero no en la contabilidad."

    return {
        "estado": estado,
        "cuenta": fila["cuenta"],
        "fecha_contab": fila["fecha"] if es_contab else pd.NaT,
        "fecha_planilla": pd.NaT if es_contab else fila["fecha"],
        "detalle_contab": fila["detalle"] if es_contab else "",
        "detalle_planilla": "" if es_contab else fila["detalle"],
        "importe_contab": importe if es_contab else np.nan,
        "importe_planilla": np.nan if es_contab else importe,
        "diferencia": importe if es_contab else -importe,
        "observacion": observacion,
    }


def conciliar(
    df_contab: pd.DataFrame,
    df_planilla: pd.DataFrame,
    comparar_abs: bool = True,
    tolerancia_dias: int = 0,
) -> pd.DataFrame:
    """Concilia los dos lados y devuelve la tabla de resultados."""
    c = df_contab.reset_index(drop=True).copy()
    p = df_planilla.reset_index(drop=True).copy()

    def claves(df):
        if not len(df):
            return pd.Series(dtype=str), pd.Series(dtype=str)
        base = df["importe"].apply(lambda v: _clave_importe(float(v), comparar_abs))
        completa = base.str.cat(df["fecha"].astype(str), sep="||")
        return completa, base

    c_completa, c_base = claves(c)
    p_completa, p_base = claves(p)

    dup_c = c_completa.duplicated(keep=False) if len(c) else pd.Series(dtype=bool)
    dup_p = p_completa.duplicated(keep=False) if len(p) else pd.Series(dtype=bool)

    # Índices de planilla por clave (colas que se van consumiendo).
    p_por_completa: dict[str, deque] = defaultdict(deque)
    p_por_base: dict[str, deque] = defaultdict(deque)
    for idx in p.index:
        p_por_completa[p_completa.iloc[idx]].append(idx)
        p_por_base[p_base.iloc[idx]].append(idx)

    c_match, p_match = set(), set()
    resultados: list[dict] = []

    def tomar(cola: deque):
        while cola:
            cand = cola.popleft()
            if cand not in p_match:
                return cand
        return None

    # Pasada 1: importe + fecha exactos.
    for idx in c.index:
        j = tomar(p_por_completa.get(c_completa.iloc[idx], deque()))
        if j is not None:
            c_match.add(idx)
            p_match.add(j)
            resultados.append(_fila_par(c.loc[idx], p.loc[j], comparar_abs, tolerancia_dias))

    # Pasada 2: mismo importe (cualquier fecha) -> diferencia de fecha.
    for idx in c.index:
        if idx in c_match:
            continue
        j = tomar(p_por_base.get(c_base.iloc[idx], deque()))
        if j is not None:
            c_match.add(idx)
            p_match.add(j)
            resultados.append(_fila_par(c.loc[idx], p.loc[j], comparar_abs, tolerancia_dias))

    # Pasada 3: mismo MES + detalle parecido, importe distinto -> diferencia de
    # importe (típico error de carga / typo). Se agrupa por mes (no por día
    # exacto) porque las fechas entre caja y contabilidad no suelen coincidir.
    # Para no inventar cruces entre conceptos distintos, solo empareja sobrantes
    # que COMPARTEN alguna palabra del detalle (ej. ABSA/TERRENO) y, entre esos,
    # toma el importe más cercano de forma mutua.
    def _mes(ts):
        return pd.Timestamp(ts).to_period("M")

    def _val(x):
        return abs(float(x)) if comparar_abs else float(x)

    def _tokens(detalle) -> set:
        # Palabras "significativas" (>= 4 caracteres) para comparar detalles.
        return {t for t in re.split(r"[^0-9A-Za-z]+", str(detalle).upper()) if len(t) >= 4}

    c_por_mes: dict = defaultdict(list)
    p_por_mes: dict = defaultdict(list)
    for idx in c.index:
        if idx not in c_match and pd.notna(c.loc[idx, "fecha"]):
            c_por_mes[_mes(c.loc[idx, "fecha"])].append(idx)
    for j in p.index:
        if j not in p_match and pd.notna(p.loc[j, "fecha"]):
            p_por_mes[_mes(p.loc[j, "fecha"])].append(j)

    for mes, idxs in c_por_mes.items():
        jdxs = p_por_mes.get(mes, [])
        if not jdxs:
            continue
        cvals = {i: _val(c.loc[i, "importe"]) for i in idxs}
        pvals = {j: _val(p.loc[j, "importe"]) for j in jdxs}
        ctok = {i: _tokens(c.loc[i, "detalle"]) for i in idxs}
        ptok = {j: _tokens(p.loc[j, "detalle"]) for j in jdxs}

        # Mejor contraparte (importe más cercano) ENTRE las que comparten palabra.
        def mejor(origen_vals, origen_tok_set, destino_ids, destino_vals, destino_tok):
            cands = [d for d in destino_ids if origen_tok_set & destino_tok[d]]
            if not cands:
                return None
            return min(cands, key=lambda d: abs(origen_vals - destino_vals[d]))

        mejor_p = {i: mejor(cvals[i], ctok[i], jdxs, pvals, ptok) for i in idxs}
        mejor_c = {j: mejor(pvals[j], ptok[j], idxs, cvals, ctok) for j in jdxs}
        for i in idxs:
            j = mejor_p[i]
            if j is not None and mejor_c.get(j) == i and i not in c_match and j not in p_match:
                c_match.add(i)
                p_match.add(j)
                resultados.append(_fila_par(c.loc[i], p.loc[j], comparar_abs, tolerancia_dias))

    # Sobrantes.
    for idx in c.index:
        if idx not in c_match:
            es_dup = bool(dup_c.iloc[idx]) if len(dup_c) else False
            resultados.append(_fila_sola(c.loc[idx], es_contab=True, es_dup=es_dup))
    for idx in p.index:
        if idx not in p_match:
            es_dup = bool(dup_p.iloc[idx]) if len(dup_p) else False
            resultados.append(_fila_sola(p.loc[idx], es_contab=False, es_dup=es_dup))

    if not resultados:
        return pd.DataFrame(columns=COLUMNAS_RESULTADO)
    return pd.DataFrame(resultados)[COLUMNAS_RESULTADO]


def resumen(df_resultado: pd.DataFrame, df_contab: pd.DataFrame, df_planilla: pd.DataFrame) -> dict:
    """Métricas para el dashboard."""
    estados = df_resultado["estado"] if len(df_resultado) else pd.Series(dtype=str)

    def n(estado):
        return int((estados == estado).sum())

    total_c = round(float(df_contab["importe"].sum()), 2) if len(df_contab) else 0.0
    total_p = round(float(df_planilla["importe"].sum()), 2) if len(df_planilla) else 0.0

    suma_dif_imp = 0.0
    if len(df_resultado):
        mask = df_resultado["estado"] == ESTADO_DIF_IMPORTE
        suma_dif_imp = round(float(df_resultado.loc[mask, "diferencia"].abs().sum()), 2)

    return {
        "Movimientos contabilidad": len(df_contab),
        "Movimientos planilla": len(df_planilla),
        "Total contabilidad": total_c,
        "Total planilla": total_p,
        "Diferencia de totales": round(total_c - total_p, 2),
        "OK": n(ESTADO_OK),
        "Diferencia de importe": n(ESTADO_DIF_IMPORTE),
        "Total diferencia de importe": suma_dif_imp,
        "Diferencia de fecha": n(ESTADO_DIF_FECHA),
        "Falta en contabilidad": n(ESTADO_FALTA_CONTAB),
        "Falta en planilla": n(ESTADO_FALTA_PLANILLA),
        "Posibles duplicados": n(ESTADO_DUPLICADO),
    }
