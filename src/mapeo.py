"""
Mapeo de columnas y armado de la lista de movimientos normalizada.

Ambos lados (contabilidad y planilla) terminan en el mismo esquema:

    fecha (Timestamp) | detalle (str) | importe (float, con signo) | cuenta (str)

Convención de signo: ingreso = positivo, egreso = negativo.

La planilla de caja puede venir de dos formas:
  - "simple": una fila por movimiento.
  - "doble":  bloque de ingresos y bloque de egresos en columnas separadas
              (típico libro de caja a dos columnas).
"""

from __future__ import annotations

import re

import pandas as pd

from . import normalizar as N

# Esquema de salida.
COLUMNAS = ["fecha", "detalle", "importe", "cuenta"]

# Pistas para autodetectar columnas por nombre.
_PISTAS = {
    "fecha": ["fecha", "date"],
    "detalle": ["detalle", "concepto", "descripcion", "glosa", "observ", "leyenda", "detall"],
    "importe": ["importe", "monto", "total", "valor", "efectivo", "neto"],
    # OJO: no incluir "haber"/"debe" acá para que no pisen las columnas Debe/Haber.
    "ingreso": ["ingreso", "entrada", "credito", "ingresos"],
    "egreso": ["egreso", "salida", "pago", "egresos"],
    "debe": ["debe"],
    "haber": ["haber"],
    "cuenta": ["cuenta", "centro", "rubro", "caja"],
}


def _columnas_que_matchean(campo: str, columnas: list[str]) -> list[str]:
    """Todas las columnas que matchean un campo, en orden de aparición."""
    pistas = _PISTAS.get(campo, [])
    res = []
    for col in columnas:
        normal = N.quitar_acentos(str(col)).lower()
        if any(p in normal for p in pistas):
            if campo == "fecha" and "venc" in normal:
                continue
            res.append(col)
    return res


def adivinar(campo: str, columnas: list[str]) -> str | None:
    matches = _columnas_que_matchean(campo, columnas)
    return matches[0] if matches else None


def adivinar_doble(columnas: list[str]) -> dict[str, str | None]:
    """Autodetecta el mapeo de un libro de caja a dos columnas.

    Toma la 1ª aparición de fecha/detalle/importe para INGRESOS y la 2ª para
    EGRESOS (ej.: FECHA/EFECTIVO -> ingresos ; FECHA.1/EFECTIVO.1 -> egresos).
    """
    fechas = _columnas_que_matchean("fecha", columnas)
    detalles = _columnas_que_matchean("detalle", columnas)
    importes = _columnas_que_matchean("importe", columnas)

    def nth(lista, n):
        return lista[n] if len(lista) > n else None

    return {
        "ing_fecha": nth(fechas, 0), "ing_detalle": nth(detalles, 0), "ing_importe": nth(importes, 0),
        "egr_fecha": nth(fechas, 1), "egr_detalle": nth(detalles, 1), "egr_importe": nth(importes, 1),
    }


# --------------------------------------------------------------------------- #
# Importe (lado "simple")
# --------------------------------------------------------------------------- #

def _serie(df: pd.DataFrame, col: str | None):
    if col and col in df.columns:
        return df[col].apply(N.parsear_importe)
    return None


def _importe_simple(df: pd.DataFrame, mapeo: dict) -> pd.Series:
    """importe directo, o debe-haber, o ingreso-egreso (en ese orden)."""
    ceros = pd.Series([0.0] * len(df), index=df.index)
    imp = _serie(df, mapeo.get("importe"))
    if imp is not None:
        return imp
    debe, haber = _serie(df, mapeo.get("debe")), _serie(df, mapeo.get("haber"))
    if debe is not None or haber is not None:
        return (debe if debe is not None else ceros) - (haber if haber is not None else ceros)
    ing, egr = _serie(df, mapeo.get("ingreso")), _serie(df, mapeo.get("egreso"))
    if ing is not None or egr is not None:
        return (ing if ing is not None else ceros) - (egr if egr is not None else ceros)
    return ceros


def _col_texto(df: pd.DataFrame, col: str | None) -> pd.Series:
    if col and col in df.columns:
        return df[col]
    return pd.Series([None] * len(df), index=df.index)


# --------------------------------------------------------------------------- #
# Construcción
# --------------------------------------------------------------------------- #

def construir_simple(df: pd.DataFrame, mapeo: dict, config: dict) -> pd.DataFrame:
    """Arma los movimientos para una planilla/tabla de una fila por movimiento."""
    out = pd.DataFrame(index=df.index)
    out["fecha"] = _col_texto(df, mapeo.get("fecha")).apply(N.parsear_fecha)
    out["detalle"] = _col_texto(df, mapeo.get("detalle")).apply(N.limpiar_texto)
    importe = _importe_simple(df, mapeo)
    if config.get("invertir_signo"):
        importe = -importe
    out["importe"] = importe.apply(N.redondear)
    out["cuenta"] = _col_texto(df, mapeo.get("cuenta")).apply(N.limpiar_texto)
    # Descarto filas sin fecha y sin importe (padding).
    out = out[~(out["fecha"].isna() & (out["importe"] == 0.0))].reset_index(drop=True)
    return out[COLUMNAS]


def construir_doble(df: pd.DataFrame, mapeo: dict, config: dict) -> pd.DataFrame:
    """Arma los movimientos para un libro de caja a dos columnas.

    ``mapeo`` usa las claves: ing_fecha, ing_detalle, ing_importe,
    egr_fecha, egr_detalle, egr_importe.
    """
    bloques = []

    for prefijo, signo in (("ing", 1), ("egr", -1)):
        imp_col = mapeo.get(f"{prefijo}_importe")
        if not imp_col:
            continue
        bloque = pd.DataFrame(index=df.index)
        bloque["fecha"] = _col_texto(df, mapeo.get(f"{prefijo}_fecha")).apply(N.parsear_fecha)
        bloque["detalle"] = _col_texto(df, mapeo.get(f"{prefijo}_detalle")).apply(N.limpiar_texto)
        bloque["importe"] = (_serie(df, imp_col) * signo).apply(N.redondear)
        bloque["cuenta"] = ""
        # Me quedo solo con filas que tienen importe (las demás son relleno).
        bloque = bloque[bloque["importe"] != 0.0]
        bloques.append(bloque)

    if not bloques:
        return pd.DataFrame(columns=COLUMNAS)

    out = pd.concat(bloques, ignore_index=True)
    if config.get("invertir_signo"):
        out["importe"] = (-out["importe"]).apply(N.redondear)
    return out[COLUMNAS]


def _excluir_por_detalle(df: pd.DataFrame, terminos: list[str]) -> pd.DataFrame:
    """Descarta filas cuyo 'detalle' contenga alguno de los términos dados.

    La comparación es sin acentos y sin distinguir mayúsculas (el detalle ya
    viene normalizado a mayúsculas/sin acentos desde la construcción).
    """
    if not len(df) or not terminos:
        return df
    limpios = [N.limpiar_texto(t) for t in terminos if str(t).strip()]
    limpios = [t for t in limpios if t]
    if not limpios:
        return df
    mascara = pd.Series(False, index=df.index)
    for termino in limpios:
        mascara = mascara | df["detalle"].str.contains(re.escape(termino), na=False)
    return df[~mascara].reset_index(drop=True)


def cortar_en(df: pd.DataFrame, marcadores: list[str]) -> pd.DataFrame:
    """Trunca el df en la PRIMERA fila que contenga alguno de los marcadores
    (en cualquier columna). Sirve para cortar la planilla donde termina la caja
    (ej. 'SUMAS IGUALES') y descartar todo lo que viene abajo: la sección de
    banco, los totales y notas sueltas.
    """
    if not len(df) or not marcadores:
        return df
    limpios = [N.limpiar_texto(m) for m in marcadores if str(m).strip()]
    limpios = [m for m in limpios if m]
    if not limpios:
        return df
    # Texto normalizado (mayúsculas/sin acentos) de toda la fila junta.
    texto_filas = df.apply(lambda fila: N.limpiar_texto(" ".join(map(str, fila.values))), axis=1)
    for pos in range(len(df)):
        if any(m in texto_filas.iloc[pos] for m in limpios):
            return df.iloc[:pos].reset_index(drop=True)
    return df


def construir(df: pd.DataFrame, mapeo: dict, config: dict, modo: str) -> pd.DataFrame:
    """Despacha según el modo ('simple' o 'doble') y aplica exclusiones."""
    if modo == "doble":
        out = construir_doble(df, mapeo, config)
    else:
        out = construir_simple(df, mapeo, config)
    return _excluir_por_detalle(out, config.get("excluir_detalle") or [])


# --------------------------------------------------------------------------- #
# Validación
# --------------------------------------------------------------------------- #

def validar_simple(mapeo: dict) -> list[str]:
    errores = []
    if not mapeo.get("fecha"):
        errores.append("Falta mapear **Fecha**.")
    if not (mapeo.get("importe") or mapeo.get("ingreso") or mapeo.get("egreso")
            or mapeo.get("debe") or mapeo.get("haber")):
        errores.append("Falta mapear el **Importe** (o Ingreso/Egreso, o Debe/Haber).")
    return errores


def validar_doble(mapeo: dict) -> list[str]:
    errores = []
    if not mapeo.get("ing_importe") and not mapeo.get("egr_importe"):
        errores.append("En modo doble columna mapeá al menos el **importe de ingresos** o el de **egresos**.")
    if mapeo.get("ing_importe") and not mapeo.get("ing_fecha"):
        errores.append("Mapeá la **fecha de ingresos**.")
    if mapeo.get("egr_importe") and not mapeo.get("egr_fecha"):
        errores.append("Mapeá la **fecha de egresos**.")
    return errores
