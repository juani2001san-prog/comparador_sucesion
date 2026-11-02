# -*- coding: utf-8 -*-
"""
Conversor Excel → TXT para presentar percepciones de Ingresos Brutos
de Rio Negro.

El Excel del sistema propio trae columnas:
    T_COMP      → "FAC" (Factura) o "NCA"/"NC" (Nota de Credito)
    N_COMP      → "A0000600040029" (letra + 4 PV + 8 Nro)
    FECHA_EMI   → fecha de emision
    COD_CLIEN   → interno
    RAZON_SOC   → razon social
    IDENTIFTRI  → CUIT con guiones (20-13779247-9)
    IMPORTE     → total (no se usa en el TXT)
    COD_IMPUE / DES_IMPUE / COD_ALICU / DES_ALICU / PORCE_IMP
    IMPUESTO    → monto de la percepcion
    IMP_GRAV    → base imponible (importe sobre el que se calculo la percepcion)

El TXT destino es un CSV sin encabezado, una linea por comprobante:

    orden,tipo,letra,pv_mas_nro,cuit,fecha,base_imp,1,percepcion,070,916

Donde:
    orden      : correlativo "00001", "00002", ...
    tipo       : 1 (Factura A) o 102 (NC-A)  — las NC quedan con signo negativo
    letra      : "A" (del N_COMP)
    pv_mas_nro : ultimos 12 digitos del N_COMP (4 PV + 8 Nro)
    cuit       : 11 digitos sin guiones
    fecha      : DD/MM/YYYY
    base_imp   : IMP_GRAV con 2 decimales (punto decimal)
    percepcion : IMPUESTO con 2 decimales (punto decimal)
    070, 916   : codigos fijos (070 = tipo de retencion/percepcion,
                                916 = jurisdiccion Rio Negro)
"""

from __future__ import annotations

import io
from datetime import datetime

import pandas as pd


def convertir(xls_bytes: bytes) -> tuple[bytes, dict]:
    """
    Lee el Excel y devuelve (txt_bytes, stats).

    stats = {
        "lineas": int,
        "importe_total": float,         # suma de IMPUESTO (con signo)
        "mes_periodo": "MM-YYYY",
        "nombre_sugerido": str,         # para el download_button
        "filas_descartadas": int,       # sin CUIT valido o fila vacia
    }
    """
    df = pd.read_excel(io.BytesIO(xls_bytes), header=0)
    df.columns = [str(c).strip() for c in df.columns]

    requeridas = ["T_COMP", "N_COMP", "FECHA_EMI", "IDENTIFTRI",
                  "IMPUESTO", "IMP_GRAV"]
    faltantes = [c for c in requeridas if c not in df.columns]
    if faltantes:
        raise ValueError(
            f"Faltan columnas en el Excel: {', '.join(faltantes)}. "
            f"Columnas encontradas: {list(df.columns)}"
        )

    lineas: list[str] = []
    importe_total = 0.0
    fechas = []
    descartadas = 0

    for _, row in df.iterrows():
        cuit_raw = str(row["IDENTIFTRI"] or "").strip()
        cuit = "".join(c for c in cuit_raw if c.isdigit())
        if len(cuit) != 11:
            descartadas += 1
            continue

        t_comp = str(row["T_COMP"] or "").strip().upper()
        n_comp = str(row["N_COMP"] or "").strip()

        try:
            imp_grav = float(row["IMP_GRAV"])
            impuesto = float(row["IMPUESTO"])
        except (TypeError, ValueError):
            descartadas += 1
            continue

        fecha = pd.to_datetime(row["FECHA_EMI"], errors="coerce")
        if pd.isna(fecha):
            descartadas += 1
            continue

        # Codigo de tipo: FAC = 1, NC = 102 (y con signo negativo)
        if "FAC" in t_comp:
            cod_tipo = "1"
        else:
            cod_tipo = "102"
            imp_grav = -abs(imp_grav)
            impuesto = -abs(impuesto)

        # Letra y digitos del comprobante
        if n_comp and not n_comp[0].isdigit():
            letra = n_comp[0]
            digitos = "".join(c for c in n_comp[1:] if c.isdigit())
        else:
            letra = "A"
            digitos = "".join(c for c in n_comp if c.isdigit())
        pv_nro = digitos[-12:].zfill(12)

        orden = f"{len(lineas) + 1:05d}"
        fecha_str = fecha.strftime("%d/%m/%Y")
        base_str = f"{imp_grav:.2f}"
        perc_str = f"{impuesto:.2f}"

        linea = (f"{orden},{cod_tipo},{letra},{pv_nro},{cuit},"
                 f"{fecha_str},{base_str},1,{perc_str},070,916")
        lineas.append(linea)
        importe_total += impuesto
        fechas.append(fecha)

    # Periodo (mes mas frecuente)
    if fechas:
        meses = pd.Series(fechas).dt.strftime("%m-%Y")
        mes_periodo = meses.mode().iloc[0]
    else:
        mes_periodo = datetime.now().strftime("%m-%Y")

    importe_str = f"{importe_total:,.2f}".replace(",", "@").replace(".", ",").replace("@", "")
    nombre_sugerido = (f"Per Ib Rn {mes_periodo}  "
                       f"{len(lineas)} lineas $ {importe_str}.txt")

    txt = "\r\n".join(lineas) + ("\r\n" if lineas else "")
    return txt.encode("latin-1", errors="replace"), {
        "lineas": len(lineas),
        "importe_total": round(importe_total, 2),
        "mes_periodo": mes_periodo,
        "nombre_sugerido": nombre_sugerido,
        "filas_descartadas": descartadas,
    }
