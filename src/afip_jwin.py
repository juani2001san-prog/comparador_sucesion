# -*- coding: utf-8 -*-
"""
Importación AFIP → JWIN.

Toma el CSV de AFIP ("Mis Comprobantes - Comprobantes Recibidos") y le agrega
la columna **Rubro** buscando el CUIT del emisor en una lista maestra de
proveedores (CUIT → código de rubro 1..17). Devuelve el archivo listo para
importar a JWIN (Excel y/o CSV) y la lista de proveedores nuevos sin rubro.

El CSV de AFIP viene con separador ';', decimales con coma y, a veces, comillas.
Los valores se preservan tal cual (no se reconvierten) para no alterar el formato
que espera JWIN; solo se agrega el rubro al final.
"""

import csv
import io
from collections import Counter
from datetime import date, datetime


# --------------------------------------------------------------------------
# Normalización previa: CSV crudo del Portal IVA (Compras) → CSV para JWIN
# --------------------------------------------------------------------------
# JWIN importa por posición de columna y no tiene conceptos para "Impuestos
# Internos" ni "Otros Tributos". Si los dejamos, esos importes se pierden y la
# factura queda descuadrada (el total de cabecera no coincide con la suma).
#
# Esta normalización se aplica ANTES del mapeo de rubros: mueve el importe de
# esas dos columnas al "No Gravado" y las pone en cero. Así el comprobante
# cuadra y JWIN lo importa completo.
#
# El layout está atado al encabezado actual del Portal IVA (32 columnas). Si
# cambia, validar_encabezado_portal_iva() lo detecta y frena el proceso.

HEADER_PORTAL_IVA = [
    "Fecha de Emisión", "Tipo de Comprobante", "Punto de Venta",
    "Número de Comprobante", "Tipo Doc. Vendedor", "Nro. Doc. Vendedor",
    "Denominación Vendedor", "Importe Total", "Moneda Original", "Tipo de Cambio",
    "Importe No Gravado", "Importe Exento", "Crédito Fiscal Computable",
    "Importe de Per. o Pagos a Cta. de Otros Imp. Nac.",
    "Importe de Percepciones de Ingresos Brutos",
    "Importe de Impuestos Municipales",
    "Importe de Percepciones o Pagos a Cuenta de IVA",
    "Importe de Impuestos Internos", "Importe Otros Tributos",
    "Neto Gravado IVA 0%",
    "Neto Gravado IVA 2,5%", "Importe IVA 2,5%",
    "Neto Gravado IVA 5%", "Importe IVA 5%",
    "Neto Gravado IVA 10,5%", "Importe IVA 10,5%",
    "Neto Gravado IVA 21%", "Importe IVA 21%",
    "Neto Gravado IVA 27%", "Importe IVA 27%",
    "Total Neto Gravado", "Total IVA",
]

# Índices (0-based) de las columnas que uso en la normalización.
_COL_FECHA = 0
_COL_CUIT_VENDEDOR = 5      # F: Nro. Doc. Vendedor
_COL_DENOMINACION = 6       # G: Denominación Vendedor
_COL_IMPORTE_TOTAL = 7      # H: Importe Total
_COL_NO_GRAVADO = 10        # K: Importe No Gravado          ← destino
_COL_EXENTO = 11            # L: Importe Exento
_COL_PERCEPCIONES = (13, 14, 15, 16)  # N-Q: cuatro columnas de percepciones
_COL_IMP_INTERNOS = 17      # R: Importe de Impuestos Internos  ← se suma a K
_COL_OTROS_TRIBUTOS = 18    # S: Importe Otros Tributos         ← se suma a K
_COL_NETO_GRAV_IVA_0 = 19   # T: Neto Gravado IVA 0% (facturas C/B usan solo esta)
_COL_TOTAL_NETO_GRAVADO = 30  # AE: suma de los netos con IVA 2,5% a 27% (NO incluye 0%)
_COL_TOTAL_IVA = 31           # AF

# Tolerancia de redondeo al validar cuadre por comprobante.
_TOLERANCIA = 0.05


def _norm_cabecera(s):
    """Normaliza un nombre de columna para comparar: sin acentos, minúsculas, un solo espacio."""
    s = "".join(ch for ch in str(s).strip() if ch not in "﻿")
    s = s.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    s = s.replace("Á", "a").replace("É", "e").replace("Í", "i").replace("Ó", "o").replace("Ú", "u")
    s = s.replace("ñ", "n").replace("Ñ", "n")
    return " ".join(s.lower().split())


def es_csv_portal_iva(csv_bytes):
    """True si el archivo parece ser el CSV crudo del Portal IVA (Compras)."""
    try:
        filas = _leer_csv_afip(csv_bytes)
    except Exception:
        return False
    if not filas:
        return False
    encab = [_norm_cabecera(c) for c in filas[0]]
    esperados = [_norm_cabecera(h) for h in HEADER_PORTAL_IVA]
    # Chequeo laxo: tienen que coincidir al menos las 5 primeras y las 3 clave.
    if encab[:5] != esperados[:5]:
        return False
    for i in (_COL_NO_GRAVADO, _COL_IMP_INTERNOS, _COL_OTROS_TRIBUTOS):
        if i >= len(encab) or encab[i] != esperados[i]:
            return False
    return True


def _validar_encabezado(encab_original):
    """Compara el encabezado del CSV con el esperado. Devuelve lista de diferencias."""
    diffs = []
    esperados_n = [_norm_cabecera(h) for h in HEADER_PORTAL_IVA]
    encab_n = [_norm_cabecera(c) for c in encab_original]
    if len(encab_n) < len(esperados_n):
        diffs.append(
            f"El archivo tiene {len(encab_n)} columnas pero el Portal IVA usa "
            f"{len(esperados_n)}. Puede ser un layout distinto."
        )
    for i, (esp, act) in enumerate(zip(esperados_n, encab_n)):
        if esp != act:
            col_letra = _letra_columna(i)
            diffs.append(
                f"Columna {col_letra} (posición {i+1}): se esperaba "
                f"«{HEADER_PORTAL_IVA[i]}» y vino «{encab_original[i]}»."
            )
    return diffs


def _letra_columna(idx):
    """0 → A, 1 → B, ..., 25 → Z, 26 → AA, 27 → AB, ..."""
    s = ""
    n = idx
    while True:
        s = chr(ord("A") + n % 26) + s
        n = n // 26 - 1
        if n < 0:
            return s


def _cuit_valido(v):
    d = "".join(ch for ch in str(v or "") if ch.isdigit())
    return d if len(d) == 11 else None


def _num_ar(v):
    """Parsea un número del CSV de ARCA (decimales con coma) a float. Vacío → 0."""
    s = str(v or "").strip()
    if not s:
        return 0.0
    # ARCA usa coma decimal. Puede venir con o sin puntos de miles.
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _fmt_ar(v):
    """Formatea un float como número estilo ARCA (coma decimal, 2 decimales)."""
    if v is None:
        return "0"
    # Redondeo a 2 y siempre coma decimal. Cero se emite como "0" (igual que ARCA).
    r = round(float(v), 2)
    if r == 0:
        return "0"
    return f"{r:.2f}".replace(".", ",")


def _detectar_mes(filas_datos):
    """Devuelve 'MM-AAAA' según la fecha más frecuente en la columna Fecha."""
    meses = []
    for fila in filas_datos:
        if _COL_FECHA >= len(fila):
            continue
        s = str(fila[_COL_FECHA] or "").strip()
        # Fechas del CSV vienen como dd/mm/aaaa.
        try:
            d = datetime.strptime(s[:10], "%d/%m/%Y")
        except ValueError:
            try:
                d = datetime.strptime(s[:10], "%Y-%m-%d")
            except ValueError:
                continue
        meses.append(f"{d.month:02d}-{d.year}")
    if not meses:
        return datetime.now().strftime("%m-%Y")
    return Counter(meses).most_common(1)[0][0]


def normalizar_portal_iva(csv_bytes):
    """
    Normaliza el CSV crudo del Portal IVA (Compras) antes de importarlo a JWIN.

    Reglas:
      1. Descarta filas sin CUIT válido de 11 dígitos (elimina la fila de
         totales y las líneas vacías que ARCA deja al final).
      2. No Gravado (col K) = No Gravado + Impuestos Internos (R) + Otros
         Tributos (S). Pone R y S en cero para que no se dupliquen si más
         adelante se mapean en JWIN.
      3. Valida por comprobante que
             Total = Neto + No Gravado + Exento + IVA + percepciones
         con tolerancia de 0,05.
      4. Reescribe con el mismo formato que ARCA (';' + coma decimal + CRLF
         + BOM UTF-8) para no romper el mapeo por columna de JWIN.

    Devuelve un dict:
      {
        "csv_bytes":       bytes | None,   # None si hay descuadres
        "encabezado":      list[str],
        "comprobantes":    int,
        "modificados":     [{fecha, cuit, denominacion, importe_movido}],
        "descartados":     [{motivo, contenido}],
        "descuadres":      [{fecha, cuit, denominacion, total, calculado, diferencia}],
        "mes_periodo":     "MM-AAAA",
        "errores_header":  list[str],       # si no está vacía, no se procesó
      }
    """
    filas = _leer_csv_afip(csv_bytes)
    if not filas:
        return {
            "csv_bytes": None, "encabezado": [], "comprobantes": 0,
            "modificados": [], "descartados": [], "descuadres": [],
            "mes_periodo": datetime.now().strftime("%m-%Y"),
            "errores_header": ["El archivo está vacío."],
        }

    encab = list(filas[0])
    diffs = _validar_encabezado(encab)
    if diffs:
        return {
            "csv_bytes": None, "encabezado": encab, "comprobantes": 0,
            "modificados": [], "descartados": [], "descuadres": [],
            "mes_periodo": datetime.now().strftime("%m-%Y"),
            "errores_header": diffs,
        }

    modificados = []
    descartados = []
    descuadres = []
    filas_out = []

    for i, fila in enumerate(filas[1:], start=2):  # start=2 → número de fila real en el archivo
        # Extiendo si vino más corta (a veces ARCA corta cuando las últimas columnas son 0).
        f = list(fila) + [""] * (len(HEADER_PORTAL_IVA) - len(fila))

        cuit = _cuit_valido(f[_COL_CUIT_VENDEDOR])
        if not cuit:
            resumen = " | ".join(str(c) for c in fila[:8] if str(c).strip())
            descartados.append({
                "fila": i,
                "motivo": "Sin CUIT válido de 11 dígitos",
                "contenido": resumen or "(fila vacía)",
            })
            continue

        no_grav = _num_ar(f[_COL_NO_GRAVADO])
        imp_int = _num_ar(f[_COL_IMP_INTERNOS])
        otros_tr = _num_ar(f[_COL_OTROS_TRIBUTOS])
        movido = round(imp_int + otros_tr, 2)

        if movido != 0:
            nuevo_no_grav = round(no_grav + imp_int + otros_tr, 2)
            f[_COL_NO_GRAVADO] = _fmt_ar(nuevo_no_grav)
            f[_COL_IMP_INTERNOS] = "0"
            f[_COL_OTROS_TRIBUTOS] = "0"
            modificados.append({
                "fecha": str(f[_COL_FECHA] or "").strip(),
                "cuit": cuit,
                "denominacion": str(f[_COL_DENOMINACION] or "").strip(),
                "importe_movido": movido,
            })

        # Validación de cuadre (con los valores ya normalizados).
        # OJO: "Total Neto Gravado" (AE) suma los netos con IVA 2,5% al 27%,
        # pero NO incluye "Neto Gravado IVA 0%" (T). Las facturas C/B a menudo
        # usan solo la columna T, así que hay que sumarla explícitamente.
        total = _num_ar(f[_COL_IMPORTE_TOTAL])
        neto_grav = _num_ar(f[_COL_TOTAL_NETO_GRAVADO])
        neto_iva0 = _num_ar(f[_COL_NETO_GRAV_IVA_0])
        no_grav_n = _num_ar(f[_COL_NO_GRAVADO])
        exento = _num_ar(f[_COL_EXENTO])
        iva = _num_ar(f[_COL_TOTAL_IVA])
        perc = sum(_num_ar(f[c]) for c in _COL_PERCEPCIONES)
        calc = round(neto_grav + neto_iva0 + no_grav_n + exento + iva + perc, 2)
        # Si calc == 0 y el total no, la fila NO tiene desglose para validar
        # (típico de facturas B/C a consumidor final: solo viene el Total y el
        # resto de columnas en cero). No es un descuadre, es una fila sin data
        # para chequear — la dejamos pasar.
        if abs(calc) > _TOLERANCIA:
            dif = round(total - calc, 2)
            if abs(dif) > _TOLERANCIA:
                descuadres.append({
                    "fila": i,
                    "fecha": str(f[_COL_FECHA] or "").strip(),
                    "cuit": cuit,
                    "denominacion": str(f[_COL_DENOMINACION] or "").strip(),
                    "total": total,
                    "calculado": calc,
                    "diferencia": dif,
                })

        filas_out.append(f)

    mes_periodo = _detectar_mes(filas_out)

    # Si hay descuadres, NO devolvemos el CSV normalizado — la UI frena la descarga.
    csv_out = None
    if not descuadres:
        buf = io.StringIO()
        w = csv.writer(buf, delimiter=";", lineterminator="\r\n")
        w.writerow(encab)
        for f in filas_out:
            w.writerow(f)
        csv_out = buf.getvalue().encode("utf-8-sig")

    return {
        "csv_bytes": csv_out,
        "encabezado": encab,
        "comprobantes": len(filas_out),
        "modificados": modificados,
        "descartados": descartados,
        "descuadres": descuadres,
        "mes_periodo": mes_periodo,
        "errores_header": [],
    }


# --------------------------------------------------------------------------
# Lectura del CSV de AFIP
# --------------------------------------------------------------------------


# --------------------------------------------------------------------------
# Lectura del CSV de AFIP
# --------------------------------------------------------------------------
def _decodificar(data):
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1", "replace")


def _leer_csv_afip(data):
    texto = _decodificar(data)
    # autodetectar separador (AFIP usa ';', pero por las dudas)
    sep = ";" if texto.count(";") >= texto.count(",") else ","
    filas = list(csv.reader(io.StringIO(texto), delimiter=sep))
    return [f for f in filas if any(str(c).strip() for c in f)]


def _celda_texto(v):
    """Convierte una celda de Excel al texto estilo AFIP (fecha ISO, coma decimal)."""
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float):
        if v.is_integer():
            return str(int(v))
        return f"{v:.2f}".replace(".", ",")
    return str(v).strip()


def _leer_excel_afip(data):
    """Lee el Excel de AFIP: busca la hoja con el encabezado correcto y devuelve
    filas (encabezado + datos) como texto, ignorando una columna 'Rubro' previa."""
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    hoja = None
    for nm in wb.sheetnames:
        ws = wb[nm]
        cab = " ".join(str(ws.cell(1, c).value or "") for c in range(1, ws.max_column + 1)).lower()
        if "nro. doc. emisor" in cab or "fecha de emisi" in cab:
            hoja = ws
            break
    if hoja is None:
        hoja = wb.active
    filas = []
    for fila in hoja.iter_rows(values_only=True):
        if any(c not in (None, "") for c in fila):
            filas.append([_celda_texto(c) for c in fila])
    return filas


def _leer_entrada(data):
    """Acepta CSV o Excel (xlsx/xls). Devuelve filas (encabezado + datos) como texto."""
    if data[:2] == b"PK" or data[:4] == b"\xd0\xcf\x11\xe0":
        return _leer_excel_afip(data)
    return _leer_csv_afip(data)


# Layout canónico que espera el importador de JWIN (formato AFIP "viejo", 30 cols).
_CANON = [
    "Fecha de Emisión", "Tipo de Comprobante", "Punto de Venta", "Número Desde",
    "Número Hasta", "Cód. Autorización", "Tipo Doc. Emisor", "Nro. Doc. Emisor",
    "Denominación Emisor", "Tipo Doc. Receptor", "Nro. Doc. Receptor", "Tipo Cambio",
    "Moneda", "Imp. Neto Gravado IVA 0%", "IVA 2,5%", "Imp. Neto Gravado IVA 2,5%",
    "IVA 5%", "Imp. Neto Gravado IVA 5%", "IVA 10,5%", "Imp. Neto Gravado IVA 10,5%",
    "IVA 21%", "Imp. Neto Gravado IVA 21%", "IVA 27%", "Imp. Neto Gravado IVA 27%",
    "Imp. Neto Gravado Total", "Imp. Neto No Gravado", "Imp. Op. Exentas",
    "Otros Tributos", "Total IVA", "Imp. Total",
]
# Nombres equivalentes en el formato AFIP "nuevo" (montos en pesos / Vendedor).
_ALT = {
    "Número Desde": ["Número de Comprobante"], "Número Hasta": ["Número de Comprobante"],
    "Tipo Doc. Emisor": ["Tipo Doc. Vendedor"], "Nro. Doc. Emisor": ["Nro. Doc. Vendedor"],
    "Denominación Emisor": ["Denominación Vendedor"], "Tipo Cambio": ["Tipo de Cambio"],
    "Moneda": ["Moneda Original"], "Imp. Neto Gravado IVA 0%": ["Neto Gravado IVA 0%"],
    "IVA 2,5%": ["Importe IVA 2,5%"], "Imp. Neto Gravado IVA 2,5%": ["Neto Gravado IVA 2,5%"],
    "IVA 5%": ["Importe IVA 5%"], "Imp. Neto Gravado IVA 5%": ["Neto Gravado IVA 5%"],
    "IVA 10,5%": ["Importe IVA 10,5%"], "Imp. Neto Gravado IVA 10,5%": ["Neto Gravado IVA 10,5%"],
    "IVA 21%": ["Importe IVA 21%"], "Imp. Neto Gravado IVA 21%": ["Neto Gravado IVA 21%"],
    "IVA 27%": ["Importe IVA 27%"], "Imp. Neto Gravado IVA 27%": ["Neto Gravado IVA 27%"],
    "Imp. Neto Gravado Total": ["Total Neto Gravado"], "Imp. Neto No Gravado": ["Importe No Gravado"],
    "Imp. Op. Exentas": ["Importe Exento"],
    "Otros Tributos": ["Importe de Percepciones de Ingresos Brutos"],
    "Imp. Total": ["Importe Total"],
}


def _norm_h(s):
    return " ".join(str(s).strip().lower().split())


def _a_layout_jwin(encab, datos):
    """Reordena cualquier formato de AFIP al layout canónico que espera JWIN.
    Para el formato viejo es identidad; el nuevo lo remapea por significado."""
    idx = {}
    for i, h in enumerate(encab):
        idx.setdefault(_norm_h(h), i)

    def buscar(canon):
        for c in [canon] + _ALT.get(canon, []):
            j = idx.get(_norm_h(c))
            if j is not None:
                return j
        return None

    cols = [buscar(c) for c in _CANON]
    nuevos = [[(f[j] if (j is not None and j < len(f)) else "") for j in cols] for f in datos]
    return list(_CANON), nuevos


def _norm_cuit(valor):
    if valor is None:
        return ""
    s = str(valor).strip()
    if s.endswith(".0"):  # por si vino como número
        s = s[:-2]
    return "".join(ch for ch in s if ch.isdigit())


# --------------------------------------------------------------------------
# Lectura del maestro (Proveedores + Rubros)
# --------------------------------------------------------------------------
def _leer_maestro(data):
    """
    Devuelve tres diccionarios:
      - proveedores: CUIT -> rubro (código o texto)
      - razones:    CUIT -> razón social (para completar filas sin denominación)
      - rubros:     código -> descripción del rubro
    """
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)

    proveedores = {}
    razones = {}
    hoja_prov = next((h for h in wb.sheetnames if "proveedor" in h.lower()), None)
    if hoja_prov:
        ws = wb[hoja_prov]
        hdr = [str(ws.cell(1, c).value or "").strip().lower() for c in range(1, ws.max_column + 1)]

        def idx(clave):
            for i, h in enumerate(hdr):
                if clave in h:
                    return i + 1
            return None

        c_cuit = idx("cuit") or 1
        c_razon = idx("razon") or idx("razón") or idx("denominacion") or idx("denominación") or 2
        c_rubro = idx("rubro") or 3
        for r in range(2, ws.max_row + 1):
            cuit = _norm_cuit(ws.cell(r, c_cuit).value)
            if not cuit:
                continue
            rub = ws.cell(r, c_rubro).value
            if rub not in (None, ""):
                try:
                    proveedores[cuit] = int(rub)
                except (TypeError, ValueError):
                    proveedores[cuit] = rub
            razon = ws.cell(r, c_razon).value
            if razon not in (None, ""):
                razones[cuit] = str(razon).strip()

    rubros = {}
    hoja_rub = next((h for h in wb.sheetnames if "rubro" in h.lower()), None)
    if hoja_rub:
        ws = wb[hoja_rub]
        for r in range(2, ws.max_row + 1):
            cod, desc = ws.cell(r, 1).value, ws.cell(r, 2).value
            if cod not in (None, ""):
                try:
                    rubros[int(cod)] = desc
                except (TypeError, ValueError):
                    pass
    return proveedores, razones, rubros


# --------------------------------------------------------------------------
# Procesamiento
# --------------------------------------------------------------------------
def procesar(csv_bytes, maestro_bytes, extra=None):
    """Devuelve (encabezado, filas_con_rubro, desconocidos, stats, rubros).

    `extra`: dict CUIT->rubro con asignaciones cargadas a mano en la app
    (proveedores nuevos que todavía no están en el maestro).
    """
    filas = _leer_entrada(csv_bytes)
    if not filas:
        raise ValueError("El archivo de AFIP está vacío.")
    encab, datos = list(filas[0]), [list(f) for f in filas[1:]]

    # Recortar del final todas las columnas que NO son parte del layout
    # Portal IVA de ARCA: 'Rubro' (si ya venía asignado), 'CAE', 'Origen',
    # o cualquier extra que un flujo previo haya sumado al final del CSV.
    #
    # Motivo: JWIN importa por posición y espera 'Rubro' en la columna AG
    # (col 33), inmediatamente después de las 32 columnas estándar de ARCA.
    # Si dejáramos las extras, el Rubro terminaría en la posición
    # equivocada y JWIN no lo encontraría.
    cols_arca_norm = {_norm_cabecera(c) for c in HEADER_PORTAL_IVA}
    while encab and _norm_cabecera(encab[-1]) not in cols_arca_norm:
        encab = encab[:-1]
        datos = [f[:len(encab)] for f in datos]

    # Dejamos el archivo tal cual viene de AFIP y solo agregamos el Rubro al final.
    def col(*nombres):
        for nm in nombres:
            for i, h in enumerate(encab):
                if nm.lower() in str(h).strip().lower():
                    return i
        return None

    # Acepta el formato viejo ('Emisor') y el nuevo 'montos en pesos' ('Vendedor').
    i_cuit = col("Nro. Doc. Emisor", "Nro. Doc. Vendedor", "Nro. Doc")
    i_deno = col("Denominación Emisor", "Denominación Vendedor", "Denominación")
    if i_cuit is None:
        raise ValueError("No encontré la columna del CUIT del proveedor "
                         "('Nro. Doc. Emisor' o 'Nro. Doc. Vendedor') en el archivo de AFIP.")

    proveedores, razones, rubros = _leer_maestro(maestro_bytes)
    # Sumar las asignaciones cargadas a mano (tienen prioridad).
    for cuit, rub in (extra or {}).items():
        c = _norm_cuit(cuit)
        if c and rub not in (None, ""):
            try:
                proveedores[c] = int(rub)
            except (TypeError, ValueError):
                proveedores[c] = rub

    salida = []
    desconocidos = {}
    asignados = 0
    autocompletados = 0
    for fila in datos:
        cuit = _norm_cuit(fila[i_cuit]) if i_cuit < len(fila) else ""

        # Si la Denominación está vacía y el CUIT figura en el maestro con
        # razón social, la completamos automáticamente. Típico caso: filas
        # que vinieron por QR de AFIP (el QR no trae razón social).
        if i_deno is not None and i_deno < len(fila) and cuit:
            deno_actual = str(fila[i_deno] or "").strip()
            if not deno_actual and cuit in razones:
                fila[i_deno] = razones[cuit]
                autocompletados += 1

        rub = proveedores.get(cuit, "")
        if rub == "":
            deno = fila[i_deno] if (i_deno is not None and i_deno < len(fila)) else ""
            if cuit:
                desconocidos.setdefault(cuit, deno)
        else:
            asignados += 1
        salida.append(list(fila) + [rub])

    encab_out = list(encab) + ["Rubro"]
    stats = {
        "comprobantes": len(datos),
        "asignados": asignados,
        "sin_rubro": len(datos) - asignados,
        "proveedores_nuevos": len(desconocidos),
        "denominaciones_autocompletadas": autocompletados,
    }
    return encab_out, salida, desconocidos, stats, rubros


# --------------------------------------------------------------------------
# Salidas
# --------------------------------------------------------------------------
# Catálogo de rubros estándar (JWIN). Sirve para armar una plantilla nueva.
RUBROS_ESTANDAR = [
    (1, "Movilidad y Viaticos"), (2, "Fletes"), (3, "Honorarios y Aranceles"),
    (4, "Repuestos y reparaciones"), (5, "Combustibles"), (6, "Internet"),
    (7, "Obra Social"), (8, "Semillas e insumos"), (9, "Telefonia"),
    (10, "Agroquimicos"), (11, "Gastos Varios"), (12, "Hacienda"),
    (13, "Gastos 10.5"), (14, "Laboreos"), (15, "Luz, gas y agua"),
    (16, "Alquileres"), (17, "Combustibles 10.5"),
]


def construir_plantilla_maestro(rubros=None):
    """Devuelve un Excel maestro vacío: hoja Rubros (catálogo) + hoja Proveedores
    (solo encabezados) + Instrucciones. Para empezar de cero un maestro."""
    import openpyxl
    from openpyxl.styles import Font

    rubros = rubros or RUBROS_ESTANDAR
    bold = Font(bold=True)
    wb = openpyxl.Workbook()

    wi = wb.active
    wi.title = "Instrucciones"
    txt = [
        "EXCEL MAESTRO DE PROVEEDORES (para AFIP -> JWIN)",
        "",
        "Hoja 'Rubros': catálogo de rubros (código 1..17). Ya viene cargado.",
        "Hoja 'Proveedores': cargá acá cada proveedor con su CUIT y el código de rubro.",
        "   - CUIT: los 11 dígitos, sin guiones.",
        "   - Rubro: el número del catálogo (mirá la hoja Rubros).",
        "",
        "Tip: en la herramienta AFIP -> JWIN, cuando aparezca un proveedor nuevo,",
        "lo cargás ahí mismo y la app te devuelve este maestro ya actualizado.",
    ]
    for i, t in enumerate(txt, start=1):
        wi.cell(i, 1, t)

    wr = wb.create_sheet("Rubros")
    wr.cell(1, 1, "Código").font = bold
    wr.cell(1, 2, "Descripción").font = bold
    for i, (cod, desc) in enumerate(rubros, start=2):
        wr.cell(i, 1, cod)
        wr.cell(i, 2, desc)

    wp = wb.create_sheet("Proveedores")
    for j, t in enumerate(["CUIT", "Razón Social", "Rubro", "Descripción Rubro"], start=1):
        wp.cell(1, j, t).font = bold
    wp.column_dimensions["A"].width = 16
    wp.column_dimensions["B"].width = 50

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def construir_maestro_actualizado(maestro_bytes, nuevos):
    """Devuelve el Excel maestro con los proveedores nuevos AGREGADOS al final
    de la hoja Proveedores. `nuevos` = lista de (cuit, denominacion, rubro, desc)."""
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(maestro_bytes))
    hoja = next((h for h in wb.sheetnames if "proveedor" in h.lower()), None)
    if hoja is None:
        hoja = "Proveedores"
        ws = wb.create_sheet(hoja)
        ws.append(["CUIT", "Razón Social", "Rubro", "Descripción Rubro"])
    ws = wb[hoja]
    fila = ws.max_row + 1
    for cuit, deno, rubro, desc in nuevos:
        ws.cell(fila, 1, str(cuit))
        ws.cell(fila, 2, deno)
        ws.cell(fila, 3, rubro)
        ws.cell(fila, 4, desc)
        fila += 1
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def construir_csv(encab, filas):
    """CSV UTF-8 con separador ';' (mismo formato que AFIP) + columna Rubro."""
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=";", lineterminator="\r\n")
    w.writerow(encab)
    for f in filas:
        w.writerow(f)
    return buf.getvalue().encode("utf-8-sig")


def construir_excel(encab, filas, desconocidos, rubros):
    import openpyxl
    from openpyxl.styles import Font, PatternFill

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "AFIP"
    bold = Font(bold=True)
    rojo = PatternFill("solid", fgColor="FFC7CE")

    for j, t in enumerate(encab, start=1):
        ws.cell(1, j, t).font = bold
    col_rubro = len(encab)
    for i, fila in enumerate(filas, start=2):
        for j, v in enumerate(fila, start=1):
            c = ws.cell(i, j, v)
            if j == col_rubro and (v == "" or v is None):
                c.fill = rojo  # sin rubro: resaltado
    ws.freeze_panes = "A2"

    # Hoja de proveedores nuevos sin rubro
    wd = wb.create_sheet("Proveedores nuevos")
    wd.cell(1, 1, "CUIT").font = bold
    wd.cell(1, 2, "Denominación").font = bold
    wd.cell(1, 3, "Rubro (completar)").font = bold
    for i, (cuit, deno) in enumerate(sorted(desconocidos.items()), start=2):
        wd.cell(i, 1, cuit)
        wd.cell(i, 2, deno)
    wd.column_dimensions["A"].width = 16
    wd.column_dimensions["B"].width = 50

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
