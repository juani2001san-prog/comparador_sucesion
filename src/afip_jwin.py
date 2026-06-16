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
    filas = list(csv.reader(io.StringIO(texto), delimiter=";"))
    # descarto filas totalmente vacías
    return [f for f in filas if any(str(c).strip() for c in f)]


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
    """Devuelve (proveedores: CUIT->rubro, rubros: cod->descripcion)."""
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)

    proveedores = {}
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
        c_rubro = idx("rubro") or 3
        for r in range(2, ws.max_row + 1):
            cuit = _norm_cuit(ws.cell(r, c_cuit).value)
            rub = ws.cell(r, c_rubro).value
            if cuit and rub not in (None, ""):
                try:
                    proveedores[cuit] = int(rub)
                except (TypeError, ValueError):
                    proveedores[cuit] = rub

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
    return proveedores, rubros


# --------------------------------------------------------------------------
# Procesamiento
# --------------------------------------------------------------------------
def procesar(csv_bytes, maestro_bytes):
    """Devuelve (encabezado, filas_con_rubro, desconocidos, stats)."""
    filas = _leer_csv_afip(csv_bytes)
    if not filas:
        raise ValueError("El CSV de AFIP está vacío.")
    encab, datos = filas[0], filas[1:]

    def col(nombre):
        for i, h in enumerate(encab):
            if nombre.lower() in str(h).strip().lower():
                return i
        return None

    i_cuit = col("Nro. Doc. Emisor")
    i_deno = col("Denominación Emisor")
    if i_cuit is None:
        raise ValueError("No encontré la columna 'Nro. Doc. Emisor' en el CSV de AFIP.")

    proveedores, rubros = _leer_maestro(maestro_bytes)

    salida = []
    desconocidos = {}
    asignados = 0
    for fila in datos:
        cuit = _norm_cuit(fila[i_cuit]) if i_cuit < len(fila) else ""
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
    }
    return encab_out, salida, desconocidos, stats, rubros


# --------------------------------------------------------------------------
# Salidas
# --------------------------------------------------------------------------
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
