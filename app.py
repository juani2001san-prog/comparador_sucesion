"""
Comparador Contabilidad vs Planilla (sucesión)
==============================================

Subís el Excel que te envían (la caja) y el que extraés de tu contabilidad
(Libro Diario), y la app te muestra las **diferencias** movimiento a movimiento,
cruzando por **importe + fecha**.

Ejecutar:
    streamlit run app.py
"""

from __future__ import annotations

import io
from datetime import datetime

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from src import conciliar as CON
from src import excel_reader as XLS
from src import mapeo as MAP
from src import pdf_banco as PDFB
from src import reporte as REP
from src import ps3_micro as PS3
from src import ventas_tango as VT
from src import afip_jwin as AJ
from src.normalizar import formato_ar

st.set_page_config(page_title="Herramientas del estudio", page_icon="🐣", layout="wide")


def _desactivar_traduccion() -> None:
    """Le indica al navegador que NO traduzca la página.

    El traductor automático de Chrome modifica el DOM y choca con cómo Streamlit
    lo redibuja, provocando el error 'removeChild ... NotFoundError'. Marcando la
    página como 'notranslate' se evita ese crash.
    """
    components.html(
        """
        <script>
        try {
            var doc = window.parent.document;
            doc.documentElement.setAttribute('translate', 'no');
            doc.documentElement.classList.add('notranslate');
            if (!doc.querySelector('meta[name=\"google\"][content=\"notranslate\"]')) {
                var m = doc.createElement('meta');
                m.name = 'google';
                m.content = 'notranslate';
                doc.head.appendChild(m);
            }
        } catch (e) {}
        </script>
        """,
        height=0,
    )


@st.cache_data(show_spinner=False)
def _hojas(contenido: bytes) -> list[str]:
    return XLS.listar_hojas(io.BytesIO(contenido))


@st.cache_data(show_spinner=False)
def _leer(contenido: bytes, hoja: str, fila: int) -> pd.DataFrame:
    return XLS.leer_hoja(io.BytesIO(contenido), hoja, fila)


def _selector_columna(etiqueta, opciones, sugerida, key):
    idx = opciones.index(sugerida) if sugerida in opciones else 0
    sel = st.selectbox(etiqueta, opciones, index=idx, key=key)
    return None if sel == "(no usar)" else sel


def _meses_presentes(df) -> set:
    """Conjunto de meses (Period 'M') con datos en el DataFrame."""
    fechas = pd.to_datetime(df["fecha"], errors="coerce").dropna()
    return set(fechas.dt.to_period("M")) if len(fechas) else set()


def _meses_disponibles(df_c, df_p) -> list[str]:
    """Lista ordenada de meses (YYYY-MM) presentes en cualquiera de los dos lados."""
    return sorted(str(m) for m in (_meses_presentes(df_c) | _meses_presentes(df_p)))


def _filtrar_por_meses(df, meses: list[str]):
    """Deja sólo las filas cuyos meses estén en la lista elegida."""
    per = pd.to_datetime(df["fecha"], errors="coerce").dt.to_period("M").astype(str)
    return df[per.isin(meses)].reset_index(drop=True)


def _filtrar_resultado_por_meses(df_res, meses: list[str]):
    """Deja las filas del RESULTADO que toquen los meses elegidos por cualquiera
    de los dos lados (contabilidad o planilla). Así, si un movimiento se cargó en
    el mes equivocado, igual aparece cuando su par cae en el mes seleccionado."""
    fc = pd.to_datetime(df_res["fecha_contab"], errors="coerce").dt.to_period("M").astype(str)
    fp = pd.to_datetime(df_res["fecha_planilla"], errors="coerce").dt.to_period("M").astype(str)
    return df_res[fc.isin(meses) | fp.isin(meses)].reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Carga genérica de archivo (devuelve df_crudo, columnas, contenido)
# --------------------------------------------------------------------------- #

def _leer_archivo(archivo, fila: int, key_hoja: str, mostrar_nombre: bool):
    """Lee un archivo: elige hoja y devuelve el DataFrame (o None si falla)."""
    contenido = archivo.getvalue()
    try:
        hojas = _hojas(contenido)
    except Exception as exc:  # noqa: BLE001
        st.error(f"No se pudo leer '{archivo.name}': {exc}")
        return None
    etiqueta = archivo.name if mostrar_nombre else "Hoja"
    hoja = st.selectbox(etiqueta, hojas, key=key_hoja)
    try:
        return _leer(contenido, hoja, int(fila) - 1)
    except Exception as exc:  # noqa: BLE001
        st.error(f"No se pudo leer la hoja de '{archivo.name}': {exc}")
        return None


def cargar_archivo(titulo: str, prefijo: str):
    """Carga uno o varios Excel (misma estructura) y los combina en un DataFrame."""
    archivos = st.file_uploader(
        f"Excel — {titulo} (podés subir varios)",
        type=["xlsx", "xlsm", "xls"], key=f"{prefijo}_file",
        accept_multiple_files=True,
    )
    if not archivos:
        st.info("Esperando archivo…")
        return None

    fila = st.number_input("Fila de encabezados", 1, 50, 1, 1, key=f"{prefijo}_header")

    dfs = []
    if len(archivos) == 1:
        df = _leer_archivo(archivos[0], fila, f"{prefijo}_hoja_0", mostrar_nombre=False)
        if df is None:
            return None
        dfs.append(df)
    else:
        with st.expander(f"Elegí la hoja de cada archivo ({len(archivos)})", expanded=True):
            for k, archivo in enumerate(archivos):
                df = _leer_archivo(archivo, fila, f"{prefijo}_hoja_{k}", mostrar_nombre=True)
                if df is None:
                    return None
                dfs.append(df)

    # Combina asumiendo que todos los archivos tienen la misma estructura.
    try:
        df = pd.concat(dfs, ignore_index=True) if len(dfs) > 1 else dfs[0]
    except Exception as exc:  # noqa: BLE001
        st.error(f"No se pudieron combinar los archivos: {exc}")
        return None

    if df.empty:
        st.warning("No hay datos.")
        return None

    st.caption(f"Archivos: {len(archivos)} · Filas totales: {len(df)} · Columnas: {len(df.columns)}")
    st.dataframe(XLS.vista_previa(df), use_container_width=True, hide_index=True)
    return df


# --------------------------------------------------------------------------- #
# Bloque CONTABILIDAD (siempre tabla simple)
# --------------------------------------------------------------------------- #

def bloque_contabilidad():
    st.subheader("1) Mi contabilidad (Libro Diario)")
    df = cargar_archivo("contabilidad", "c")
    if df is None:
        return None

    cols = list(df.columns)
    opc = ["(no usar)"] + cols
    sug = {k: MAP.adivinar(k, cols) for k in ("fecha", "detalle", "importe", "ingreso", "egreso", "debe", "haber", "cuenta")}

    st.markdown("**Mapeo de columnas**")
    g = st.columns(3)
    mapeo = {}
    with g[0]:
        mapeo["fecha"] = _selector_columna("Fecha *", opc, sug["fecha"], "c_fecha")
        mapeo["detalle"] = _selector_columna("Detalle", opc, sug["detalle"], "c_detalle")
        mapeo["cuenta"] = _selector_columna("Cuenta / Centro (filtrar)", opc, sug["cuenta"], "c_cuenta")
    with g[1]:
        mapeo["debe"] = _selector_columna("Debe", opc, sug["debe"], "c_debe")
        mapeo["haber"] = _selector_columna("Haber", opc, sug["haber"], "c_haber")
    with g[2]:
        mapeo["importe"] = _selector_columna("Importe (con signo)", opc, sug["importe"], "c_importe")
        mapeo["ingreso"] = _selector_columna("Ingreso", opc, sug["ingreso"], "c_ing")
        mapeo["egreso"] = _selector_columna("Egreso", opc, sug["egreso"], "c_egr")
    st.caption("Para el importe usá **una** opción: Debe/Haber, o Importe (con signo), o Ingreso/Egreso.")

    invertir = st.checkbox("Invertir signo (contabilidad)", key="c_inv")

    # Filtro opcional por cuenta/centro de costo.
    df_filtrado = df
    if mapeo["cuenta"] and mapeo["cuenta"] in df.columns:
        valores = sorted({str(v) for v in df[mapeo["cuenta"]].dropna().unique()})
        elegidos = st.multiselect(
            f"Filtrar por '{mapeo['cuenta']}' (vacío = todo)", valores, key="c_filtro"
        )
        if elegidos:
            df_filtrado = df[df[mapeo["cuenta"]].astype(str).isin(elegidos)]
            st.caption(f"Filtrado: {len(df_filtrado)} de {len(df)} filas.")

    errores = MAP.validar_simple(mapeo)
    for e in errores:
        st.warning(e)

    return {"df": df_filtrado, "mapeo": mapeo, "config": {"invertir_signo": invertir},
            "modo": "simple", "errores": errores}


# --------------------------------------------------------------------------- #
# Bloque PLANILLA (simple o doble columna)
# --------------------------------------------------------------------------- #

def bloque_planilla():
    st.subheader("2) Planilla que me envían (caja)")
    df = cargar_archivo("planilla", "p")
    if df is None:
        return None

    cols = list(df.columns)
    opc = ["(no usar)"] + cols

    modo = st.radio(
        "Formato de la planilla",
        ["doble", "simple"],
        index=0,
        format_func=lambda x: "Una fila por movimiento" if x == "simple"
        else "Doble columna (ingresos | egresos)",
        horizontal=True,
        help="Para la caja de la sucesión usá 'Doble columna' (viene elegido por defecto).",
    )

    mapeo = {}
    if modo == "simple":
        sug = {k: MAP.adivinar(k, cols) for k in ("fecha", "detalle", "importe", "ingreso", "egreso", "debe", "haber")}
        st.markdown("**Mapeo de columnas**")
        g = st.columns(3)
        with g[0]:
            mapeo["fecha"] = _selector_columna("Fecha *", opc, sug["fecha"], "p_fecha")
            mapeo["detalle"] = _selector_columna("Detalle", opc, sug["detalle"], "p_detalle")
        with g[1]:
            mapeo["importe"] = _selector_columna("Importe (con signo)", opc, sug["importe"], "p_importe")
            mapeo["ingreso"] = _selector_columna("Ingreso", opc, sug["ingreso"], "p_ing")
        with g[2]:
            mapeo["egreso"] = _selector_columna("Egreso", opc, sug["egreso"], "p_egr")
        errores = MAP.validar_simple(mapeo)
    else:
        sd = MAP.adivinar_doble(cols)
        st.markdown("**Bloque de INGRESOS**")
        g1 = st.columns(3)
        with g1[0]:
            mapeo["ing_fecha"] = _selector_columna("Fecha ingresos *", opc, sd["ing_fecha"], "p_ing_fecha")
        with g1[1]:
            mapeo["ing_detalle"] = _selector_columna("Detalle ingresos", opc, sd["ing_detalle"], "p_ing_det")
        with g1[2]:
            mapeo["ing_importe"] = _selector_columna("Importe ingresos *", opc, sd["ing_importe"], "p_ing_imp")
        st.markdown("**Bloque de EGRESOS**")
        g2 = st.columns(3)
        with g2[0]:
            mapeo["egr_fecha"] = _selector_columna("Fecha egresos *", opc, sd["egr_fecha"], "p_egr_fecha")
        with g2[1]:
            mapeo["egr_detalle"] = _selector_columna("Detalle egresos", opc, sd["egr_detalle"], "p_egr_det")
        with g2[2]:
            mapeo["egr_importe"] = _selector_columna("Importe egresos *", opc, sd["egr_importe"], "p_egr_imp")
        errores = MAP.validar_doble(mapeo)

    invertir = st.checkbox("Invertir signo (planilla)", key="p_inv")
    for e in errores:
        st.warning(e)

    return {"df": df, "mapeo": mapeo, "config": {"invertir_signo": invertir},
            "modo": modo, "errores": errores}


# --------------------------------------------------------------------------- #
# Resultado simple
# --------------------------------------------------------------------------- #

def mostrar_resultado(df_res, res):
    total = len(df_res)
    ok = res["OK"]
    con_dif = total - ok

    # Resumen corto: lo que coincide vs lo que tiene diferencia.
    st.subheader("Resultado")
    c = st.columns(3)
    c[0].metric("✅ Coinciden (OK)", ok)
    c[1].metric("⚠️ Con diferencia", con_dif)
    c[2].metric("Diferencia de importe $", formato_ar(res["Total diferencia de importe"]))

    # Por defecto muestro SOLO los que tienen algo (no los OK).
    ver_todos = st.checkbox("Mostrar también los que coinciden (OK)", value=False)
    buscar = st.text_input("Buscar en el detalle (opcional)")

    d = df_res if ver_todos else df_res[df_res["estado"] != CON.ESTADO_OK]
    if buscar:
        t = buscar.upper()
        d = d[d["detalle_contab"].str.contains(t, na=False) | d["detalle_planilla"].str.contains(t, na=False)]

    if not len(d):
        st.success("¡No hay diferencias! Todo coincide. 🎉")
        return

    # Vista simple: qué es, fecha, detalle, cuánto de cada lado y la diferencia.
    fecha = d["fecha_contab"].where(d["fecha_contab"].notna(), d["fecha_planilla"])
    detalle = d["detalle_contab"].where(d["detalle_contab"] != "", d["detalle_planilla"])
    vista = pd.DataFrame({
        "Qué pasó": d["estado"],
        "Fecha": fecha,
        "Detalle": detalle,
        "Contabilidad $": d["importe_contab"],
        "Caja $": d["importe_planilla"],
        "Diferencia $": d["diferencia"],
    })
    st.caption(f"{len(vista)} movimientos con diferencia.")
    st.dataframe(
        vista, use_container_width=True, hide_index=True,
        column_config={
            "Fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
            "Contabilidad $": st.column_config.NumberColumn(format="%.2f"),
            "Caja $": st.column_config.NumberColumn(format="%.2f"),
            "Diferencia $": st.column_config.NumberColumn(format="%.2f"),
        },
    )


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #

def seccion_comparador():
    st.title("🧮 Contabilidad vs Planilla — Control de diferencias")
    st.caption("Cruce movimiento a movimiento por **importe + fecha**.")

    col1, col2 = st.columns(2)
    with col1:
        contab = bloque_contabilidad()
    with col2:
        planilla = bloque_planilla()

    st.divider()
    st.subheader("3) Opciones de cruce")
    o1, o2 = st.columns(2)
    with o1:
        comparar_abs = st.checkbox(
            "Comparar por valor absoluto (ignorar signo)", value=True,
            help="Útil cuando un lado guarda los egresos en negativo y el otro en positivo.",
        )
    with o2:
        ignorar_fechas = st.checkbox(
            "Ignorar diferencias de fecha (comparar por importe)", value=True,
            help="La caja suele fechar los movimientos distinto a la contabilidad. "
                 "Con esto tildado, si el importe coincide se considera OK aunque la fecha "
                 "difiera. Destildalo si querés controlar también las fechas.",
        )
        tolerancia = st.number_input(
            "Tolerancia de fecha (días)", 0, 60, 0, 1, disabled=ignorar_fechas,
            help="Solo aplica si NO estás ignorando las fechas: si el importe coincide y la "
                 "fecha difiere por hasta N días, se considera OK.",
        )

    st.caption(
        "ℹ️ La planilla se compara como **hoja completa del mes** (la solapa que elegís, "
        "ej. '2026 02'), sin importar la fecha de cada fila. La **contabilidad** sí se "
        "filtra por el mes elegido abajo. Así, si la caja fechó un movimiento en otro mes "
        "(ej. 31/01) pero está en la hoja de febrero, igual se aparea."
    )

    excluir_texto = st.text_input(
        "Excluir movimientos cuyo detalle contenga (separá con comas)",
        value="DIFERENCIA DE CAJA, TOTAL INGRESOS, TOTAL EGRESOS, SALDO ANTERIOR, "
              "SALDO ACTUAL, SUMAS IGUALES",
        help="Descarta líneas de cuadre/totales que no son movimientos reales. "
             "Se aplica a los dos lados. (La sección de banco del final ya se descarta "
             "con el corte de abajo, así que no hace falta excluir 'BANCO DEL CHUBUT'.)",
    )

    cortar_texto = st.text_input(
        "Cortar la planilla al llegar a (fin de la caja)",
        value="SUMAS IGUALES",
        help="La caja real suele terminar en 'SUMAS IGUALES'. Todo lo que esté DEBAJO "
             "de esa línea (la sección de banco, totales, notas) se ignora. Dejalo vacío "
             "para no cortar. Solo afecta a la planilla, no a la contabilidad.",
    )

    listo = (
        contab is not None and planilla is not None
        and not contab["errores"] and not planilla["errores"]
    )

    # Construyo los movimientos apenas el mapeo es válido, para poder ofrecer el
    # selector de meses (aunque los archivos tengan todos los meses cargados).
    df_c = df_p = None
    meses_sel: list[str] = []
    if listo:
        try:
            excluir = [t.strip() for t in excluir_texto.split(",") if t.strip()]
            contab["config"]["excluir_detalle"] = excluir
            planilla["config"]["excluir_detalle"] = excluir
            # Corto la planilla donde termina la caja (ej. "SUMAS IGUALES"), así
            # descarto la sección de banco y los totales que vienen abajo.
            cortar = [t.strip() for t in cortar_texto.split(",") if t.strip()]
            planilla_df = MAP.cortar_en(planilla["df"], cortar) if cortar else planilla["df"]
            df_c = MAP.construir(contab["df"], contab["mapeo"], contab["config"], contab["modo"])
            df_p = MAP.construir(planilla_df, planilla["mapeo"], planilla["config"], planilla["modo"])
        except Exception as exc:  # noqa: BLE001
            st.error(f"Error al preparar los datos: {exc}")
            df_c = df_p = None

    if df_c is not None and df_p is not None:
        meses_disp = _meses_disponibles(df_c, df_p)
        comunes = sorted(str(m) for m in (_meses_presentes(df_c) & _meses_presentes(df_p)))
        # Por defecto, UN SOLO mes (el más reciente en común) para no mezclar meses.
        if comunes:
            default = [comunes[-1]]
        elif meses_disp:
            default = [meses_disp[-1]]
        else:
            default = []
        meses_sel = st.multiselect(
            "📅 Meses a comparar", meses_disp, default=default,
            help="Elegí el/los mes(es) a cruzar. Por defecto viene UN solo mes (el más "
                 "reciente). Agregá más si querés comparar varios; vaciá la lista para todo.",
        )

    if st.button("🔍 Comparar", type="primary", disabled=not listo, use_container_width=True):
        try:
            if df_c is None or df_p is None:
                raise RuntimeError("No se pudieron preparar los datos. Revisá el mapeo.")
            # La planilla es la hoja de un mes: se usa COMPLETA (sin filtrar por la
            # fecha de cada fila). Solo se filtra la CONTABILIDAD por el mes elegido.
            # Así, si la caja fechó algo en otro mes pero está en la hoja, igual aparea.
            df_c_cmp = _filtrar_por_meses(df_c, meses_sel) if meses_sel else df_c
            df_p_cmp = df_p
            # Si se ignoran las fechas, uso una tolerancia enorme: alcanza con que el
            # importe coincida para considerarlo OK.
            tol = 10**9 if ignorar_fechas else int(tolerancia)
            df_res = CON.conciliar(df_c_cmp, df_p_cmp, comparar_abs=comparar_abs, tolerancia_dias=tol)
            st.session_state["res"] = {"df_res": df_res, "df_c": df_c_cmp, "df_p": df_p_cmp, "meses": meses_sel}
        except Exception as exc:  # noqa: BLE001
            st.error(f"Error al comparar: {exc}")
            st.exception(exc)

    if "res" not in st.session_state:
        st.info("Cargá ambos Excel, completá el mapeo y tocá **Comparar**.")
        return

    r = st.session_state["res"]
    df_res, df_c, df_p = r["df_res"], r["df_c"], r["df_p"]
    if df_res.empty:
        st.warning("No se generaron resultados.")
        return

    meses = r.get("meses")
    if meses:
        st.success(f"Meses comparados: {', '.join(meses)}")
    else:
        st.info("Se compararon todos los meses (no se filtró por mes).")

    res = CON.resumen(df_res, df_c, df_p)
    st.divider()
    mostrar_resultado(df_res, res)
    st.divider()

    st.subheader("Exportar")
    excel = REP.exportar_excel(df_res, df_c, df_p, REP.resumen_df(res))
    st.download_button(
        "⬇️ Descargar Excel de diferencias",
        data=excel,
        file_name=f"diferencias_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


# --------------------------------------------------------------------------- #
# Sección: PDF de banco → Excel
# --------------------------------------------------------------------------- #

def seccion_pdf_banco():
    st.title("🏦 PDF de banco → Excel")
    st.caption("Convertí un extracto bancario en PDF a una tabla. Optimizado para "
               "**Banco del Chubut** (también tiene un lector genérico de respaldo).")

    archivo = st.file_uploader("Subí el extracto en PDF", type=["pdf"], key="pdfbanco_file")
    if archivo is None:
        st.info("Esperando un PDF… (tiene que ser un PDF con texto, no escaneado/foto).")
        return

    try:
        df, parser = PDFB.pdf_a_dataframe(archivo.getvalue())
    except PDFB.ErrorConversion as exc:
        st.error(str(exc))
        return
    except Exception as exc:  # noqa: BLE001
        st.error(f"No se pudo convertir el PDF: {exc}")
        return

    st.success(f"Listo: {len(df)} movimientos detectados (lector: {parser}).")
    st.dataframe(
        df, use_container_width=True, hide_index=True,
        column_config={
            "Débito": st.column_config.NumberColumn(format="%.2f"),
            "Crédito": st.column_config.NumberColumn(format="%.2f"),
            "Saldo": st.column_config.NumberColumn(format="%.2f"),
        },
    )
    st.download_button(
        "⬇️ Descargar Excel",
        data=PDFB.dataframe_a_excel(df),
        file_name=f"extracto_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("Tip: después podés subir este Excel en el **Comparador** para cruzarlo "
               "contra tu contabilidad.")


# --------------------------------------------------------------------------- #
# Sección: JWIN → PS3 (MICROENV)
# --------------------------------------------------------------------------- #

@st.cache_data(show_spinner=False)
def _ps3_plan():
    """Carga el plan de cuentas fijo de MICROENV (CO válidos + correcciones)."""
    return PS3.cargar_plan()


def _ps3_bytes(resultado):
    """El .ps3/.txt va en codificación latin-1 con saltos CRLF."""
    return resultado.texto_ps3().encode("latin-1", "replace")


def _ps3_construir_reporte(resultados):
    """Arma el reporte de control en texto plano."""
    lineas = ["REPORTE DE CONTROL - GENERACION PS3 MICROENV",
              "Fecha de corrida: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"), ""]
    for res in resultados:
        lineas.append("=" * 60)
        lineas.append(f"EMPRESA: {res.empresa}   PERIODO: {res.periodo}")
        lineas.append(f"  Renglones: {res.cant_renglones}   Asientos: {res.cant_asientos}")
        lineas.append(f"  Total DEBE : {res.total_debe:,.2f}")
        lineas.append(f"  Total HABER: {res.total_haber:,.2f}")
        estado = "OK (balanceado)" if res.balanceado else f"DESCUADRE: {res.diferencia:,.2f}"
        lineas.append(f"  Control DEBE=HABER: {estado}")
        if res.correcciones_aplicadas:
            lineas.append("  Correcciones de cuenta aplicadas:")
            for cod, (nuevo, veces) in res.correcciones_aplicadas.items():
                lineas.append(f"    {cod} -> {nuevo}  ({veces} renglones)")
        if res.inexistentes:
            lineas.append("  CUENTAS INEXISTENTES SIN RESOLVER:")
            for cod, info in res.inexistentes.items():
                lineas.append(f"    {cod} (CO {info['co']})  {info['veces']} renglones")
        else:
            lineas.append("  Cuentas inexistentes: ninguna")
        lineas.append("")
    return "\n".join(lineas)


def _ps3_guardar_evidencia(resultados, reporte_txt, nombres_entrada):
    """Deja .txt, Excel, reporte.txt y log.txt en evidencia_microenv/AAAA-MM/.

    En la nube el disco es efímero/de solo lectura; si no se puede escribir, no
    rompe (las descargas siguen funcionando). Devuelve la ruta o None.
    """
    import os
    try:
        base = os.path.join(os.path.dirname(os.path.abspath(__file__)), "evidencia_microenv")
        periodo = next((r.periodo for r in resultados if r.periodo), "")
        carpeta = (periodo[3:] + "-" + periodo[:2]) if periodo else datetime.now().strftime("%Y-%m")
        destino = os.path.join(base, carpeta)
        os.makedirs(destino, exist_ok=True)

        archivos = []
        for res in resultados:
            nombre = PS3.nombre_archivo_ps3(res.empresa, res.periodo)
            with open(os.path.join(destino, nombre), "wb") as f:
                f.write(_ps3_bytes(res))
            archivos.append(nombre)
            nombre_xlsx = PS3.nombre_archivo_xlsx(res.empresa, res.periodo)
            with open(os.path.join(destino, nombre_xlsx), "wb") as f:
                f.write(PS3.construir_diario_xlsx(res))
            archivos.append(nombre_xlsx)

        with open(os.path.join(destino, "reporte.txt"), "w", encoding="utf-8") as f:
            f.write(reporte_txt)

        log_linea = (
            f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] "
            f"entrada={nombres_entrada} -> generados={archivos} "
            + "; ".join(
                f"{r.empresa}:reng={r.cant_renglones},balanceado={r.balanceado}"
                for r in resultados
            )
        )
        with open(os.path.join(destino, "log.txt"), "a", encoding="utf-8") as f:
            f.write(log_linea + "\n")
        return destino
    except OSError:
        return None  # en la nube no siempre se puede escribir; no es crítico


def _ps3_mostrar_controles(res):
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Renglones", f"{res.cant_renglones:,}")
    c2.metric("Asientos", f"{res.cant_asientos:,}")
    c3.metric("Total DEBE", f"{res.total_debe:,.2f}")
    c4.metric("Total HABER", f"{res.total_haber:,.2f}")
    if res.balanceado:
        st.success(f"DEBE = HABER (período {res.periodo})")
    else:
        st.error(f"DESCUADRE: la diferencia DEBE-HABER es {res.diferencia:,.2f}. Revisá el JWIN.")
    if res.correcciones_aplicadas:
        txt = ", ".join(f"{c} → {n} ({v})" for c, (n, v) in res.correcciones_aplicadas.items())
        st.info(f"Correcciones de cuenta aplicadas: {txt}")


def seccion_ps3():
    st.title("📒 JWIN → PS3 · MICROENV")
    st.caption(
        "Subí los JWIN crudos del mes. La app arma los renglones, valida las cuentas "
        "contra el plan, controla que DEBE = HABER y genera los archivos para tu sistema "
        "contable (Excel igual al *-ps3 + .txt de ancho fijo)."
    )

    co_validos, correcciones_base = _ps3_plan()

    with st.expander("Plan de cuentas (fijo)"):
        st.write(f"Cuentas válidas (CO): **{len(co_validos)}**")
        st.write("Correcciones fijas:")
        for k, v in correcciones_base.items():
            st.write(f"• {k} → {v}")
        st.caption("Para actualizar el plan, reemplazá src/plan_cuentas_microenv.json y reiniciá la app.")

    col_m, col_p = st.columns(2)
    with col_m:
        up_micro = st.file_uploader("JWIN Micro (.xlsx o .xls)", type=["xlsx", "xls"], key="ps3_micro")
    with col_p:
        up_pruebas = st.file_uploader("JWIN Pruebas (.xlsx o .xls)", type=["xlsx", "xls"], key="ps3_pruebas")

    entradas = []
    if up_micro is not None:
        entradas.append(("Micro", up_micro.name, up_micro.getvalue()))
    if up_pruebas is not None:
        entradas.append(("Pruebas", up_pruebas.name, up_pruebas.getvalue()))

    if not entradas:
        st.info("Esperando archivos. Subí al menos un JWIN para empezar.")
        return

    def procesar(file_bytes, empresa, correcciones):
        return PS3.procesar_jwin(io.BytesIO(file_bytes), empresa, co_validos, correcciones)

    # Primer procesamiento (solo correcciones fijas) para detectar inexistentes.
    previos = [(emp, nombre, procesar(data, emp, correcciones_base)) for emp, nombre, data in entradas]

    inexistentes_todos = {}
    for _, _, res in previos:
        for cod, info in res.inexistentes.items():
            d = inexistentes_todos.setdefault(cod, {"co": info["co"], "veces": 0, "empresas": set()})
            d["veces"] += info["veces"]
            d["empresas"].add(res.empresa)

    correcciones_manuales = {}
    if inexistentes_todos:
        st.warning(
            f"Se encontraron {len(inexistentes_todos)} código(s) de cuenta que no están en el plan. "
            "Ingresá el código correcto (9 dígitos) para cada uno antes de generar."
        )
        with st.form("ps3_correcciones"):
            for cod, info in inexistentes_todos.items():
                emps = ", ".join(sorted(info["empresas"]))
                nuevo = st.text_input(
                    f"Cuenta inexistente {cod}  (CO {info['co']}, {info['veces']} renglones, en {emps})",
                    key=f"ps3_fix_{cod}",
                    placeholder="Código correcto de 9 dígitos",
                )
                if nuevo.strip():
                    correcciones_manuales[cod] = nuevo.strip()
            st.form_submit_button("Aplicar correcciones")

    # Correcciones efectivas = fijas + tipeadas a mano.
    correcciones = dict(correcciones_base)
    correcciones.update(correcciones_manuales)

    resultados = [(emp, nombre, procesar(data, emp, correcciones)) for emp, nombre, data in entradas]

    sin_resolver = {}
    for _, _, res in resultados:
        for cod, info in res.inexistentes.items():
            sin_resolver[cod] = info

    for cod, nuevo in correcciones_manuales.items():
        if nuevo[:8] not in co_validos:
            st.error(f"El código {nuevo} que ingresaste para {cod} tampoco existe en el plan (CO {nuevo[:8]}).")

    st.divider()
    st.subheader("Controles")
    for emp, nombre, res in resultados:
        st.markdown(f"### {emp} — `{nombre}`")
        _ps3_mostrar_controles(res)

    st.divider()
    hay_descuadre = any(not res.balanceado for _, _, res in resultados)
    puede_generar = not sin_resolver

    if sin_resolver:
        st.error(
            "Todavía hay cuentas inexistentes sin resolver: "
            + ", ".join(sin_resolver.keys())
            + ". Completá las correcciones de arriba para poder generar."
        )
    if hay_descuadre:
        st.warning("Hay descuadre DEBE≠HABER en al menos una empresa. Revisá antes de subir al sistema.")

    if st.button("Generar archivos", type="primary", disabled=not puede_generar, key="ps3_generar"):
        solo_res = [res for _, _, res in resultados]
        nombres_entrada = [nombre for _, nombre, _ in resultados]
        reporte_txt = _ps3_construir_reporte(solo_res)
        destino = _ps3_guardar_evidencia(solo_res, reporte_txt, nombres_entrada)
        payload = {"reporte": reporte_txt, "destino": destino, "archivos": []}
        for res in solo_res:
            payload["archivos"].append({
                "empresa": res.empresa,
                "xlsx_nombre": PS3.nombre_archivo_xlsx(res.empresa, res.periodo),
                "xlsx_bytes": PS3.construir_diario_xlsx(res),
                "ps3_nombre": PS3.nombre_archivo_ps3(res.empresa, res.periodo),
                "ps3_bytes": _ps3_bytes(res),
            })
        st.session_state["ps3_generado"] = payload

    # Render persistente de las descargas (FUERA del if del botón), para que no
    # se borren al apretar un botón de descarga.
    generado = st.session_state.get("ps3_generado")
    if generado:
        destino = generado.get("destino")
        if destino:
            st.success(f"Listo. Evidencia guardada en: {destino}")
        else:
            st.success("Listo. Descargá los archivos abajo.")
        st.subheader("Descargar archivos")
        for a in generado["archivos"]:
            cxa, cxb = st.columns(2)
            cxa.download_button(
                f"⬇️ {a['xlsx_nombre']}  (Excel)",
                data=a["xlsx_bytes"],
                file_name=a["xlsx_nombre"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"ps3_dlx_{a['empresa']}",
            )
            cxb.download_button(
                f"⬇️ {a['ps3_nombre']}  (.ps3)",
                data=a["ps3_bytes"],
                file_name=a["ps3_nombre"],
                mime="application/octet-stream",
                key=f"ps3_dl_{a['empresa']}",
            )
        st.download_button(
            "⬇️ reporte.txt",
            data=generado["reporte"].encode("utf-8"),
            file_name="reporte.txt",
            mime="text/plain",
            key="ps3_dl_rep",
        )
        with st.expander("Ver reporte de control"):
            st.code(generado["reporte"])


# --------------------------------------------------------------------------- #
# Sección: Ventas por actividad (Tango) → IVA / Convenio Multilateral
# --------------------------------------------------------------------------- #

def seccion_ventas():
    st.title("🧾 Ventas por actividad (Tango)")
    st.caption(
        "Subí el export **'IVA por actividad'** de Tango (`F 2002 ventas…`). La app "
        "calcula, por actividad y categoría de IVA, el **neto y el IVA netos** "
        "(ventas − notas de crédito), listo para la DJ de IVA y Convenio Multilateral."
    )

    archivo = st.file_uploader("Export de ventas de Tango (.xls o .xlsx)",
                               type=["xls", "xlsx"], key="vt_file")
    if archivo is None:
        st.info("Esperando el archivo de Tango. Subí el reporte 'IVA por actividad'.")
        return

    try:
        detalle, por_act, resumen, totales = VT.procesar(archivo.getvalue())
    except Exception as exc:  # noqa: BLE001
        st.error(f"No pude procesar el archivo: {exc}")
        return

    c1, c2, c3 = st.columns(3)
    c1.metric("Neto (ventas − NC)", formato_ar(totales["Neto"]))
    c2.metric("IVA (ventas − NC)", formato_ar(totales["IVA"]))
    c3.metric("Facturado neto", formato_ar(totales["Facturado"]))

    st.subheader("Por actividad y categoría de IVA")
    st.caption("Total = Factura − Nota de crédito.")
    st.dataframe(
        por_act, use_container_width=True, hide_index=True,
        column_config={col: st.column_config.NumberColumn(format="%.2f")
                       for col in ["NC Neto", "NC IVA", "Factura Neto", "Factura IVA",
                                   "Total Neto", "Total IVA"]},
    )

    st.subheader("Resumen por actividad")
    st.dataframe(
        resumen, use_container_width=True, hide_index=True,
        column_config={"Neto Grav.+Ex.": st.column_config.NumberColumn(format="%.2f"),
                       "IVA": st.column_config.NumberColumn(format="%.2f")},
    )

    with st.expander("Ver detalle (renglón por renglón)"):
        st.dataframe(detalle.drop(columns=["Cat"]), use_container_width=True, hide_index=True)

    st.download_button(
        "⬇️ Descargar Excel (Por actividad + Resumen + Detalle)",
        data=VT.construir_excel(detalle, por_act, resumen, totales),
        file_name=f"ventas_por_actividad_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


# --------------------------------------------------------------------------- #
# Sección: AFIP (Portal IVA) → JWIN con rubros
# --------------------------------------------------------------------------- #

def seccion_afip():
    st.title("📥 AFIP (Portal IVA) → JWIN con rubros")
    st.caption(
        "Subí el **CSV de AFIP** (Mis Comprobantes → Recibidos) y tu **Excel maestro** "
        "(con las hojas Proveedores y Rubros). La app le agrega la columna **Rubro** "
        "buscando el CUIT del emisor, y te deja el archivo listo para importar a JWIN."
    )

    c1, c2 = st.columns(2)
    with c1:
        up_csv = st.file_uploader("Archivo de AFIP (.csv, .xlsx o .xls)",
                                  type=["csv", "xlsx", "xls"], key="afip_csv")
    with c2:
        up_maestro = st.file_uploader("Excel maestro (Proveedores/Rubros) (.xlsx)",
                                      type=["xlsx"], key="afip_maestro")

    with st.expander("¿No tenés el Excel maestro? Descargá una plantilla para empezar"):
        st.caption(
            "Viene con el catálogo de Rubros cargado y la hoja Proveedores vacía. "
            "Tip: subí esta plantilla + el CSV de AFIP y la app te va a listar TODOS los "
            "proveedores para que les asignes el rubro; después bajás el maestro completo."
        )
        st.download_button(
            "⬇️ Descargar plantilla de Excel maestro",
            data=AJ.construir_plantilla_maestro(),
            file_name="Maestro proveedores.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if up_csv is None or up_maestro is None:
        st.info("Subí los dos archivos: el CSV de AFIP y tu Excel maestro de proveedores.")
        return

    csv_bytes = up_csv.getvalue()
    maestro_bytes = up_maestro.getvalue()
    try:
        encab, filas, desconocidos, stats, rubros = AJ.procesar(csv_bytes, maestro_bytes)
    except Exception as exc:  # noqa: BLE001
        st.error(f"No pude procesar: {exc}")
        return

    # Si hay proveedores nuevos, dejar cargar el rubro ahí mismo.
    nuevos = []          # (cuit, deno, rubro, desc) para actualizar el maestro
    if desconocidos:
        st.warning(
            f"Hay {len(desconocidos)} proveedor(es) que NO están en tu lista. "
            "Asignales el rubro acá abajo (se aplica al toque y podés bajar tu maestro actualizado)."
        )
        opciones = [""] + [f"{c} - {d}" for c, d in sorted(rubros.items())]
        base = pd.DataFrame([{"CUIT": c, "Denominación": d, "Rubro": ""}
                             for c, d in sorted(desconocidos.items())])
        editado = st.data_editor(
            base, hide_index=True, use_container_width=True, key="afip_editor",
            column_config={
                "CUIT": st.column_config.TextColumn(disabled=True),
                "Denominación": st.column_config.TextColumn(disabled=True),
                "Rubro": st.column_config.SelectboxColumn("Rubro", options=opciones),
            },
        )
        extra = {}
        for _, row in editado.iterrows():
            sel = str(row["Rubro"]).strip()
            if sel:
                cod = int(sel.split(" - ")[0])
                cuit = AJ._norm_cuit(row["CUIT"])
                extra[cuit] = cod
                nuevos.append((cuit, row["Denominación"], cod, rubros.get(cod, "")))
        if extra:
            # Reprocesar con los rubros recién cargados.
            encab, filas, desconocidos, stats, rubros = AJ.procesar(csv_bytes, maestro_bytes, extra)

    m1, m2, m3 = st.columns(3)
    m1.metric("Comprobantes", stats["comprobantes"])
    m2.metric("Con rubro", stats["asignados"])
    m3.metric("Sin rubro", stats["sin_rubro"])

    if stats["sin_rubro"] == 0:
        st.success("Todos los comprobantes quedaron con su rubro. 🎉")

    with st.expander("Ver tabla (CUIT, emisor, rubro)"):
        try:
            i_cuit = encab.index(next(h for h in encab if "Nro. Doc. Emisor" in h))
            i_deno = encab.index(next(h for h in encab if "Denominación Emisor" in h))
        except StopIteration:
            i_cuit, i_deno = 7, 8
        vista = pd.DataFrame([
            {"CUIT": f[i_cuit] if i_cuit < len(f) else "",
             "Emisor": f[i_deno] if i_deno < len(f) else "",
             "Rubro": f[-1]}
            for f in filas
        ])
        st.dataframe(vista, use_container_width=True, hide_index=True)

    cda, cdb = st.columns(2)
    cda.download_button(
        "⬇️ CSV para importar a JWIN",
        data=AJ.construir_csv(encab, filas),
        file_name=f"Importacion JWIN {datetime.now():%m-%Y}.csv",
        mime="text/csv", use_container_width=True,
    )
    cdb.download_button(
        "⬇️ Excel (para revisar)",
        data=AJ.construir_excel(encab, filas, desconocidos, rubros),
        file_name=f"Importacion AFIP JWIN {datetime.now():%m-%Y}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    if nuevos:
        st.divider()
        st.info(
            "Cargaste proveedores nuevos. Bajá tu **maestro actualizado** (les quedan "
            "agregados a la hoja Proveedores) y guardalo para usarlo el mes que viene."
        )
        st.download_button(
            "⬇️ Excel maestro ACTUALIZADO (con los proveedores nuevos)",
            data=AJ.construir_maestro_actualizado(maestro_bytes, nuevos),
            file_name=up_maestro.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


# --------------------------------------------------------------------------- #
# Programa: menú de herramientas
# --------------------------------------------------------------------------- #

# Herramientas del programa: (clave, etiqueta del menú, función).
_HERRAMIENTAS = [
    ("comparador", "🧮  Comparador", lambda: seccion_comparador()),
    ("pdf", "🏦  PDF de banco → Excel", lambda: seccion_pdf_banco()),
    ("ps3", "📒  JWIN → PS3 (MICROENV)", lambda: seccion_ps3()),
    ("ventas", "🧾  Ventas por actividad (Tango)", lambda: seccion_ventas()),
    ("afip", "📥  AFIP → JWIN (rubros)", lambda: seccion_afip()),
]


def _password_ok() -> bool:
    """Pantalla de contraseña. La clave se define en st.secrets['password']
    (en Streamlit Cloud se carga en Settings → Secrets). Si no hay clave
    configurada (uso local), no pide nada."""
    try:
        correcta = st.secrets["password"]
    except Exception:
        correcta = None
    if not correcta:
        return True  # sin clave configurada → uso local sin pedir nada
    if st.session_state.get("auth_ok"):
        return True

    st.title("🔒 Herramientas del estudio")
    st.caption("Ingresá la contraseña para acceder.")
    with st.form("login"):
        pwd = st.text_input("Contraseña", type="password")
        entrar = st.form_submit_button("Entrar")
    if entrar:
        if pwd == correcta:
            st.session_state["auth_ok"] = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")
    return False


def main():
    _desactivar_traduccion()

    if not _password_ok():
        return

    if "seccion" not in st.session_state:
        st.session_state["seccion"] = _HERRAMIENTAS[0][0]

    with st.sidebar:
        st.markdown("### 🐣 Herramientas del estudio")
        st.caption("Elegí una herramienta")
        for clave, etiqueta, _ in _HERRAMIENTAS:
            activo = st.session_state["seccion"] == clave
            if st.button(
                etiqueta, key=f"nav_{clave}", use_container_width=True,
                type="primary" if activo else "secondary",
            ):
                st.session_state["seccion"] = clave
                st.rerun()
        st.divider()
        st.caption("Programa de uso diario.\nSe irán agregando más herramientas.")

    # Renderiza la sección activa.
    for clave, _, render in _HERRAMIENTAS:
        if st.session_state["seccion"] == clave:
            render()
            break


if __name__ == "__main__":
    main()
