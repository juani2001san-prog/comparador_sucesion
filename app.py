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
from src import rango_compras as RC
from src import monotributo as MT
from src import iva as IVA
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

def bloque_contabilidad(prefijo: str = "c", titulo: str = "1) Mi contabilidad (Libro Diario)"):
    st.subheader(titulo)
    df = cargar_archivo("contabilidad", prefijo)
    if df is None:
        return None

    cols = list(df.columns)
    opc = ["(no usar)"] + cols
    sug = {k: MAP.adivinar(k, cols) for k in ("fecha", "detalle", "importe", "ingreso", "egreso", "debe", "haber", "cuenta")}

    st.markdown("**Mapeo de columnas**")
    g = st.columns(3)
    mapeo = {}
    with g[0]:
        mapeo["fecha"] = _selector_columna("Fecha *", opc, sug["fecha"], f"{prefijo}_fecha")
        mapeo["detalle"] = _selector_columna("Detalle", opc, sug["detalle"], f"{prefijo}_detalle")
        mapeo["cuenta"] = _selector_columna("Cuenta / Centro (filtrar)", opc, sug["cuenta"], f"{prefijo}_cuenta")
    with g[1]:
        mapeo["debe"] = _selector_columna("Debe", opc, sug["debe"], f"{prefijo}_debe")
        mapeo["haber"] = _selector_columna("Haber", opc, sug["haber"], f"{prefijo}_haber")
    with g[2]:
        mapeo["importe"] = _selector_columna("Importe (con signo)", opc, sug["importe"], f"{prefijo}_importe")
        mapeo["ingreso"] = _selector_columna("Ingreso", opc, sug["ingreso"], f"{prefijo}_ing")
        mapeo["egreso"] = _selector_columna("Egreso", opc, sug["egreso"], f"{prefijo}_egr")
    st.caption("Para el importe usá **una** opción: Debe/Haber, o Importe (con signo), o Ingreso/Egreso.")

    invertir = st.checkbox("Invertir signo (contabilidad)", key=f"{prefijo}_inv")

    # Filtro opcional por cuenta/centro de costo.
    df_filtrado = df
    if mapeo["cuenta"] and mapeo["cuenta"] in df.columns:
        valores = sorted({str(v) for v in df[mapeo["cuenta"]].dropna().unique()})
        elegidos = st.multiselect(
            f"Filtrar por '{mapeo['cuenta']}' (vacío = todo)", valores, key=f"{prefijo}_filtro"
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

def mostrar_resultado(df_res, res, etiqueta_b: str = "Caja", key_prefix: str = ""):
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
    ver_todos = st.checkbox("Mostrar también los que coinciden (OK)", value=False, key=f"{key_prefix}ver_todos")
    buscar = st.text_input("Buscar en el detalle (opcional)", key=f"{key_prefix}buscar")

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
    col_b = f"{etiqueta_b} $"
    vista = pd.DataFrame({
        "Qué pasó": d["estado"],
        "Fecha": fecha,
        "Detalle": detalle,
        "Contabilidad $": d["importe_contab"],
        col_b: d["importe_planilla"],
        "Diferencia $": d["diferencia"],
    })
    st.caption(f"{len(vista)} movimientos con diferencia.")
    st.dataframe(
        vista, use_container_width=True, hide_index=True,
        column_config={
            "Fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
            "Contabilidad $": st.column_config.NumberColumn(format="%.2f"),
            col_b: st.column_config.NumberColumn(format="%.2f"),
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
    st.caption("Convertí un extracto bancario en PDF a una tabla. Soporta "
               "**Banco del Chubut**, **Banco de la Nación** y un lector genérico de respaldo.")

    col_arch, col_banco = st.columns([2, 1])
    with col_arch:
        archivo = st.file_uploader("Subí el extracto en PDF", type=["pdf"], key="pdfbanco_file")
    with col_banco:
        AUTO = "Detectar automáticamente"
        opciones = [AUTO] + PDFB.nombres_parsers()
        eleccion = st.selectbox(
            "Banco",
            opciones,
            index=0,
            key="pdfbanco_banco",
            help="Dejá en automático y la app intenta reconocerlo sola. "
                 "Si la detección falla, elegí el banco a mano.",
        )

    if archivo is None:
        st.info("Esperando un PDF… (tiene que ser un PDF con texto, no escaneado/foto).")
        return

    forzado = None if eleccion == AUTO else eleccion
    try:
        df, parser = PDFB.pdf_a_dataframe(archivo.getvalue(), parser_nombre=forzado)
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
# Sección: Banco vs Contabilidad
# --------------------------------------------------------------------------- #

# Palabras del extracto bancario que NO son transferencias reales
# (impuestos automáticos, comisiones, IVA sobre comisiones, etc.).
_FILTRO_BANCO_DEFAULT = "GRAVAMEN LEY, COMIS.TRANSF, COMIS.COMPENSACION, I.V.A."


def _df_banco_a_conciliar(df_pdf, palabras_excluir_csv: str):
    """
    Convierte el DataFrame del PDF de banco (Fecha/Descripción/Débito/Crédito/Saldo)
    al esquema que usa ``conciliar``: ``fecha, detalle, importe, cuenta``.

    - Importe: ``Crédito - Débito`` (ingreso positivo, egreso negativo).
    - Filtra filas cuyo detalle contenga las palabras especificadas (impuestos, comisiones).
    """
    from src import normalizar as N

    df = pd.DataFrame()
    df["fecha"] = df_pdf["Fecha"].apply(N.parsear_fecha)
    df["detalle"] = df_pdf["Descripción"].fillna("").apply(N.limpiar_texto)
    debito = df_pdf["Débito"].fillna(0)
    credito = df_pdf["Crédito"].fillna(0)
    df["importe"] = (credito - debito).apply(N.redondear)
    df["cuenta"] = ""
    df = df[df["importe"] != 0.0].reset_index(drop=True)

    palabras = [p.strip() for p in (palabras_excluir_csv or "").split(",") if p.strip()]
    if palabras:
        df = MAP._excluir_por_detalle(df, palabras).reset_index(drop=True)
    return df


def seccion_banco_contab():
    st.title("🔄 Banco vs Contabilidad")
    st.caption("Cruzá el extracto del banco (PDF) contra el mayor contable, "
               "matcheando por **importe** (y opcionalmente fecha). "
               "Se filtran impuestos y comisiones del lado banco.")

    c_izq, c_der = st.columns(2)

    # ---- Lado contabilidad ----
    with c_izq:
        contab = bloque_contabilidad(prefijo="bcc", titulo="1) Mayor contable")

    # ---- Lado banco ----
    with c_der:
        st.subheader("2) Extracto del banco (PDF)")
        archivo_pdf = st.file_uploader("PDF del extracto", type=["pdf"], key="bc_pdf")
        AUTO = "Detectar automáticamente"
        eleccion = st.selectbox(
            "Banco", [AUTO] + PDFB.nombres_parsers(), key="bc_banco_sel",
        )

        df_banco_raw = None
        parser_usado = None
        if archivo_pdf is not None:
            try:
                forzado = None if eleccion == AUTO else eleccion
                df_banco_raw, parser_usado = PDFB.pdf_a_dataframe(
                    archivo_pdf.getvalue(), parser_nombre=forzado
                )
                st.success(f"✓ {len(df_banco_raw)} movimientos extraídos ({parser_usado}).")
                st.dataframe(
                    df_banco_raw.head(6),
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Débito": st.column_config.NumberColumn(format="%.2f"),
                        "Crédito": st.column_config.NumberColumn(format="%.2f"),
                        "Saldo": st.column_config.NumberColumn(format="%.2f"),
                    },
                )
            except Exception as exc:  # noqa: BLE001
                st.error(f"No se pudo procesar el PDF: {exc}")

    st.divider()
    st.subheader("3) Filtros y opciones")

    palabras_filtro = st.text_input(
        "Excluir del banco las filas cuyo detalle contenga (separá con comas)",
        value=_FILTRO_BANCO_DEFAULT,
        key="bc_filtro_palabras",
        help="Líneas del extracto que NO son transferencias reales: impuestos "
             "automáticos (GRAVAMEN LEY), comisiones (COMIS.TRANSF, COMIS.COMPENSACION) "
             "e IVA sobre comisiones. Se descartan antes de comparar.",
    )

    palabras_filtro_c = st.text_input(
        "Excluir de la contabilidad las filas cuyo detalle contenga (separá con comas)",
        value="SALDO ANTERIOR, SALDO FINAL, TRANSPORTE",
        key="bc_filtro_palabras_c",
        help="Líneas del mayor que NO son movimientos reales (saldos de apertura, "
             "totales, líneas de transporte).",
    )

    o1, o2 = st.columns(2)
    with o1:
        ignorar_fechas = st.checkbox(
            "Ignorar diferencias de fecha (comparar solo por importe)",
            value=True, key="bc_ig_fechas",
            help="Tildado: alcanza con que el importe coincida (las fechas pueden "
                 "estar mal cargadas en cualquiera de los dos lados).",
        )
    with o2:
        tolerancia = st.number_input(
            "Tolerancia de fecha (días)", 0, 60, 7, 1,
            disabled=ignorar_fechas, key="bc_tol",
            help="Solo aplica si NO ignorás las fechas.",
        )

    listo = (
        df_banco_raw is not None
        and contab is not None and not contab["errores"]
    )

    if st.button("🔍 Comparar", type="primary",
                 disabled=not listo, use_container_width=True, key="bc_btn"):
        try:
            df_c_full = MAP.construir(contab["df"], contab["mapeo"], contab["config"], contab["modo"])
            palabras_c = [p.strip() for p in (palabras_filtro_c or "").split(",") if p.strip()]
            df_c = MAP._excluir_por_detalle(df_c_full, palabras_c).reset_index(drop=True) if palabras_c else df_c_full
            df_b = _df_banco_a_conciliar(df_banco_raw, palabras_filtro)
            if df_b.empty:
                st.warning("Tras el filtro no quedaron movimientos en el banco.")
                return
            tol = 10**9 if ignorar_fechas else int(tolerancia)
            df_res = CON.conciliar(df_c, df_b, comparar_abs=True, tolerancia_dias=tol)
            st.session_state["bc_res"] = {
                "df_res": df_res, "df_c": df_c, "df_b": df_b,
                "filtradas_b": len(df_banco_raw) - len(df_b),
                "filtradas_c": len(df_c_full) - len(df_c),
            }
        except Exception as exc:  # noqa: BLE001
            st.error(f"Error al comparar: {exc}")
            st.exception(exc)

    if "bc_res" not in st.session_state:
        st.info("Cargá el PDF del banco y el Excel del mayor, mapeá las columnas "
                "y tocá **Comparar**.")
        return

    r = st.session_state["bc_res"]
    df_res, df_c, df_b = r["df_res"], r["df_c"], r["df_b"]

    if df_res.empty:
        st.warning("No se generaron resultados.")
        return

    msgs = []
    if r.get("filtradas_b"):
        msgs.append(f"{r['filtradas_b']} fila(s) del banco (impuestos / comisiones)")
    if r.get("filtradas_c"):
        msgs.append(f"{r['filtradas_c']} fila(s) de contabilidad (saldos / transportes)")
    if msgs:
        st.caption("Se filtraron: " + "; ".join(msgs) + ".")

    resumen = CON.resumen(df_res, df_c, df_b)
    st.divider()
    mostrar_resultado(df_res, resumen, etiqueta_b="Banco", key_prefix="bc_")

    st.divider()
    st.subheader("Exportar")
    excel = REP.exportar_excel(df_res, df_c, df_b, REP.resumen_df(resumen))
    st.download_button(
        "⬇️ Descargar Excel de diferencias",
        data=excel,
        file_name=f"banco_vs_contab_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="bc_dl",
    )


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
        up_csv = st.file_uploader("Archivo a procesar (.csv, .xlsx o .xls)",
                                  type=["csv", "xlsx", "xls"], key="afip_csv",
                                  help="Podés subir el CSV de ARCA (Portal IVA) "
                                       "o un CSV generado por Facturas por foto.")
    with c2:
        up_maestro = st.file_uploader("Excel maestro (Proveedores/Rubros) (.xlsx)",
                                      type=["xlsx"], key="afip_maestro")

    up_excluir = st.file_uploader(
        "🔍 CSV con lo que YA cargaste en JWIN — opcional, para excluir duplicados",
        type=["csv"], key="afip_excluir",
        help="Si ya importaste un CSV a JWIN y ahora querés subir SOLO lo que "
             "falta, subilo acá. La app compara y saca del CSV a descargar "
             "todos los comprobantes que ya estén en éste (match por CAE, o "
             "por CUIT + fecha + importe con tolerancia $0,05).",
    )

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

    # -------------------------------------------------------------------- #
    # Paso previo: si es el CSV crudo del Portal IVA, normalizarlo antes
    # de mapear rubros. Mueve Impuestos Internos y Otros Tributos al No
    # Gravado para que JWIN no descuadre las facturas (JWIN importa por
    # posición y no tiene esas dos columnas).
    # -------------------------------------------------------------------- #
    mes_periodo = None
    if AJ.es_csv_portal_iva(csv_bytes):
        norm = AJ.normalizar_portal_iva(csv_bytes)

        if norm["errores_header"]:
            st.error("El encabezado del CSV no coincide con el layout del Portal IVA. "
                     "ARCA capaz cambió el archivo — revisá los cambios antes de seguir.")
            for e in norm["errores_header"]:
                st.write(f"• {e}")
            return

        # Reporte visual: comprobantes / modificados / descartados.
        st.subheader("🧾 Normalización del CSV")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Comprobantes", norm["comprobantes"])
        c2.metric("Modificados (R+S → K)", len(norm["modificados"]))
        c3.metric("Filas descartadas", len(norm["descartados"]))
        c4.metric("Descuadres", len(norm["descuadres"]))

        if norm["modificados"]:
            with st.expander(f"Ver comprobantes modificados ({len(norm['modificados'])})",
                             expanded=True):
                st.dataframe(
                    pd.DataFrame(norm["modificados"]),
                    use_container_width=True, hide_index=True,
                    column_config={
                        "importe_movido": st.column_config.NumberColumn(
                            "Importe movido a No Gravado", format="%.2f"),
                    },
                )
                total_movido = sum(m["importe_movido"] for m in norm["modificados"])
                st.caption(f"Total movido a No Gravado: **${total_movido:,.2f}** "
                           "(antes se perdía porque JWIN no mapea esas columnas).")

        if norm["descartados"]:
            with st.expander(f"Ver filas descartadas ({len(norm['descartados'])})"):
                st.caption("Filas sin CUIT válido de 11 dígitos: totales, vacías o mal formadas.")
                st.dataframe(pd.DataFrame(norm["descartados"]),
                             use_container_width=True, hide_index=True)

        # Si hay descuadres, NO habilitamos la descarga: hay que revisar a mano.
        if norm["descuadres"]:
            st.error(
                f"❌ Hay {len(norm['descuadres'])} comprobante(s) que no cuadran "
                "(Total ≠ Neto + No Gravado + Exento + IVA + percepciones, "
                f"tolerancia $0,05). No se puede bajar el archivo hasta que se resuelvan."
            )
            st.dataframe(
                pd.DataFrame(norm["descuadres"]),
                use_container_width=True, hide_index=True,
                column_config={
                    "total": st.column_config.NumberColumn("Total cabecera", format="%.2f"),
                    "calculado": st.column_config.NumberColumn("Suma columnas", format="%.2f"),
                    "diferencia": st.column_config.NumberColumn("Diferencia", format="%.2f"),
                },
            )
            return

        # Todo OK: seguimos el flujo con el CSV normalizado.
        csv_bytes = norm["csv_bytes"]
        mes_periodo = norm["mes_periodo"]
        st.success(f"✅ CSV normalizado sin descuadres. Período detectado: **{mes_periodo}**")

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

    # ----------------------------------------------------------------- #
    # Excluir los comprobantes que YA se cargaron a JWIN (dedup opcional)
    # ----------------------------------------------------------------- #
    excluidos_por_dedup = 0
    if up_excluir is not None:
        try:
            filas_excluir_indices = _detectar_duplicados_excluir(
                encab, filas, up_excluir.getvalue())
        except Exception as exc:  # noqa: BLE001
            st.error(f"No se pudo leer el CSV a excluir: {exc}")
        else:
            if filas_excluir_indices:
                excluidos_por_dedup = len(filas_excluir_indices)
                filas = [f for i, f in enumerate(filas)
                         if i not in filas_excluir_indices]
                st.info(f"🔍 Se excluyeron **{excluidos_por_dedup}** comprobante(s) "
                        "que ya estaban en el CSV a excluir. Solo quedan los que "
                        "faltan cargar en JWIN.")
                # Recalcular stats (asignados/sin_rubro/comprobantes) con la lista filtrada
                asignados_nuevo = sum(1 for f in filas if str(f[-1]).strip())
                stats["comprobantes"] = len(filas)
                stats["asignados"] = asignados_nuevo
                stats["sin_rubro"] = len(filas) - asignados_nuevo
            else:
                st.success("No hay comprobantes duplicados con el CSV a excluir.")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Comprobantes", stats["comprobantes"])
    m2.metric("Con rubro", stats["asignados"])
    m3.metric("Sin rubro", stats["sin_rubro"])
    m4.metric("Excluidos (dedup)", excluidos_por_dedup)

    if stats["sin_rubro"] == 0 and stats["comprobantes"] > 0:
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
    sufijo_mes = mes_periodo or datetime.now().strftime("%m-%Y")
    cda.download_button(
        "⬇️ CSV para importar a JWIN",
        data=AJ.construir_csv(encab, filas),
        file_name=f"Importacion_JWIN_{sufijo_mes}.csv",
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
# Sección: Rango — Compras (Paradigma) → Ordenado
# --------------------------------------------------------------------------- #

def seccion_rango():
    st.title("📦 Rango — Compras (Paradigma) → Ordenado")
    st.caption(
        "Subí el **Listado** de Libro IVA Compras del sistema de Rango + el **maestro "
        "de proveedores**. La app lo reordena, le pone el **rubro contable** (por CUIT), "
        "calcula NG+NO GR e IVA TOTAL, y detecta el **mes de proceso** del encabezado."
    )

    c1, c2 = st.columns(2)
    with c1:
        up_list = st.file_uploader("Listado de compras (Paradigma) (.xlsx)",
                                   type=["xlsx", "xls"], key="rango_list")
    with c2:
        up_maestro = st.file_uploader("Maestro de proveedores - Rango (.xlsx)",
                                      type=["xlsx"], key="rango_maestro")

    with st.expander("¿Necesitás la plantilla del maestro de proveedores?"):
        st.caption("Trae los rubros contables cargados y la hoja Proveedores vacía.")
        st.download_button(
            "⬇️ Descargar plantilla (rubros-proveedores)",
            data=RC.construir_plantilla_maestro(),
            file_name="Maestro proveedores - Rango (plantilla).xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if up_list is None or up_maestro is None:
        st.info("Subí el Listado del mes y el maestro de proveedores de Rango.")
        return

    list_bytes = up_list.getvalue()
    maestro_bytes = up_maestro.getvalue()
    try:
        filas, desconocidos, stats, proc_det, rubros, rubro_codigo = RC.procesar(list_bytes, maestro_bytes)
    except Exception as exc:  # noqa: BLE001
        st.error(f"No pude procesar: {exc}")
        return

    # Mes de proceso: detectado, con opción de corregir.
    import datetime as _dt
    cmes, cano = st.columns(2)
    mes_def = proc_det.month if proc_det else _dt.date.today().month
    ano_def = proc_det.year if proc_det else _dt.date.today().year
    meses_nom = ["", "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
                 "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    mes = cmes.selectbox("Mes de proceso", list(range(1, 13)),
                         index=mes_def - 1, format_func=lambda m: meses_nom[m])
    ano = cano.number_input("Año de proceso", 2020, 2100, ano_def, 1)
    proceso = _dt.date(int(ano), int(mes), 1)

    # Reprocesar con proveedores nuevos cargados y el proceso elegido.
    nuevos = []
    if desconocidos:
        st.warning(
            f"Hay {len(desconocidos)} proveedor(es) sin rubro. La app ya te sugiere uno "
            "(revisá y corregí si hace falta), después podés bajar tu maestro actualizado."
        )
        # Sugerencia automática: aprende del maestro + palabras clave.
        aprendido = RC.aprender_rubros(RC.leer_maestro_pares(maestro_bytes))
        opciones = [""] + list(rubros)
        base = pd.DataFrame([
            {"CUIT": c, "Proveedor": n,
             "Rubro": RC.sugerir_rubro(n, aprendido, set(rubros))}
            for c, n in sorted(desconocidos.items())
        ])
        editado = st.data_editor(
            base, hide_index=True, use_container_width=True, key="rango_editor",
            column_config={
                "CUIT": st.column_config.TextColumn(disabled=True),
                "Proveedor": st.column_config.TextColumn(disabled=True),
                "Rubro": st.column_config.SelectboxColumn("Rubro contable", options=opciones),
            },
        )
        extra = {}
        for _, row in editado.iterrows():
            sel = str(row["Rubro"]).strip()
            if sel:
                cuit = RC._norm_cuit(row["CUIT"])
                extra[cuit] = (sel, "")
                nuevos.append((cuit, row["Proveedor"], sel, ""))
        # Reprocesar con el mes elegido y aplicar las asignaciones manuales.
        filas, desconocidos, stats, proc_det, rubros, rubro_codigo = RC.procesar(
            list_bytes, maestro_bytes, proceso)
        if extra:
            for f in filas:
                if f["rubro contable"] == "" and f["CUIT"] in extra:
                    f["rubro contable"] = extra[f["CUIT"]][0]
            stats["asignados"] = sum(1 for f in filas if f["rubro contable"])
            stats["sin_rubro"] = len(filas) - stats["asignados"]
    else:
        filas, desconocidos, stats, proc_det, rubros, rubro_codigo = RC.procesar(
            list_bytes, maestro_bytes, proceso)

    m1, m2, m3 = st.columns(3)
    m1.metric("Comprobantes", stats["comprobantes"])
    m2.metric("Con rubro", stats["asignados"])
    m3.metric("Sin rubro", stats["sin_rubro"])
    if stats["sin_rubro"] == 0:
        st.success("Todos los comprobantes quedaron con su rubro contable. 🎉")

    # --- Control: lo procesado vs los totales de la hoja 'Comprobante' de Paradigma ---
    control_sis = RC.leer_control(list_bytes)
    if control_sis:
        detalle, cuadra = RC.control_totales(filas, control_sis)
        st.divider()
        st.subheader("✅ Control contra Paradigma (hoja Comprobante)")
        if cuadra:
            st.success("Cuadra: los totales procesados coinciden con los del sistema. 👌")
        else:
            st.error("¡No cuadra! Hay diferencias contra los totales de Paradigma. Revisá:")
        st.dataframe(
            pd.DataFrame(detalle), use_container_width=True, hide_index=True,
            column_config={c: st.column_config.NumberColumn(format="%.2f")
                           for c in ["Sistema (Paradigma)", "Procesado", "Diferencia"]},
        )
    else:
        st.caption("ℹ️ No encontré la hoja 'Comprobante' en el Listado, así que no pude "
                   "hacer el control automático de totales.")

    st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)

    st.caption("El Excel trae: Ordenado (Tabla) · iva+retenciones · TD-RUBROS-ASIENTOS "
               "(el asiento con gastos bancarios + tarjetas a cobrar + IVA + retenciones + 'a caja') "
               "· TARJETAS · RESUMEN.")
    st.download_button(
        "⬇️ Descargar libro Rango (Excel con todas las hojas)",
        data=RC.construir_libro(filas, maestro_bytes, rubro_codigo),
        file_name=f"ASIENTOS COMPRAS Rango {proceso:%m-%Y}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    if nuevos:
        st.info("Cargaste proveedores nuevos: bajá tu maestro actualizado y guardalo.")
        st.download_button(
            "⬇️ Maestro Rango ACTUALIZADO",
            data=RC.construir_maestro_actualizado(maestro_bytes, nuevos),
            file_name=up_maestro.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


# --------------------------------------------------------------------------- #
# Sección: Monotributo — recategorización
# --------------------------------------------------------------------------- #

def seccion_monotributo():
    st.title("📊 Monotributo — Recategorización")
    st.caption(
        "Subí la planilla de facturación del cliente. La app calcula el acumulado de "
        "los últimos 12 meses, la categoría que le corresponde, cuánto puede facturar "
        "para mantenerse y cuánto le falta para pasarse."
    )

    escala = MT.cargar_escala()
    with st.expander(f"Escala vigente ({escala.get('vigencia','')}) — actualizable"):
        st.caption("Para actualizar cada semestre, reemplazá src/escala_monotributo.json con los topes vigentes de ARCA.")
        st.dataframe(pd.DataFrame(escala["servicios"]).rename(
            columns={"cat": "Cat.", "tope": "Tope ingresos 12m", "cuota": "Cuota mensual"}),
            use_container_width=True, hide_index=True,
            column_config={"Tope ingresos 12m": st.column_config.NumberColumn(format="%.2f"),
                           "Cuota mensual": st.column_config.NumberColumn(format="%.2f")})

    archivo = st.file_uploader("Planilla de facturación (.xlsx)", type=["xlsx"], key="mt_file")
    if archivo is None:
        st.info("Subí la planilla de facturación del monotributista.")
        return

    try:
        serie = MT.leer_facturacion(archivo.getvalue())
    except Exception as exc:  # noqa: BLE001
        st.error(f"No pude leer la planilla: {exc}")
        return
    if not serie:
        st.warning("No encontré la facturación mensual (esperaba una hoja 'ACUMULADOS' "
                   "o hojas mensuales con una columna 'MONTO').")
        return

    meses = MT.meses_disponibles(serie)
    ini_a, ini_m = meses[0]
    ult_a, ult_m = meses[-1]
    c1, c2, c3 = st.columns(3)
    actividad = c1.selectbox("Actividad", ["servicios", "bienes"],
                             format_func=lambda a: "Servicios / locaciones" if a == "servicios"
                             else "Venta de cosas muebles")
    mes = c2.selectbox("Mes hasta (cierre del período)", list(range(1, 13)),
                       index=ult_m - 1, format_func=lambda m: MT._MESNOM[m])
    anio = c3.number_input("Año", 2020, 2100, ult_a, 1)

    masde12 = st.radio("¿Lleva más de 12 meses inscripto?", ["Sí", "No"], horizontal=True,
                       help="Si lleva menos de 12 meses, ARCA anualiza: ingresos ÷ meses × 12.")
    nuevo = masde12 == "No"
    if nuevo:
        cci, ccj = st.columns(2)
        mes_ini = cci.selectbox("¿Cuándo arrancó? Mes", list(range(1, 13)),
                                index=ini_m - 1, format_func=lambda m: MT._MESNOM[m], key="mt_ini_mes")
        anio_ini = ccj.number_input("Año de inicio", 2020, 2100, ini_a, 1, key="mt_ini_anio")
        real, n_meses, detalle = MT.acumulado_rango(serie, int(anio_ini), int(mes_ini), int(anio), int(mes))
        if n_meses < 6:
            st.warning(
                f"Lleva **{n_meses}** mes(es) de actividad (menos de 6) → **NO corresponde "
                f"recategorizar** todavía. Mantiene la categoría con la que se inscribió. "
                f"(Facturado hasta ahora: {formato_ar(real)}.)"
            )
            with st.expander("Ver los meses considerados"):
                st.dataframe(
                    pd.DataFrame([{"Período": f"{MT._MESNOM[m]} {y}", "Facturado": v} for y, m, v in detalle]),
                    use_container_width=True, hide_index=True,
                    column_config={"Facturado": st.column_config.NumberColumn(format="%.2f")})
            return
        acumulado = MT.anualizar(real, n_meses)
        st.info(f"Anualización ({n_meses} meses, supera los 6): facturó **{formato_ar(real)}** → "
                f"proyectado a 12 meses = **{formato_ar(acumulado)}**.")
    else:
        acumulado, detalle = MT.acumulado_12m(serie, int(anio), int(mes))

    a = MT.analizar(acumulado, actividad, escala)

    st.divider()
    m1, m2, m3 = st.columns(3)
    m1.metric("Anualizado (proyectado)" if nuevo else "Facturado últimos 12 meses",
              formato_ar(acumulado))
    m2.metric("Categoría que le corresponde", a["categoria"])
    m3.metric("Cuota mensual", formato_ar(a["cuota"]) if a["cuota"] else "—")

    if a["categoria"] == "EXCLUIDO":
        st.error(f"⚠️ Supera el tope máximo (categoría K) por {formato_ar(a['excede_por'])}. "
                 "Queda EXCLUIDO del monotributo → pasa al Régimen General.")
    else:
        st.success(
            f"✅ Categoría **{a['categoria']}**. Para **mantenerse** puede facturar hasta "
            f"**{formato_ar(a['margen_mantenerse'])}** más (en los próximos meses del período). "
            f"Si factura más que eso, **pasa a {a['siguiente']}**."
        )
        cc1, cc2 = st.columns(2)
        cc1.metric(f"Margen para mantenerse en {a['categoria']}", formato_ar(a["margen_mantenerse"]))
        cc2.metric(f"Si supera eso → pasa a {a['siguiente']}",
                   f"tope {formato_ar(a['tope_siguiente'])}" if a["tope_siguiente"] else "EXCLUIDO")

    with st.expander("Ver los 12 meses considerados"):
        st.dataframe(
            pd.DataFrame([{"Período": f"{MT._MESNOM[m]} {y}", "Facturado": v} for y, m, v in detalle]),
            use_container_width=True, hide_index=True,
            column_config={"Facturado": st.column_config.NumberColumn(format="%.2f")})


# --------------------------------------------------------------------------- #
# Tareas y checklist (compartido entre computadoras vía Supabase)
# --------------------------------------------------------------------------- #

def _miembros_equipo() -> list[str]:
    """Nombres del equipo para el selector '¿Quién sos?'.

    Se configuran en secrets:  [equipo]  miembros = ["Juani", "..."]
    Si no están, usa una lista por defecto editable.
    """
    try:
        miembros = list(st.secrets["equipo"]["miembros"])
        if miembros:
            return miembros
    except Exception:
        pass
    return ["Juani"]


def _quien_soy() -> str | None:
    """Selector de identidad: quién está usando la app ahora."""
    opciones = ["— elegí tu nombre —"] + _miembros_equipo()
    actual = st.session_state.get("quien")
    idx = opciones.index(actual) if actual in opciones else 0
    sel = st.selectbox("👤 ¿Quién sos?", opciones, index=idx, key="sel_quien")
    st.session_state["quien"] = sel
    return None if sel == opciones[0] else sel


def _fila_tarea(TDB, t: dict, quien: str | None) -> None:
    """Una fila de la lista de tareas pendientes/hechas."""
    c1, c2, c3 = st.columns([0.08, 0.77, 0.15])
    nueva = c1.checkbox(
        "hecha", value=t["hecha"], key=f"chk_t_{t['id']}_{int(t['hecha'])}",
        label_visibility="collapsed",
    )
    if nueva != t["hecha"]:
        if not quien:
            st.warning("Elegí tu nombre arriba para tachar tareas.")
        else:
            TDB.marcar_tarea(t["id"], nueva, quien)
            st.rerun()

    texto = t["titulo"]
    if t.get("monto"):
        texto += f"  ·  **${formato_ar(t['monto'])}**"
    c2.markdown(f"~~{texto}~~" if t["hecha"] else texto)

    detalle = []
    if t.get("nota"):
        detalle.append(t["nota"])
    if t.get("creada_por"):
        detalle.append(f"📌 {t['creada_por']}")
    if t["hecha"] and t.get("hecha_por"):
        detalle.append(f"✔️ {t['hecha_por']}")
    if detalle:
        c2.caption("  ·  ".join(detalle))

    if c3.button("🗑️", key=f"del_t_{t['id']}", help="Borrar tarea"):
        TDB.borrar_tarea(t["id"])
        st.rerun()


def _ui_tareas(TDB, quien: str | None) -> None:
    with st.form("nueva_tarea", clear_on_submit=True):
        st.markdown("**➕ Nueva tarea**")
        c1, c2 = st.columns([3, 1])
        titulo = c1.text_input("¿Qué hay que hacer?", placeholder="Ej: hacer FC a Juani")
        monto = c2.number_input("Monto $ (opcional)", min_value=0.0, value=0.0, step=1000.0)
        nota = st.text_input("Nota (opcional)", placeholder="Aclaración, detalle…")
        agregar = st.form_submit_button("Agregar", type="primary", use_container_width=True)
    if agregar:
        if not titulo.strip():
            st.warning("Escribí qué hay que hacer.")
        elif not quien:
            st.warning("Primero elegí tu nombre arriba.")
        else:
            TDB.crear_tarea(titulo.strip(), nota.strip(), monto or None, quien)
            st.rerun()

    tareas = TDB.listar_tareas()
    pendientes = [t for t in tareas if not t["hecha"]]
    hechas = [t for t in tareas if t["hecha"]]

    st.markdown(f"### Pendientes ({len(pendientes)})")
    if not pendientes:
        st.caption("🎉 No hay tareas pendientes.")
    for t in pendientes:
        _fila_tarea(TDB, t, quien)

    if hechas:
        with st.expander(f"✔️ Hechas ({len(hechas)})"):
            for t in hechas:
                _fila_tarea(TDB, t, quien)


def _ui_rutinas(TDB, quien: str | None) -> None:
    GRUPOS = {"diaria": "🗓️ Diarias", "semanal": "📅 Semanales", "mensual": "🈷️ Mensuales"}
    OPCIONES = {"diaria": "Todos los días", "semanal": "Cada semana", "mensual": "Cada mes"}

    with st.expander("➕ Agregar tarea rutinaria"):
        with st.form("nueva_rutina", clear_on_submit=True):
            titulo = st.text_input("Tarea rutinaria", placeholder="Ej: Cargar movimientos del banco")
            frec = st.selectbox(
                "¿Cada cuánto?", list(OPCIONES), format_func=lambda f: OPCIONES[f])
            if st.form_submit_button("Agregar rutina", type="primary"):
                if titulo.strip():
                    TDB.crear_rutina(titulo.strip(), frec)
                    st.rerun()

    rutinas = TDB.listar_rutinas()
    if not rutinas:
        st.caption("Todavía no hay tareas rutinarias. Agregá la primera con el botón de arriba.")
        return

    periodos = {r["frecuencia"]: TDB.periodo_actual(r["frecuencia"]) for r in rutinas}
    estados = TDB.estados_periodo(list(periodos.values()))

    for frec, etiqueta in GRUPOS.items():
        grupo = [r for r in rutinas if r["frecuencia"] == frec]
        if not grupo:
            continue
        per = periodos[frec]
        hechas = sum((r["id"], per) in estados for r in grupo)
        st.markdown(f"### {etiqueta}  ·  {hechas}/{len(grupo)}")
        for r in grupo:
            fila = estados.get((r["id"], per))
            hecha = fila is not None
            c1, c2, c3 = st.columns([0.08, 0.77, 0.15])
            nueva = c1.checkbox(
                "hecha", value=hecha, key=f"chk_r_{r['id']}_{per}_{int(hecha)}",
                label_visibility="collapsed",
            )
            if nueva != hecha:
                if not quien:
                    st.warning("Elegí tu nombre arriba para marcar rutinas.")
                else:
                    TDB.marcar_rutina(r["id"], per, nueva, quien)
                    st.rerun()
            c2.markdown(f"~~{r['titulo']}~~" if hecha else r["titulo"])
            if hecha and fila.get("hecha_por"):
                c2.caption(f"✔️ {fila['hecha_por']}")
            if c3.button("🗑️", key=f"del_r_{r['id']}", help="Quitar rutina"):
                TDB.borrar_rutina(r["id"])
                st.rerun()


# --------------------------------------------------------------------------- #
# Sección: Clientes del estudio
# --------------------------------------------------------------------------- #

def _form_cliente(prefijo: str, valores: dict | None = None) -> dict | None:
    """Formulario de carga/edición de cliente. Devuelve los datos si se valida y submitea.

    Los campos cambian según el tipo de contribuyente:
      - Monotributo: categoría (A-K)
      - Sociedad: día y mes de cierre de balance
      - RI / Exento / Otro: sin campos extra
    """
    from src import clientes_db as CDB

    v = valores or {}

    # El selector de tipo va FUERA del form para que los campos condicionales
    # debajo reaccionen en tiempo real (los widgets dentro de st.form no
    # disparan rerun hasta que se submitea).
    tipo_inicial = v.get("tipo") if v.get("tipo") in CDB.TIPOS else "Monotributo"
    tipo = st.selectbox(
        "Tipo de contribuyente",
        CDB.TIPOS,
        index=CDB.TIPOS.index(tipo_inicial),
        key=f"{prefijo}_tipo",
    )

    with st.form(key=f"{prefijo}_form", clear_on_submit=(valores is None)):
        c1, c2 = st.columns(2)
        with c1:
            razon = st.text_input("Razón social *", value=v.get("razon_social", "") or "")
            cuit = st.text_input("CUIT *", value=v.get("cuit", "") or "",
                                 help="11 dígitos, con o sin guiones.")
            email = st.text_input("Email", value=v.get("email", "") or "")
            telefono = st.text_input("Teléfono", value=v.get("telefono", "") or "")
        with c2:
            domicilio = st.text_input("Domicilio", value=v.get("domicilio", "") or "")
            cat = None
            dia = None
            mes = None
            if tipo == "Monotributo":
                cat_opciones = [""] + CDB.CATEGORIAS_MONO
                cat = st.selectbox(
                    "Categoría monotributo",
                    cat_opciones,
                    index=cat_opciones.index(v.get("categoria_mono") or ""),
                )
            elif tipo == "Sociedad":
                ca, cb = st.columns(2)
                with ca:
                    dia = st.number_input(
                        "Día de cierre", 0, 31, int(v.get("dia_cierre") or 0),
                        help="Día del mes en que cierra el balance. 0 = sin definir.",
                    )
                with cb:
                    mes = st.number_input(
                        "Mes de cierre", 0, 12, int(v.get("mes_cierre") or 0),
                        help="Mes del año en que cierra el balance. 0 = sin definir.",
                    )
            else:
                st.caption("ℹ️ Sin campos extra para este tipo de contribuyente.")

        obs = st.text_area("Observaciones", value=v.get("observaciones", "") or "", height=80)

        guardar = st.form_submit_button("💾 Guardar" if valores else "➕ Crear cliente",
                                         type="primary", use_container_width=True)
        if not guardar:
            return None

        errores = []
        if not razon.strip():
            errores.append("Falta la razón social.")
        if not CDB.validar_cuit(cuit):
            errores.append("CUIT inválido (deben ser 11 dígitos con dígito verificador correcto).")
        for e in errores:
            st.error(e)
        if errores:
            return None

        return {
            "razon_social": razon.strip(),
            "cuit": "".join(c for c in cuit if c.isdigit()),
            "tipo": tipo,
            "categoria_mono": cat or None,
            "email": email or None,
            "telefono": telefono or None,
            "domicilio": domicilio or None,
            "dia_cierre": int(dia) if dia else None,
            "mes_cierre": int(mes) if mes else None,
            "observaciones": obs or None,
        }


def _panel_arca_cliente(cliente: dict):
    """Botón 'Actualizar desde ARCA' + vista de los datos del padrón ya guardados."""
    from src import clientes_db as CDB
    from src.arca import certs as ACERTS, padron as APADRON, wsaa as AWSAA

    st.subheader("🔄 Padrón ARCA")
    pdata = cliente.get("padron_data") or {}
    pact = cliente.get("padron_actualizado_en")
    if pdata and pact:
        st.caption(f"Última consulta: {pact}")
    elif pdata:
        st.caption("Datos cargados desde una consulta anterior.")
    else:
        st.caption("Todavía no se consultó el padrón para este cliente.")

    if st.button("🔄 Consultar / actualizar desde ARCA",
                 key=f"arca_{cliente['id']}", use_container_width=True, type="primary"):
        try:
            with st.spinner("Conectando con ARCA…"):
                cert_pem, key_pem, cuit_titular = ACERTS.cargar()
                cred = AWSAA.obtener_credenciales(
                    servicio=APADRON.SERVICIO,
                    cert_pem=cert_pem, key_pem=key_pem,
                    cuit_titular=cuit_titular,
                )
                datos = APADRON.consultar(
                    token=cred.token, sign=cred.sign,
                    cuit_titular=cuit_titular,
                    cuit_a_consultar=cliente["cuit"],
                )
            CDB.guardar_padron(cliente["id"], datos)
            st.success("Padrón actualizado.")
            st.rerun()
        except FileNotFoundError as exc:
            st.error(str(exc))
            st.caption("Tip: subí el `.crt` y la `.key` a `certificados/` (local) "
                       "o configurá `st.secrets['arca']` en Streamlit Cloud.")
        except Exception as exc:  # noqa: BLE001
            st.error(f"No se pudo consultar ARCA: {exc}")
            st.exception(exc)

    if not pdata:
        return

    # Mostrar los datos del padrón en formato amigable
    nombre = pdata.get("razon_social") or " ".join(
        x for x in (pdata.get("nombre"), pdata.get("apellido")) if x
    )
    c1, c2, c3 = st.columns(3)
    c1.metric("Estado clave", pdata.get("estado_clave") or "—")
    c2.metric("Tipo persona", pdata.get("tipo_persona") or "—")
    c3.metric("Tipo clave", pdata.get("tipo_clave") or "—")
    if nombre:
        st.markdown(f"**{nombre}**")

    dom = pdata.get("domicilio") or {}
    if dom:
        partes = [dom.get("direccion"), dom.get("localidad"),
                  dom.get("provincia"), dom.get("codigo_postal")]
        st.caption("📍 " + ", ".join(p for p in partes if p))

    if pdata.get("categorias"):
        st.markdown("**Monotributo / categorías**")
        st.dataframe(pd.DataFrame(pdata["categorias"]),
                     use_container_width=True, hide_index=True)

    if pdata.get("impuestos"):
        st.markdown("**Impuestos**")
        st.dataframe(pd.DataFrame(pdata["impuestos"]),
                     use_container_width=True, hide_index=True)

    if pdata.get("actividades"):
        st.markdown("**Actividades económicas**")
        st.dataframe(pd.DataFrame(pdata["actividades"]),
                     use_container_width=True, hide_index=True)

    with st.expander("🧾 Ver respuesta completa (JSON)"):
        st.json(pdata)


def seccion_clientes():
    from src import clientes_db as CDB

    st.title("👥 Clientes")

    if not CDB.hay_conexion():
        st.warning(
            "Esta sección usa Supabase y todavía no están configuradas las credenciales. "
            "Cargalas en `st.secrets['supabase']` (o en Streamlit Cloud → Settings → Secrets) "
            "y volvé a entrar."
        )
        return

    try:
        incluir_inactivos = st.toggle("Mostrar también archivados", value=False, key="cli_inactivos")
        clientes = CDB.listar(incluir_inactivos=incluir_inactivos)
    except Exception as exc:  # noqa: BLE001
        st.error("No se pudo leer la tabla `clientes` de Supabase. "
                 "Asegurate de haber corrido `sql/clientes_setup.sql` en el SQL Editor del proyecto.")
        st.exception(exc)
        return

    # --- Lista
    if clientes:
        df = pd.DataFrame(clientes)
        buscar = st.text_input("🔎 Buscar", placeholder="razón social, CUIT, email…",
                               key="cli_buscar")
        if buscar:
            t = buscar.strip().lower()
            df = df[df.apply(lambda r: t in " ".join(
                str(r.get(c, "") or "").lower() for c in ("razon_social", "cuit", "email", "telefono")
            ), axis=1)]

        df_view = df.copy()
        df_view["CUIT"] = df_view["cuit"].apply(CDB.formato_cuit)
        df_view["Activo"] = df_view["activo"].map({True: "✓", False: "—"})
        df_view = df_view.rename(columns={
            "razon_social": "Razón social",
            "tipo": "Tipo",
            "categoria_mono": "Cat.",
            "email": "Email",
            "telefono": "Teléfono",
        })
        cols_mostrar = ["Razón social", "CUIT", "Tipo", "Cat.", "Email", "Teléfono", "Activo"]
        st.dataframe(df_view[cols_mostrar], use_container_width=True, hide_index=True)
        st.caption(f"Total: {len(df)} cliente(s).")
    else:
        st.info("Todavía no hay clientes cargados. Usá **'Nuevo cliente'** abajo para empezar.")

    st.divider()
    tab_nuevo, tab_editar = st.tabs(["➕ Nuevo cliente", "✏️ Editar / archivar"])

    with tab_nuevo:
        datos = _form_cliente("nuevo")
        if datos:
            try:
                existente = CDB.buscar_por_cuit(datos["cuit"])
                if existente:
                    st.error(f"Ya existe un cliente con CUIT {CDB.formato_cuit(datos['cuit'])}: "
                             f"**{existente['razon_social']}**.")
                else:
                    CDB.crear(datos)
                    st.success(f"Cliente creado: {datos['razon_social']}.")
                    st.rerun()
            except Exception as exc:  # noqa: BLE001
                st.error(f"No se pudo crear: {exc}")

    with tab_editar:
        if not clientes:
            st.info("Cargá un cliente primero para poder editarlo.")
        else:
            opciones = {f"{c['razon_social']} ({CDB.formato_cuit(c['cuit'])})": c["id"] for c in clientes}
            label = st.selectbox("Cliente a editar", list(opciones.keys()), key="cli_editar_sel")
            cliente_id = opciones[label]
            cliente = next(c for c in clientes if c["id"] == cliente_id)

            datos = _form_cliente(f"edit_{cliente_id}", valores=cliente)
            if datos:
                try:
                    CDB.actualizar(cliente_id, datos)
                    st.success("Cliente actualizado.")
                    st.rerun()
                except Exception as exc:  # noqa: BLE001
                    st.error(f"No se pudo actualizar: {exc}")

            st.divider()
            _panel_arca_cliente(cliente)

            st.divider()
            c1, c2 = st.columns(2)
            with c1:
                if cliente["activo"]:
                    if st.button("📦 Archivar", key=f"arch_{cliente_id}", use_container_width=True,
                                  help="Saca al cliente de la lista activa pero conserva sus datos."):
                        CDB.archivar(cliente_id, True)
                        st.success("Archivado.")
                        st.rerun()
                else:
                    if st.button("♻️ Reactivar", key=f"react_{cliente_id}", use_container_width=True):
                        CDB.archivar(cliente_id, False)
                        st.success("Reactivado.")
                        st.rerun()
            with c2:
                with st.expander("🗑️ Borrar definitivamente"):
                    st.caption("Esto borra el registro de la base. No se puede deshacer. "
                               "Usá **Archivar** salvo que sea un error de carga.")
                    if st.button("Confirmar borrar", type="primary",
                                 key=f"del_{cliente_id}", use_container_width=True):
                        CDB.borrar(cliente_id)
                        st.success("Borrado.")
                        st.rerun()


# --------------------------------------------------------------------------- #
# Sección: Posición IVA
# --------------------------------------------------------------------------- #

def _bloque_carga_iva(label: str, key: str):
    """Sube un archivo 'Mis Comprobantes' de AFIP y devuelve el DataFrame normalizado."""
    st.subheader(label)
    archivo = st.file_uploader(
        "Excel descargado de AFIP — 'Mis Comprobantes'",
        type=["xlsx", "xls"],
        key=f"iva_{key}_file",
    )
    if archivo is None:
        st.info("Esperando archivo…")
        return None
    try:
        df = IVA.cargar_archivo_afip(archivo.getvalue())
    except Exception as exc:  # noqa: BLE001
        st.error(f"No se pudo leer el archivo: {exc}")
        return None

    n_nc = int(df["es_nc"].sum())
    st.success(f"✓ {len(df)} comprobantes ({n_nc} notas de crédito).")
    with st.expander("Vista previa", expanded=False):
        st.dataframe(
            df[["fecha", "tipo", "punto_venta", "numero", "denominacion", "total_iva", "imp_total"]].head(20),
            use_container_width=True, hide_index=True,
            column_config={
                "fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
                "total_iva": st.column_config.NumberColumn("Total IVA", format="%.2f"),
                "imp_total": st.column_config.NumberColumn("Imp. Total", format="%.2f"),
            },
        )
    return df


def seccion_iva():
    st.title("📊 Posición IVA")
    st.caption("Subí los Excel de 'Mis Comprobantes' Emitidos y Recibidos descargados de AFIP. "
               "Las **notas de crédito** se descuentan automáticamente. Indicá el saldo a favor "
               "del período anterior (si lo hay) y obtené la posición.")

    c1, c2 = st.columns(2)
    with c1:
        df_emit = _bloque_carga_iva("1) Comprobantes emitidos", "emit")
    with c2:
        df_recib = _bloque_carga_iva("2) Comprobantes recibidos", "recib")

    if df_emit is None or df_recib is None:
        st.info("Subí los dos archivos para calcular la posición.")
        return

    # Selector de período si hay más de uno presente en cualquiera de los dos archivos.
    periodos = sorted(set(IVA.periodos_presentes(df_emit)) | set(IVA.periodos_presentes(df_recib)))
    if len(periodos) > 1:
        TODOS = "Todos los períodos del archivo"
        opcion = st.selectbox(
            "📅 Período a calcular",
            [TODOS] + periodos,
            index=len(periodos),  # default: el último período
            key="iva_periodo",
            help="Los archivos traen más de un mes — elegí cuál querés calcular.",
        )
        periodo_sel = None if opcion == TODOS else opcion
    elif len(periodos) == 1:
        periodo_sel = periodos[0]
        st.caption(f"Período detectado: **{periodo_sel}**")
    else:
        periodo_sel = None

    df_emit_per = IVA.filtrar_por_periodo(df_emit, periodo_sel)
    df_recib_per = IVA.filtrar_por_periodo(df_recib, periodo_sel)

    st.divider()
    st.subheader("3) Saldos del período anterior")
    s1, s2 = st.columns(2)
    with s1:
        saldo_tec = st.number_input(
            "Saldo SJ anterior técnico (a favor)",
            value=0.0, step=1000.0, format="%.2f", key="iva_saldo_tec",
            help="Saldo a favor TÉCNICO arrastrado del período anterior. Se RESTA a la posición.",
        )
    with s2:
        saldo_ld = st.number_input(
            "Saldo SJ anterior LD / retenciones",
            value=0.0, step=1000.0, format="%.2f", key="iva_saldo_ld",
            help="Saldo a favor de LIBRE DISPONIBILIDAD (retenciones, percepciones) del período "
                 "anterior. Se RESTA a la posición.",
        )

    pos = IVA.calcular_posicion(df_emit_per, df_recib_per, saldo_tec, saldo_ld)
    df_lado = pos["debito_fiscal"]
    cf_lado = pos["credito_fiscal"]

    st.divider()
    st.subheader("Resultado")

    m1, m2, m3 = st.columns(3)
    m1.metric("IVA Débito Fiscal (a)", formato_ar(df_lado["neto"]))
    m2.metric("IVA Crédito Fiscal (b)", formato_ar(cf_lado["neto"]))
    if pos["posicion"] > 0:
        m3.metric("📤 A pagar", formato_ar(pos["posicion"]))
    elif pos["posicion"] < 0:
        m3.metric("✅ Saldo a favor (sigue al próximo)", formato_ar(abs(pos["posicion"])))
    else:
        m3.metric("Posición", "0,00")

    # Detalle del cálculo
    detalle = pd.DataFrame([
        ["IVA Débito Fiscal — Facturas + ND emitidos",   df_lado["fc"],   df_lado["cant_fact_nd"]],
        ["IVA Débito Fiscal — Notas de Crédito emitidas", -df_lado["nc"],  df_lado["cant_nc"]],
        ["IVA DF neto (a)",                               df_lado["neto"], df_lado["cant_fact_nd"] + df_lado["cant_nc"]],
        ["", None, None],
        ["IVA Crédito Fiscal — Facturas + ND recibidos",  cf_lado["fc"],   cf_lado["cant_fact_nd"]],
        ["IVA Crédito Fiscal — Notas de Crédito recibidas", -cf_lado["nc"], cf_lado["cant_nc"]],
        ["IVA CF neto (b)",                               cf_lado["neto"], cf_lado["cant_fact_nd"] + cf_lado["cant_nc"]],
        ["", None, None],
        ["(-) Saldo SJ anterior técnico (c)",            -pos["saldo_anterior_tecnico"], None],
        ["(-) Saldo SJ anterior LD / retenciones (d)",   -pos["saldo_anterior_ld"], None],
        ["", None, None],
        ["POSICIÓN = a - b - c - d",                      pos["posicion"], None],
    ], columns=["Concepto", "Importe", "Cantidad"])

    st.dataframe(
        detalle, use_container_width=True, hide_index=True,
        column_config={
            "Importe": st.column_config.NumberColumn(format="%.2f"),
            "Cantidad": st.column_config.NumberColumn(format="%d"),
        },
    )

    st.divider()
    st.subheader("Exportar")
    excel = IVA.exportar_excel(df_emit_per, df_recib_per, pos, periodo=periodo_sel)
    nombre = f"posicion_iva_{periodo_sel or 'completo'}_{datetime.now():%Y%m%d_%H%M}.xlsx"
    st.download_button(
        "⬇️ Descargar Excel (cálculo + detalle de comprobantes)",
        data=excel,
        file_name=nombre,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="iva_dl",
    )


# --------------------------------------------------------------------------- #
# Sección: Facturas por foto (lectura de QR de AFIP)
# --------------------------------------------------------------------------- #

# Layout de salida = exactamente el que ARCA usa en el Portal IVA (Compras),
# tal cual lo importa JWIN por posición de columna. Se reutiliza la constante
# de afip_jwin para no duplicar y para que si ARCA cambia el layout, se cambia
# en un solo lugar.
from src.afip_jwin import HEADER_PORTAL_IVA as _COLS_ARCA
_FACTURAI_COLUMNAS = list(_COLS_ARCA) + ["CAE", "Origen"]


def _fmt_fecha_ar(iso: str) -> str:
    """'2026-08-24' → '24/08/2026'. Devuelve la fecha original si no matchea."""
    if not iso or not isinstance(iso, str):
        return iso or ""
    try:
        y, m, d = iso[:10].split("-")
        return f"{int(d):02d}/{int(m):02d}/{y}"
    except (ValueError, IndexError):
        return iso


def _fmt_num_ar(v) -> str:
    """
    123.45 → '123,45', 0 → '0', '' → '0'. Formato ARCA (coma decimal).
    Se emite como texto para preservar el mismo output que trae ARCA.
    """
    if v is None or v == "":
        return "0"
    try:
        r = round(float(v), 2)
    except (TypeError, ValueError):
        return str(v)
    if r == 0:
        return "0"
    return f"{r:.2f}".replace(".", ",")


def _fmt_codigo_tipo(codigo) -> str:
    """Devuelve el código del tipo de comprobante como string ('1', '6', '11', '81'...)."""
    if codigo is None or codigo == "":
        return ""
    try:
        return str(int(codigo))
    except (TypeError, ValueError):
        return str(codigo)


def _facturai_fila(datos: dict) -> dict:
    """
    Convierte los datos de un comprobante (QR o Gemini Vision) a una fila con
    el layout EXACTO del CSV que emite el Portal IVA de ARCA. Formato:
    fechas dd/mm/aaaa, decimales con coma, tipo de comprobante como código
    numérico. Así el archivo se puede subir directo a AFIP → JWIN sin tocar.

    Fallback para Facturas A que quedaron sin desglose: si Gemini devolvió
    neto=0 e IVA=0 en una Factura A con total > 0, se asume alícuota 21%
    y se calcula Neto = Total / 1,21 e IVA = Total - Neto.
    """
    codigo_tipo = datos.get("tipo_comprobante_codigo")
    if codigo_tipo in (1, 2, 3, 4, 5):  # Factura/ND/NC/Recibo A
        try:
            total = float(datos.get("importe_total") or 0)
            neto = float(datos.get("importe_neto_gravado") or 0)
            iva = float(datos.get("importe_iva") or 0)
        except (TypeError, ValueError):
            total = neto = iva = 0.0
        if total > 0 and neto == 0 and iva == 0:
            alic = datos.get("alicuota_iva") or 21
            try:
                alic_f = float(alic)
            except (TypeError, ValueError):
                alic_f = 21.0
            neto_calc = round(total / (1 + alic_f / 100), 2)
            iva_calc = round(total - neto_calc, 2)
            datos = {**datos,
                     "importe_neto_gravado": neto_calc,
                     "importe_iva": iva_calc,
                     "alicuota_iva": alic_f}

    fila = {c: "" for c in _FACTURAI_COLUMNAS}

    fila["Fecha de Emisión"] = _fmt_fecha_ar(datos.get("fecha") or "")
    fila["Tipo de Comprobante"] = _fmt_codigo_tipo(datos.get("tipo_comprobante_codigo"))
    fila["Punto de Venta"] = str(datos.get("punto_venta") or "")
    fila["Número de Comprobante"] = str(datos.get("numero") or "")
    fila["Tipo Doc. Vendedor"] = "80"
    fila["Nro. Doc. Vendedor"] = str(datos.get("cuit_emisor") or "")
    fila["Denominación Vendedor"] = datos.get("razon_social_emisor") or ""
    fila["Importe Total"] = _fmt_num_ar(datos.get("importe_total"))
    fila["Moneda Original"] = datos.get("moneda") or "PES"
    fila["Tipo de Cambio"] = _fmt_num_ar(datos.get("cotizacion") or 1)

    # JWIN no tiene columnas separadas para Impuestos Internos ni Otros
    # Tributos, así que los ITC/IDC y otros tributos los sumamos DIRECTO
    # al No Gravado desde acá — así el CSV que baja de Facturai ya está
    # listo para importar sin pasar por otra normalización.
    no_gravado_base = float(datos.get("importe_no_gravado") or 0)
    imp_internos = float(datos.get("importe_impuestos_internos") or 0)
    otros_trib = float(datos.get("importe_otros_tributos") or 0)
    fila["Importe No Gravado"] = _fmt_num_ar(
        no_gravado_base + imp_internos + otros_trib)
    fila["Importe Exento"] = _fmt_num_ar(datos.get("importe_exento") or 0)
    fila["Crédito Fiscal Computable"] = "0"
    fila["Importe de Per. o Pagos a Cta. de Otros Imp. Nac."] = "0"
    # Percepciones IIBB en su columna. Compatible con el campo viejo.
    fila["Importe de Percepciones de Ingresos Brutos"] = _fmt_num_ar(
        datos.get("importe_percepciones_iibb")
        or datos.get("importe_percepciones")
        or 0)
    fila["Importe de Impuestos Municipales"] = "0"
    fila["Importe de Percepciones o Pagos a Cuenta de IVA"] = "0"
    # Estas dos SIEMPRE en cero — el importe ya fue trasladado a No Gravado.
    fila["Importe de Impuestos Internos"] = "0"
    fila["Importe Otros Tributos"] = "0"

    # Netos e IVA por alícuota. Como Gemini reporta la alícuota predominante,
    # pongo neto+IVA en la columna que corresponde a esa alícuota.
    neto = datos.get("importe_neto_gravado")
    iva = datos.get("importe_iva")
    alicuota = datos.get("alicuota_iva")
    alic_map = {
        0: ("Neto Gravado IVA 0%", None),
        2.5: ("Neto Gravado IVA 2,5%", "Importe IVA 2,5%"),
        5: ("Neto Gravado IVA 5%", "Importe IVA 5%"),
        10.5: ("Neto Gravado IVA 10,5%", "Importe IVA 10,5%"),
        21: ("Neto Gravado IVA 21%", "Importe IVA 21%"),
        27: ("Neto Gravado IVA 27%", "Importe IVA 27%"),
    }
    # Todos los importes por alícuota arrancan en "0"
    for col_neto, col_iva in alic_map.values():
        fila[col_neto] = "0"
        if col_iva:
            fila[col_iva] = "0"
    if neto is not None and alicuota is not None and alicuota in alic_map:
        col_neto, col_iva = alic_map[alicuota]
        fila[col_neto] = _fmt_num_ar(neto)
        if col_iva and iva is not None:
            fila[col_iva] = _fmt_num_ar(iva)

    fila["Total Neto Gravado"] = _fmt_num_ar(neto or 0)
    fila["Total IVA"] = _fmt_num_ar(iva or 0)

    # Extras al final (no rompen la importación por posición porque son
    # columnas 33 y 34, después de las 32 estándar del Portal IVA).
    fila["CAE"] = str(datos.get("codigo_autorizacion") or "")
    fila["Origen"] = "Vision" if datos.get("fuente") == "gemini" else "QR"
    return fila


def _detectar_duplicados_excluir(encab: list, filas: list, csv_excluir_bytes: bytes) -> set[int]:
    """
    Devuelve el set de índices de ``filas`` (post-procesar) que ya están en
    el CSV a excluir. Usa la misma regla que ``_detectar_duplicados_arca``:
    match por CAE si hay, o por CUIT + fecha + importe (tolerancia 0,05).

    ``encab`` y ``filas`` vienen de ``AJ.procesar`` — layout ARCA + Rubro.
    """
    import csv as _csv
    import io as _io

    try:
        texto = csv_excluir_bytes.decode("utf-8-sig")
    except UnicodeDecodeError:
        texto = csv_excluir_bytes.decode("latin-1")

    filas_arca = list(_csv.reader(_io.StringIO(texto), delimiter=";"))
    if not filas_arca:
        return set()

    header_arca = [str(h).strip().lower() for h in filas_arca[0]]

    def _idx(hdr, *claves):
        for k in claves:
            k = k.lower()
            for i, h in enumerate(hdr):
                if k in h:
                    return i
        return None

    # Índices en el CSV a excluir
    x_fecha = _idx(header_arca, "fecha de emisi", "fecha")
    x_cuit = _idx(header_arca, "nro. doc. vendedor", "nro doc vendedor", "nro. doc")
    x_total = _idx(header_arca, "importe total")
    x_cae = _idx(header_arca, "cae", "cod. autorizaci")

    # Índices en las filas a filtrar (usan el encab de AJ.procesar)
    e_lower = [str(h).strip().lower() for h in encab]
    f_fecha = _idx(e_lower, "fecha de emisi", "fecha")
    f_cuit = _idx(e_lower, "nro. doc. vendedor", "nro. doc. emisor", "nro doc")
    f_total = _idx(e_lower, "importe total")
    f_cae = _idx(e_lower, "cae", "cod. autorizaci")

    if x_cuit is None or x_total is None or f_cuit is None or f_total is None:
        return set()

    def _num(s):
        s = str(s or "").strip()
        if not s:
            return 0.0
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return 0.0

    def _cuit_d(s):
        return "".join(c for c in str(s or "") if c.isdigit())

    caes_excluir: set[str] = set()
    claves_excluir: set[tuple[str, str, float]] = set()
    for fila in filas_arca[1:]:
        if not fila or len(fila) <= max(x_cuit, x_total):
            continue
        cuit = _cuit_d(fila[x_cuit])
        if len(cuit) != 11:
            continue
        fecha = str(fila[x_fecha] or "").strip() if x_fecha is not None else ""
        total = round(_num(fila[x_total]), 2)
        claves_excluir.add((cuit, fecha, total))
        if x_cae is not None and x_cae < len(fila):
            cae = _cuit_d(fila[x_cae])
            if cae:
                caes_excluir.add(cae)

    duplicados: set[int] = set()
    for i, fila in enumerate(filas):
        cuit = _cuit_d(fila[f_cuit] if f_cuit < len(fila) else "")
        if len(cuit) != 11:
            continue
        # 1) Chequear CAE si está
        if f_cae is not None and f_cae < len(fila):
            cae = _cuit_d(fila[f_cae])
            if cae and cae in caes_excluir:
                duplicados.add(i)
                continue
        # 2) Chequear por CUIT + fecha + importe
        fecha = str(fila[f_fecha] or "").strip() if f_fecha is not None else ""
        total = round(_num(fila[f_total]), 2)
        if (cuit, fecha, total) in claves_excluir:
            duplicados.add(i)
    return duplicados


def _cargar_facturai_csv(csv_bytes: bytes):
    """
    Lee un CSV con el layout Portal IVA (o similar) y lo carga en el estado
    de la sesión como si viniera de procesar fotos. Sirve para retomar un
    CSV ya generado por Facturai (o incluso el CSV mensual de ARCA) y seguir
    con el cruce de duplicados y el mapeo de rubros sin re-procesar imágenes.
    """
    import csv as _csv
    import io as _io

    # Decodificar
    try:
        texto = csv_bytes.decode("utf-8-sig")
    except UnicodeDecodeError:
        texto = csv_bytes.decode("latin-1")

    filas = list(_csv.reader(_io.StringIO(texto), delimiter=";"))
    if not filas:
        return 0

    header = [str(h).strip().lower() for h in filas[0]]

    def _idx(*claves):
        for k in claves:
            k = k.lower()
            for i, h in enumerate(header):
                if k in h:
                    return i
        return None

    i_fecha = _idx("fecha de emisi", "fecha")
    i_tipo = _idx("tipo de comprobante", "tipo")
    i_pv = _idx("punto de venta")
    i_num = _idx("número de comprobante", "numero de comprobante", "número desde",
                 "numero desde")
    i_cuit = _idx("nro. doc. vendedor", "nro doc vendedor", "nro. doc")
    i_deno = _idx("denominación vendedor", "denominacion vendedor")
    i_total = _idx("importe total")
    i_moneda = _idx("moneda original", "moneda")
    i_ctz = _idx("tipo de cambio", "cotización", "cotizacion")
    i_cae = _idx("cae", "cod. autorizaci")
    i_origen = _idx("origen")

    if i_cuit is None or i_total is None:
        raise ValueError("El CSV no parece del layout Portal IVA "
                         "(faltan columnas Nro. Doc. Vendedor o Importe Total).")

    def _num(s):
        s = str(s or "").strip()
        if not s:
            return 0.0
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return 0.0

    def _cuit(s):
        return "".join(c for c in str(s or "") if c.isdigit())

    def _fecha_iso(dmy):
        s = str(dmy or "").strip()
        try:
            d, m, y = s[:10].split("/")
            return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"
        except (ValueError, IndexError):
            return s

    cargadas = 0
    for fila in filas[1:]:
        if not fila or len(fila) <= i_cuit:
            continue
        cuit = _cuit(fila[i_cuit])
        if len(cuit) != 11:
            continue

        datos = {
            "fecha": _fecha_iso(fila[i_fecha]) if i_fecha is not None else "",
            "tipo_comprobante_codigo": int(_num(fila[i_tipo])) if i_tipo is not None else None,
            "tipo_comprobante": "",  # no viene en el CSV, se puede reconstruir después
            "punto_venta": int(_num(fila[i_pv])) if i_pv is not None else None,
            "numero": int(_num(fila[i_num])) if i_num is not None else None,
            "cuit_emisor": cuit,
            "razon_social_emisor": (fila[i_deno].strip() if i_deno is not None
                                    and i_deno < len(fila) else ""),
            "importe_total": _num(fila[i_total]),
            "moneda": (fila[i_moneda] if i_moneda is not None
                       and i_moneda < len(fila) else "PES") or "PES",
            "cotizacion": _num(fila[i_ctz]) if i_ctz is not None else 1,
            "codigo_autorizacion": (fila[i_cae].strip() if i_cae is not None
                                    and i_cae < len(fila) else ""),
            "fuente": None,     # no sabemos si vino de QR o Vision
        }
        # Si venía la columna Origen del Excel previo, la conservamos.
        if i_origen is not None and i_origen < len(fila):
            origen = str(fila[i_origen] or "").strip().lower()
            if "vision" in origen or "gemini" in origen:
                datos["fuente"] = "gemini"

        st.session_state["facturai_datos"].append(datos)
        st.session_state["facturai_lista"].append(_facturai_fila(datos))
        st.session_state.setdefault("facturai_incluir", []).append(True)
        cargadas += 1
    return cargadas


def _detectar_duplicados_arca(datos_lista: list[dict], csv_arca_bytes: bytes) -> set[int]:
    """
    Cruza las facturas procesadas por Facturai contra un CSV de ARCA
    (Portal IVA - Compras). Devuelve el set de índices que ya están en
    ARCA y no habría que exportar de nuevo.

    Regla de match:
      1. Si el CAE coincide → duplicado seguro.
      2. Si no hay CAE: CUIT emisor + fecha + importe total (tolerancia 0,05).
    """
    import csv as _csv
    import io as _io

    # Detectar encoding del CSV de ARCA
    try:
        texto = csv_arca_bytes.decode("utf-8-sig")
    except UnicodeDecodeError:
        texto = csv_arca_bytes.decode("latin-1")

    filas = list(_csv.reader(_io.StringIO(texto), delimiter=";"))
    if not filas:
        return set()

    header = [str(h).strip().lower() for h in filas[0]]

    def _idx(*claves):
        for k in claves:
            k = k.lower()
            for i, h in enumerate(header):
                if k in h:
                    return i
        return None

    i_fecha = _idx("fecha de emisi", "fecha")
    i_cuit = _idx("nro. doc. vendedor", "nro doc vendedor", "nro. doc")
    i_total = _idx("importe total")
    i_cae = _idx("cae", "cod. autorizaci", "cod autorizac")

    if i_cuit is None or i_total is None:
        return set()

    def _num(s):
        s = str(s or "").strip()
        if not s:
            return 0.0
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return 0.0

    def _cuit_digs(s):
        return "".join(c for c in str(s or "") if c.isdigit())

    caes_arca: set[str] = set()
    claves_arca: set[tuple[str, str, float]] = set()   # (cuit, fecha, importe)
    for fila in filas[1:]:
        if not fila or len(fila) <= max(i_cuit, i_total):
            continue
        cuit = _cuit_digs(fila[i_cuit])
        if len(cuit) != 11:
            continue
        fecha = str(fila[i_fecha] or "").strip() if i_fecha is not None else ""
        total = round(_num(fila[i_total]), 2)
        claves_arca.add((cuit, fecha, total))
        if i_cae is not None and i_cae < len(fila):
            cae = _cuit_digs(fila[i_cae])
            if cae:
                caes_arca.add(cae)

    def _fecha_ar(iso):
        try:
            y, m, d = iso[:10].split("-")
            return f"{int(d):02d}/{int(m):02d}/{y}"
        except (ValueError, IndexError, AttributeError):
            return iso

    duplicados: set[int] = set()
    for i, datos in enumerate(datos_lista):
        cae = _cuit_digs(datos.get("codigo_autorizacion"))
        if cae and cae in caes_arca:
            duplicados.add(i)
            continue
        cuit = _cuit_digs(datos.get("cuit_emisor"))
        fecha_ar = _fecha_ar(datos.get("fecha") or "")
        try:
            total = round(float(datos.get("importe_total") or 0), 2)
        except (TypeError, ValueError):
            continue
        if (cuit, fecha_ar, total) in claves_arca:
            duplicados.add(i)
    return duplicados


def _facturai_a_csv(facturas: list[dict]) -> bytes:
    """
    CSV con el layout exacto del Portal IVA (';' + coma decimal + CRLF + BOM
    UTF-8). Se puede subir directo en la sección AFIP → JWIN con rubros.
    """
    import csv as _csv
    import io as _io

    buf = _io.StringIO()
    w = _csv.writer(buf, delimiter=";", lineterminator="\r\n")
    # Solo las 32 columnas estándar. CAE y Origen quedan para el Excel de vista.
    encab = list(_COLS_ARCA)
    w.writerow(encab)
    for f in facturas:
        w.writerow([f.get(c, "") for c in encab])
    return buf.getvalue().encode("utf-8-sig")


def _facturai_a_excel(facturas: list[dict], columnas: list[str] | None = None) -> bytes:
    """
    Genera un Excel con las facturas en el layout Portal IVA (32 cols de ARCA).
    Se puede pasar ``columnas`` para forzar otro layout (ej. cuando se aplica
    el maestro y viene 32 + 'Rubro').
    """
    import io as _io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    cols = columnas if columnas is not None else list(_COLS_ARCA)

    wb = Workbook()
    ws = wb.active
    ws.title = "Facturas"

    bold = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="1F4E78")
    for j, col in enumerate(cols, start=1):
        c = ws.cell(1, j, col)
        c.font = bold
        c.fill = fill

    for i, f in enumerate(facturas, start=2):
        # ``facturas`` puede venir como lista de dicts (Facturai-style) o de
        # listas (AJ.procesar-style). Manejo los dos casos.
        for j, col in enumerate(cols, start=1):
            if isinstance(f, dict):
                ws.cell(i, j, f.get(col, ""))
            elif isinstance(f, (list, tuple)) and j - 1 < len(f):
                ws.cell(i, j, f[j - 1])

    # Anchos razonables
    anchos = {"A": 12, "B": 22, "C": 8, "D": 12, "E": 14, "F": 14,
              "G": 40, "H": 14, "I": 10, "J": 10, "AF": 20, "AG": 10}
    for letra, ancho in anchos.items():
        ws.column_dimensions[letra].width = ancho
    ws.freeze_panes = "A2"

    buf = _io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def seccion_facturai():
    from src.facturai import qr as QR

    st.title("📸 Facturas por foto → Excel")
    st.caption("Subí una foto o PDF de una factura, la app lee el **QR de AFIP** y le "
               "extrae los datos oficiales del comprobante. Cada factura procesada se "
               "suma a la tabla — al final bajás un Excel con todo.")

    # Estado en la sesión: lista de facturas acumuladas.
    if "facturai_lista" not in st.session_state:
        st.session_state["facturai_lista"] = []      # cada item: dict con la fila
        st.session_state["facturai_datos"] = []      # cada item: dict con los datos crudos del QR

    # Modo alternativo: retomar un CSV ya procesado antes.
    # Útil si bajaste el CSV en una sesión anterior y ahora querés
    # cruzarlo con ARCA o mapear rubros sin reprocesar las fotos.
    lista_actual = st.session_state.get("facturai_lista", [])
    with st.expander("📄 ¿Ya tenés un CSV procesado antes? Cargalo acá",
                     expanded=not lista_actual):
        prev_csv = st.file_uploader(
            "CSV con el layout Portal IVA (32 columnas ARCA)",
            type=["csv"], key="facturai_prev_csv",
            help="Podés subir un CSV que hayas bajado antes de esta misma "
                 "sección, o cualquier CSV con layout ARCA — la app lo carga "
                 "en la sesión y podés seguir con cruce y rubros.",
        )
        if prev_csv is not None:
            if st.button("📥 Cargar en la sesión", key="facturai_cargar_prev",
                         type="primary", use_container_width=True):
                try:
                    n = _cargar_facturai_csv(prev_csv.getvalue())
                except Exception as exc:  # noqa: BLE001
                    st.error(f"No se pudo cargar el CSV: {exc}")
                else:
                    st.success(f"Se cargaron {n} factura(s). "
                               "Ya podés hacer el cruce con ARCA y bajar el CSV.")
                    st.rerun()

    st.divider()
    st.subheader("O procesar fotos/PDFs nuevos")
    archivos = st.file_uploader(
        "Foto(s) o PDF(s) de facturas",
        type=["jpg", "jpeg", "png", "webp", "bmp", "tif", "tiff", "pdf"],
        accept_multiple_files=True,
        key="facturai_up",
    )

    procesar = st.button("🔍 Procesar archivos", type="primary",
                         disabled=not archivos, use_container_width=True)

    if procesar and archivos:
        from src.facturai import vision as VISION
        gemini_ok = VISION.hay_conexion()

        errores = []

        def _clave(datos):
            cae = str(datos.get("codigo_autorizacion") or "")
            return cae or f"{datos.get('cuit_emisor')}-{datos.get('punto_venta')}-{datos.get('numero')}"

        # Barra de progreso: total = suma de páginas de todos los archivos.
        total_paginas = 0
        for arch in archivos:
            d = arch.getvalue()
            if d[:4] == b"%PDF":
                try:
                    total_paginas += QR.pdf_paginas(d)
                except Exception:  # noqa: BLE001
                    total_paginas += 1
            else:
                total_paginas += 1
        barra = st.progress(0.0, text="Preparando…")
        procesadas = 0

        for arch in archivos:
            nombre_arch = arch.name
            data = arch.getvalue()

            # PDF: iteramos página por página con un generador (no cargamos
            # todas las páginas en memoria a la vez — el proceso se cae si
            # el PDF tiene 20+ páginas escaneadas en resolución alta).
            if data[:4] == b"%PDF":
                try:
                    iter_paginas = QR.pdf_a_imagenes(data)
                    total_pag_este_pdf = QR.pdf_paginas(data)
                except Exception as exc:  # noqa: BLE001
                    errores.append((nombre_arch, f"PDF ilegible: {exc}"))
                    procesadas += 1
                    continue
            else:
                iter_paginas = iter([data])
                total_pag_este_pdf = 1

            for i, img_bytes in enumerate(iter_paginas, start=1):
                nombre_pag = (nombre_arch if total_pag_este_pdf == 1
                              else f"{nombre_arch} · pág {i}/{total_pag_este_pdf}")
                barra.progress(procesadas / max(total_paginas, 1),
                               text=f"Procesando {nombre_pag}…")

                facturas_encontradas: list[dict] = []

                # 1) Intentar QR primero — es oficial, no consume API.
                try:
                    url = QR.detectar_qr_en_imagen(img_bytes)
                except Exception:  # noqa: BLE001
                    url = None
                if url:
                    d = QR.parsear_url_afip(url)
                    if d:
                        facturas_encontradas.append(d)

                # 2) Si no encontró nada por QR, cae a Gemini Vision
                #    (una imagen puede tener varios comprobantes: Vision los
                #    devuelve todos como lista).
                if not facturas_encontradas:
                    if not gemini_ok:
                        errores.append((nombre_pag,
                                       "Sin QR y Gemini no configurado."))
                        procesadas += 1
                        continue
                    try:
                        raws = VISION.extraer_datos(img_bytes)
                        for raw in raws:
                            facturas_encontradas.append(VISION.a_formato_qr(raw))
                    except Exception as exc:  # noqa: BLE001
                        errores.append((nombre_pag, f"Gemini: {exc}"))
                        procesadas += 1
                        continue
                    if not facturas_encontradas:
                        errores.append((nombre_pag,
                                       "No se detectó ninguna factura."))
                        procesadas += 1
                        continue

                # 3) Agregar cada factura encontrada, chequeando duplicados.
                for datos in facturas_encontradas:
                    clave = _clave(datos)
                    claves_existentes = {_clave(d)
                                         for d in st.session_state["facturai_datos"]}
                    if clave in claves_existentes:
                        fuente = "Vision" if datos.get("fuente") == "gemini" else "QR"
                        errores.append((nombre_pag,
                                       f"Duplicada — ya estaba cargada ({fuente})"))
                        continue
                    st.session_state["facturai_datos"].append(datos)
                    st.session_state["facturai_lista"].append(_facturai_fila(datos))
                    st.session_state.setdefault("facturai_incluir", []).append(True)

                procesadas += 1

        barra.progress(1.0, text=f"Listo: {procesadas} página(s) procesadas.")
        if errores:
            with st.expander(f"⚠️ {len(errores)} caso(s) no se procesaron",
                             expanded=True):
                for nombre, motivo in errores:
                    st.write(f"• **{nombre}**: {motivo}")

        if not gemini_ok:
            st.caption("💡 Tip: si configurás una API key gratuita de Gemini en "
                       "`st.secrets['gemini']['api_key']`, las facturas sin QR "
                       "(tickets viejos, escaneos con varios tickets juntos) se "
                       "leen automáticamente con IA de visión.")

    # Tabla acumulada
    lista = st.session_state["facturai_lista"]
    if not lista:
        st.info("Todavía no hay facturas cargadas. Subí una o varias imágenes/PDFs y tocá **Procesar**.")
        return

    st.divider()
    st.subheader(f"📋 Facturas acumuladas ({len(lista)})")

    # ---------------------------------------------------------------- #
    # Cruce opcional con el CSV de ARCA: para no duplicar facturas
    # electrónicas que ya vienen del Portal IVA en la baja mensual.
    # ---------------------------------------------------------------- #
    up_arca = st.file_uploader(
        "CSV de ARCA (opcional): detecta duplicados con lo que ya bajaste del Portal IVA",
        type=["csv"], key="facturai_arca_dedup",
    )
    dup_indices: set[int] = set()
    if up_arca is not None:
        try:
            dup_indices = _detectar_duplicados_arca(
                st.session_state["facturai_datos"], up_arca.getvalue()
            )
        except Exception as exc:  # noqa: BLE001
            st.error(f"No se pudo leer el CSV de ARCA: {exc}")
        else:
            if dup_indices:
                st.info(f"🔍 Se detectaron **{len(dup_indices)}** factura(s) "
                        "que ya están en el CSV de ARCA — vienen destildadas "
                        "abajo. Chequealas y ajustá si querés.")
            else:
                st.success("No se detectó ningún cruce — todas las facturas "
                           "procesadas son distintas a las que trae el CSV de ARCA.")

    # Inicializo/ajusto la lista de "incluidos" en session_state.
    if len(st.session_state.get("facturai_incluir", [])) != len(lista):
        st.session_state["facturai_incluir"] = [True] * len(lista)
    # Aplicar los duplicados detectados: quedan destildados por defecto,
    # pero el usuario puede volver a tildarlos si quiere igual.
    if up_arca is not None:
        st.session_state["facturai_incluir"] = [
            i not in dup_indices for i in range(len(lista))
        ]

    # Tabla editable con checkbox "Incluir"
    vista_df = pd.DataFrame([{
        "Incluir": st.session_state["facturai_incluir"][i],
        "Fecha": f["Fecha de Emisión"],
        "Tipo": f["Tipo de Comprobante"],
        "Emisor": f["Denominación Vendedor"] or QR.formato_cuit(f["Nro. Doc. Vendedor"]),
        "CUIT": QR.formato_cuit(f["Nro. Doc. Vendedor"]),
        "Comprobante": f"{str(f['Punto de Venta']).zfill(5)}-{str(f['Número de Comprobante']).zfill(8)}"
                        if f["Punto de Venta"] and f["Número de Comprobante"] else "",
        "Importe": f["Importe Total"],
        "Origen": f["Origen"],
    } for i, f in enumerate(lista)])

    editado = st.data_editor(
        vista_df, use_container_width=True, hide_index=True,
        key="facturai_editor_vista",
        column_config={
            "Incluir": st.column_config.CheckboxColumn(
                "Incluir",
                help="Destildá las filas que no querés exportar (duplicadas, "
                     "cargadas mal, etc.)",
                default=True,
            ),
        },
        disabled=[c for c in vista_df.columns if c != "Incluir"],
    )
    st.session_state["facturai_incluir"] = editado["Incluir"].tolist()

    incluidas = [f for i, f in enumerate(lista)
                 if st.session_state["facturai_incluir"][i]]
    excluidas = len(lista) - len(incluidas)

    def _num_ar_parse(s):
        s = str(s or "0").replace(".", "").replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return 0.0

    total = sum(_num_ar_parse(f["Importe Total"]) for f in incluidas)
    if excluidas:
        st.caption(f"✅ {len(incluidas)} para exportar   ·   "
                   f"❌ {excluidas} destildadas   ·   "
                   f"Suma exportada: **${total:,.2f}**")
    else:
        st.caption(f"Suma total: **${total:,.2f}**")

    # ---------------------------------------------------------------- #
    # Mapeo de rubros con el mismo maestro que ya usa AFIP → JWIN.
    # Es opcional: si no subís maestro, el CSV sale sin la columna Rubro.
    # ---------------------------------------------------------------- #
    st.divider()
    st.subheader("🏷️ Asignar rubros para JWIN")
    st.caption("Si subís tu **Excel maestro de proveedores** (el mismo que usás en "
               "la sección AFIP → JWIN), la app le pone el rubro a cada factura "
               "y te deja el CSV **listo para importar a JWIN**. Si no lo subís, "
               "podés bajar el CSV crudo sin rubros abajo.")

    up_maestro = st.file_uploader(
        "Excel maestro (Proveedores + Rubros) — opcional",
        type=["xlsx"], key="facturai_maestro",
    )

    # A partir de acá se usa 'incluidas' (solo las tildadas), no 'lista'.
    lista = incluidas

    csv_puro = _facturai_a_csv(lista)
    csv_final = csv_puro    # si no hay maestro, es lo que se descarga
    stats_map = None
    nuevos = []             # (cuit, denom, rubro, desc) para actualizar el maestro

    if up_maestro is not None:
        maestro_bytes = up_maestro.getvalue()
        try:
            encab_final, filas_final, desconocidos, stats_map, rubros_disp = \
                AJ.procesar(csv_puro, maestro_bytes)
        except Exception as exc:  # noqa: BLE001
            st.error(f"No se pudo leer el maestro: {exc}")
            desconocidos = {}
        else:
            # Interfaz para asignar rubro a proveedores nuevos
            if desconocidos:
                st.warning(
                    f"Hay {len(desconocidos)} proveedor(es) sin rubro en tu maestro. "
                    "Asignáselos ahora — se aplica al toque y quedan cargados en el "
                    "maestro actualizado que podés bajar abajo."
                )
                opciones = [""] + [f"{c} - {d}" for c, d in sorted(rubros_disp.items())]
                base = pd.DataFrame([{"CUIT": c, "Denominación": d, "Rubro": ""}
                                     for c, d in sorted(desconocidos.items())])
                editado = st.data_editor(
                    base, hide_index=True, use_container_width=True,
                    key="facturai_editor",
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
                        nuevos.append((cuit, row["Denominación"], cod,
                                       rubros_disp.get(cod, "")))
                if extra:
                    # Reprocesar con los rubros recién cargados
                    encab_final, filas_final, desconocidos, stats_map, rubros_disp = \
                        AJ.procesar(csv_puro, maestro_bytes, extra)

            csv_final = AJ.construir_csv(encab_final, filas_final)

            m1, m2, m3 = st.columns(3)
            m1.metric("Comprobantes", stats_map["comprobantes"])
            m2.metric("Con rubro", stats_map["asignados"])
            m3.metric("Sin rubro", stats_map["sin_rubro"])
            if stats_map["sin_rubro"] == 0:
                st.success("Todos los comprobantes tienen su rubro. 🎉")

    st.divider()
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        c1.download_button(
            "⬇️ CSV para JWIN" + (" (con rubro)" if up_maestro is not None else ""),
            data=csv_final,
            file_name=f"Facturas_procesadas_{datetime.now():%Y%m%d_%H%M}.csv",
            mime="text/csv",
            use_container_width=True,
            key="facturai_dl_csv",
            type="primary",
        )
    with c2:
        # Excel espeja el CSV: si hay maestro, usa la salida de AJ.procesar
        # (32 columnas ARCA + Rubro). Si no, solo las 32 de ARCA. Nunca
        # CAE ni Origen, así JWIN encuentra el Rubro en la columna AG.
        if up_maestro is not None and stats_map is not None:
            excel_bytes = _facturai_a_excel(filas_final, columnas=encab_final)
            label_excel = "⬇️ Excel (con rubro)"
        else:
            excel_bytes = _facturai_a_excel(lista)
            label_excel = "⬇️ Excel (revisar)"
        c2.download_button(
            label_excel,
            data=excel_bytes,
            file_name=f"Facturas_procesadas_{datetime.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="facturai_dl_xlsx",
        )
    with c3:
        # Botón para bajar el maestro actualizado con los proveedores nuevos.
        if up_maestro is not None and nuevos:
            c3.download_button(
                "⬇️ Maestro actualizado",
                data=AJ.construir_maestro_actualizado(up_maestro.getvalue(), nuevos),
                file_name=up_maestro.name.replace(".xlsx", "_actualizado.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="facturai_dl_maestro",
            )
    with c4:
        if c4.button("🗑️ Vaciar lista", use_container_width=True, key="facturai_clear"):
            st.session_state["facturai_lista"] = []
            st.session_state["facturai_datos"] = []
            st.session_state["facturai_incluir"] = []
            st.rerun()

    with st.expander("Ver JSON de la última factura cargada"):
        st.json(st.session_state["facturai_datos"][-1])

    st.caption("**Origen** de cada fila: **QR** = leído del código QR de AFIP "
               "(oficial, sin desglose de netos). **Vision** = leído por Gemini "
               "cuando no hay QR o no se pudo decodificar (trae razón social, "
               "neto e IVA discriminados).")


# --------------------------------------------------------------------------- #
# Sección: Tareas y checklist (Supabase)
# --------------------------------------------------------------------------- #

def seccion_tareas():
    from src import tareas_db as TDB

    st.title("✅ Tareas y checklist")

    if not TDB.hay_conexion():
        st.warning(
            "Todavía no está configurada la base de datos compartida (Supabase). "
            "Cargá las credenciales en los *secrets* para empezar a usar esta sección."
        )
        st.caption(
            "Necesitás un bloque `[supabase]` con `url` y `key` en "
            "`.streamlit/secrets.toml` (local) o en Settings → Secrets (Streamlit Cloud)."
        )
        return

    top1, top2 = st.columns([3, 1])
    with top1:
        quien = _quien_soy()
    with top2:
        st.write("")
        if st.button("🔄 Actualizar", use_container_width=True, help="Traer lo último"):
            st.rerun()

    st.divider()

    tab_tareas, tab_rutinas = st.tabs(["📌 Tareas pendientes", "🔁 Rutinas"])
    with tab_tareas:
        _ui_tareas(TDB, quien)
    with tab_rutinas:
        _ui_rutinas(TDB, quien)


# --------------------------------------------------------------------------- #
# Programa: menú de herramientas
# --------------------------------------------------------------------------- #

# Herramientas agrupadas: cada grupo es (etiqueta, [(clave, etiqueta, render), ...]).
_GRUPOS = [
    # Grupo "🏢 Gestión" deshabilitado temporalmente. Los archivos siguen en el
    # repo (seccion_clientes, seccion_tareas, src/clientes_db.py, src/tareas_db.py,
    # src/arca/*); para reactivar cualquiera, descomentar la línea correspondiente.
    # ("🏢 Gestión", [
    #     ("clientes", "👥  Clientes",           lambda: seccion_clientes()),
    #     ("tareas",   "✅  Tareas y checklist", lambda: seccion_tareas()),
    # ]),
    ("🏦 Bancos", [
        ("pdf",          "🏦  PDF de banco → Excel",                lambda: seccion_pdf_banco()),
        ("banco_contab", "🔄  Banco vs Contabilidad",               lambda: seccion_banco_contab()),
    ]),
    ("💲 Impuestos", [
        ("iva",          "💲  Posición IVA",                        lambda: seccion_iva()),
        ("monotributo",  "📊  Monotributo — Recategorización",      lambda: seccion_monotributo()),
    ]),
    ("🔧 Conversiones entre sistemas", [
        ("facturai",     "📸  Facturas por foto → Excel",           lambda: seccion_facturai()),
        ("ps3",          "📒  JWIN → PS3 (MICROENV)",               lambda: seccion_ps3()),
        ("afip",         "📥  AFIP → JWIN (rubros)",                lambda: seccion_afip()),
        ("ventas",       "🧾  Ventas por actividad (Tango)",        lambda: seccion_ventas()),
        ("rango",        "📦  Rango — Compras (Paradigma)",         lambda: seccion_rango()),
    ]),
    ("🧮 Otros", [
        ("comparador",   "🧮  Comparador (Contab vs Caja sucesión)", lambda: seccion_comparador()),
    ]),
]

# Lista plana (clave, etiqueta, render) derivada de los grupos — para resolver
# rápido qué render corresponde a la sección activa.
_HERRAMIENTAS = [item for _, items in _GRUPOS for item in items]


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
        for i, (grupo, items) in enumerate(_GRUPOS):
            if i > 0:
                st.markdown("")  # separación leve entre grupos
            st.markdown(f"**{grupo}**")
            for clave, etiqueta, _ in items:
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
