"""
Tareas y rutinas — base de datos compartida (Supabase)
======================================================

Las tareas se guardan en Supabase (una base de datos Postgres en la nube) para
que se **compartan entre todas las computadoras** que abren la app. Si la app se
reinicia o se actualiza, los datos siguen ahí (a diferencia de un archivo local,
que en Streamlit Cloud se borra).

La conexión usa las credenciales de ``st.secrets["supabase"]``:

    [supabase]
    url = "https://xxxxxxxx.supabase.co"
    key = "clave-service_role"

Si no hay credenciales cargadas, ``hay_conexion()`` devuelve False y la sección
muestra un aviso amable en lugar de romperse.
"""

from __future__ import annotations

from datetime import date, datetime, timezone

import streamlit as st


def _ahora() -> str:
    """Marca de tiempo actual en formato ISO (UTC)."""
    return datetime.now(timezone.utc).isoformat()


def hay_conexion() -> bool:
    """True si están cargadas las credenciales de Supabase en los secrets."""
    try:
        cfg = st.secrets["supabase"]
        return bool(cfg.get("url") and cfg.get("key"))
    except Exception:
        return False


@st.cache_resource(show_spinner=False)
def _cliente():
    """Cliente de Supabase (se crea una sola vez y se reutiliza)."""
    from supabase import create_client

    cfg = st.secrets["supabase"]
    return create_client(cfg["url"], cfg["key"])


# --------------------------------------------------------------------------- #
# Tareas puntuales (de una vez)
# --------------------------------------------------------------------------- #

def listar_tareas() -> list[dict]:
    """Todas las tareas: primero las pendientes, luego las más nuevas."""
    res = (
        _cliente().table("tareas").select("*")
        .order("hecha").order("creada_en", desc=True).execute()
    )
    return res.data or []


def crear_tarea(titulo: str, nota: str | None, monto, creada_por: str) -> None:
    _cliente().table("tareas").insert({
        "titulo": titulo,
        "nota": nota or None,
        "monto": monto,
        "creada_por": creada_por,
    }).execute()


def marcar_tarea(tarea_id: int, hecha: bool, quien: str) -> None:
    """Tacha (o destacha) una tarea, guardando quién y cuándo."""
    _cliente().table("tareas").update({
        "hecha": hecha,
        "hecha_por": quien if hecha else None,
        "hecha_en": _ahora() if hecha else None,
    }).eq("id", tarea_id).execute()


def borrar_tarea(tarea_id: int) -> None:
    _cliente().table("tareas").delete().eq("id", tarea_id).execute()


# --------------------------------------------------------------------------- #
# Tareas rutinarias (se repiten y se reinician por período)
# --------------------------------------------------------------------------- #

def periodo_actual(frecuencia: str) -> str:
    """Identificador del período en curso según la frecuencia.

    - diaria  -> '2026-06-19'
    - semanal -> '2026-W25'
    - mensual -> '2026-06'

    Cuando cambia el período, la rutina vuelve a aparecer sin tachar.
    """
    hoy = date.today()
    if frecuencia == "semanal":
        anio, semana, _ = hoy.isocalendar()
        return f"{anio}-W{semana:02d}"
    if frecuencia == "mensual":
        return hoy.strftime("%Y-%m")
    return hoy.strftime("%Y-%m-%d")


def listar_rutinas() -> list[dict]:
    res = (
        _cliente().table("rutinas").select("*")
        .eq("activa", True).order("frecuencia").order("orden").order("id").execute()
    )
    return res.data or []


def crear_rutina(titulo: str, frecuencia: str) -> None:
    _cliente().table("rutinas").insert({
        "titulo": titulo,
        "frecuencia": frecuencia,
    }).execute()


def borrar_rutina(rutina_id: int) -> None:
    _cliente().table("rutinas").delete().eq("id", rutina_id).execute()


def estados_periodo(periodos: list[str]) -> dict:
    """Marcas de completado para los períodos dados.

    Devuelve un diccionario ``{(rutina_id, periodo): fila}``. Que exista una
    fila significa que la rutina está hecha en ese período.
    """
    periodos = list({p for p in periodos})
    if not periodos:
        return {}
    res = (
        _cliente().table("rutinas_estado").select("*")
        .in_("periodo", periodos).execute()
    )
    return {(r["rutina_id"], r["periodo"]): r for r in (res.data or [])}


def marcar_rutina(rutina_id: int, periodo: str, hecha: bool, quien: str) -> None:
    """Marca (crea fila) o desmarca (borra fila) una rutina en un período."""
    tabla = _cliente().table("rutinas_estado")
    if hecha:
        tabla.upsert(
            {
                "rutina_id": rutina_id,
                "periodo": periodo,
                "hecha_por": quien,
                "hecha_en": _ahora(),
            },
            on_conflict="rutina_id,periodo",
        ).execute()
    else:
        tabla.delete().eq("rutina_id", rutina_id).eq("periodo", periodo).execute()
