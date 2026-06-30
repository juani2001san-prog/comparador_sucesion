"""
Clientes del estudio — capa de acceso a Supabase
================================================

Misma idea que ``tareas_db``: si no hay credenciales de Supabase configuradas,
``hay_conexion()`` devuelve False y la sección muestra un aviso amable en lugar
de explotar.
"""

from __future__ import annotations

from datetime import datetime, timezone

import streamlit as st


TIPOS = [
    "Monotributo",
    "Responsable Inscripto",
    "Exento",
    "Sociedad",
    "Otro",
]

CATEGORIAS_MONO = list("ABCDEFGHIJK")


def _ahora() -> str:
    return datetime.now(timezone.utc).isoformat()


def hay_conexion() -> bool:
    """True si están cargadas las credenciales de Supabase."""
    try:
        cfg = st.secrets["supabase"]
        return bool(cfg.get("url") and cfg.get("key"))
    except Exception:
        return False


@st.cache_resource(show_spinner=False)
def _cliente():
    from supabase import create_client
    cfg = st.secrets["supabase"]
    return create_client(cfg["url"], cfg["key"])


# --------------------------------------------------------------------------- #
# CRUD
# --------------------------------------------------------------------------- #

def listar(incluir_inactivos: bool = False) -> list[dict]:
    """Lista de clientes. Por defecto, solo los activos."""
    q = _cliente().table("clientes").select("*")
    if not incluir_inactivos:
        q = q.eq("activo", True)
    res = q.order("razon_social").execute()
    return res.data or []


def obtener(cliente_id: int) -> dict | None:
    res = _cliente().table("clientes").select("*").eq("id", cliente_id).limit(1).execute()
    filas = res.data or []
    return filas[0] if filas else None


def buscar_por_cuit(cuit: str) -> dict | None:
    res = _cliente().table("clientes").select("*").eq("cuit", cuit).limit(1).execute()
    filas = res.data or []
    return filas[0] if filas else None


def crear(datos: dict) -> dict:
    """Crea un cliente. Devuelve el registro insertado."""
    res = _cliente().table("clientes").insert(_limpiar(datos)).execute()
    return (res.data or [{}])[0]


def actualizar(cliente_id: int, datos: dict) -> None:
    _cliente().table("clientes").update(_limpiar(datos)).eq("id", cliente_id).execute()


def archivar(cliente_id: int, archivar_si: bool = True) -> None:
    """Baja lógica: deja el registro pero lo saca de la lista visible."""
    _cliente().table("clientes").update({"activo": not archivar_si}).eq("id", cliente_id).execute()


def borrar(cliente_id: int) -> None:
    """Baja DEFINITIVA — usalo solo para errores de carga."""
    _cliente().table("clientes").delete().eq("id", cliente_id).execute()


def guardar_padron(cliente_id: int, padron_data: dict) -> None:
    """Persiste el snapshot del padrón ARCA + timestamp de actualización."""
    _cliente().table("clientes").update({
        "padron_data": padron_data,
        "padron_actualizado_en": _ahora(),
    }).eq("id", cliente_id).execute()


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #

def _limpiar(datos: dict) -> dict:
    """Convierte strings vacíos a None y deja sólo claves no nulas para insertar/actualizar."""
    out = {}
    for k, v in datos.items():
        if isinstance(v, str):
            v = v.strip() or None
        out[k] = v
    return out


def validar_cuit(cuit: str) -> bool:
    """Valida el dígito verificador de un CUIT argentino (11 dígitos)."""
    s = "".join(c for c in (cuit or "") if c.isdigit())
    if len(s) != 11:
        return False
    pesos = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2]
    suma = sum(int(s[i]) * pesos[i] for i in range(10))
    dv = 11 - (suma % 11)
    if dv == 11:
        dv = 0
    elif dv == 10:
        return False
    return dv == int(s[10])


def formato_cuit(cuit: str) -> str:
    """Devuelve el CUIT con guiones (XX-XXXXXXXX-X). Si es inválido, lo deja igual."""
    s = "".join(c for c in (cuit or "") if c.isdigit())
    if len(s) != 11:
        return cuit or ""
    return f"{s[:2]}-{s[2:10]}-{s[10]}"
