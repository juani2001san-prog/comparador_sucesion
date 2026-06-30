"""
Carga del certificado digital y clave privada del estudio.

Busca las credenciales en este orden:
  1. ``st.secrets["arca"]`` — para producción (Streamlit Cloud).
  2. Carpeta ``certificados/`` del repo — para desarrollo local.

El CUIT del titular se extrae automáticamente del subject del certificado.
"""

from __future__ import annotations

from pathlib import Path

try:
    import streamlit as st
except ImportError:  # pragma: no cover
    st = None  # type: ignore


def _desde_secrets() -> tuple[bytes, bytes, str] | None:
    if st is None:
        return None
    try:
        cfg = st.secrets["arca"]
    except Exception:
        return None
    cert_pem = cfg.get("certificado")
    key_pem = cfg.get("clave_privada")
    cuit = str(cfg.get("cuit_representada") or "")
    if not (cert_pem and key_pem and cuit):
        return None
    return (
        cert_pem.encode("utf-8") if isinstance(cert_pem, str) else cert_pem,
        key_pem.encode("utf-8") if isinstance(key_pem, str) else key_pem,
        "".join(c for c in cuit if c.isdigit()),
    )


def _desde_archivos(carpeta: str = "certificados") -> tuple[bytes, bytes, str] | None:
    base = Path(carpeta)
    if not base.is_dir():
        return None

    certs = sorted(base.glob("*.crt"))
    if not certs:
        return None

    cert_file = certs[0]
    key_file = cert_file.with_suffix(".key")
    if not key_file.exists():
        return None

    cert_pem = cert_file.read_bytes()
    key_pem = key_file.read_bytes()
    cuit = _cuit_del_certificado(cert_pem)
    return cert_pem, key_pem, cuit


def _cuit_del_certificado(cert_pem: bytes) -> str:
    """Extrae el CUIT del campo serialNumber del subject del certificado."""
    from cryptography import x509
    from cryptography.x509.oid import NameOID

    cert = x509.load_pem_x509_certificate(cert_pem)
    for attr in cert.subject:
        if attr.oid == NameOID.SERIAL_NUMBER:
            cuit = "".join(c for c in attr.value if c.isdigit())
            if len(cuit) == 11:
                return cuit
    raise ValueError("El certificado no tiene un CUIT (serialNumber) válido en su subject.")


def cargar() -> tuple[bytes, bytes, str]:
    """
    Devuelve ``(certificado_pem, clave_privada_pem, cuit_representada)``.
    Lanza ``RuntimeError`` si no encuentra credenciales en ningún lado.
    """
    for fuente in (_desde_secrets, _desde_archivos):
        resultado = fuente()
        if resultado is not None:
            return resultado
    raise RuntimeError(
        "No se encontró el certificado de ARCA. "
        "Configurá ``st.secrets['arca']`` o dejá los archivos `.crt` y `.key` "
        "en la carpeta `certificados/`."
    )
