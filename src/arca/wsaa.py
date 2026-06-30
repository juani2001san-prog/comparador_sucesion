"""
WSAA — Web Service de Autenticación y Autorización de ARCA.

Antes de poder consultar cualquier web service (Padrón, Constancia, Mis
Comprobantes, etc.) hay que pedirle a este servicio un **token + sign**, que
duran 12 horas.

El flujo es:

  1. Armar un **TRA** (Ticket Request Access): un XML con un id único, hora de
     generación, hora de expiración y el nombre del servicio que vamos a usar.
  2. Firmarlo con la **clave privada + certificado** del estudio, usando CMS
     (PKCS#7) en formato DER.
  3. Codificarlo en base64 y mandarlo a WSAA dentro de un sobre SOAP.
  4. Parsear la respuesta y quedarse con ``token`` y ``sign``.

El token se cachea en memoria mientras dure el proceso (no hay sentido pedir
otro si no expiró). Si en producción levantamos varios workers, cada uno se
hace su propio cache — no importa mucho porque AFIP no tiene rate limit estricto.
"""

from __future__ import annotations

import base64
import datetime as _dt
import xml.etree.ElementTree as ET
from dataclasses import dataclass

import requests
from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.serialization.pkcs7 import (
    PKCS7Options,
    PKCS7SignatureBuilder,
)


WSAA_URL_PRODUCCION = "https://wsaa.afip.gov.ar/ws/services/LoginCms"
WSAA_URL_HOMOLOGACION = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms"


@dataclass
class Credenciales:
    """Token + sign devueltos por WSAA. Duran ~12 horas."""
    token: str
    sign: str
    expira_en: _dt.datetime  # UTC

    def vigente(self, margen_segundos: int = 60) -> bool:
        ahora = _dt.datetime.now(_dt.timezone.utc)
        return ahora < self.expira_en - _dt.timedelta(seconds=margen_segundos)


# Cache simple en memoria: (servicio, cuit) → Credenciales.
_cache: dict[tuple[str, str], Credenciales] = {}


def _generar_tra(servicio: str, expira_en_segundos: int = 600) -> bytes:
    ahora = _dt.datetime.now(_dt.timezone.utc)
    gen = ahora.strftime("%Y-%m-%dT%H:%M:%SZ")
    exp = (ahora + _dt.timedelta(seconds=expira_en_segundos)).strftime("%Y-%m-%dT%H:%M:%SZ")
    unique_id = int(ahora.timestamp())
    xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<loginTicketRequest version="1.0">'
        f"<header><uniqueId>{unique_id}</uniqueId>"
        f"<generationTime>{gen}</generationTime>"
        f"<expirationTime>{exp}</expirationTime></header>"
        f"<service>{servicio}</service>"
        "</loginTicketRequest>"
    )
    return xml.encode("utf-8")


def _firmar(tra: bytes, cert_pem: bytes, key_pem: bytes) -> bytes:
    cert = x509.load_pem_x509_certificate(cert_pem)
    clave = serialization.load_pem_private_key(key_pem, password=None)
    return (
        PKCS7SignatureBuilder()
        .set_data(tra)
        .add_signer(cert, clave, hashes.SHA256())
        .sign(serialization.Encoding.DER, [PKCS7Options.Binary, PKCS7Options.NoCapabilities])
    )


def _parsear_respuesta(soap_xml: str) -> Credenciales:
    """Extrae token + sign + expirationTime del SOAP que devuelve WSAA."""
    root = ET.fromstring(soap_xml)

    # Buscamos el <loginCmsReturn> sin importar el namespace.
    contenido = None
    for el in root.iter():
        if el.tag.endswith("loginCmsReturn"):
            contenido = el.text
            break
    if not contenido:
        # Capaz vino un soap:Fault
        for el in root.iter():
            if el.tag.endswith("faultstring"):
                raise RuntimeError(f"WSAA devolvió error: {el.text}")
        raise RuntimeError(f"Respuesta inesperada de WSAA:\n{soap_xml[:500]}")

    interno = ET.fromstring(contenido)
    token = (interno.findtext(".//token") or "").strip()
    sign = (interno.findtext(".//sign") or "").strip()
    exp_text = (interno.findtext(".//header/expirationTime") or "").strip()
    if not (token and sign and exp_text):
        raise RuntimeError(f"Respuesta de WSAA incompleta: {contenido}")

    # exp_text viene en formato ISO con timezone (ej. 2026-06-30T18:30:00.000-03:00).
    expira = _dt.datetime.fromisoformat(exp_text).astimezone(_dt.timezone.utc)
    return Credenciales(token=token, sign=sign, expira_en=expira)


def obtener_credenciales(
    servicio: str,
    cert_pem: bytes,
    key_pem: bytes,
    cuit_titular: str,
    url: str = WSAA_URL_PRODUCCION,
) -> Credenciales:
    """
    Devuelve un par token+sign válido para usar con ``servicio``.

    Reutiliza el cache si todavía no expiraron. Si no, pide uno nuevo a WSAA.
    """
    clave_cache = (servicio, cuit_titular)
    cred = _cache.get(clave_cache)
    if cred is not None and cred.vigente():
        return cred

    tra = _generar_tra(servicio)
    cms = _firmar(tra, cert_pem, key_pem)
    cms_b64 = base64.b64encode(cms).decode("ascii")

    sobre = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" '
        'xmlns:wsaa="http://wsaa.view.sua.dvadac.desein.afip.gov">'
        "<soapenv:Header/>"
        "<soapenv:Body>"
        "<wsaa:loginCms><wsaa:in0>" + cms_b64 + "</wsaa:in0></wsaa:loginCms>"
        "</soapenv:Body>"
        "</soapenv:Envelope>"
    )

    resp = requests.post(
        url,
        data=sobre.encode("utf-8"),
        headers={
            "Content-Type": "text/xml; charset=utf-8",
            "SOAPAction": "",
        },
        timeout=30,
    )
    resp.raise_for_status()
    cred = _parsear_respuesta(resp.text)
    _cache[clave_cache] = cred
    return cred
