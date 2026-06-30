"""
Consulta de Constancia de Inscripción / Padrón A5 a ARCA.

El servicio que autorizamos en el portal se llama
``ws_sr_constancia_inscripcion`` y devuelve, dado un CUIT, los mismos datos
que ves en la "Constancia de Inscripción" del sitio web:

- Razón social / nombre y apellido
- Tipo de persona (Física / Jurídica)
- Estado de la clave fiscal (Activo / Inactivo)
- Domicilio fiscal
- Categoría de monotributo (si aplica)
- Condición frente al IVA
- Actividades económicas registradas
- Impuestos inscritos
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from typing import Any

import requests


CONSTANCIA_URL = "https://aws.afip.gob.ar/sr-padron/webservices/personaServiceA5"
SERVICIO = "ws_sr_constancia_inscripcion"


def consultar(
    token: str,
    sign: str,
    cuit_titular: str,
    cuit_a_consultar: str,
    url: str = CONSTANCIA_URL,
) -> dict[str, Any]:
    """
    Consulta los datos del padrón para un CUIT determinado.

    ``cuit_titular`` es el CUIT del estudio (vos), y ``cuit_a_consultar`` es
    el CUIT que querés ver (vos o un cliente).
    """
    sobre = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" '
        'xmlns:a5="http://a5.soap.ws.server.puc.sr/">'
        "<soapenv:Header/>"
        "<soapenv:Body>"
        "<a5:getPersona_v2>"
        f"<token>{token}</token>"
        f"<sign>{sign}</sign>"
        f"<cuitRepresentada>{cuit_titular}</cuitRepresentada>"
        f"<idPersona>{cuit_a_consultar}</idPersona>"
        "</a5:getPersona_v2>"
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
    return _parsear_persona(resp.text, cuit_a_consultar)


# --------------------------------------------------------------------------- #
# Parser de la respuesta
# --------------------------------------------------------------------------- #

def _texto(el: ET.Element, tag: str) -> str | None:
    """Busca el primer descendiente con esa tag (ignora namespace)."""
    for child in el.iter():
        if child.tag.split("}")[-1] == tag:
            return (child.text or "").strip() or None
    return None


def _lista(el: ET.Element, tag: str) -> list[ET.Element]:
    return [c for c in el.iter() if c.tag.split("}")[-1] == tag]


def _parsear_persona(xml_str: str, cuit_consultado: str) -> dict[str, Any]:
    root = ET.fromstring(xml_str)

    # Buscamos el elemento <persona> en cualquier parte del árbol.
    persona = next((el for el in root.iter() if el.tag.split("}")[-1] == "persona"), None)
    if persona is None:
        # Capaz vino un soap:Fault
        for el in root.iter():
            if el.tag.split("}")[-1] == "faultstring":
                raise RuntimeError(f"ARCA devolvió error: {el.text}")
        raise RuntimeError(f"ARCA no devolvió datos para CUIT {cuit_consultado}.\n{xml_str[:500]}")

    datos: dict[str, Any] = {
        "cuit": _texto(persona, "idPersona") or cuit_consultado,
        "tipo_persona": _texto(persona, "tipoPersona"),
        "tipo_clave": _texto(persona, "tipoClave"),
        "estado_clave": _texto(persona, "estadoClave"),
        "razon_social": _texto(persona, "razonSocial"),
        "nombre": _texto(persona, "nombre"),
        "apellido": _texto(persona, "apellido"),
        "fecha_nacimiento": _texto(persona, "fechaNacimiento"),
        "fecha_inscripcion": _texto(persona, "fechaInscripcion"),
        "mes_cierre": _texto(persona, "mesCierre"),
    }

    # Domicilio fiscal
    domicilios = _lista(persona, "domicilio")
    if domicilios:
        d = domicilios[0]
        datos["domicilio"] = {
            "tipo": _texto(d, "tipoDomicilio"),
            "direccion": _texto(d, "direccion"),
            "localidad": _texto(d, "localidad"),
            "provincia": _texto(d, "descripcionProvincia"),
            "codigo_postal": _texto(d, "codPostal"),
        }

    # Categorías de monotributo y otros impuestos
    categorias = []
    for cat in _lista(persona, "categoria"):
        categorias.append({
            "idImpuesto": _texto(cat, "idImpuesto"),
            "descripcion": _texto(cat, "descripcionCategoria"),
            "estado": _texto(cat, "estado"),
            "periodo": _texto(cat, "periodo"),
        })
    if categorias:
        datos["categorias"] = categorias

    # Impuestos (RI, etc.)
    impuestos = []
    for imp in _lista(persona, "impuesto"):
        impuestos.append({
            "idImpuesto": _texto(imp, "idImpuesto"),
            "descripcion": _texto(imp, "descripcionImpuesto"),
            "estado": _texto(imp, "estado"),
            "periodo": _texto(imp, "periodo"),
        })
    if impuestos:
        datos["impuestos"] = impuestos

    # Actividades económicas
    actividades = []
    for act in _lista(persona, "actividad"):
        actividades.append({
            "id": _texto(act, "idActividad"),
            "descripcion": _texto(act, "descripcionActividad"),
            "nomenclador": _texto(act, "nomenclador"),
            "periodo": _texto(act, "periodo"),
            "orden": _texto(act, "orden"),
        })
    if actividades:
        datos["actividades"] = actividades

    return datos
