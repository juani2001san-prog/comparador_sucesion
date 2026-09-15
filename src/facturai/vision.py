"""
Extracción de datos de facturas por Gemini Vision.

Se usa como fallback cuando el detector de QR falla:

  - Factura con QR ilegible por borrosidad, papel arrugado o QR muy chico.
  - Ticket fiscal viejo emitido por controlador fiscal (no tiene QR).
  - Comprobantes manuales, tickets no fiscales, facturas escaneadas viejas.

Gemini "mira" la imagen entera y devuelve un JSON estructurado con los
campos que necesitamos para armar la fila del Excel tipo Portal IVA.

Requiere una API key gratuita de Google AI Studio configurada en
``st.secrets["gemini"]["api_key"]``. Tier gratuito de Gemini 2.5 Flash:
15 requests/minuto, más que suficiente para el volumen mensual del estudio.
"""

from __future__ import annotations

import io
import json
import re

try:
    import streamlit as st
except ImportError:  # pragma: no cover
    st = None  # type: ignore


MODELO_DEFAULT = "gemini-3.6-flash"

_PROMPT = """Analizá esta imagen de una factura o ticket fiscal argentino y devolveme
SOLO un JSON válido con la siguiente estructura, sin texto adicional, sin markdown,
sin backticks:

{
  "tipo_comprobante": "Factura A" | "Factura B" | "Factura C" | "Tique Factura A" | "Tique Factura B" | "Tique Factura C" | "Ticket" | "Nota de Crédito A" | "Nota de Crédito B" | "Nota de Crédito C" | "Nota de Débito A" | "Nota de Débito B" | "Nota de Débito C" | "Factura M" | "Recibo" | "Otro",
  "codigo_tipo": 1 (código AFIP: 1=Fact A, 6=Fact B, 11=Fact C, 3/8/13=NC, 2/7/12=ND, 81/82/83=Tique, 111=Tique C, etc.),
  "cuit_emisor": "20111746525" (solo los 11 dígitos, sin guiones),
  "razon_social_emisor": "Denominación del emisor tal cual aparece",
  "condicion_iva_emisor": "Responsable Inscripto" | "Monotributista" | "Exento" | "Consumidor Final" | null,
  "punto_venta": 5 (número entero),
  "numero": 486 (número entero del comprobante),
  "fecha": "2026-08-24" (formato YYYY-MM-DD),
  "cuit_receptor": "20437400432" o null si es consumidor final,
  "razon_social_receptor": "Nombre del receptor" o null,
  "importe_neto_gravado": 22314.05 (subtotal sin IVA, si discrimina),
  "importe_iva": 4685.95 (IVA total sumado de todas las alícuotas),
  "alicuota_iva": 21 (porcentaje predominante: 21, 10.5, 27, 5, 2.5, 0),
  "importe_no_gravado": 0,
  "importe_exento": 0,
  "importe_percepciones": 0 (suma de todas las percepciones si aparecen),
  "importe_total": 27000.00 (importe final total),
  "cae": "86349754031008" o null si no aparece,
  "moneda": "PES"
}

Reglas:
- Los importes son números decimales con PUNTO decimal (no coma). Sin símbolo $.
- Si algún campo no aparece o no es legible, devolvé null.
- Si es un ticket B o C que no discrimina IVA, el importe_iva puede ser 0 y el neto igual al total.
- El CUIT del emisor es SIEMPRE quien emite (arriba de la factura), no el receptor.
- Devolvé exclusivamente el JSON, nada más.
"""


def hay_conexion() -> bool:
    """True si la API key de Gemini está configurada en los secrets."""
    if st is None:
        return False
    try:
        return bool(st.secrets["gemini"]["api_key"])
    except Exception:
        return False


def _configurar_cliente():
    import google.generativeai as genai
    if st is None:
        raise RuntimeError("Streamlit no está disponible para leer la API key.")
    api_key = st.secrets["gemini"]["api_key"]
    genai.configure(api_key=api_key)
    return genai


def extraer_datos(image_bytes: bytes, modelo: str = MODELO_DEFAULT) -> dict:
    """
    Envía la imagen a Gemini y devuelve los datos del comprobante como dict.

    Lanza ``RuntimeError`` si la API responde con error o el JSON es inválido.
    """
    from PIL import Image

    genai = _configurar_cliente()
    modelo_ia = genai.GenerativeModel(modelo)

    img = Image.open(io.BytesIO(image_bytes))
    # Reducir tamaño si es enorme (Gemini acepta hasta 20MB pero se cobra más).
    max_lado = 2000
    if max(img.size) > max_lado:
        img.thumbnail((max_lado, max_lado))

    resp = modelo_ia.generate_content([_PROMPT, img])
    texto = (resp.text or "").strip()
    return _parsear_json_respuesta(texto)


def _parsear_json_respuesta(texto: str) -> dict:
    """Parsea el JSON que devuelve Gemini, tolerando envoltorios de markdown."""
    if not texto:
        raise RuntimeError("Gemini devolvió una respuesta vacía.")

    # A veces envuelve la respuesta en ```json ... ``` a pesar del prompt.
    m = re.search(r"```(?:json)?\s*(.*?)\s*```", texto, re.DOTALL)
    if m:
        texto = m.group(1).strip()

    try:
        return json.loads(texto)
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            f"No se pudo parsear la respuesta de Gemini como JSON: {exc}\n"
            f"Texto recibido (primeros 300 chars): {texto[:300]}"
        )


# --------------------------------------------------------------------------- #
# Adaptación al formato interno (mismo shape que devuelve qr.parsear_url_afip)
# --------------------------------------------------------------------------- #

def _cuit_solo_digitos(v) -> str:
    return "".join(c for c in str(v or "") if c.isdigit())


def a_formato_qr(datos_vision: dict) -> dict:
    """
    Convierte la respuesta de Gemini al mismo shape que devuelve el parser de QR,
    así el resto del código (armado de fila, tabla, Excel) no se entera de qué
    vino de dónde.
    """
    cuit_emisor = _cuit_solo_digitos(datos_vision.get("cuit_emisor"))
    codigo_tipo = datos_vision.get("codigo_tipo")

    return {
        "cuit_emisor": cuit_emisor,
        "razon_social_emisor": datos_vision.get("razon_social_emisor"),
        "condicion_iva_emisor": datos_vision.get("condicion_iva_emisor"),
        "punto_venta": datos_vision.get("punto_venta"),
        "numero": datos_vision.get("numero"),
        "tipo_comprobante_codigo": codigo_tipo,
        "tipo_comprobante": datos_vision.get("tipo_comprobante"),
        "fecha": datos_vision.get("fecha"),
        "importe_total": datos_vision.get("importe_total"),
        "moneda": datos_vision.get("moneda") or "PES",
        "cotizacion": 1,
        "nro_doc_receptor": _cuit_solo_digitos(datos_vision.get("cuit_receptor")),
        "razon_social_receptor": datos_vision.get("razon_social_receptor"),
        "codigo_autorizacion": datos_vision.get("cae"),
        # Campos que NO vienen del QR pero sí de Vision:
        "importe_neto_gravado": datos_vision.get("importe_neto_gravado"),
        "importe_iva": datos_vision.get("importe_iva"),
        "alicuota_iva": datos_vision.get("alicuota_iva"),
        "importe_no_gravado": datos_vision.get("importe_no_gravado"),
        "importe_exento": datos_vision.get("importe_exento"),
        "importe_percepciones": datos_vision.get("importe_percepciones"),
        "fuente": "gemini",
    }
