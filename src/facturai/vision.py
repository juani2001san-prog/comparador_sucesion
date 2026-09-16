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


# Google renombra modelos seguido y va deprecando los anteriores.
# Probamos en cascada — el primero que responde OK se usa. Podés
# también forzar uno específico configurándolo en st.secrets['gemini']['modelo'].
MODELOS_FALLBACK = [
    "gemini-flash-latest",         # alias oficial que sigue al modelo Flash actual
    "gemini-2.0-flash",            # estable y con free tier propio
    "gemini-2.0-flash-exp",        # cuota generosa historicamente
    "gemini-2.0-flash-lite",
    "gemini-1.5-flash-latest",     # muy estable
    "gemini-1.5-flash",
    "gemini-1.5-flash-8b",         # variante lite con su propia cuota
    "gemini-2.5-flash",
    "gemini-2.5-flash-lite",
    "gemini-3.5-flash-lite",
]
MODELO_DEFAULT = MODELOS_FALLBACK[0]

_PROMPT = """Analizá esta imagen: puede contener UNA o VARIAS facturas/tickets fiscales
argentinos (típico caso: varios tickets escaneados juntos en una misma hoja).

Devolveme SOLO un JSON válido, sin markdown ni backticks, con esta estructura:

{
  "facturas": [
    {
      "tipo_comprobante": "Factura A" | "Factura B" | "Factura C" | "Tique Factura A" | "Tique Factura B" | "Tique Factura C" | "Ticket" | "Nota de Crédito A" | "Nota de Crédito B" | "Nota de Crédito C" | "Nota de Débito A" | "Nota de Débito B" | "Nota de Débito C" | "Factura M" | "Recibo" | "Otro",
      "codigo_tipo": 1,
      "cuit_emisor": "20111746525",
      "razon_social_emisor": "Denominación del emisor tal cual aparece",
      "condicion_iva_emisor": "Responsable Inscripto" | "Monotributista" | "Exento" | "Consumidor Final" | null,
      "punto_venta": 5,
      "numero": 486,
      "fecha": "2026-08-24",
      "cuit_receptor": "20437400432",
      "razon_social_receptor": "Nombre del receptor",
      "importe_neto_gravado": 22314.05,
      "importe_iva": 4685.95,
      "alicuota_iva": 21,
      "importe_no_gravado": 0,
      "importe_exento": 0,
      "importe_percepciones_iibb": 0,
      "importe_impuestos_internos": 0,
      "importe_otros_tributos": 0,
      "importe_total": 27000.00,
      "cae": "86349754031008",
      "moneda": "PES"
    }
  ]
}

Códigos AFIP (codigo_tipo): 1=Fact A, 6=Fact B, 11=Fact C, 3/8/13=NC A/B/C,
2/7/12=ND A/B/C, 81=Tique Fact A, 82=Tique Fact B, 111=Tique Fact C,
51/52/53=Fact/ND/NC M, 4/9/15=Recibo A/B/C.

CLASIFICACIÓN DE IMPUESTOS Y PERCEPCIONES — MUY IMPORTANTE:

- **importe_percepciones_iibb**: SOLO las percepciones de INGRESOS BRUTOS
  (típicamente aparecen como "Percep. IIBB", "Percepción IB", "IIBB",
  "Ingresos Brutos" seguido de un porcentaje).
- **importe_impuestos_internos**: los IMPUESTOS INTERNOS. Ejemplos típicos
  en tickets de COMBUSTIBLES:
    · ITC (Impuesto a la Transferencia de Combustibles).
    · IDC (Impuesto al Dióxido de Carbono).
    · Otros impuestos internos (art. 24, IIL, tabaco, etc.).
  Estos NUNCA son percepciones — van a Impuestos Internos.
- **importe_otros_tributos**: cualquier otro tributo que no encaja en las
  categorías anteriores (impuestos municipales, tasas, contribuciones).

Ejemplo concreto de un ticket de nafta con importes:
  ITC: 3911,45  →  importe_impuestos_internos suma 3911.45
  IDC: 445,81   →  importe_impuestos_internos suma 445.81
  (no hay percepciones IIBB en este ticket → importe_percepciones_iibb = 0)

Reglas:
- Devolvé SIEMPRE la clave 'facturas' con un array (aunque haya solo un comprobante).
- Si la imagen tiene VARIAS facturas visibles (tickets pegados uno al lado
  del otro en un scan), listalas todas — una por elemento del array.
- Los importes son números decimales con PUNTO decimal (no coma). Sin símbolo $.
- Si un campo no aparece o no es legible, devolvé null.
- El CUIT emisor es SIEMPRE quien emite (arriba del ticket), no el receptor.

FACTURAS A (código 1, 2, 3) - MUY IMPORTANTE:
- Las Facturas A SIEMPRE discriminan Neto Gravado + IVA. NUNCA devuelvas
  importe_neto_gravado=0 e importe_iva=0 en una Factura A que tenga total.
- Si el desglose no se ve claro en la imagen pero SÍ ves el importe total
  y la alícuota, CALCULÁ los valores tú:
    importe_neto_gravado = importe_total / (1 + alicuota/100)
    importe_iva = importe_total - importe_neto_gravado
  Ejemplo: Factura A al 21%, total 45000. Entonces:
    importe_neto_gravado = 45000 / 1.21 = 37190.08
    importe_iva = 45000 - 37190.08 = 7809.92
- Si no se ve la alícuota, asumí 21% (la más común).

TICKETS B (81, 82) y C (11, 111) - MUY IMPORTANTE:
- Los tickets a consumidor final NO discriminan IVA en el ticket.
- Para estos, importe_iva puede ser 0 y importe_neto_gravado = importe_total.

Devolvé exclusivamente el JSON, nada más.
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


def _modelos_a_probar():
    """Modelo forzado en secrets (si existe) + cascada de fallback."""
    modelo_forzado = None
    if st is not None:
        try:
            m = st.secrets["gemini"].get("modelo")
            if m:
                modelo_forzado = str(m).strip()
        except Exception:
            pass
    if modelo_forzado:
        return [modelo_forzado] + [m for m in MODELOS_FALLBACK if m != modelo_forzado]
    return list(MODELOS_FALLBACK)


# Cache en memoria: cuando un modelo respondió OK, se usa ese en las
# siguientes llamadas — evita perder tiempo probando modelos deprecados.
_ultimo_modelo_ok: str | None = None


def extraer_datos(image_bytes: bytes, modelo: str | None = None) -> list[dict]:
    """
    Envía la imagen a Gemini y devuelve la lista de comprobantes encontrados.

    Siempre devuelve una lista (aunque sea de un solo elemento), porque la
    imagen puede tener varias facturas escaneadas juntas.

    Prueba varios modelos en cascada — si el primero está deprecado (Google
    los va renombrando), pasa al siguiente. El que funciona queda cacheado.
    """
    global _ultimo_modelo_ok

    from PIL import Image

    genai = _configurar_cliente()

    img = Image.open(io.BytesIO(image_bytes))
    max_lado = 2000
    if max(img.size) > max_lado:
        img.thumbnail((max_lado, max_lado))

    if modelo is not None:
        candidatos = [modelo]
    else:
        candidatos = _modelos_a_probar()
        if _ultimo_modelo_ok and _ultimo_modelo_ok in candidatos:
            # Mover al frente el último que funcionó
            candidatos.remove(_ultimo_modelo_ok)
            candidatos.insert(0, _ultimo_modelo_ok)

    ultimo_error = None
    for nombre in candidatos:
        try:
            modelo_ia = genai.GenerativeModel(nombre)
            resp = modelo_ia.generate_content([_PROMPT, img])
            texto = (resp.text or "").strip()
            data = _parsear_json_respuesta(texto)
            _ultimo_modelo_ok = nombre
            break
        except Exception as exc:  # noqa: BLE001
            msg = str(exc).lower()
            # Saltar al siguiente si es error de modelo (404), deprecación,
            # cuota agotada (429), o límite del free tier. En el free tier
            # cada modelo tiene su propia cuota diaria — es razonable ir
            # probando distintos si uno se agotó.
            recuperable = (
                "404" in msg
                or "not found" in msg
                or "no longer available" in msg
                or "429" in msg
                or "quota" in msg
                or "resource_exhausted" in msg
                or "rate limit" in msg
                or "free_tier" in msg
            )
            if recuperable:
                ultimo_error = exc
                continue
            raise
    else:
        raise RuntimeError(
            f"Ningún modelo de Gemini respondió OK. Último error: {ultimo_error}"
        )

    # Aceptamos varios shapes por si Gemini responde ligeramente distinto:
    #   {"facturas": [...]}  → normal
    #   [...]                → array pelado
    #   {..campos..}         → un solo dict (formato viejo)
    if isinstance(data, dict) and "facturas" in data:
        facturas = data["facturas"]
    elif isinstance(data, list):
        facturas = data
    elif isinstance(data, dict):
        facturas = [data]
    else:
        facturas = []

    return [f for f in facturas if isinstance(f, dict)]


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
        # Compatibilidad con respuestas viejas del prompt anterior.
        "importe_percepciones_iibb": (
            datos_vision.get("importe_percepciones_iibb")
            or datos_vision.get("importe_percepciones")
        ),
        "importe_impuestos_internos": datos_vision.get("importe_impuestos_internos"),
        "importe_otros_tributos": datos_vision.get("importe_otros_tributos"),
        "fuente": "gemini",
    }
