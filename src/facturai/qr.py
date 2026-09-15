"""
Lectura del QR de AFIP en fotos y PDFs de facturas.

Toda factura electrónica argentina emitida desde 2020 lleva un QR en la
esquina inferior derecha con una URL de AFIP tipo:

    https://www.afip.gob.ar/fe/qr/?p=<base64url del JSON del comprobante>

Este módulo:
  1. Extrae la imagen (o convierte cada página del PDF a imagen).
  2. Detecta el QR con OpenCV.
  3. Parsea la URL de AFIP y devuelve los datos del comprobante en un dict.

Para las facturas SIN QR (tickets viejos, comprobantes manuales) se planea
un fallback con LLM Vision en una segunda etapa.
"""

from __future__ import annotations

import base64
import io
import json
from urllib.parse import parse_qs, urlparse


# --------------------------------------------------------------------------- #
# Tablas de códigos de AFIP (los que aparecen en el QR)
# --------------------------------------------------------------------------- #

TIPOS_COMPROBANTE = {
    1:   "Factura A",
    2:   "Nota de Débito A",
    3:   "Nota de Crédito A",
    4:   "Recibo A",
    5:   "Nota de Venta al Contado A",
    6:   "Factura B",
    7:   "Nota de Débito B",
    8:   "Nota de Crédito B",
    9:   "Recibo B",
    10:  "Nota de Venta al Contado B",
    11:  "Factura C",
    12:  "Nota de Débito C",
    13:  "Nota de Crédito C",
    15:  "Recibo C",
    19:  "Factura E (Exportación)",
    20:  "Nota de Débito E",
    21:  "Nota de Crédito E",
    51:  "Factura M",
    52:  "Nota de Débito M",
    53:  "Nota de Crédito M",
    54:  "Recibo M",
    81:  "Tique Factura A",
    82:  "Tique Factura B",
    83:  "Tique",
    111: "Tique Factura C",
    118: "Tique Nota de Crédito C",
    201: "Factura de Crédito Electrónica MiPyMEs A",
    202: "Nota de Débito Electrónica MiPyMEs A",
    203: "Nota de Crédito Electrónica MiPyMEs A",
    206: "Factura de Crédito Electrónica MiPyMEs B",
    207: "Nota de Débito Electrónica MiPyMEs B",
    208: "Nota de Crédito Electrónica MiPyMEs B",
    211: "Factura de Crédito Electrónica MiPyMEs C",
    212: "Nota de Débito Electrónica MiPyMEs C",
    213: "Nota de Crédito Electrónica MiPyMEs C",
}

TIPOS_DOC_RECEPTOR = {
    80: "CUIT",
    86: "CUIL",
    87: "CDI",
    89: "LE",
    90: "LC",
    91: "CI Extranjera",
    92: "En trámite",
    93: "Acta de Nacimiento",
    94: "Pasaporte",
    95: "CI Bs. As. RNP",
    96: "DNI",
    99: "Consumidor Final",
}

MONEDAS = {
    "PES": "Pesos argentinos",
    "DOL": "Dólares estadounidenses",
    "EUR": "Euros",
}


class QrError(Exception):
    """Error al procesar una imagen o PDF de factura."""


# --------------------------------------------------------------------------- #
# Parseo de la URL de AFIP
# --------------------------------------------------------------------------- #

def parsear_url_afip(url: str) -> dict | None:
    """
    Dada la URL codificada en el QR de una factura AFIP, devuelve los datos
    del comprobante en un dict. Devuelve None si la URL no es de AFIP o no
    se puede decodificar.
    """
    if not url or "afip" not in url.lower():
        return None
    try:
        query = urlparse(url).query
        params = parse_qs(query)
        payload_b64 = (params.get("p") or [None])[0]
        if not payload_b64:
            return None
        # Base64 URL-safe: agregar padding si falta
        payload_b64 = payload_b64 + "=" * ((-len(payload_b64)) % 4)
        raw = base64.urlsafe_b64decode(payload_b64)
        data = json.loads(raw)
    except (json.JSONDecodeError, ValueError, TypeError):
        return None

    tipo_cmp = data.get("tipoCmp")
    tipo_doc = data.get("tipoDocRec")

    return {
        "cuit_emisor": _cuit_texto(data.get("cuit")),
        "punto_venta": data.get("ptoVta"),
        "numero": data.get("nroCmp"),
        "tipo_comprobante_codigo": tipo_cmp,
        "tipo_comprobante": TIPOS_COMPROBANTE.get(tipo_cmp) or f"Tipo {tipo_cmp}",
        "fecha": data.get("fecha"),
        "importe_total": data.get("importe"),
        "moneda": data.get("moneda"),
        "moneda_descripcion": MONEDAS.get(data.get("moneda"), data.get("moneda")),
        "cotizacion": data.get("ctz"),
        "tipo_doc_receptor_codigo": tipo_doc,
        "tipo_doc_receptor": TIPOS_DOC_RECEPTOR.get(tipo_doc) or f"Tipo {tipo_doc}",
        "nro_doc_receptor": _cuit_texto(data.get("nroDocRec")),
        "tipo_codigo_autorizacion": data.get("tipoCodAut"),  # 'E' = CAE, 'A' = CAEA
        "codigo_autorizacion": data.get("codAut"),
        "version_qr": data.get("ver"),
        "url_qr": url,
    }


def _cuit_texto(v) -> str:
    if v is None or v == "":
        return ""
    return str(v)


def formato_cuit(cuit: str) -> str:
    s = "".join(c for c in str(cuit or "") if c.isdigit())
    if len(s) == 11:
        return f"{s[:2]}-{s[2:10]}-{s[10]}"
    return str(cuit or "")


# --------------------------------------------------------------------------- #
# Detección del QR en imagen y PDF
# --------------------------------------------------------------------------- #

def _detectar_qr_en_imagen_bytes(image_bytes: bytes) -> str | None:
    """Corre el detector de OpenCV sobre una imagen. Devuelve la URL del QR o None."""
    import cv2
    import numpy as np
    from PIL import Image

    img = Image.open(io.BytesIO(image_bytes))
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    arr = np.array(img)
    if arr.ndim == 3:
        # PIL RGB → OpenCV BGR
        arr = cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)

    detector = cv2.QRCodeDetector()
    data, _points, _straight = detector.detectAndDecode(arr)
    if data:
        return data

    # Si falla, probar con el detector "multi" que a veces es más robusto
    ok, datos, _pts, _straight = detector.detectAndDecodeMulti(arr)
    if ok and datos:
        for d in datos:
            if d:
                return d
    return None


def _pdf_a_imagenes(pdf_bytes: bytes, dpi: int = 200) -> list[bytes]:
    """Convierte cada página del PDF a PNG. Devuelve lista de bytes."""
    import fitz  # PyMuPDF

    imagenes: list[bytes] = []
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        for page in doc:
            pix = page.get_pixmap(dpi=dpi)
            imagenes.append(pix.tobytes("png"))
    return imagenes


def procesar_archivo(data: bytes) -> dict:
    """
    Procesa una imagen (PNG/JPG/WebP) o un PDF de una factura.

    Detecta el QR de AFIP y devuelve un dict con:
        - ok: bool
        - datos: dict con los campos del comprobante (si ok)
        - error: str (si no ok)
        - imagen_preview: bytes de la primera imagen procesada
        - pagina_detectada: int (para PDF multi-página, 1-based)
        - url_qr: str raw del QR
    """
    if not data:
        return {"ok": False, "error": "Archivo vacío.", "imagen_preview": None}

    es_pdf = data[:4] == b"%PDF"

    if es_pdf:
        try:
            imagenes = _pdf_a_imagenes(data)
        except Exception as exc:  # noqa: BLE001
            return {
                "ok": False,
                "error": f"No se pudo leer el PDF: {exc}",
                "imagen_preview": None,
            }
        if not imagenes:
            return {
                "ok": False,
                "error": "El PDF no tiene páginas.",
                "imagen_preview": None,
            }
    else:
        imagenes = [data]

    for i, img_bytes in enumerate(imagenes, start=1):
        try:
            url = _detectar_qr_en_imagen_bytes(img_bytes)
        except Exception as exc:  # noqa: BLE001
            return {
                "ok": False,
                "error": f"Error al leer la imagen: {exc}",
                "imagen_preview": img_bytes,
            }
        if url:
            datos = parsear_url_afip(url)
            if datos:
                return {
                    "ok": True,
                    "datos": datos,
                    "url_qr": url,
                    "imagen_preview": img_bytes,
                    "pagina_detectada": i,
                }
            # QR encontrado pero no es de AFIP
            return {
                "ok": False,
                "error": f"Se detectó un QR pero no parece de AFIP:\n{url[:200]}",
                "imagen_preview": img_bytes,
                "url_qr": url,
                "pagina_detectada": i,
            }

    return {
        "ok": False,
        "error": ("No se detectó QR de AFIP en la imagen. "
                  "Puede ser una foto poco nítida, una factura muy vieja "
                  "sin QR obligatorio, o un ticket no fiscal."),
        "imagen_preview": imagenes[0] if imagenes else None,
    }
