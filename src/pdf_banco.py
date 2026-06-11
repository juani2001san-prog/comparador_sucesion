"""
Conversión de extractos bancarios PDF a tabla.

Reutiliza la lógica del proyecto `pdf-to-excel` (pdfplumber + parsers por banco),
consolidada en un solo archivo para que el programa sea autocontenido.

Soporta PDFs con texto seleccionable (no escaneados). Incluye un parser
específico de **Banco del Chubut** y un parser genérico de respaldo.
"""

from __future__ import annotations

import io
import re
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from typing import List, Optional, Tuple

import pandas as pd

try:
    import pdfplumber
except ImportError:  # pragma: no cover
    pdfplumber = None


# --------------------------------------------------------------------------- #
# Modelos
# --------------------------------------------------------------------------- #

@dataclass
class Movimiento:
    fecha: str
    descripcion: str
    debito: Optional[Decimal] = None
    credito: Optional[Decimal] = None
    saldo: Optional[Decimal] = None


Tabla = List[List[Optional[str]]]


@dataclass
class PdfExtraido:
    lines: List[str]
    pages_tables: list = field(default_factory=list)


class ErrorConversion(Exception):
    """Error que se muestra al usuario tal cual."""


# --------------------------------------------------------------------------- #
# Extracción
# --------------------------------------------------------------------------- #

def extraer_pdf(pdf_bytes: bytes) -> PdfExtraido:
    """Extrae líneas de texto y tablas del PDF. Falla si es escaneado (sin texto)."""
    if pdfplumber is None:
        raise ErrorConversion("Falta instalar pdfplumber (pip install pdfplumber).")

    lines: list[str] = []
    pages_tables: list = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            texto = page.extract_text() or ""
            for raw in texto.splitlines():
                if raw.strip():
                    lines.append(raw.strip())
            pages_tables.append(page.extract_tables() or [])

    if not lines:
        raise ErrorConversion("El PDF no contiene texto legible. Puede ser un PDF escaneado.")
    return PdfExtraido(lines=lines, pages_tables=pages_tables)


# --------------------------------------------------------------------------- #
# Helpers numéricos
# --------------------------------------------------------------------------- #

def _norm(s: Optional[str]) -> str:
    return (s or "").strip()


def _a_decimal_ar(raw: Optional[str]) -> Optional[Decimal]:
    """Convierte un número en formato argentino (1.234,56) a Decimal."""
    s = _norm(raw)
    if not s:
        return None
    negativo = False
    if s.startswith("(") and s.endswith(")"):
        negativo, s = True, s[1:-1]
    if s.endswith("-"):
        negativo, s = True, s[:-1]
    if s.startswith("-"):
        negativo, s = True, s[1:]
    s = s.replace(".", "").replace(",", ".")
    try:
        valor = Decimal(s)
    except InvalidOperation:
        return None
    if valor == 0:
        return None
    return -valor if negativo else valor


# --------------------------------------------------------------------------- #
# Parsers
# --------------------------------------------------------------------------- #

class ParserBanco(ABC):
    name = "base"

    @abstractmethod
    def puede_parsear(self, pdf: PdfExtraido) -> bool: ...

    @abstractmethod
    def parsear(self, pdf: PdfExtraido) -> List[Movimiento]: ...


_DATE_RE = re.compile(r"^\d{1,2}/\d{1,2}/\d{2,4}$")


def _indices_columnas(header: List[Optional[str]]) -> Optional[dict]:
    """Mapea nombre de columna -> índice. None si no es el header esperado."""
    mapa: dict = {}
    for i, cell in enumerate(header):
        text = _norm(cell).upper()
        if not text:
            continue
        if "FECHA" in text and "fecha" not in mapa:
            mapa["fecha"] = i
        elif "CONCEPTO" in text:
            mapa["concepto"] = i
        elif "REFERENCIA" in text:
            mapa["referencia"] = i
        elif "DÉBITO" in text or "DEBITO" in text:
            mapa["debito"] = i
        elif "CRÉDITO" in text or "CREDITO" in text:
            mapa["credito"] = i
        elif "SALDO" in text:
            mapa["saldo"] = i
    if {"fecha", "debito", "credito", "saldo"}.issubset(mapa):
        return mapa
    return None


class ParserChubut(ParserBanco):
    """Banco del Chubut: lee la tabla de movimientos por columnas."""

    name = "Banco del Chubut"

    def puede_parsear(self, pdf: PdfExtraido) -> bool:
        text = "\n".join(pdf.lines[:80]).upper()
        return "BANCO DEL CHUBUT" in text or "BANCOCHUBUT.COM.AR" in text

    def parsear(self, pdf: PdfExtraido) -> List[Movimiento]:
        movs: List[Movimiento] = []
        for tablas in pdf.pages_tables:
            for tabla in tablas:
                if not tabla or len(tabla) < 2:
                    continue
                cols = _indices_columnas(tabla[0])
                if not cols:
                    continue
                movs.extend(self._parsear_tabla(tabla, cols))
        return movs

    @staticmethod
    def _parsear_tabla(tabla: Tabla, cols: dict) -> List[Movimiento]:
        result: List[Movimiento] = []
        for row in tabla[1:]:
            fecha = _norm(row[cols["fecha"]]) if cols["fecha"] < len(row) else ""
            if not _DATE_RE.match(fecha):
                continue
            concepto = _norm(row[cols["concepto"]]) if "concepto" in cols and cols["concepto"] < len(row) else ""
            referencia = _norm(row[cols["referencia"]]) if "referencia" in cols and cols["referencia"] < len(row) else ""
            descripcion = " ".join(p for p in (concepto, referencia) if p)
            result.append(Movimiento(
                fecha=fecha,
                descripcion=descripcion,
                debito=_a_decimal_ar(row[cols["debito"]]) if cols["debito"] < len(row) else None,
                credito=_a_decimal_ar(row[cols["credito"]]) if cols["credito"] < len(row) else None,
                saldo=_a_decimal_ar(row[cols["saldo"]]) if cols["saldo"] < len(row) else None,
            ))
        return result


_DATE_RE_GEN = re.compile(r"^(\d{1,2}[/-]\d{1,2}(?:[/-]\d{2,4})?)\b")
_NUMBER_RE = re.compile(r"-?\(?\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})\)?-?")


def _a_decimal_libre(raw: str) -> Optional[Decimal]:
    s = raw.strip()
    if not s:
        return None
    negativo = False
    if s.startswith("(") and s.endswith(")"):
        negativo, s = True, s[1:-1]
    if s.endswith("-"):
        negativo, s = True, s[:-1]
    if s.startswith("-"):
        negativo, s = True, s[1:]
    if s.rfind(",") > s.rfind("."):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        valor = Decimal(s)
    except InvalidOperation:
        return None
    return -valor if negativo else valor


class ParserGenerico(ParserBanco):
    """Genérico: FECHA DESCRIPCION [DEBITO] [CREDITO] SALDO."""

    name = "Genérico"

    def puede_parsear(self, pdf: PdfExtraido) -> bool:
        return True

    def parsear(self, pdf: PdfExtraido) -> List[Movimiento]:
        movs: List[Movimiento] = []
        actual: Optional[Movimiento] = None
        for line in pdf.lines:
            m = _DATE_RE_GEN.match(line)
            if m:
                if actual is not None:
                    movs.append(actual)
                actual = self._parsear_linea(line, m)
            elif actual is not None and not self._es_header(line):
                actual.descripcion = f"{actual.descripcion} {line}".strip()
        if actual is not None:
            movs.append(actual)
        return movs

    @staticmethod
    def _es_header(line: str) -> bool:
        low = line.lower()
        return any(h in low for h in ("saldo anterior", "saldo inicial", "página", "page ", "fecha"))

    def _parsear_linea(self, line: str, m: re.Match) -> Movimiento:
        fecha = m.group(1)
        rest = line[m.end():].strip()
        nums = [(_a_decimal_libre(x.group(0)), x.start()) for x in _NUMBER_RE.finditer(rest)]
        nums = [(v, s) for v, s in nums if v is not None]
        if not nums:
            return Movimiento(fecha=fecha, descripcion=rest)
        descripcion = rest[:nums[0][1]].strip()
        vals = [v for v, _ in nums]
        if len(vals) == 1:
            return Movimiento(fecha=fecha, descripcion=descripcion, saldo=vals[0])
        if len(vals) == 2:
            monto, saldo = vals
            if monto < 0:
                return Movimiento(fecha=fecha, descripcion=descripcion, debito=-monto, saldo=saldo)
            return Movimiento(fecha=fecha, descripcion=descripcion, credito=monto, saldo=saldo)
        debito, credito, saldo = vals[-3], vals[-2], vals[-1]
        return Movimiento(
            fecha=fecha, descripcion=descripcion,
            debito=debito if debito != 0 else None,
            credito=credito if credito != 0 else None,
            saldo=saldo,
        )


_PARSERS: List[ParserBanco] = [ParserChubut(), ParserGenerico()]


def _elegir_parser(pdf: PdfExtraido) -> ParserBanco:
    for p in _PARSERS:
        if p.puede_parsear(pdf):
            return p
    return _PARSERS[-1]


# --------------------------------------------------------------------------- #
# API pública
# --------------------------------------------------------------------------- #

def pdf_a_dataframe(pdf_bytes: bytes) -> Tuple[pd.DataFrame, str]:
    """Convierte un extracto PDF a DataFrame. Devuelve (df, nombre_parser)."""
    pdf = extraer_pdf(pdf_bytes)
    parser = _elegir_parser(pdf)
    movs = parser.parsear(pdf)
    if not movs:
        raise ErrorConversion("No se encontraron movimientos en el PDF.")

    def f(d: Optional[Decimal]):
        return float(d) if d is not None else None

    df = pd.DataFrame([{
        "Fecha": m.fecha,
        "Descripción": m.descripcion,
        "Débito": f(m.debito),
        "Crédito": f(m.credito),
        "Saldo": f(m.saldo),
    } for m in movs])
    return df, parser.name


def dataframe_a_excel(df: pd.DataFrame) -> bytes:
    """Exporta el DataFrame de movimientos a un Excel descargable."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Movimientos")
        ws = writer.sheets["Movimientos"]
        enc = writer.book.add_format({"bold": True, "bg_color": "#1F4E78",
                                      "font_color": "white", "border": 1})
        num = writer.book.add_format({"num_format": "#,##0.00"})
        for j, col in enumerate(df.columns):
            ws.write(0, j, str(col), enc)
            ancho = 16 if col in ("Débito", "Crédito", "Saldo") else 30
            ws.set_column(j, j, ancho, num if col in ("Débito", "Crédito", "Saldo") else None)
        ws.freeze_panes(1, 0)
    buf.seek(0)
    return buf.getvalue()
