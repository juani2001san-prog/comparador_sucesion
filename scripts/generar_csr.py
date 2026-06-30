"""
Genera una clave privada y una solicitud de certificado (CSR) para ARCA.

Uso (desde la carpeta raíz del repo, con el venv activado):

    python scripts/generar_csr.py

El script te va a preguntar:
  - Tu CUIT (el del titular del certificado: vos o el estudio).
  - El nombre del estudio (organización).
  - Un alias para el certificado (un nombre cualquiera, ej. "estudiojuani").

Genera dos archivos en la carpeta `certificados/`:
  - <alias>.key  →  clave privada. NUNCA la compartas. NO subir al repo.
  - <alias>.csr  →  solicitud de certificado. Esto SÍ se sube a ARCA.

Después de generar el CSR:
  1. Entrá a ARCA con Clave Fiscal nivel 3.
  2. Buscá el servicio "Administración de Certificados Digitales".
  3. "Agregar Alias" → ponés el mismo alias y subís el archivo .csr.
  4. ARCA te devuelve un certificado .crt — lo guardás en la carpeta `certificados/`.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509 import CertificateSigningRequestBuilder, Name, NameAttribute
from cryptography.x509.oid import NameOID


CARPETA_SALIDA = Path("certificados")


def _solo_digitos(s: str) -> str:
    return "".join(c for c in s if c.isdigit())


def _pedir_cuit() -> str:
    while True:
        valor = input("CUIT del titular del certificado (11 dígitos): ").strip()
        d = _solo_digitos(valor)
        if len(d) == 11:
            return d
        print(f"  ⚠️  Tiene que tener 11 dígitos (vos pasaste {len(d)}). Volvé a intentar.")


def _pedir_no_vacio(prompt: str) -> str:
    while True:
        valor = input(prompt).strip()
        if valor:
            return valor
        print("  ⚠️  No puede estar vacío.")


def _pedir_alias() -> str:
    while True:
        valor = input("Alias para el certificado (sin espacios, ej. 'estudiojuani'): ").strip()
        if re.fullmatch(r"[a-z0-9_-]+", valor):
            return valor
        print("  ⚠️  Solo letras minúsculas, números, guión y guión bajo.")


def main() -> int:
    print("=" * 60)
    print("Generación de CSR para ARCA")
    print("=" * 60)
    print()

    cuit = _pedir_cuit()
    organizacion = _pedir_no_vacio("Nombre del estudio (organización): ")
    alias = _pedir_alias()

    CARPETA_SALIDA.mkdir(exist_ok=True)

    archivo_key = CARPETA_SALIDA / f"{alias}.key"
    archivo_csr = CARPETA_SALIDA / f"{alias}.csr"

    if archivo_key.exists() or archivo_csr.exists():
        respuesta = input(f"\n⚠️  Ya existen archivos con el alias '{alias}'. ¿Sobreescribir? [s/N]: ")
        if respuesta.strip().lower() not in ("s", "si", "sí", "y", "yes"):
            print("Cancelado.")
            return 1

    print("\nGenerando clave privada RSA de 2048 bits…")
    clave = rsa.generate_private_key(public_exponent=65537, key_size=2048)

    print("Guardando clave privada en", archivo_key)
    archivo_key.write_bytes(clave.private_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PrivateFormat.TraditionalOpenSSL,
        encryption_algorithm=serialization.NoEncryption(),
    ))

    print("Generando solicitud de certificado (CSR)…")
    csr = (
        CertificateSigningRequestBuilder()
        .subject_name(Name([
            NameAttribute(NameOID.COUNTRY_NAME, "AR"),
            NameAttribute(NameOID.ORGANIZATION_NAME, organizacion),
            NameAttribute(NameOID.COMMON_NAME, alias),
            NameAttribute(NameOID.SERIAL_NUMBER, f"CUIT {cuit}"),
        ]))
        .sign(clave, hashes.SHA256())
    )

    print("Guardando CSR en", archivo_csr)
    archivo_csr.write_bytes(csr.public_bytes(serialization.Encoding.PEM))

    print()
    print("=" * 60)
    print("✓ Listo")
    print("=" * 60)
    print(f"  Clave privada:  {archivo_key}   ← NO la compartas ni la subas al repo.")
    print(f"  CSR para ARCA:  {archivo_csr}   ← esto subís a ARCA.")
    print()
    print("Próximos pasos:")
    print("  1. Entrá a ARCA con Clave Fiscal nivel 3.")
    print("  2. Andá a 'Administración de Certificados Digitales'.")
    print(f"  3. Agregá un alias con el mismo nombre: '{alias}'.")
    print(f"  4. Subí el archivo '{archivo_csr.name}'.")
    print("  5. ARCA te devuelve un .crt — bajalo y guardalo en la carpeta 'certificados/'.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
