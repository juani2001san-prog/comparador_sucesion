"""
Genera el bloque TOML que tenés que pegar en Streamlit Cloud Secrets.

Lee tu certificado y clave privada de la carpeta `certificados/` y los formatea
correctamente para que Streamlit los lea como secrets.

Uso (desde la raíz del repo, con venv activado):

    python scripts/generar_secrets_arca.py

Te imprime el bloque listo. Copialo y pegalo en:
    https://share.streamlit.io → tu app → Settings (engranaje) → Secrets
    (al final de lo que ya tenés ahí).
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from src.arca.certs import _cuit_del_certificado


def main() -> int:
    base = Path("certificados")
    certs = sorted(base.glob("*.crt"))
    if not certs:
        print("✗ No se encontró ningún .crt en la carpeta certificados/")
        return 1
    cert_file = certs[0]
    key_file = cert_file.with_suffix(".key")
    if not key_file.exists():
        print(f"✗ Falta {key_file} (la clave privada).")
        return 1

    cert_pem = cert_file.read_text().strip()
    key_pem = key_file.read_text().strip()
    cuit = _cuit_del_certificado(cert_pem.encode("utf-8"))

    print()
    print("=" * 70)
    print("Copiá TODO lo que está entre las dos líneas '---' (sin las líneas '---')")
    print("y pegalo al final de los Secrets en Streamlit Cloud.")
    print("=" * 70)
    print()
    print("---")
    print()
    print("[arca]")
    print(f'cuit_representada = "{cuit}"')
    print('certificado = """')
    print(cert_pem)
    print('"""')
    print('clave_privada = """')
    print(key_pem)
    print('"""')
    print()
    print("---")
    print()
    print(f"(CUIT detectado en el certificado: {cuit})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
