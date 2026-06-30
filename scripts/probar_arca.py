"""
Script de prueba — verifica que el certificado de ARCA esté funcionando.

Lo que hace:
  1. Lee tu certificado y clave privada de la carpeta ``certificados/``.
  2. Le pide un token a WSAA (autenticación).
  3. Consulta tu propio CUIT en el Padrón.
  4. Imprime el resultado.

Uso (desde la raíz del repo, con el venv activado):

    python scripts/probar_arca.py
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

# Permitir importar desde src/
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.arca import certs, padron, wsaa


def main() -> int:
    print("== Cargando certificado ==")
    try:
        cert_pem, key_pem, cuit_titular = certs.cargar()
    except Exception as exc:
        print(f"  ✗ Error: {exc}")
        return 1
    print(f"  ✓ Certificado cargado. CUIT del titular: {cuit_titular}")

    print("\n== Pidiendo token a WSAA ==")
    print("  (Esto puede tardar 3-5 segundos la primera vez)")
    try:
        cred = wsaa.obtener_credenciales(
            servicio=padron.SERVICIO,
            cert_pem=cert_pem,
            key_pem=key_pem,
            cuit_titular=cuit_titular,
        )
    except Exception as exc:
        print(f"  ✗ Error de WSAA: {exc}")
        return 1
    print(f"  ✓ Token obtenido. Expira: {cred.expira_en.isoformat()}")
    print(f"     Token (primeros 50 chars):  {cred.token[:50]}…")
    print(f"     Sign  (primeros 50 chars):  {cred.sign[:50]}…")

    print(f"\n== Consultando padrón de TU PROPIO CUIT ({cuit_titular}) ==")
    try:
        datos = padron.consultar(
            token=cred.token,
            sign=cred.sign,
            cuit_titular=cuit_titular,
            cuit_a_consultar=cuit_titular,
        )
    except Exception as exc:
        print(f"  ✗ Error de consulta: {exc}")
        return 1

    print("  ✓ Datos recibidos:")
    print(json.dumps(datos, indent=2, ensure_ascii=False, default=str))
    return 0


if __name__ == "__main__":
    sys.exit(main())
