# Herramientas del estudio

App (Python + Streamlit) con varias herramientas del estudio contable:

- 🧮 **Comparador** — Contabilidad vs Planilla (caja): cruce movimiento a movimiento por importe + fecha.
- 🏦 **PDF de banco → Excel** — convierte extractos en PDF (optimizado Banco del Chubut).
- 📒 **JWIN → PS3 (MICROENV)** — genera los `.ps3` y el Excel del sistema contable.
- 🧾 **Ventas por actividad (Tango)** — neto e IVA por actividad y categoría de IVA (ventas − NC), para la DJ de IVA / Convenio Multilateral.

---

## Correr en la PC (local)

1. Doble clic en **`Instalar.bat`** (una sola vez).
2. Doble clic en **`Abrir programa.bat`** → se abre en el navegador.

(o manual: `python -m venv .venv`, `pip install -r requirements.txt`, `streamlit run app.py`)

---

## Publicar en la web (Streamlit Community Cloud · gratis · con contraseña)

1. **Cuenta de GitHub** (github.com) → crear un **repositorio privado** (ej. `herramientas-estudio`).
2. Subir este proyecto al repo (pasos de git abajo).
3. **Cuenta en Streamlit Cloud** → https://share.streamlit.io (iniciá sesión con GitHub).
4. **New app** → elegí el repo, rama `main`, archivo principal `app.py` → **Deploy**.
5. En la app: **Settings → Secrets** y pegá:
   ```toml
   password = "TU-CLAVE-SECRETA"
   ```
   Esa es la contraseña que la app pide al entrar.
6. Te queda un link `https://....streamlit.app` protegido por contraseña.

### Subir el código a GitHub (una vez)
```bash
git init
git add .
git commit -m "App herramientas del estudio"
git branch -M main
git remote add origin https://github.com/TU-USUARIO/herramientas-estudio.git
git push -u origin main
```
Para actualizar luego: `git add . && git commit -m "cambios" && git push`
(Streamlit Cloud redepliega solo al detectar el push.)

### Seguridad
- La contraseña va en **Secrets**, nunca en el código.
- `.gitignore` evita subir datos de clientes (evidencia, Excel) y los secretos.
- Mantené el repositorio **privado**.

---

## Estructura

```
comparador_sucesion/
├── app.py                       # UI Streamlit (menú de herramientas + login)
├── requirements.txt
└── src/
    ├── excel_reader.py          # leer Excel
    ├── normalizar.py            # fechas, importes AR, texto
    ├── mapeo.py                 # movimientos (simple / doble columna) + corte
    ├── conciliar.py             # cruce por importe + fecha y estados
    ├── reporte.py               # resumen y exportación
    ├── pdf_banco.py             # PDF de banco → tabla
    ├── ps3_micro.py             # JWIN → PS3 (MICROENV)
    ├── plan_cuentas_microenv.json
    └── ventas_tango.py          # ventas por actividad (Tango)
```
