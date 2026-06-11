@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
  echo Abriendo el programa... se abre solo en el navegador.
  echo NO cierres esta ventana mientras lo usas.
  ".venv\Scripts\python.exe" -m streamlit run app.py
) else (
  echo [!] Todavia no esta instalado.
  echo     Primero hace doble clic en  "Instalar.bat"
  echo.
  pause
)
