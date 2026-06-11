@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo    Instalador - Herramientas del estudio
echo ============================================
echo.

where python >nul 2>nul
if errorlevel 1 (
  echo [!] No se encontro Python en esta computadora.
  echo.
  echo     1^) Instala Python desde:  https://www.python.org/downloads/
  echo     2^) MUY IMPORTANTE: en la primera pantalla del instalador
  echo        tilda la casilla  "Add python.exe to PATH".
  echo     3^) Volve a hacer doble clic en este  Instalar.bat
  echo.
  pause
  exit /b 1
)

echo Creando el entorno e instalando lo necesario...
echo (La primera vez tarda unos minutos. No cierres esta ventana.)
echo.

python -m venv .venv
call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
pip install -r requirements.txt

echo.
echo ============================================
echo    Listo! Ya podes abrir el programa con
echo    "Abrir programa.bat"
echo ============================================
pause
