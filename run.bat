@echo off
cd /d "%~dp0"
echo Instalando dependencias...
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo Error al instalar dependencias. Verifica que Python este instalado.
  pause
  exit /b 1
)
echo.
echo Iniciando SGR tools en http://127.0.0.1:5055
echo Abri esa URL en el navegador (no abras los HTML directo).
echo.
python app.py
pause
