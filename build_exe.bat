@echo off
REM ===========================================================================
REM  Construye Cyclope.exe  (ejecutar en WINDOWS, una sola vez)
REM ---------------------------------------------------------------------------
REM  Requisitos antes de correr este .bat:
REM    - Python instalado (https://www.python.org/downloads/  -> marca "Add to PATH")
REM    - Esta carpeta debe contener:
REM         cyclope_core.py
REM         cyclope_gui.py
REM         requirements.txt
REM         Requirements\   (la carpeta con las 7 plantillas .docx/.pdf/.xlsx)
REM
REM  Resultado:  dist\Cyclope.exe   (eso es lo que repartes a tu equipo)
REM ===========================================================================

echo.
echo [1/3] Instalando dependencias...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller
if errorlevel 1 goto error

echo.
echo [2/3] Empaquetando Cyclope.exe (esto puede tardar 1-3 minutos)...
pyinstaller --onefile --windowed --name Cyclope ^
  --add-data "Requirements;Requirements" ^
  --collect-all pdfplumber ^
  --collect-all pdfminer ^
  cyclope_gui.py
if errorlevel 1 goto error

echo.
echo [3/3] LISTO.
echo El ejecutable esta en:   dist\Cyclope.exe
echo Reparte ese unico archivo a tu equipo (no necesitan Python ni nada mas).
echo.
pause
exit /b 0

:error
echo.
echo *** Ocurrio un error. Revisa los mensajes de arriba. ***
echo Sugerencia: si el .exe no abre, prueba cambiar --windowed por --console
echo en este .bat para ver el detalle del error.
echo.
pause
exit /b 1
