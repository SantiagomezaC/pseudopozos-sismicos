@echo off
REM ==========================================================================
REM  Instalacion reproducible de "Pseudopozos sismicos"
REM  Crea un entorno virtual aislado (.venv) junto a este archivo e instala
REM  las versiones exactas de requirements.txt. Se ejecuta una sola vez.
REM ==========================================================================
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo No se encontro Python. Instale Python 3.12 o superior desde https://www.python.org
    echo marcando la opcion "Add python.exe to PATH" y vuelva a ejecutar este archivo.
    pause
    exit /b 1
)

echo [1/3] Creando el entorno virtual .venv ...
python -m venv .venv || goto :error

echo [2/3] Instalando dependencias (versiones fijadas en requirements.txt) ...
".venv\Scripts\python.exe" -m pip install --upgrade pip || goto :error
".venv\Scripts\python.exe" -m pip install -r requirements.txt || goto :error

echo [3/3] Verificando la instalacion con la prueba reproducible ...
".venv\Scripts\python.exe" pruebas\prueba_reproducible.py || goto :error

echo.
echo Instalacion completa. Abra el programa con ejecutar.bat
pause
exit /b 0

:error
echo.
echo La instalacion se detuvo por un error. Revise los mensajes anteriores.
pause
exit /b 1
