@echo off
REM Abre "Pseudopozos sismicos" con el entorno creado por instalar.bat
cd /d "%~dp0"
if not exist ".venv\Scripts\pythonw.exe" (
    echo Primero ejecute instalar.bat
    pause
    exit /b 1
)
start "" ".venv\Scripts\pythonw.exe" pseudopozos_sismicos.py
