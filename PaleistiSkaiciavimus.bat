@echo off
chcp 65001 > nul
setlocal

:: Kelias iki šio aplanko
cd /d "%~dp0"

:: 1. Tikriname, ar aplinka išvis egzistuoja
if not exist "venv" (
    echo [KLAIDA] Virtuali aplinka nerasta! 
    echo Paleiskite instaliavimo skriptą arba sukurkite venv rankiniu būdu.
    pause
    exit /b
)

:: 2. TIK aktyvuojame aplinką (jokio diegimo ar atnaujinimo)
call venv\Scripts\activate.bat

:: 3. Paleidžiame jūsų skriptą
echo [INFO] Paleidžiamas Python skriptas...
python "Skaiciavimai.py"

echo.
echo [INFO] Darbas baigtas.
pause