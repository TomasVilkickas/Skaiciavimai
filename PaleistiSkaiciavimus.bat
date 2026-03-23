@echo off
chcp 65001 > nul
setlocal

:: Tikriname, ar jau paleista maksimizuotame lange
if "%1" == "maximized" goto START_PROCESAS

:: Paleidžiame iš naujo maksimizuotai
start /max "" "%~f0" maximized
exit /b

:START_PROCESAS
:: Kelias iki šio aplanko (naudojame kabutes dėl tarpų kelyje)
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