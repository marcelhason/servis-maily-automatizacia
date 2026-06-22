@echo off
REM Build sap_nahraj.py -> sap_nahraj.exe (v roote projektu)
REM Spusti z adresara C:\moje_py\lucka_data\
REM
REM Poznamka: distpath je C:\Temp\sap_build aby sa vyhlo
REM   zamykaniu exe Windows Defenderom pocas buildu v projektu.

pip install pyinstaller pywin32 >nul 2>&1

if not exist C:\Temp\sap_build mkdir C:\Temp\sap_build

pyinstaller --onefile --noconsole --name sap_nahraj ^
            --distpath C:\Temp\sap_build ^
            sap_nahraj.py

echo.
if exist C:\Temp\sap_build\sap_nahraj.exe (
    echo BUILD OK
    copy C:\Temp\sap_build\sap_nahraj.exe . >nul
    echo Skopirovaný sap_nahraj.exe do rootu projektu
) else (
    echo BUILD ZLYHALO
)
pause
