@echo off
setlocal
cd /d "%~dp0"
py -3 -m venv .venv
if errorlevel 1 goto error
call .venv\Scripts\activate.bat
python -m pip install -r requirements.txt
if errorlevel 1 goto error
python -m PyInstaller --noconfirm --clean --windowed --onedir --name LCDCatalogueDesktop app.py
if errorlevel 1 goto error
if not exist dist\LCDCatalogueDesktop\data mkdir dist\LCDCatalogueDesktop\data
copy /Y data\lcd_catalogue.json dist\LCDCatalogueDesktop\data\lcd_catalogue.json
if errorlevel 1 goto error
echo.
echo BUILD SUCCESSFUL: dist\LCDCatalogueDesktop\LCDCatalogueDesktop.exe
pause
exit /b 0
:error
echo BUILD FAILED. Review errors above.
pause
exit /b 1
