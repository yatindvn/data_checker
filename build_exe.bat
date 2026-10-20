@echo off
REM Builds dist\ExcelCompare\ExcelCompare.exe and zips it to dist\ExcelCompare.zip.
REM Share the zip: users unzip it and double-click ExcelCompare.exe (no Python needed).
REM A folder build is used because it starts in ~2s; a --onefile exe takes ~20s
REM every launch because it unpacks pandas to a temp folder each time.
cd /d "%~dp0"
python -m pip install -r requirements.txt pyinstaller || exit /b 1
python -m PyInstaller --noconfirm --clean --onedir --windowed ^
  --name ExcelCompare ^
  --collect-all python_calamine ^
  --hidden-import xlsxwriter ^
  --exclude-module pytest --exclude-module matplotlib --exclude-module IPython ^
  --exclude-module scipy --exclude-module PIL --exclude-module notebook ^
  app.py || exit /b 1
copy /y README.md dist\ExcelCompare\README.md >nul || exit /b 1
powershell -NoProfile -Command "Compress-Archive -Force -Path dist\ExcelCompare -DestinationPath dist\ExcelCompare.zip" || exit /b 1
echo.
echo Built: %~dp0dist\ExcelCompare.zip
