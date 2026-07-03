@echo off
setlocal

cd /d "%~dp0"

echo Installing required packages if needed...
py -m pip install --upgrade pyinstaller pypdf pandas pillow reportlab openpyxl

echo.
echo Building CatalogGenerator.exe...
py -m PyInstaller ^
  --noconfirm ^
  --onefile ^
  --windowed ^
  --name CatalogGenerator ^
  --hidden-import openpyxl ^
  --hidden-import pypdf ^
  --collect-submodules pandas ^
  --collect-submodules PIL ^
  --collect-submodules reportlab ^
  catalog_generator_app.py

if exist "dist\CatalogGenerator.exe" (
  copy /Y "dist\CatalogGenerator.exe" "CatalogGenerator.exe" >nul
  echo.
  echo Done: %~dp0CatalogGenerator.exe
) else (
  echo.
  echo Build failed. Please check the messages above.
)

pause
