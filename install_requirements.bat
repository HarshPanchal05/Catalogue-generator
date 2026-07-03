@echo off
setlocal

echo Installing Catalog Generator requirements...
py -m pip install --upgrade pandas pillow reportlab openpyxl pypdf

echo.
echo Done. You can now run catalog_generator_app.py again.
pause
