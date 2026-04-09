@echo off
cd /d "%~dp0"
echo Verificando dependencias...
python -m pip install flask openpyxl xlrd --quiet
echo.
echo ===============================================
echo   Polizas de Provision de Gastos
echo   GBC Business Consulting
echo ===============================================
echo.
echo   Abre tu navegador en:
echo   ^> http://localhost:5051
echo.
python app.py
pause
