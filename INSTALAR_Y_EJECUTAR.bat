@echo off
cd /d "%~dp0"
echo Instalando/actualizando dependencias...
py -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo ERROR: No se pudo instalar. Verifique que Python este instalado.
  pause
  exit /b 1
)
echo.
echo Abriendo Verificador Vial Caminera...
py verificador_vial.py
pause
