@echo off
cd /d "%~dp0"
py verificador_vial.py
if errorlevel 1 pause
