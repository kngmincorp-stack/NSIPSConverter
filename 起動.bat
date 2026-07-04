@echo off
chcp 65001 >nul
cd /d "%~dp0"
start "" pythonw "%~dp0nsips_converter.py"
