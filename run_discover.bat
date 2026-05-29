@echo off
cd /d "%~dp0"
echo Installing dependencies...
python -m pip install -r requirements.txt
echo.
echo Running discovery...
python -m src.discover
echo.
pause
