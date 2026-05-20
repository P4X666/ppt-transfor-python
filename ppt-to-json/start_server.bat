@echo off
REM PPTX 转换服务启动脚本 (Windows)

set "VENV_PATH=I:\wefor\ai\knowledge\venv"
set "APP_PATH=I:\wefor\ai\knowledge\ppt-to-json"

echo Activating virtual environment...
call "%VENV_PATH%\Scripts\activate.bat"

echo Changing to app directory...
cd /d "%APP_PATH%"

echo Starting PPTX Converter API Server...
uvicorn main:app --host 0.0.0.0 --port 8000 --workers 4

echo Server stopped.
pause