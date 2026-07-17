@echo off
cd /d "%~dp0"
if exist "%~dp0\.venv\Scripts\python.exe" (
    "%~dp0\.venv\Scripts\python.exe" "%~dp0generar_dashboard.py" --refresh --publish --watch
) else (
    python "%~dp0generar_dashboard.py" --refresh --publish --watch
)
pause