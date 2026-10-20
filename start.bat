@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
py -3 -c "import openpyxl, docx, tkinter" >nul 2>&1
if not errorlevel 1 (
    py -3 app.py %*
    goto end
)
python -c "import openpyxl, docx, tkinter" >nul 2>&1
if not errorlevel 1 (
    python app.py %*
    goto end
)
set "TICKET_PYTHON=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if not exist "%TICKET_PYTHON%" goto missing_dependency
"%TICKET_PYTHON%" -c "import openpyxl, docx, tkinter" >nul 2>&1
if errorlevel 1 goto missing_dependency
"%TICKET_PYTHON%" app.py %*
goto end
:missing_dependency
echo No Python with openpyxl, python-docx and tkinter found. Install dependencies:
echo py -3 -m pip install -r requirements.txt
pause
exit /b 1
:end
if errorlevel 1 pause
endlocal
