@echo off
cd /d "%~dp0"

rem If needed, you can hardcode your Python path here instead:
rem "C:\Path\To\Python\python.exe" -m streamlit run app.py
rem pause
rem goto :eof

if exist ".venv\Scripts\python.exe" (
    echo Using local virtual environment: .venv\Scripts\python.exe
    ".venv\Scripts\python.exe" -m streamlit run app.py
    if errorlevel 1 (
        echo.
        echo Failed to start Streamlit from .venv.
        pause
    )
    goto :eof
)

where python >nul 2>nul
if %errorlevel%==0 (
    echo Using Python from PATH
    python -m streamlit run app.py
    if errorlevel 1 (
        echo.
        echo Failed to start Streamlit from Python in PATH.
        echo Make sure Streamlit is installed in this interpreter.
        echo You can also open excel-link-cleaner.bat and replace the launch command with the full path to your Python interpreter.
        pause
    )
    goto :eof
)

echo.
echo Python not found. Install Python or create .venv first.
echo You can also open excel-link-cleaner.bat and replace the launch command with the full path to your Python interpreter.
pause
