@echo off
py -3 mrf_parse\empty.py >nul 2>nul
if %errorlevel% equ 0 (
    py -3 mrf_parse
    goto fin
)
python -V 2>&1 | findstr " 3" >nul 2>nul
if %errorlevel% equ 0 (
    python mrf_parse
    goto fin
)
echo Python 3 not installed
echo Install latest Python 3 release from https://www.python.org/downloads/windows/
echo.
goto fin

:fin
pause
