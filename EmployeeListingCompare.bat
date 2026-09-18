@echo off

if "%~2"=="" (
    echo.
    echo ==============================================
    echo   EMPLOYEE FILE COMPARISON
    echo ==============================================
    echo.
    echo Please drag TWO Excel files onto this BAT file.
    echo.
    pause
    exit /b
)

py "%~dp0employeelistingcompare.py" "%~1" "%~2"

pause