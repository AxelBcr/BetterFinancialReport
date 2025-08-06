@echo off
echo ========================================
echo Dashboard Refresher Script
echo ========================================
echo.

:loop
echo [%date% %time%] Python script execution...
python "easybourse_valorisation.py"

if %errorlevel% neq 0 (
    echo [%date% %time%] ERROR: The python script has encountered an error (code: %errorlevel%)
) else (
    echo [%date% %time%] Successfully executed the script.
)

echo.
echo Next execution in 1 hour...
echo Use Ctrl+C to stop the script
echo ----------------------------------------

REM Wait 3600 seconds (1 hour)
timeout /t 3600 /nobreak > nul

goto loop