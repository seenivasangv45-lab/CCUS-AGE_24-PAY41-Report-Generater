@echo off
echo ============================================
echo  AGE_24 Web Processor - Setup ^& Launch
echo ============================================
echo.

REM Install required Python packages
echo Installing required packages...
pip install flask pandas openpyxl

echo.
echo Starting the web server...
echo.
echo *** Share the address shown below with your coworkers ***
echo.

REM Run the app
python app.py

pause
