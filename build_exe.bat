@echo off
echo ============================================
echo  Filedge GST - EXE Builder
echo ============================================
echo.

echo [1/3] Installing required packages...
pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo ERROR: pip install failed. Make sure Python is installed
    echo and added to PATH, then try again.
    pause
    exit /b 1
)

echo.
echo [2/3] Building Filedge_GST.exe (this may take a few minutes)...
pyinstaller --onefile --name Filedge_GST --console ^
    --collect-all selenium ^
    --collect-all webdriver_manager ^
    --hidden-import selenium.webdriver.chrome.webdriver ^
    --hidden-import selenium.webdriver.chrome.service ^
    --hidden-import selenium.webdriver.chrome.options ^
    --hidden-import selenium.webdriver.common.by ^
    --hidden-import selenium.webdriver.support.ui ^
    --hidden-import selenium.webdriver.support.expected_conditions ^
    gst_checker.py

echo.
echo [3/3] Done!
echo ============================================
echo Your exe is at:  dist\Filedge_GST.exe
echo Copy that single file to a folder and share it.
echo ============================================
pause
