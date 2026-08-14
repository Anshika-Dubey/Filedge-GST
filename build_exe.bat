@echo off
REM ============================================================
REM  GST Filing Status Checker — One-Click EXE Builder
REM ============================================================
echo.
echo ============================================================
echo   GST Filing Status Checker — EXE Builder
echo ============================================================
echo.

echo [1/5] Installing required Python packages...
pip install selenium openpyxl pandas flask webdriver-manager pyinstaller --quiet
if errorlevel 1 ( echo ERROR: pip install failed. & pause & exit /b 1 )

echo [2/5] Writing spec file...
python write_spec.py
if errorlevel 1 ( echo ERROR: Could not write spec file. & pause & exit /b 1 )

echo [3/5] Creating runtime hook...
python write_hook.py
if errorlevel 1 ( echo ERROR: Could not write hook. & pause & exit /b 1 )

echo [4/5] Building EXE from spec (takes 2-3 mins)...
pyinstaller GSTChecker.spec --noconfirm
if errorlevel 1 ( echo. & echo ERROR: Build failed. & pause & exit /b 1 )

echo [5/5] Cleaning up...
rmdir /s /q build 2>nul
del /q hook_wdm.py write_spec.py write_hook.py GSTChecker.spec 2>nul

echo.
echo ============================================================
echo   SUCCESS!  dist\GSTChecker.exe is ready.
echo   Your friend only needs Google Chrome installed.
echo ============================================================
echo.
pause
