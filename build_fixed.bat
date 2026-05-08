@echo off
REM Set console code page to UTF-8
chcp 65001 >nul 2>&1

echo ========================================
echo Packaging Batch Import Tool
echo ========================================
echo.

REM Check if PyInstaller is installed
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo [1/5] Installing PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo ERROR: PyInstaller installation failed, please install manually: pip install pyinstaller
        pause
        exit /b 1
    )
)

echo [2/5] Checking dependencies...
pip install -r requirements.txt
if errorlevel 1 (
    echo WARNING: Dependencies installation may have issues, please check
)

echo [3/5] Checking required files...
if not exist "价格.json" (
    echo ERROR: Price.json file not found!
    pause
    exit /b 1
)
if not exist "links" (
    echo ERROR: Links folder not found!
    pause
    exit /b 1
)
if not exist "使用说明.md" (
    echo WARNING: Usage.md file not found, will skip
)

echo [4/5] Starting packaging...
pyinstaller --clean 批量导入.spec

if errorlevel 1 (
    echo ERROR: Packaging failed!
    pause
    exit /b 1
)

echo [5/5] Copying data files to dist directory...
REM Copy price.json
if exist "价格.json" (
    copy /Y "价格.json" "dist\价格.json" >nul
    echo   Copied price.json
) else (
    echo   price.json does not exist, skipping
)

REM Copy usage.md
if exist "使用说明.md" (
    copy /Y "使用说明.md" "dist\使用说明.md" >nul
    echo   Copied usage.md
) else (
    echo   usage.md does not exist, skipping
)

REM Copy usage.txt
if exist "使用说明.txt" (
    copy /Y "使用说明.txt" "dist\使用说明.txt" >nul
    echo   Copied usage.txt
) else (
    echo   usage.txt does not exist, skipping
)

REM Copy links folder
if exist "links" (
    if exist "dist\links" (
        rmdir /S /Q "dist\links"
    )
    xcopy /E /I /Y "links" "dist\links" >nul
    echo   Copied links folder
) else (
    echo   links folder does not exist, skipping
)

echo.
echo Packaging complete!
echo.
echo Output location: dist\
echo.
echo Package contents:
echo    - 批量导入.exe (main program)
echo    - 价格.json (price configuration file)
echo    - links folder (link data folder)
echo    - 使用说明.md (usage documentation)
echo    - 使用说明.txt (usage documentation)
echo.
echo Deployment instructions:
echo 1. Copy all files in the dist directory to the target computer
echo 2. Make sure the target computer has the "批量导入" window open
echo 3. Double-click to run 批量导入.exe
echo.
echo Notes:
echo - First run may be blocked by antivirus software, need to add to trust
echo - Make sure the target computer is Windows system
echo - No need to install Python or other dependencies
echo.
pause




