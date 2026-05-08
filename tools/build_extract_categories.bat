@echo off
REM 打包 tools\extract_categories.py 为独立 exe（PyInstaller 单文件）
chcp 65001 >nul 2>&1

cd /d "%~dp0"

echo ========================================
echo 打包：类目提取 extract_categories
echo ========================================
echo.

python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo [1/4] 正在安装 PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo 错误：PyInstaller 安装失败，请手动执行: pip install pyinstaller
        pause
        exit /b 1
    )
) else (
    echo [1/4] PyInstaller 已可用
)

echo [2/4] 安装项目依赖 ^(uiautomation、pywinauto 等^)...
if exist "%~dp0..\requirements.txt" (
    pip install -r "%~dp0..\requirements.txt"
) else (
    pip install uiautomation pywinauto
)
if errorlevel 1 (
    echo 警告：依赖安装可能不完整，请检查网络或 pip 环境
)

echo [3/4] 开始 PyInstaller 打包...
pyinstaller --clean --onefile --console --name extract_categories ^
    --hidden-import=comtypes.gen.UIAutomationClient ^
    --hidden-import=pywinauto ^
    "%~dp0extract_categories.py"

if errorlevel 1 (
    echo 错误：打包失败
    pause
    exit /b 1
)

echo [4/4] 完成
echo.
echo 生成文件: %~dp0dist\extract_categories.exe
echo.
echo 使用说明:
echo - 目标电脑无需安装 Python。
echo - 运行前请先打开「类目修改」窗口。
echo - 文本.json 会生成在**当前工作目录**（从哪双击/在哪启动命令行运行，cwd 即为该目录^)；若需固定目录，可在该文件夹里打开 cmd 再运行 exe。
echo.
pause
