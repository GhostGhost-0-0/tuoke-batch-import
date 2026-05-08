@echo off
REM 尝试设置控制台编码为UTF-8
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion

echo ========================================
echo 正在打包批量导入工具
echo ========================================
echo.

REM 检查是否安装了 PyInstaller
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo [1/5] 正在安装 PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo PyInstaller 安装失败，请手动安装: pip install pyinstaller
        pause
        exit /b 1
    )
)

echo [2/5] 检查依赖项...
pip install -r requirements.txt
if errorlevel 1 (
    echo 依赖项安装可能有问题，请检查
)

echo [3/5] 检查必要文件...
if not exist "价格.json" (
    echo 未找到 价格.json 文件！
    pause
    exit /b 1
)
if not exist "links" (
    echo 未找到 links 文件夹！
    pause
    exit /b 1
)
if not exist "使用说明.md" (
    echo 未找到 使用说明.md 文件，将跳过
)

echo [4/5] 开始打包...
pyinstaller --clean 批量导入.spec

if errorlevel 1 (
    echo 打包失败！
    pause
    exit /b 1
)

echo [5/5] 复制数据文件到 dist 目录...
REM 复制价格.json
if exist "价格.json" (
    copy /Y "价格.json" "dist\价格.json" >nul
    echo   已复制 价格.json
) else (
    echo   价格.json 不存在，跳过
)

REM 复制使用说明.md
if exist "使用说明.md" (
    copy /Y "使用说明.md" "dist\使用说明.md" >nul
    echo   已复制 使用说明.md
) else (
    echo   使用说明.md 不存在，跳过
)

REM 复制使用说明.txt
if exist "使用说明.txt" (
    copy /Y "使用说明.txt" "dist\使用说明.txt" >nul
    echo   已复制 使用说明.txt
) else (
    echo   使用说明.txt 不存在，跳过
)

REM 复制 links 文件夹
if exist "links" (
    if exist "dist\links" (
        rmdir /S /Q "dist\links"
    )
    xcopy /E /I /Y "links" "dist\links" >nul
    echo   已复制 links 文件夹
) else (
    echo   links 文件夹不存在，跳过
)

echo.
echo 打包完成！
echo.
echo 打包输出位置: dist\
echo.
echo 打包内容包含：
echo    - 批量导入.exe（主程序）
echo    - 价格.json（价格配置文件）
echo    - links文件夹（链接数据文件夹）
echo    - 使用说明.md（使用说明文档）
echo    - 使用说明.txt（使用说明文档）
echo.
echo 部署说明：
echo 1. 将 dist 目录下的所有文件复制到目标电脑
echo 2. 确保目标电脑已打开"批量导入"窗口
echo 3. 双击运行 批量导入.exe 即可
echo.
echo 注意事项：
echo - 首次运行可能被杀毒软件拦截，需要添加信任
echo - 确保目标电脑是 Windows 系统
echo - 无需安装 Python 或其他依赖
echo.
pause




