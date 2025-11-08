@echo off
chcp 65001 > nul
REM 邮件群发工具 - 自动安装脚本 (Windows)

echo =========================================
echo   邮件群发工具 - 自动安装
echo =========================================
echo.

REM 检查 Python 版本
echo [1/5] 检查 Python 环境...
python --version > nul 2>&1
if errorlevel 1 (
    echo 错误：未找到 Python，请先安装 Python 3.6 或更高版本
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)

for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo ✓ 检测到 Python %PYTHON_VERSION%
echo.

REM 创建虚拟环境
echo [2/5] 创建虚拟环境...
if exist email_venv\ (
    echo 虚拟环境已存在，跳过创建
) else (
    python -m venv email_venv
    echo ✓ 虚拟环境创建成功
)
echo.

REM 激活虚拟环境
echo [3/5] 激活虚拟环境...
call email_venv\Scripts\activate.bat
echo ✓ 虚拟环境已激活
echo.

REM 升级 pip
echo [4/5] 升级 pip...
python -m pip install --upgrade pip -q
echo ✓ pip 已升级到最新版本
echo.

REM 安装依赖
echo [5/5] 安装项目依赖...
pip install -r requirements.txt -q
echo ✓ 依赖安装完成
echo.

REM 配置 .env 文件
if not exist .env (
    if exist .env.example (
        echo ----------------------------------------
        echo   配置邮箱信息
        echo ----------------------------------------
        copy .env.example .env > nul
        echo ✓ 已创建 .env 文件
        echo.
        echo ⚠️  请编辑 .env 文件，填写您的邮箱配置：
        echo    - SENDER_EMAIL ^(发件人邮箱^)
        echo    - SENDER_PASSWORD ^(邮箱密码或授权码^)
        echo    - SMTP_SERVER ^(SMTP服务器地址^)
        echo.
    )
) else (
    echo ✓ .env 配置文件已存在
    echo.
)

echo =========================================
echo   安装完成！
echo =========================================
echo.
echo 下一步操作：
echo 1. 编辑 .env 文件配置您的邮箱信息
echo 2. 运行脚本：
echo    - 通用附件群发: python 通用附件群发\send_bulk_emails.py
echo    - 证书群发: python 捐赠证书群发\send_certificates.py
echo.
echo 提示：每次运行前需要先激活虚拟环境
echo       email_venv\Scripts\activate.bat
echo.
pause
