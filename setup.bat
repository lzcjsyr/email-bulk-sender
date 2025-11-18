@echo off
chcp 65001 > nul
REM 邮件群发工具 - 自动安装脚本 (Windows)

echo =========================================
echo   邮件群发工具 - 自动安装
echo =========================================
echo.

REM 检查 Python 版本
echo [1/6] 检查 Python 环境...
python --version > nul 2>&1
if errorlevel 1 (
    echo 错误：未找到 Python，请先安装 Python 3.13
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)

for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo ✓ 检测到 Python %PYTHON_VERSION%

REM 检查是否为推荐版本
echo %PYTHON_VERSION% | findstr /B "3.13" > nul
if %errorlevel% equ 0 (
    echo ✓ 使用推荐版本 Python 3.13
) else (
    echo %PYTHON_VERSION% | findstr /B "3\." > nul
    if %errorlevel% equ 0 (
        echo ⚠️  当前版本可用，但推荐使用 Python 3.13
    ) else (
        echo 错误：需要 Python 3.6 或更高版本（推荐 3.13）
        pause
        exit /b 1
    )
)
echo.

REM 创建虚拟环境
echo [2/6] 创建虚拟环境...
if exist email_venv\ (
    echo 虚拟环境已存在，跳过创建
) else (
    python -m venv email_venv
    echo ✓ 虚拟环境创建成功
)
echo.

REM 激活虚拟环境
echo [3/6] 激活虚拟环境...
call email_venv\Scripts\activate.bat
echo ✓ 虚拟环境已激活
echo.

REM 升级 pip 和安装依赖
echo [4/6] 安装项目依赖...

REM 检测是否使用 uv
where uv >nul 2>&1
if %errorlevel% equ 0 (
    echo 检测到 uv 环境，使用 uv pip 安装...
    uv pip install -r requirements.txt --python .\email_venv\Scripts\python.exe -q
    echo ✓ 依赖安装完成 ^(使用 uv^)
) else (
    python -m pip --version >nul 2>&1
    if %errorlevel% equ 0 (
        echo 升级 pip...
        python -m pip install --upgrade pip -q
        echo 安装依赖...
        pip install -r requirements.txt -q
        echo ✓ 依赖安装完成 ^(使用 pip^)
    ) else (
        echo ⚠️  未检测到 pip 或 uv，请手动安装依赖:
        echo    uv pip install -r requirements.txt --python .\email_venv\Scripts\python.exe
        echo    或
        echo    python -m pip install -r requirements.txt
    )
)
echo.

REM 验证安装
echo [5/6] 验证安装...
python -c "import pandas, openpyxl, dotenv, dkimpy, dns.resolver, premailer, html2text" 2>nul
if %errorlevel% equ 0 (
    echo ✓ 所有依赖库已正确安装
) else (
    echo ⚠️  部分依赖库可能未安装，请检查上面的输出
)
echo.

REM 配置 VSCode
echo [6/6] 配置 VSCode 开发环境...
if not exist .vscode\ (
    mkdir .vscode
    echo ✓ 创建 .vscode 目录
)

REM 检查并创建 settings.json（如果不存在）
if not exist .vscode\settings.json (
    (
        echo {
        echo     "python.defaultInterpreterPath": "${workspaceFolder}/email_venv/bin/python",
        echo     "python.terminal.activateEnvironment": true,
        echo     "files.exclude": {
        echo         "**/__pycache__": true,
        echo         "**/*.pyc": true
        echo     },
        echo     "python.analysis.autoImportCompletions": true,
        echo     "python.analysis.typeCheckingMode": "basic",
        echo     "python.languageServer": "Pylance",
        echo     "[python]": {
        echo         "editor.defaultFormatter": "ms-python.python",
        echo         "editor.formatOnSave": false,
        echo         "editor.codeActionsOnSave": {
        echo             "source.organizeImports": "never"
        echo         }
        echo     },
        echo     "files.encoding": "utf8",
        echo     "files.autoGuessEncoding": false
        echo }
    ) > .vscode\settings.json
    echo ✓ 创建 VSCode settings.json
)

REM 检查并创建 launch.json（如果不存在）
if not exist .vscode\launch.json (
    (
        echo {
        echo     "version": "0.2.0",
        echo     "configurations": [
        echo         {
        echo             "name": "通用邮件群发",
        echo             "type": "debugpy",
        echo             "request": "launch",
        echo             "program": "${workspaceFolder}/通用邮件群发/send_bulk_emails.py",
        echo             "console": "integratedTerminal",
        echo             "python": "${workspaceFolder}/email_venv/bin/python",
        echo             "cwd": "${workspaceFolder}/通用邮件群发",
        echo             "justMyCode": true,
        echo             "env": {
        echo                 "PYTHONPATH": "${workspaceFolder}"
        echo             }
        echo         },
        echo         {
        echo             "name": "捐赠证书群发",
        echo             "type": "debugpy",
        echo             "request": "launch",
        echo             "program": "${workspaceFolder}/捐赠证书群发/send_certificates.py",
        echo             "console": "integratedTerminal",
        echo             "python": "${workspaceFolder}/email_venv/bin/python",
        echo             "cwd": "${workspaceFolder}/捐赠证书群发",
        echo             "justMyCode": true,
        echo             "env": {
        echo                 "PYTHONPATH": "${workspaceFolder}"
        echo             }
        echo         },
        echo         {
        echo             "name": "解析证书文件名",
        echo             "type": "debugpy",
        echo             "request": "launch",
        echo             "program": "${workspaceFolder}/捐赠证书群发/parse_certificates.py",
        echo             "console": "integratedTerminal",
        echo             "python": "${workspaceFolder}/email_venv/bin/python",
        echo             "cwd": "${workspaceFolder}/捐赠证书群发",
        echo             "justMyCode": true,
        echo             "env": {
        echo                 "PYTHONPATH": "${workspaceFolder}"
        echo             }
        echo         }
        echo     ]
        echo }
    ) > .vscode\launch.json
    echo ✓ 创建 VSCode launch.json（调试配置）
)
echo ✓ VSCode 配置完成
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
echo   ✓ 安装完成！
echo =========================================
echo.
echo 📋 下一步操作：
echo.
echo 1. 编辑 .env 文件配置您的邮箱信息
echo.
echo 2. 使用 VSCode 运行脚本（推荐）：
echo    - 用 VSCode 打开项目文件夹
echo    - 打开任意 Python 文件（如 send_bulk_emails.py）
echo    - 点击右上角的 ▶️ 运行按钮
echo    - 或按 F5 开始调试
echo.
echo 3. 或使用命令行运行：
echo    email_venv\Scripts\activate.bat
echo    python 通用邮件群发\send_bulk_emails.py
echo.
echo 💡 提示：VSCode 已自动配置好 Python 解释器路径
echo    无需手动选择虚拟环境，开箱即用！
echo.
pause
