#!/bin/bash
# 邮件群发工具 - 自动安装脚本 (Mac/Linux)

set -e  # 遇到错误立即退出

echo "========================================="
echo "  邮件群发工具 - 自动安装"
echo "========================================="
echo ""

# 检查 Python 版本
echo "[1/6] 检查 Python 环境..."
if ! command -v python3 &> /dev/null; then
    echo "错误：未找到 python3，请先安装 Python 3.6 或更高版本"
    exit 1
fi

PYTHON_VERSION=$(python3 --version 2>&1 | awk '{print $2}')
echo "✓ 检测到 Python $PYTHON_VERSION"
echo ""

# 创建虚拟环境
echo "[2/6] 创建虚拟环境..."
if [ -d "email_venv" ]; then
    echo "虚拟环境已存在，跳过创建"
else
    python3 -m venv email_venv
    echo "✓ 虚拟环境创建成功"
fi
echo ""

# 激活虚拟环境
echo "[3/6] 激活虚拟环境..."
source email_venv/bin/activate
echo "✓ 虚拟环境已激活"
echo ""

# 升级 pip
echo "[4/6] 升级 pip..."
python -m pip install --upgrade pip -q
echo "✓ pip 已升级到最新版本"
echo ""

# 安装依赖
echo "[5/6] 安装项目依赖..."
pip install -r requirements.txt -q
echo "✓ 依赖安装完成"
echo ""

# 配置 VSCode
echo "[6/6] 配置 VSCode 开发环境..."
if [ ! -d ".vscode" ]; then
    mkdir -p .vscode
    echo "✓ 创建 .vscode 目录"
fi

# 检查并创建 settings.json（如果不存在）
if [ ! -f ".vscode/settings.json" ]; then
    cat > .vscode/settings.json << 'EOF'
{
    "python.defaultInterpreterPath": "${workspaceFolder}/email_venv/bin/python",
    "python.terminal.activateEnvironment": true,
    "files.exclude": {
        "**/__pycache__": true,
        "**/*.pyc": true
    },
    "python.analysis.autoImportCompletions": true,
    "python.analysis.typeCheckingMode": "basic",
    "python.languageServer": "Pylance",
    "[python]": {
        "editor.defaultFormatter": "ms-python.python",
        "editor.formatOnSave": false,
        "editor.codeActionsOnSave": {
            "source.organizeImports": "never"
        }
    },
    "files.encoding": "utf8",
    "files.autoGuessEncoding": false
}
EOF
    echo "✓ 创建 VSCode settings.json"
fi

# 检查并创建 launch.json（如果不存在）
if [ ! -f ".vscode/launch.json" ]; then
    cat > .vscode/launch.json << 'EOF'
{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "通用附件群发",
            "type": "debugpy",
            "request": "launch",
            "program": "${workspaceFolder}/通用附件群发/send_bulk_emails.py",
            "console": "integratedTerminal",
            "python": "${workspaceFolder}/email_venv/bin/python",
            "cwd": "${workspaceFolder}/通用附件群发",
            "justMyCode": true,
            "env": {
                "PYTHONPATH": "${workspaceFolder}"
            }
        },
        {
            "name": "捐赠证书群发",
            "type": "debugpy",
            "request": "launch",
            "program": "${workspaceFolder}/捐赠证书群发/send_certificates.py",
            "console": "integratedTerminal",
            "python": "${workspaceFolder}/email_venv/bin/python",
            "cwd": "${workspaceFolder}/捐赠证书群发",
            "justMyCode": true,
            "env": {
                "PYTHONPATH": "${workspaceFolder}"
            }
        },
        {
            "name": "解析证书文件名",
            "type": "debugpy",
            "request": "launch",
            "program": "${workspaceFolder}/捐赠证书群发/parse_certificates.py",
            "console": "integratedTerminal",
            "python": "${workspaceFolder}/email_venv/bin/python",
            "cwd": "${workspaceFolder}/捐赠证书群发",
            "justMyCode": true,
            "env": {
                "PYTHONPATH": "${workspaceFolder}"
            }
        }
    ]
}
EOF
    echo "✓ 创建 VSCode launch.json（调试配置）"
fi
echo "✓ VSCode 配置完成"
echo ""

# 配置 .env 文件
if [ ! -f ".env" ]; then
    if [ -f ".env.example" ]; then
        echo "----------------------------------------"
        echo "  配置邮箱信息"
        echo "----------------------------------------"
        cp .env.example .env
        echo "✓ 已创建 .env 文件"
        echo ""
        echo "⚠️  请编辑 .env 文件，填写您的邮箱配置："
        echo "   - SENDER_EMAIL (发件人邮箱)"
        echo "   - SENDER_PASSWORD (邮箱密码或授权码)"
        echo "   - SMTP_SERVER (SMTP服务器地址)"
        echo ""
    fi
else
    echo "✓ .env 配置文件已存在"
    echo ""
fi

echo "========================================="
echo "  ✓ 安装完成！"
echo "========================================="
echo ""
echo "📋 下一步操作："
echo ""
echo "1. 编辑 .env 文件配置您的邮箱信息"
echo ""
echo "2. 使用 VSCode 运行脚本（推荐）："
echo "   - 用 VSCode 打开项目文件夹"
echo "   - 打开任意 Python 文件（如 send_bulk_emails.py）"
echo "   - 点击右上角的 ▶️ 运行按钮"
echo "   - 或按 F5 开始调试"
echo ""
echo "3. 或使用命令行运行："
echo "   source email_venv/bin/activate"
echo "   python3 通用附件群发/send_bulk_emails.py"
echo ""
echo "💡 提示：VSCode 已自动配置好 Python 解释器路径"
echo "   无需手动选择虚拟环境，开箱即用！"
echo ""
