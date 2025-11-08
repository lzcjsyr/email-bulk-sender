#!/bin/bash
# 邮件群发工具 - 自动安装脚本 (Mac/Linux)

set -e  # 遇到错误立即退出

echo "========================================="
echo "  邮件群发工具 - 自动安装"
echo "========================================="
echo ""

# 检查 Python 版本
echo "[1/5] 检查 Python 环境..."
if ! command -v python3 &> /dev/null; then
    echo "错误：未找到 python3，请先安装 Python 3.6 或更高版本"
    exit 1
fi

PYTHON_VERSION=$(python3 --version 2>&1 | awk '{print $2}')
echo "✓ 检测到 Python $PYTHON_VERSION"
echo ""

# 创建虚拟环境
echo "[2/5] 创建虚拟环境..."
if [ -d "email_venv" ]; then
    echo "虚拟环境已存在，跳过创建"
else
    python3 -m venv email_venv
    echo "✓ 虚拟环境创建成功"
fi
echo ""

# 激活虚拟环境
echo "[3/5] 激活虚拟环境..."
source email_venv/bin/activate
echo "✓ 虚拟环境已激活"
echo ""

# 升级 pip
echo "[4/5] 升级 pip..."
python -m pip install --upgrade pip -q
echo "✓ pip 已升级到最新版本"
echo ""

# 安装依赖
echo "[5/5] 安装项目依赖..."
pip install -r requirements.txt -q
echo "✓ 依赖安装完成"
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
echo "  安装完成！"
echo "========================================="
echo ""
echo "下一步操作："
echo "1. 编辑 .env 文件配置您的邮箱信息"
echo "2. 运行脚本："
echo "   - 通用附件群发: python3 通用附件群发/send_bulk_emails.py"
echo "   - 证书群发: python3 捐赠证书群发/send_certificates.py"
echo ""
echo "提示：每次运行前需要先激活虚拟环境"
echo "      source email_venv/bin/activate"
echo ""
