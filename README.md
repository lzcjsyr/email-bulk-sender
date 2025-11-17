# 邮件群发工具

一个简单易用的 Python 邮件批量发送工具，支持附件发送和个性化邮件内容。

---

## 📋 功能特性

本项目提供两个独立的邮件群发功能：

### 功能1：捐赠证书邮件群发
- 从证书文件名自动提取收件人信息
- 自动发送个性化感谢邮件
- 证书图片作为附件发送

### 功能2：通用邮件群发
- 从 Excel 读取收件人列表
- 支持自定义邮件模板
- 灵活的附件管理
- 邮件内容支持变量替换

### 🚀 增强特性（v2.0）

**提高冗余度：**
- ✅ SMTP连接池：复用连接，提高发送效率和稳定性
- ✅ 自动重试机制：发送失败自动重试最多3次，支持自定义
- ✅ 详细日志记录：所有操作记录到logs目录，便于追踪问题
- ✅ Excel自动备份：更新状态前自动备份，防止数据丢失
- ✅ 发送前验证：检查邮箱格式、附件存在性，避免无效发送

**提升便捷性：**
- ✅ 进度条显示：实时查看发送进度和预计剩余时间
- ✅ 命令行参数：支持 --dry-run、--test-smtp 等便捷操作
- ✅ 模拟模式：测试邮件配置而不实际发送
- ✅ SMTP测试：一键测试邮箱连接是否正常
- ✅ 批量操作：智能批次管理，避免触发邮件服务商限制

---

## 🚀 快速开始

### 第一步：克隆或下载项目

```bash
git clone https://github.com/lzcjsyr/email-bulk-sender.git
cd email-bulk-sender
```

或直接下载 ZIP 文件并解压。

### 第二步：一键安装环境

#### macOS / Linux
```bash
./setup.sh
```

#### Windows
双击运行 `setup.bat` 或在命令行执行：
```cmd
setup.bat
```

安装脚本会自动：
- ✅ 检查 Python 环境
- ✅ 创建虚拟环境
- ✅ 安装所有依赖
- ✅ 配置 VSCode 开发环境
- ✅ 生成配置文件模板

### 第三步：配置邮箱信息

编辑 `.env` 文件，填写您的邮箱配置（详见下方"配置邮箱"章节）

### 第四步：运行脚本

#### 使用 VSCode（推荐）

**运行 `setup.sh` 或 `setup.bat` 后，VSCode 已自动配置完成！**

1. **用 VSCode 打开项目文件夹**
2. **打开任意 Python 文件**（如 `send_bulk_emails.py`）
3. **点击右上角 ▶️ 运行按钮** 或按 `F5`

✨ **无需手动选择 Python 解释器，无需激活虚拟环境，开箱即用！**

> setup 脚本已自动配置好：
> - Python 解释器路径：`email_venv/bin/python`
> - 调试配置：3 个预设运行配置
> - 文件编码：UTF-8（支持中文）

#### 使用命令行（备选）

```bash
# 激活虚拟环境
source email_venv/bin/activate          # macOS/Linux
email_venv\Scripts\activate.bat         # Windows

# 运行脚本
python 通用邮件群发/send_bulk_emails.py
python 捐赠证书群发/send_certificates.py
```

---

<details>
<summary><b>手动安装（点击展开）</b></summary>

```bash
# 1. 创建虚拟环境
python3 -m venv email_venv

# 2. 激活虚拟环境
source email_venv/bin/activate          # macOS/Linux
# 或
email_venv\Scripts\activate.bat         # Windows

# 3. 安装依赖
pip install -r requirements.txt

# 4. 配置环境变量
cp .env.example .env
# 编辑 .env 文件填写邮箱配置
```

</details>

---

## ⚙️ 配置邮箱

编辑 `.env` 文件，填写您的邮箱信息：

```env
SMTP_SERVER=smtp.exmail.qq.com    # SMTP 服务器地址
SMTP_PORT=587                     # SMTP 端口
SENDER_EMAIL=your@email.com       # 发件人邮箱
SENDER_PASSWORD=your_auth_code    # 邮箱密码或授权码
```

### 常见邮箱配置

| 邮箱服务商 | SMTP 服务器 | 端口 | 说明 |
|-----------|------------|------|------|
| QQ邮箱 | smtp.qq.com | 587 | 需要授权码 |
| 163邮箱 | smtp.163.com | 587 | 需要授权码 |
| Gmail | smtp.gmail.com | 587 | 需要应用专用密码 |
| 腾讯企业邮箱 | smtp.exmail.qq.com | 587 | 使用邮箱密码 |
| 阿里企业邮箱 | smtp.qiye.aliyun.com | 587 | 使用邮箱密码 |

### 如何获取授权码？

**QQ邮箱 / 163邮箱：**
1. 登录邮箱网页版
2. 进入 设置 → 账户
3. 找到 POP3/SMTP/IMAP 设置
4. 开启 SMTP 服务
5. 点击"生成授权码"并复制

**Gmail：**
1. 开启两步验证
2. 生成应用专用密码
3. 使用该密码而非 Google 账户密码

---

## 📖 使用说明

### 功能 1：捐赠证书邮件群发

#### 步骤1：准备证书文件

将证书图片放入 `捐赠证书群发/捐赠证书/` 文件夹

文件名格式：`姓名_邮箱.jpg`（例如：`张三_zhang@example.com.jpg`）

#### 步骤2：解析证书文件名

```bash
python3 捐赠证书群发/parse_certificates.py
```

这将生成 `证书信息确认表.xlsx`，请检查并确认信息

#### 步骤3：发送证书

**基本用法：**
```bash
python3 捐赠证书群发/send_certificates.py
```

**高级用法：**
```bash
# 测试SMTP连接
python3 捐赠证书群发/send_certificates.py --test-smtp

# 模拟发送（不实际发送邮件，用于测试）
python3 捐赠证书群发/send_certificates.py --dry-run

# 跳过确认直接发送
python3 捐赠证书群发/send_certificates.py --yes

# 自定义重试次数
python3 捐赠证书群发/send_certificates.py --max-retries 5
```

---

### 功能 2：通用邮件群发

#### 步骤1：配置邮件模板

编辑 `通用邮件群发/prompts.py`：

```python
# 发件人配置
SENDER_NAME = "您的姓名"

# 邮件标题（可用变量：{var1}, {var2}, {var3}, {sender_name}）
EMAIL_SUBJECT = "主题示例：关于 {var1} 的通知"

# 邮件正文（可用变量：{var1}, {var2}, {var3}, {sender_name}）
EMAIL_BODY = """尊敬的收件人：

您好！

这是一封示例邮件。您可以在此处使用以下变量：
- 第一个变量：{var1}
- 第二个变量：{var2}
- 第三个变量：{var3}

[在这里编辑您的邮件内容]

此致
敬礼

{sender_name}
"""
```

**支持的变量：**
- `{var1}` - 第一个自定义变量（对应 Excel 中的 var1 列）
- `{var2}` - 第二个自定义变量（对应 Excel 中的 var2 列）
- `{var3}` - 第三个自定义变量（对应 Excel 中的 var3 列）
- `{sender_name}` - 发件人姓名

#### 步骤2：准备附件

将附件文件放入 `通用邮件群发/attachments/` 文件夹

#### 步骤3：填写发送列表

编辑 `通用邮件群发/发送列表.xlsx`：

**Excel 列格式：**

| 收件邮箱 | var1 | var2 | var3 | 附件名称1 | 附件名称2 | 发送情况 |
|---------|------|------|------|----------|----------|---------|
| zhang@example.com | 张三 | 100元 | 2024-01-20 | report.pdf | | 0 |
| li@example.com | 李四 | 200元 | 2024-01-21 | contract.docx | invoice.pdf | 0 |

**列说明：**
- **收件邮箱**：收件人的电子邮箱地址
- **var1, var2, var3**：自定义变量，对应模板中的 {var1}, {var2}, {var3}
- **附件名称1, 附件名称2**：附件文件名（可选，留空表示无附件）
- **发送情况**：发送状态（0 = 未发送，1 = 已发送）

#### 步骤4：运行脚本

**基本用法：**
```bash
python3 通用邮件群发/send_bulk_emails.py
```

**高级用法：**
```bash
# 测试SMTP连接
python3 通用邮件群发/send_bulk_emails.py --test-smtp

# 模拟发送（不实际发送邮件，用于测试）
python3 通用邮件群发/send_bulk_emails.py --dry-run

# 跳过确认直接发送
python3 通用邮件群发/send_bulk_emails.py --yes

# 自定义重试次数
python3 通用邮件群发/send_bulk_emails.py --max-retries 5

# 查看帮助信息
python3 通用邮件群发/send_bulk_emails.py --help
```

---

## 📁 项目结构

```
邮件群发/
├── .env                      # 邮箱配置（需自己创建）
├── .env.example              # 配置模板
├── requirements.txt          # Python 依赖
├── setup.sh                  # macOS/Linux 安装脚本
├── setup.bat                 # Windows 安装脚本
├── README.md                 # 本文档
├── email_utils.py            # 共享工具模块（连接池、重试、日志等）
├── email_venv/               # 虚拟环境（运行 setup.sh 后自动创建）
├── logs/                     # 日志文件夹（自动创建）
│
├── 通用邮件群发/
│   ├── send_bulk_emails.py   # 群发脚本（已增强）
│   ├── prompts.py            # 邮件模板
│   ├── 发送列表.xlsx         # 收件人列表
│   └── attachments/          # 附件文件夹
│
└── 捐赠证书群发/
    ├── parse_certificates.py  # 解析证书文件名
    ├── send_certificates.py   # 发送证书邮件（已增强）
    ├── 捐赠证书/              # 证书文件夹
    └── 证书信息确认表.xlsx    # 自动生成的确认表
```

---

## 🔧 常见问题

### 1. 找不到 python3 命令

**Windows 用户**：使用 `python` 而非 `python3`

**macOS 用户**：确保已安装 Python 3
```bash
brew install python3
```

### 2. 邮件发送失败：535 Authentication failed

- 检查 `.env` 中的邮箱和密码是否正确
- QQ/163 邮箱需要使用授权码，不是登录密码
- 确认 SMTP 服务已开启

### 3. 邮件发送失败：smtplib.SMTPException

- 检查网络连接
- 确认 SMTP 服务器地址和端口正确
- 某些网络可能屏蔽 SMTP 端口

### 4. Excel 文件读取错误

- 确保 Excel 文件格式为 `.xlsx`（不是 `.xls`）
- 检查列名是否正确（区分大小写）
- 关闭 Excel 文件后再运行脚本

### 5. 虚拟环境激活失败

**Windows PowerShell 执行策略限制**：
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### 6. 附件找不到

- **功能1**：检查证书文件是否在 `捐赠证书群发/捐赠证书/` 文件夹中
- **功能2**：检查附件是否在 `通用邮件群发/attachments/` 文件夹中，且文件名与Excel中完全一致

### 7. 中文文件名乱码

脚本已处理中文编码，如果仍有问题：
- 确保系统编码为 UTF-8
- Windows 用户可以在命令行执行 `chcp 65001` 切换到 UTF-8

---

## ⚠️ 注意事项

### 安全提示

1. **不要分享 `.env` 文件**，其中包含您的邮箱密码
2. 使用 Git 时，`.gitignore` 已配置忽略敏感文件
3. 定期更换邮箱授权码
4. 不要将项目上传到公共代码仓库（如果包含真实数据）

### 发送建议

1. **测试先行**
   - 先用少量数据测试
   - 确认邮件内容和附件正确后再批量发送

2. **控制发送速度**
   - 建议设置发送间隔为 5-10 秒
   - 可在 `.env` 文件中调整参数：
     ```env
     DELAY_BETWEEN_EMAILS=5    # 每批邮件后等待秒数
     EMAILS_PER_BATCH=10       # 每批发送数量
     ```

3. **附件大小限制**
   - 不同邮箱服务商对附件大小限制不同（一般 10-25MB）
   - 建议附件文件大小不超过 5MB

4. **备份数据**
   - 建议备份原始文件和 Excel 文件
   - 保留发送记录便于后续查询

---

## 🛠️ 高级配置

### 调整发送速度

编辑 `.env` 文件：

```env
DELAY_BETWEEN_EMAILS=5    # 每批邮件后等待秒数
EMAILS_PER_BATCH=10       # 每批发送数量
```

### 使用不同的 SMTP 端口

```env
SMTP_PORT=465  # 使用 SSL/TLS
```

对应代码需要修改：
```python
server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port)  # 465端口
```

---

## 📞 技术支持

### 系统要求

- **Python**: 3.6 或更高版本
- **操作系统**: Windows / macOS / Linux
- **网络**: 能够访问 SMTP 邮件服务器

### 依赖库

- pandas >= 1.3.0
- openpyxl >= 3.0.0
- python-dotenv >= 0.19.0

### 故障排查

如果遇到问题：
1. 检查本文档的"常见问题"部分
2. 查看脚本输出的错误信息
3. 确认 `.env` 配置正确
4. 测试邮箱 SMTP 连接是否正常

---

## 📖 新功能详细说明

### 命令行参数

所有脚本现在支持以下命令行参数：

| 参数 | 说明 |
|------|------|
| `--dry-run` | 模拟发送模式，不实际发送邮件，用于测试 |
| `--test-smtp` | 测试SMTP连接是否正常，无需发送邮件 |
| `--yes`, `-y` | 跳过确认提示，直接开始发送 |
| `--max-retries N` | 设置发送失败后的最大重试次数（默认3次） |
| `--help`, `-h` | 显示帮助信息 |

### 日志系统

- 所有操作都会记录到 `logs/` 目录
- 每次运行生成独立的日志文件，文件名包含时间戳
- 日志包含详细的错误信息，便于问题排查
- 日志示例：`logs/bulk_email_sender_20250112_143000.log`

### SMTP连接池

- 自动复用SMTP连接，减少连接开销
- 默认每个连接发送50封邮件后重新建立
- 提高发送效率，减少网络波动影响
- 发生错误时自动重置连接

### 自动重试机制

- 发送失败自动重试，默认最多3次
- 使用指数退避策略：2秒 → 4秒 → 8秒
- 可通过 `--max-retries` 参数自定义重试次数
- 详细的重试日志记录

### Excel自动备份

- 更新发送状态前自动备份Excel文件
- 备份文件名包含时间戳，如：`发送列表_backup_20250112_143000.xlsx`
- 防止意外操作导致数据丢失
- 可追溯历史发送记录

### 发送前验证

**邮箱格式验证：**
- 自动检查邮箱地址格式是否有效
- 无效邮箱会被跳过并记录日志
- 在发送报告中统计无效邮箱数量

**附件存在性检查：**
- 发送前预检查所有附件是否存在
- 缺失的附件会提前警告
- 避免发送过程中才发现附件缺失

### 进度显示

如果安装了 `tqdm` 库（已包含在requirements.txt中）：
- 显示美观的进度条
- 实时更新发送速度
- 显示预计剩余时间

如果未安装 `tqdm`：
- 使用文本方式显示进度
- 显示当前进度和预计剩余时间

---

## 📄 许可证

本项目仅供学习和内部使用，请遵守相关法律法规和邮箱服务商的使用条款。

**注意事项：**
- 不要用于垃圾邮件发送
- 遵守邮箱服务商的发送频率限制
- 确保收件人同意接收邮件

---

## 🔄 更新日志

### v2.0 (2025-01-12)

**冗余度提升：**
- 新增 SMTP 连接池，提高发送效率和稳定性
- 新增自动重试机制，支持自定义重试次数
- 新增详细日志记录系统
- 新增 Excel 自动备份功能
- 新增发送前邮箱格式和附件存在性验证

**便捷性提升：**
- 新增命令行参数支持
- 新增 `--dry-run` 模拟发送模式
- 新增 `--test-smtp` SMTP连接测试
- 新增进度条显示（需要 tqdm 库）
- 新增预计剩余时间显示

**Bug修复：**
- 优化错误处理和异常捕获
- 改进中文文件名编码处理

### v1.0 (2025-01-10)
- 初始版本发布
- 支持通用邮件群发
- 支持捐赠证书群发
