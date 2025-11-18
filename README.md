# 邮件群发工具

一个简单易用的 Python 邮件批量发送工具，支持附件发送、个性化邮件内容和智能发送管理。

---

## 📋 功能特性

### 核心功能
- ✅ 从 Excel 读取收件人列表
- ✅ 支持自定义邮件模板（标题和正文）
- ✅ 支持双附件发送
- ✅ 邮件内容支持 4 个变量替换
- ✅ 自动记录发送状态到 Excel

### 🔒 邮件安全与送达率增强（NEW!）
- ✅ **DKIM数字签名**：为邮件添加DKIM签名，大幅提高Gmail等邮箱的送达率
- ✅ **SPF/DMARC验证**：发送前自动检查DNS配置，确保邮件身份验证
- ✅ **关键邮件头部**：自动添加List-Unsubscribe、Reply-To等合规头部
- ✅ **内容安全检查**：自动检测垃圾邮件关键词和可疑URL，降低被拒风险
- ✅ **IP信誉检查**：发送前检查发送IP是否在黑名单中
- ✅ **退订机制**：符合CAN-SPAM等法规要求的一键退订功能
- ✅ **RFC5322标准**：完全符合邮件格式规范，确保各大邮箱兼容

### 🚀 智能发送与容错
- ✅ **SMTP连接池**：复用连接，提高发送效率和稳定性
- ✅ **指数退避重试**：真正的指数退避算法，智能处理临时性错误
- ✅ **错误分类处理**：区分永久性/临时性错误，避免无效重试
- ✅ **反弹检测**：自动识别硬反弹(邮箱不存在)和软反弹(临时故障)
- ✅ **速率限制检测**：自动识别并应对邮件服务商的速率限制
- ✅ **详细日志记录**：所有操作记录到logs目录，便于追踪问题
- ✅ **Excel自动备份**：更新状态前自动备份，防止数据丢失
- ✅ **发送前验证**：检查邮箱格式、附件存在性，避免无效发送

### 便捷特性
- ✅ **进度条显示**：实时查看发送进度和预计剩余时间
- ✅ **模拟模式**：测试邮件配置而不实际发送(DRY_RUN=true)
- ✅ **批量操作**：智能批次管理，避免触发邮件服务商限制
- ✅ **简洁设计**：纯文本邮件，简单高效，兼容所有邮箱

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
chmod +x setup.sh
./setup.sh
```

#### Windows
双击运行 `setup.bat` 或在命令行执行：
```cmd
setup.bat
```

安装脚本会自动：
- ✅ 检查 Python 环境（需要 Python 3.6+）
- ✅ 创建虚拟环境
- ✅ 安装所有依赖
- ✅ 配置 VSCode 开发环境
- ✅ 生成配置文件模板（.env）

### 第三步：配置邮箱信息

编辑 `.env` 文件，填写您的邮箱配置：

```env
# 基础配置（必需）
SMTP_SERVER=smtp.exmail.qq.com    # SMTP 服务器地址
SMTP_PORT=587                     # SMTP 端口
SENDER_EMAIL=your@email.com       # 发件人邮箱
SENDER_PASSWORD=your_auth_code    # 邮箱密码或授权码

# 增强功能（可选，但强烈建议配置以提高Gmail送达率）
ENABLE_PRE_SEND_CHECKS=true       # 启用发送前安全检查
# DKIM_DOMAIN=example.com         # DKIM域名（需配置DNS）
# DKIM_SELECTOR=default           # DKIM选择器
# DKIM_PRIVATE_KEY_PATH=dkim_private.key  # DKIM私钥路径
```

详细配置说明见下方"配置邮箱"章节。

### 第四步：运行脚本

#### 使用 VSCode（推荐）

**运行 `setup.sh` 或 `setup.bat` 后，VSCode 已自动配置完成！**

1. **用 VSCode 打开项目文件夹**
2. **打开 Python 文件**（如 `通用邮件群发/send_bulk_emails.py`）
3. **点击右上角 ▶️ 运行按钮** 或按 `F5`

✨ **无需手动选择 Python 解释器，无需激活虚拟环境，开箱即用！**

> setup 脚本已自动配置好：
> - Python 解释器路径：`email_venv/bin/python`
> - 调试配置：预设运行配置
> - 文件编码：UTF-8（支持中文）

#### 使用命令行（备选）

```bash
# 激活虚拟环境
source email_venv/bin/activate          # macOS/Linux
email_venv\Scripts\activate.bat         # Windows

# 运行脚本
python 通用邮件群发/send_bulk_emails.py
```

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

### 步骤1：配置邮件模板

编辑 `通用邮件群发/template.py`：

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

### 步骤2：准备附件（可选）

将附件文件放入 `通用邮件群发/attachments/` 文件夹

### 步骤3：填写发送列表

编辑 `通用邮件群发/发送列表.xlsx`：

**Excel 列格式：**

| 收件邮箱 | var1 | var2 | var3 | 附件名称1 | 附件名称2 | 发送情况 |
|---------|------|------|------|----------|----------|---------|
| zhang@example.com | 张三 | 100元 | 2024-01-20 | report.pdf | | 0 |
| li@example.com | 李四 | 200元 | 2024-01-21 | contract.docx | invoice.pdf | 0 |

**列说明：**
- **收件邮箱**：收件人的电子邮箱地址（必填）
- **var1, var2, var3**：自定义变量，对应模板中的 {var1}, {var2}, {var3}（可选）
- **附件名称1, 附件名称2**：附件文件名（可选，留空表示无附件）
- **发送情况**：发送状态（0 = 未发送，1 = 已发送，脚本自动更新）

### 步骤4：运行脚本

**基本用法：**
```bash
python 通用邮件群发/send_bulk_emails.py
```

**高级用法：**
```bash
# 测试SMTP连接
python 通用邮件群发/send_bulk_emails.py --test-smtp

# 模拟发送（不实际发送邮件，用于测试）
python 通用邮件群发/send_bulk_emails.py --dry-run

# 跳过确认直接发送
python 通用邮件群发/send_bulk_emails.py --yes

# 自定义重试次数
python 通用邮件群发/send_bulk_emails.py --max-retries 5

# 查看帮助信息
python 通用邮件群发/send_bulk_emails.py --help
```

---

## 📁 项目结构

```
email-bulk-sender/
├── .env                      # 邮箱配置（需自己创建）
├── .env.example              # 配置模板
├── requirements.txt          # Python 依赖
├── setup.sh                  # macOS/Linux 安装脚本
├── setup.bat                 # Windows 安装脚本
├── README.md                 # 本文档
│
├── core/                     # 核心代码模块
│   ├── __init__.py           # 模块初始化
│   ├── email_security.py     # 【NEW】邮件安全模块(DKIM/SPF/内容检查)
│   ├── email_enhanced.py     # 【NEW】增强发送模块(智能重试/错误分类)
│   └── email_utils.py        # 工具函数
│
├── email_venv/               # 虚拟环境（运行 setup 后自动创建）
├── logs/                     # 日志文件夹（自动创建）
├── dkim_private.key          # DKIM私钥（可选，需自己生成）
├── dkim_public.key           # DKIM公钥（可选，需自己生成）
│
└── 通用邮件群发/
    ├── send_bulk_emails.py   # 群发脚本（已集成增强功能）
    ├── template.py           # 邮件模板配置（支持HTML）
    ├── 发送列表.xlsx         # 收件人列表
    ├── email_sender.log      # 发送日志（自动生成）
    └── attachments/          # 附件文件夹
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

- 检查附件是否在 `通用邮件群发/attachments/` 文件夹中
- 确保文件名与 Excel 中完全一致（包括扩展名）
- 检查文件路径中是否有特殊字符

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
   - 使用 `--dry-run` 模式测试邮件内容
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
   - 建议单个附件文件大小不超过 5MB
   - 建议总附件大小不超过 10MB

4. **备份数据**
   - 建议备份原始文件和 Excel 文件
   - 脚本会自动备份 Excel（带时间戳）
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

对应代码需要修改 `send_bulk_emails.py`：
```python
server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port)  # 465端口
```

---

## 📖 命令行参数说明

| 参数 | 说明 |
|------|------|
| `--dry-run` | 模拟发送模式，不实际发送邮件，用于测试 |
| `--test-smtp` | 测试SMTP连接是否正常，无需发送邮件 |
| `--yes`, `-y` | 跳过确认提示，直接开始发送 |
| `--max-retries N` | 设置发送失败后的最大重试次数（默认3次） |
| `--help`, `-h` | 显示帮助信息 |

---

## 📋 功能详解

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

## 📞 技术支持

### 系统要求

- **Python**: 3.6 或更高版本
- **操作系统**: Windows / macOS / Linux
- **网络**: 能够访问 SMTP 邮件服务器

### 依赖库

- pandas >= 1.3.0
- openpyxl >= 3.0.0
- python-dotenv >= 0.19.0
- tqdm >= 4.62.0（可选，用于进度条显示）

### 故障排查

如果遇到问题：
1. 检查本文档的"常见问题"部分
2. 查看脚本输出的错误信息
3. 查看 `logs/` 目录中的日志文件
4. 确认 `.env` 配置正确
5. 使用 `--test-smtp` 测试邮箱连接是否正常

---

## 📄 许可证

本项目仅供学习和内部使用，请遵守相关法律法规和邮箱服务商的使用条款。

**注意事项：**
- 不要用于垃圾邮件发送
- 遵守邮箱服务商的发送频率限制
- 确保收件人同意接收邮件
- 尊重收件人的隐私权

---

## 🎯 提高Gmail等邮箱送达率配置指南

针对Gmail、Outlook等严格的邮件服务商，本工具提供了完整的送达率优化方案。

### 为什么邮件会进垃圾箱?

Gmail等邮箱会检查以下几点:
1. **身份验证**: 是否配置了SPF、DKIM、DMARC
2. **发送者信誉**: IP地址是否被拉黑
3. **邮件内容**: 是否包含垃圾邮件特征
4. **邮件格式**: 是否符合标准(HTML+纯文本)
5. **合规性**: 是否提供退订链接

### 完整配置步骤

#### 步骤1: 配置SPF记录（必需）

在您的域名DNS管理中添加TXT记录:

```
记录类型: TXT
主机记录: @
记录值: v=spf1 include:_spf.google.com ~all
```

如果使用其他邮件服务商,参考:
- **腾讯企业邮**: `v=spf1 include:spf.mail.qq.com ~all`
- **阿里企业邮**: `v=spf1 include:spf.mxhichina.com ~all`
- **自建服务器**: `v=spf1 ip4:your-server-ip ~all`

验证SPF: 运行脚本时会自动检查

#### 步骤2: 配置DKIM签名（强烈推荐）

**1) 生成DKIM密钥对:**

```bash
# 生成私钥
openssl genrsa -out dkim_private.key 2048

# 提取公钥
openssl rsa -in dkim_private.key -pubout -out dkim_public.key

# 获取公钥内容(用于DNS配置)
cat dkim_public.key
```

**2) 配置DNS记录:**

```
记录类型: TXT
主机记录: default._domainkey
记录值: v=DKIM1; k=rsa; p=<公钥内容>
```

公钥内容格式: 去除`-----BEGIN PUBLIC KEY-----`和`-----END PUBLIC KEY-----`以及换行符

**3) 配置.env文件:**

```env
DKIM_DOMAIN=example.com
DKIM_SELECTOR=default
DKIM_PRIVATE_KEY_PATH=dkim_private.key
```

**4) 验证DKIM:**

运行脚本时会自动检查DNS配置是否正确

#### 步骤3: 配置DMARC策略（推荐）

在域名DNS中添加:

```
记录类型: TXT
主机记录: _dmarc
记录值: v=DMARC1; p=none; rua=mailto:dmarc@example.com
```

DMARC策略说明:
- `p=none`: 监控模式(推荐新手)
- `p=quarantine`: 隔离可疑邮件
- `p=reject`: 拒绝未通过验证的邮件

#### 步骤4: 启用HTML邮件

在`.env`中配置:

```env
ENABLE_HTML_EMAIL=true
```

在`template.py`中编辑`EMAIL_BODY_HTML`,同时保留`EMAIL_BODY`纯文本版本

#### 步骤5: 启用发送前检查

在`.env`中配置:

```env
ENABLE_PRE_SEND_CHECKS=true
```

脚本会在发送前自动检查:
- ✓ SPF记录是否存在
- ✓ DMARC策略是否配置
- ✓ 发送IP是否在黑名单
- ✓ 邮件内容是否有垃圾特征
- ✓ 是否包含可疑URL

### 配置验证清单

运行脚本前,确认以下各项:

- [ ] SPF记录已添加到DNS
- [ ] DKIM密钥已生成并配置到DNS
- [ ] DMARC策略已配置(可选,但推荐)
- [ ] .env中DKIM相关配置已填写
- [ ] ENABLE_HTML_EMAIL=true
- [ ] ENABLE_PRE_SEND_CHECKS=true
- [ ] template.py中已配置EMAIL_BODY_HTML
- [ ] 邮件内容不含垃圾关键词

### DNS配置示例(完整)

```
# SPF记录
@ TXT "v=spf1 include:_spf.google.com ~all"

# DKIM记录
default._domainkey TXT "v=DKIM1; k=rsa; p=MIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQC..."

# DMARC记录
_dmarc TXT "v=DMARC1; p=none; rua=mailto:dmarc@example.com"
```

### 测试送达率

配置完成后:

1. **测试DNS配置:**
   ```bash
   # 检查SPF
   dig txt example.com

   # 检查DKIM
   dig txt default._domainkey.example.com

   # 检查DMARC
   dig txt _dmarc.example.com
   ```

2. **发送测试邮件:**
   ```bash
   # 先用DRY RUN模式测试
   python 通用邮件群发/send_bulk_emails.py
   # 会显示所有检查结果

   # 发送给自己的Gmail测试
   # 在Excel中只添加一条记录(你的Gmail地址)
   ```

3. **检查Gmail原始邮件头:**
   - 打开收到的邮件
   - 点击"更多" → "显示原始邮件"
   - 查找以下内容:
     - `Authentication-Results`: 应显示`spf=pass`, `dkim=pass`
     - `Received-SPF`: 应显示`pass`

### 常见问题

**Q: 配置后还是进垃圾箱?**

A: 检查以下几点:
1. DNS记录是否生效(可能需要等待几小时)
2. 邮件内容是否包含垃圾关键词
3. 发送IP是否被拉黑(运行脚本时会检查)
4. 是否发送过快(调低EMAILS_PER_BATCH)
5. 域名是否有良好的发送历史

**Q: 没有域名怎么办?**

A: 如果使用企业邮箱(如腾讯企业邮、Gmail企业版):
- SPF/DMARC由邮箱服务商管理
- 启用HTML邮件和内容检查仍然有效
- 送达率会比个人邮箱好

**Q: DKIM配置太复杂?**

A: DKIM虽然配置复杂,但对Gmail送达率提升最明显。建议:
- 先配置SPF(简单)
- 再配置HTML邮件(简单)
- 最后配置DKIM(有技术能力时)

---

## 🔄 更新日志

### v3.0 (2025-01-17) - 送达率增强版

**🔒 邮件安全与送达率:**
- ✅ 新增DKIM数字签名支持
- ✅ 新增SPF/DMARC自动验证
- ✅ 新增HTML邮件支持(multipart/alternative)
- ✅ 新增List-Unsubscribe等关键邮件头
- ✅ 新增垃圾邮件内容检查
- ✅ 新增IP黑名单检查
- ✅ 新增退订机制

**🚀 智能重试与容错:**
- ✅ 实现真正的指数退避重试(之前仅为固定延迟)
- ✅ 新增错误分类系统(永久性/临时性/速率限制)
- ✅ 新增SMTP反弹检测(硬反弹/软反弹)
- ✅ 智能识别邮箱不存在、速率限制等错误
- ✅ 根据错误类型自动调整重试策略

**📦 新增模块:**
- `core/email_security.py` - 安全检查模块
- `core/email_enhanced.py` - 增强发送模块
- `core/email_utils.py` - 工具函数模块

### v2.0 (2025-01-12)

**冗余度提升:**
- 新增 SMTP 连接池，提高发送效率和稳定性
- 新增自动重试机制，支持自定义重试次数
- 新增详细日志记录系统
- 新增 Excel 自动备份功能
- 新增发送前邮箱格式和附件存在性验证

**便捷性提升:**
- 新增命令行参数支持
- 新增 `--dry-run` 模拟发送模式
- 新增 `--test-smtp` SMTP连接测试
- 新增进度条显示（需要 tqdm 库）
- 新增预计剩余时间显示

### v1.0 (2025-01-10)
- 初始版本发布
- 支持通用邮件群发
- 支持 Excel 数据导入
- 支持附件发送
- 支持邮件模板变量替换
