#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
邮件模板配置文件
用户可以在此文件中编辑邮件的标题和正文
支持使用变量：{var1}, {var2}, {var3}
"""

# ============================================================
# 发件人配置
# ============================================================
SENDER_NAME = "您的姓名"

# ============================================================
# 邮件模板
# ============================================================
# 邮件标题
# 可用变量：{var1}, {var2}, {var3}, {sender_name}
EMAIL_SUBJECT = "主题示例：关于 {var1} 的通知"

# 纯文本邮件正文
# 可用变量：{var1}, {var2}, {var3}, {sender_name}
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

# HTML邮件正文（可选，提高邮件送达率和显示效果）
# 可用变量：{var1}, {var2}, {var3}, {sender_name}
# 如果设置为None，则只发送纯文本邮件
EMAIL_BODY_HTML = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        body {{
            font-family: "Microsoft YaHei", "Helvetica Neue", Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 600px;
            margin: 0 auto;
            padding: 20px;
        }}
        .container {{
            background-color: #ffffff;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            padding: 30px;
        }}
        h1 {{
            color: #2c3e50;
            font-size: 24px;
            margin-bottom: 20px;
        }}
        .content {{
            margin: 20px 0;
        }}
        .variable {{
            background-color: #f8f9fa;
            padding: 10px;
            border-left: 3px solid #4CAF50;
            margin: 10px 0;
        }}
        .signature {{
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #e0e0e0;
            color: #666;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>尊敬的收件人</h1>

        <div class="content">
            <p>您好！</p>

            <p>这是一封示例邮件。您可以在此处使用以下变量：</p>

            <div class="variable">
                <strong>第一个变量：</strong>{var1}
            </div>
            <div class="variable">
                <strong>第二个变量：</strong>{var2}
            </div>
            <div class="variable">
                <strong>第三个变量：</strong>{var3}
            </div>

            <p>[在这里编辑您的邮件内容]</p>
        </div>

        <div class="signature">
            <p>此致<br>敬礼</p>
            <p><strong>{sender_name}</strong></p>
        </div>
    </div>
</body>
</html>
"""

# ============================================================
# 使用说明
# ============================================================
# 1. 将 SENDER_NAME 改为您的真实姓名
#
# 2. 编辑 EMAIL_SUBJECT、EMAIL_BODY 和 EMAIL_BODY_HTML，可以使用以下变量：
#    - {var1}: 将被替换为Excel中"var1"列的值
#    - {var2}: 将被替换为Excel中"var2"列的值
#    - {var3}: 将被替换为Excel中"var3"列的值
#    - {sender_name}: 将被替换为发件人姓名（SENDER_NAME）
#
# 3. HTML邮件说明：
#    - EMAIL_BODY_HTML 用于发送更美观的HTML格式邮件
#    - 同时也会发送纯文本版本作为备用
#    - 如果不想使用HTML,可以将 EMAIL_BODY_HTML 设置为 None
#    - HTML中的 {{ 和 }} 是转义的大括号,不会被变量替换
#
# 4. Excel文件列顺序：
#    收件邮箱 | var1 | var2 | var3 | 附件名称1 | 附件名称2 | 发送情况
#    发送情况列的值：0 = 未发送，1 = 已发送
#
# 5. 变量使用提示：
#    - 如果某个变量在Excel中为空，将被替换为空字符串
#    - 变量名必须完全匹配（包括大括号），例如：{var1}
#    - 如果不需要使用某个变量，可以不在邮件模板中使用它
#
# 6. 保存此文件后，运行 send_bulk_emails.py 发送邮件
# ============================================================
