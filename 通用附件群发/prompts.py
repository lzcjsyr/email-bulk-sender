#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
邮件模板配置文件
用户可以在此文件中编辑邮件的标题和正文
支持使用变量：{recipient_name} 和 {sender_name}
"""

# 发件人姓名
SENDER_NAME = "您的姓名"

# 邮件标题（支持变量：{recipient_name}, {sender_name}）
EMAIL_SUBJECT = "您好 {recipient_name}"

# 邮件正文（支持变量：{recipient_name}, {sender_name}）
EMAIL_BODY = """尊敬的 {recipient_name}：

您好！

[在这里编辑您的邮件内容]

此致
敬礼

{sender_name}
"""

# 使用说明：
# 1. 将 SENDER_NAME 改为您的真实姓名
# 2. 编辑 EMAIL_SUBJECT 和 EMAIL_BODY，可以使用以下变量：
#    - {recipient_name}: 将被替换为收件人姓名（从Excel中读取）
#    - {sender_name}: 将被替换为发件人姓名（SENDER_NAME）
# 3. 保存此文件后，运行 send_bulk_emails.py 发送邮件
