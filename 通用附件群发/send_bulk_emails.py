#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通用附件群发邮件工具
从Excel文件读取收件人列表，发送带附件的邮件
"""

import os
import smtplib
import time
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from dotenv import load_dotenv
import sys

# 导入邮件模板配置
from prompts import SENDER_NAME, EMAIL_SUBJECT, EMAIL_BODY

# 加载环境变量（从上级目录读取.env文件）
script_dir = os.path.dirname(os.path.abspath(__file__))
env_path = os.path.join(script_dir, '..', '.env')
load_dotenv(dotenv_path=env_path)

class BulkEmailSender:
    """
    通用附件群发邮件的类
    """

    def __init__(self):
        # 邮箱配置
        self.smtp_server = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
        self.smtp_port = int(os.getenv('SMTP_PORT', '587'))
        self.sender_email = os.getenv('SENDER_EMAIL')
        self.sender_password = os.getenv('SENDER_PASSWORD')

        # 发件人姓名（从prompts.py导入）
        self.sender_name = SENDER_NAME

        # 邮件模板（从prompts.py导入）
        self.email_subject_template = EMAIL_SUBJECT
        self.email_body_template = EMAIL_BODY

        # 发送配置
        self.delay_between_emails = int(os.getenv('DELAY_BETWEEN_EMAILS', '5'))
        self.emails_per_batch = int(os.getenv('EMAILS_PER_BATCH', '10'))

        # 文件配置
        self.excel_file = '发送列表.xlsx'
        self.attachments_folder = 'attachments'

        # 验证必要配置
        if not self.sender_email or not self.sender_password:
            raise ValueError("请设置 SENDER_EMAIL 和 SENDER_PASSWORD 环境变量")

        if self.sender_name == "您的姓名":
            print("警告：请在 prompts.py 中设置 SENDER_NAME")

    def load_recipients_from_excel(self):
        """
        从Excel文件加载收件人信息
        """
        try:
            # 读取Excel文件
            df = pd.read_excel(self.excel_file)

            # 检查必要的列
            required_columns = ['姓名', '邮箱', '附件名称']
            for col in required_columns:
                if col not in df.columns:
                    print(f"错误：Excel文件缺少'{col}'列")
                    return []

            # 过滤出需要发送的邮件（状态为'待发送'或空）
            recipients = []
            for index, row in df.iterrows():
                status = row.get('状态', '待发送')
                if status == '待发送' or pd.isna(status) or status == '':
                    recipients.append({
                        'name': row.get('姓名', ''),
                        'email': row.get('邮箱', ''),
                        'attachment': row.get('附件名称', ''),
                        'row_index': index
                    })

            return recipients

        except FileNotFoundError:
            print(f"错误：找不到Excel文件 {self.excel_file}")
            print("请先创建发送列表Excel文件")
            return []
        except Exception as e:
            print(f"读取Excel文件时出错：{e}")
            return []

    def format_email_content(self, template, recipient_name):
        """
        替换邮件模板中的变量
        """
        content = template.replace('{recipient_name}', recipient_name)
        content = content.replace('{sender_name}', self.sender_name)
        return content

    def send_email_with_attachment(self, recipient_name, recipient_email, attachment_filename):
        """
        发送单封带附件的邮件
        """
        try:
            # 创建邮件对象
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = recipient_email

            # 替换邮件标题中的变量
            subject = self.format_email_content(self.email_subject_template, recipient_name)
            msg['Subject'] = subject

            # 添加必要的邮件头信息
            from email.utils import make_msgid, formatdate
            msg['Message-ID'] = make_msgid()
            msg['Date'] = formatdate(localtime=True)
            msg['MIME-Version'] = '1.0'

            # 替换邮件正文中的变量
            body = self.format_email_content(self.email_body_template, recipient_name)

            # 添加邮件正文
            msg.attach(MIMEText(body, 'plain', 'utf-8'))

            # 添加附件（如果提供了附件名称）
            if attachment_filename and attachment_filename.strip():
                attachment_path = os.path.join(self.attachments_folder, attachment_filename.strip())

                if os.path.exists(attachment_path):
                    with open(attachment_path, 'rb') as attachment:
                        # 根据文件扩展名设置正确的MIME类型
                        import mimetypes
                        content_type, encoding = mimetypes.guess_type(attachment_path)
                        if content_type is None or encoding is not None:
                            content_type = 'application/octet-stream'

                        main_type, sub_type = content_type.split('/', 1)
                        part = MIMEBase(main_type, sub_type)
                        part.set_payload(attachment.read())
                        encoders.encode_base64(part)

                        # 获取文件名并处理编码
                        filename = os.path.basename(attachment_path)
                        # 使用RFC 2231编码处理中文文件名
                        from email.header import Header

                        # 对文件名进行编码，支持中文
                        encoded_filename = Header(filename, 'utf-8').encode()
                        part.add_header(
                            'Content-Disposition',
                            f'attachment; filename*={encoded_filename}'
                        )
                        msg.attach(part)
                else:
                    print(f"警告：附件文件不存在 {attachment_path}")
                    return False

            # 连接SMTP服务器并发送邮件
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            server.login(self.sender_email, self.sender_password)

            text = msg.as_string()
            server.sendmail(self.sender_email, recipient_email, text)
            server.quit()

            return True

        except Exception as e:
            print(f"发送邮件失败：{e}")
            return False

    def send_all_emails(self):
        """
        发送所有邮件
        """
        # 从Excel文件加载收件人信息
        recipients = self.load_recipients_from_excel()

        if not recipients:
            print("没有需要发送的邮件")
            return

        total_recipients = len(recipients)
        print(f"准备发送 {total_recipients} 封邮件")

        success_count = 0
        fail_count = 0

        # 用于记录实际发送结果，用于后续更新状态
        send_results = []

        for index, recipient in enumerate(recipients):
            print(f"\n[{index + 1}/{total_recipients}] 发送给：{recipient['name']} ({recipient['email']})")

            # 发送邮件
            send_success = self.send_email_with_attachment(
                recipient['name'],
                recipient['email'],
                recipient['attachment']
            )

            if send_success:
                print(f"✓ 发送成功")
                success_count += 1
                send_results.append({'row_index': recipient['row_index'], 'status': '已发送'})
            else:
                print(f"✗ 发送失败")
                fail_count += 1
                send_results.append({'row_index': recipient['row_index'], 'status': '发送失败'})

            # 批量发送间隔
            if (index + 1) % self.emails_per_batch == 0 and index < total_recipients - 1:
                print(f"批量发送 {self.emails_per_batch} 封邮件完成，等待 {self.delay_between_emails} 秒...")
                time.sleep(self.delay_between_emails)
            elif index < total_recipients - 1:
                # 每封邮件之间的短暂延迟
                time.sleep(1)

        # 发送统计
        print(f"\n=== 发送完成 ===")
        print(f"总计：{total_recipients}")
        print(f"成功：{success_count}")
        print(f"失败：{fail_count}")

        # 更新Excel文件状态
        self.update_excel_status(send_results)

    def update_excel_status(self, send_results):
        """
        更新Excel文件中的发送状态
        """
        try:
            # 读取原始Excel文件
            df = pd.read_excel(self.excel_file)

            # 确保有状态列
            if '状态' not in df.columns:
                df['状态'] = ''

            # 更新状态列
            for result in send_results:
                df.at[result['row_index'], '状态'] = result['status']

            # 保存回原文件
            df.to_excel(self.excel_file, index=False)
            print(f"状态已更新到：{self.excel_file}")

        except Exception as e:
            print(f"更新Excel状态失败：{e}")

def main():
    """
    主函数
    """
    # 切换到脚本所在目录，确保相对路径正确
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    print("=== 通用附件群发邮件工具 ===")
    print("功能：从Excel文件读取收件人列表并发送带附件的邮件")
    print()

    # 检查Excel文件是否存在
    excel_file = '发送列表.xlsx'
    if not os.path.exists(excel_file):
        print(f"错误：找不到Excel文件 {excel_file}")
        print("请先创建发送列表Excel文件（包含：姓名、邮箱、附件名称三列）")
        return

    # 检查附件文件夹
    attachments_folder = 'attachments'
    if not os.path.exists(attachments_folder):
        print(f"警告：attachments 文件夹不存在，将创建")
        os.makedirs(attachments_folder)

    # 确认发送
    print("准备发送邮件：")
    print(f"Excel文件：{excel_file}")
    print(f"附件文件夹：{attachments_folder}")
    print(f"邮件模板：prompts.py")
    print()

    confirm = input("确认开始发送吗？(y/N): ").strip().lower()
    if confirm != 'y':
        print("已取消发送")
        return

    try:
        # 创建发送器
        sender = BulkEmailSender()

        # 发送所有邮件
        sender.send_all_emails()

    except Exception as e:
        print(f"发生错误：{e}")
        import traceback
        traceback.print_exc()
        return

    print("\n=== 全部完成 ===")

if __name__ == "__main__":
    main()
