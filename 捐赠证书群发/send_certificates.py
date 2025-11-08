#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从Excel文件发送证书邮件
读取证书信息确认表.xlsx中的数据，发送证书邮件
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

# 加载环境变量（从上级目录读取.env文件）
script_dir = os.path.dirname(os.path.abspath(__file__))
env_path = os.path.join(script_dir, '..', '.env')
load_dotenv(dotenv_path=env_path)

class ExcelCertificateSender:
    """
    从Excel文件发送证书邮件的类
    """
    
    def __init__(self):
        # 邮箱配置
        self.smtp_server = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
        self.smtp_port = int(os.getenv('SMTP_PORT', '587'))
        self.sender_email = os.getenv('SENDER_EMAIL')
        self.sender_password = os.getenv('SENDER_PASSWORD')
        
        # 邮件内容配置（直接写在代码中，避免.env文件多行文本问题）
        self.email_subject = os.getenv('EMAIL_SUBJECT', '您的捐赠证书')
        self.email_body = """尊敬的捐赠人{recipient_name}：

        您好！感谢您捐出的爱心物资，感谢您的支持。

        请查收您的捐赠证书。谢谢！

        善淘
"""
        
        # 发送配置
        self.delay_between_emails = int(os.getenv('DELAY_BETWEEN_EMAILS', '5'))
        self.emails_per_batch = int(os.getenv('EMAILS_PER_BATCH', '10'))
        
        # Excel文件配置
        self.excel_file = os.getenv('EXCEL_FILE', '证书信息确认表.xlsx')
        self.certificates_folder = os.getenv('CERTIFICATES_FOLDER', '捐赠证书')
        
        # 验证必要配置
        if not self.sender_email or not self.sender_password:
            raise ValueError("请设置 SENDER_EMAIL 和 SENDER_PASSWORD 环境变量")
    
    def load_recipients_from_excel(self):
        """
        从Excel文件加载收件人信息
        """
        try:
            # 读取Excel文件
            df = pd.read_excel(self.excel_file)
            
            # 过滤出需要发送的证书（状态为'待确认'）
            recipients = []
            for index, row in df.iterrows():
                if row.get('状态', '') == '待确认':
                    recipients.append({
                        'name': row.get('姓名', ''),
                        'email': row.get('邮箱', ''),
                        'filename': row.get('文件名', ''),
                        'filepath': row.get('文件路径', '')
                    })
            
            return recipients
            
        except FileNotFoundError:
            print(f"错误：找不到Excel文件 {self.excel_file}")
            print("请先运行 parse_certificates.py 生成Excel文件")
            return []
        except Exception as e:
            print(f"读取Excel文件时出错：{e}")
            return []
    
    def send_certificate_email(self, recipient_name, recipient_email, certificate_path):
        """
        发送单个证书邮件
        """
        try:
            # 创建邮件对象
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = recipient_email
            msg['Subject'] = self.email_subject
            
            # 添加必要的邮件头信息，符合Gmail要求
            from email.utils import make_msgid, formatdate
            msg['Message-ID'] = make_msgid()
            msg['Date'] = formatdate(localtime=True)
            msg['MIME-Version'] = '1.0'
            msg['Content-Type'] = 'multipart/mixed; charset=utf-8'
            
            # 替换邮件正文中的占位符
            body = self.email_body.replace('{recipient_name}', recipient_name)
            
            # 添加邮件正文
            msg.attach(MIMEText(body, 'plain', 'utf-8'))
            
            # 添加证书附件
            if os.path.exists(certificate_path):
                with open(certificate_path, 'rb') as attachment:
                    # 根据文件扩展名设置正确的MIME类型
                    import mimetypes
                    content_type, encoding = mimetypes.guess_type(certificate_path)
                    if content_type is None or encoding is not None:
                        content_type = 'application/octet-stream'
                    
                    main_type, sub_type = content_type.split('/', 1)
                    part = MIMEBase(main_type, sub_type)
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    
                    # 获取文件名并处理编码
                    filename = os.path.basename(certificate_path)
                    # 使用RFC 2231编码处理中文文件名
                    from email.utils import formataddr
                    from email.header import Header
                    
                    # 对文件名进行编码，支持中文
                    encoded_filename = Header(filename, 'utf-8').encode()
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename*={encoded_filename}'
                    )
                    msg.attach(part)
            else:
                print(f"警告：证书文件不存在 {certificate_path}")
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
    
    def send_all_certificates(self):
        """
        发送所有证书
        """
        # 从Excel文件加载收件人信息
        recipients = self.load_recipients_from_excel()
        
        if not recipients:
            print("没有需要发送的证书")
            return
        
        total_recipients = len(recipients)
        print(f"准备发送 {total_recipients} 封证书邮件")
        
        success_count = 0
        fail_count = 0
        
        # 用于记录实际发送结果，用于后续更新状态
        send_results = []
        
        for index, recipient in enumerate(recipients):
            print(f"\n[{index + 1}/{total_recipients}] 发送给：{recipient['name']} ({recipient['email']})")
            
            # 检查证书文件路径
            certificate_path = recipient.get('filepath', '')
            if not certificate_path:
                # 如果Excel中没有文件路径，尝试从文件名构建
                certificate_path = os.path.join(self.certificates_folder, recipient['filename'])
            
            # 发送邮件
            send_success = self.send_certificate_email(
                recipient['name'], 
                recipient['email'], 
                certificate_path
            )
            
            if send_success:
                print(f"✓ 发送成功")
                success_count += 1
                send_results.append({'name': recipient['name'], 'email': recipient['email'], 'status': '已发送'})
            else:
                print(f"✗ 发送失败")
                fail_count += 1
                send_results.append({'name': recipient['name'], 'email': recipient['email'], 'status': '发送失败'})
            
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
            
            # 创建状态映射
            status_map = {}
            for result in send_results:
                key = (result['name'], result['email'])
                status_map[key] = result['status']
            
            # 更新状态列
            for index, row in df.iterrows():
                if row.get('状态') == '待确认':
                    key = (row.get('姓名'), row.get('邮箱'))
                    if key in status_map:
                        df.at[index, '状态'] = status_map[key]
            
            # 直接保存回原文件，不再生成新文件
            df.to_excel(self.excel_file, index=False)
            print(f"状态已更新到：{self.excel_file}")
            
        except Exception as e:
            print(f"更新Excel状态失败：{e}")

def main():
    """
    主函数
    """
    print("=== 证书邮件发送工具（Excel版）===")
    print("功能：从Excel文件读取证书信息并发送邮件")
    print()
    
    # 检查Excel文件是否存在
    excel_file = '证书信息确认表.xlsx'
    if not os.path.exists(excel_file):
        print(f"错误：找不到Excel文件 {excel_file}")
        print("请先运行 parse_certificates.py 生成Excel文件")
        return
    
    # 确认发送
    print("准备从以下文件发送证书：")
    print(f"Excel文件：{excel_file}")
    print(f"证书文件夹：捐赠证书")
    print()
    
    confirm = input("确认开始发送吗？(y/N): ").strip().lower()
    if confirm != 'y':
        print("已取消发送")
        return
    
    try:
        # 创建发送器
        sender = ExcelCertificateSender()
        
        # 发送所有证书
        sender.send_all_certificates()
        
    except Exception as e:
        print(f"发生错误：{e}")
        return
    
    print("\n=== 全部完成 ===")

if __name__ == "__main__":
    main()