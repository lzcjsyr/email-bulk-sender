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
import argparse
from datetime import datetime

# 导入工具模块
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from core.email_utils import (
    SMTPConnectionPool, retry_on_failure, validate_email,
    setup_logger, test_smtp_connection, format_time_remaining
)
try:
    from tqdm import tqdm
    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False

# 加载环境变量（从上级目录读取.env文件）
script_dir = os.path.dirname(os.path.abspath(__file__))
env_path = os.path.join(script_dir, '..', '.env')
load_dotenv(dotenv_path=env_path)

class ExcelCertificateSender:
    """
    从Excel文件发送证书邮件的类
    """

    def __init__(self, dry_run=False, max_retries=3):
        # 邮箱配置
        self.smtp_server = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
        self.smtp_port = int(os.getenv('SMTP_PORT', '587'))
        self.sender_email = os.getenv('SENDER_EMAIL')
        self.sender_password = os.getenv('SENDER_PASSWORD')

        # 邮件内容配置
        self.email_subject = os.getenv('EMAIL_SUBJECT', '您的捐赠证书')
        self.email_body = """尊敬的捐赠人{recipient_name}：

        您好！感谢您捐出的爱心物资，感谢您的支持。

        请查收您的捐赠证书。谢谢！

        善淘
"""

        # 发送配置
        self.delay_between_emails = int(os.getenv('DELAY_BETWEEN_EMAILS', '5'))
        self.emails_per_batch = int(os.getenv('EMAILS_PER_BATCH', '10'))
        self.max_retries = max_retries
        self.dry_run = dry_run

        # Excel文件配置
        self.excel_file = os.getenv('EXCEL_FILE', '证书信息确认表.xlsx')
        self.certificates_folder = os.getenv('CERTIFICATES_FOLDER', '捐赠证书')

        # 初始化日志
        self.logger = setup_logger('certificate_sender')

        # 验证必要配置
        if not self.sender_email or not self.sender_password:
            raise ValueError("请设置 SENDER_EMAIL 和 SENDER_PASSWORD 环境变量")

        # SMTP连接池
        self.connection_pool = None
    
    def load_recipients_from_excel(self):
        """
        从Excel文件加载收件人信息
        """
        try:
            # 读取Excel文件
            df = pd.read_excel(self.excel_file)

            # 过滤出需要发送的证书（状态为'待确认'）
            recipients = []
            invalid_emails = []

            for index, row in df.iterrows():
                if row.get('状态', '') == '待确认':
                    email = row.get('邮箱', '').strip()

                    # 验证邮箱格式
                    if not validate_email(email):
                        invalid_emails.append((index + 2, email))
                        self.logger.warning(f"第{index + 2}行邮箱格式无效: {email}")
                        continue

                    recipients.append({
                        'name': row.get('姓名', ''),
                        'email': email,
                        'filename': row.get('文件名', ''),
                        'filepath': row.get('文件路径', '')
                    })

            if invalid_emails:
                self.logger.warning(f"发现 {len(invalid_emails)} 个无效邮箱地址，已跳过")

            return recipients

        except FileNotFoundError:
            self.logger.error(f"找不到Excel文件 {self.excel_file}")
            self.logger.error("请先运行 parse_certificates.py 生成Excel文件")
            return []
        except Exception as e:
            self.logger.error(f"读取Excel文件时出错：{e}")
            return []
    
    @retry_on_failure(max_retries=3, delay=2, backoff=2)
    def send_certificate_email(self, recipient_name, recipient_email, certificate_path):
        """
        发送单个证书邮件（带重试机制）
        """
        if self.dry_run:
            self.logger.info(f"[模拟模式] 将发送给: {recipient_name} <{recipient_email}>")
            return True

        try:
            # 创建邮件对象
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = recipient_email
            msg['Subject'] = self.email_subject

            # 添加必要的邮件头信息
            from email.utils import make_msgid, formatdate
            msg['Message-ID'] = make_msgid()
            msg['Date'] = formatdate(localtime=True)
            msg['MIME-Version'] = '1.0'

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
                    from email.header import Header

                    # 对文件名进行编码，支持中文
                    encoded_filename = Header(filename, 'utf-8').encode()
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename*={encoded_filename}'
                    )
                    msg.attach(part)
            else:
                self.logger.error(f"证书文件不存在: {certificate_path}")
                return False

            # 使用连接池发送邮件
            self.connection_pool.send_email(msg, recipient_email)
            return True

        except Exception as e:
            self.logger.error(f"发送邮件失败 ({recipient_email}): {e}")
            raise  # 让retry装饰器处理重试
    
    def send_all_certificates(self):
        """
        发送所有证书
        """
        # 从Excel文件加载收件人信息
        recipients = self.load_recipients_from_excel()

        if not recipients:
            self.logger.info("没有需要发送的证书")
            return

        total_recipients = len(recipients)
        self.logger.info(f"准备发送 {total_recipients} 封证书邮件")

        # 预检查证书文件
        self.logger.info("正在检查证书文件...")
        missing_certificates = []
        for recipient in recipients:
            certificate_path = recipient.get('filepath', '')
            if not certificate_path:
                certificate_path = os.path.join(self.certificates_folder, recipient['filename'])

            if not os.path.exists(certificate_path):
                missing_certificates.append((recipient['name'], recipient['filename']))

        if missing_certificates:
            self.logger.warning(f"发现 {len(missing_certificates)} 个缺失的证书:")
            for name, filename in missing_certificates[:5]:
                self.logger.warning(f"  - {name}: {filename}")
            if len(missing_certificates) > 5:
                self.logger.warning(f"  ... 还有 {len(missing_certificates) - 5} 个")

        success_count = 0
        fail_count = 0
        send_results = []
        start_time = time.time()

        # 创建SMTP连接池
        with SMTPConnectionPool(
            self.smtp_server, self.smtp_port,
            self.sender_email, self.sender_password
        ) as pool:
            self.connection_pool = pool

            # 使用进度条（如果可用）
            iterator = tqdm(recipients, desc="发送证书", unit="封") if HAS_TQDM else recipients

            for index, recipient in enumerate(iterator):
                if not HAS_TQDM:
                    elapsed = time.time() - start_time
                    avg_time = elapsed / (index + 1) if index > 0 else 0
                    remaining = avg_time * (total_recipients - index - 1)
                    self.logger.info(
                        f"[{index + 1}/{total_recipients}] {recipient['name']} ({recipient['email']}) "
                        f"- 预计剩余: {format_time_remaining(remaining)}"
                    )

                # 检查证书文件路径
                certificate_path = recipient.get('filepath', '')
                if not certificate_path:
                    certificate_path = os.path.join(self.certificates_folder, recipient['filename'])

                # 发送邮件
                try:
                    send_success = self.send_certificate_email(
                        recipient['name'],
                        recipient['email'],
                        certificate_path
                    )
                except Exception:
                    send_success = False

                if send_success:
                    success_count += 1
                    send_results.append({'name': recipient['name'], 'email': recipient['email'], 'status': '已发送'})
                else:
                    fail_count += 1
                    send_results.append({'name': recipient['name'], 'email': recipient['email'], 'status': '发送失败'})

                # 批量发送间隔
                if (index + 1) % self.emails_per_batch == 0 and index < total_recipients - 1:
                    self.logger.info(f"已发送 {index + 1} 封，等待 {self.delay_between_emails} 秒...")
                    time.sleep(self.delay_between_emails)
                elif index < total_recipients - 1:
                    time.sleep(1)

        # 发送统计
        elapsed_time = time.time() - start_time
        self.logger.info(f"\n=== 发送完成 ===")
        self.logger.info(f"总计：{total_recipients} 封")
        self.logger.info(f"成功：{success_count} 封")
        self.logger.info(f"失败：{fail_count} 封")
        self.logger.info(f"用时：{format_time_remaining(elapsed_time)}")

        if not self.dry_run:
            # 更新Excel文件状态
            self.update_excel_status(send_results)
    
    def update_excel_status(self, send_results):
        """
        更新Excel文件中的发送状态（带备份）
        """
        try:
            # 读取原始Excel文件
            df = pd.read_excel(self.excel_file)

            # 备份原文件
            backup_file = f"{os.path.splitext(self.excel_file)[0]}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            df.to_excel(backup_file, index=False)
            self.logger.info(f"已备份原文件到: {backup_file}")

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

            # 保存回原文件
            df.to_excel(self.excel_file, index=False)
            self.logger.info(f"状态已更新到：{self.excel_file}")

        except Exception as e:
            self.logger.error(f"更新Excel状态失败：{e}")
            self.logger.error("发送结果未能保存到Excel，请手动更新")

def main():
    """
    主函数
    """
    # 解析命令行参数
    parser = argparse.ArgumentParser(
        description='证书邮件发送工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
使用示例:
  %(prog)s                        # 正常发送模式
  %(prog)s --dry-run              # 模拟发送，不实际发送邮件
  %(prog)s --test-smtp            # 测试SMTP连接
  %(prog)s --yes                  # 跳过确认直接发送
        '''
    )
    parser.add_argument('--dry-run', action='store_true',
                        help='模拟发送模式，不实际发送邮件')
    parser.add_argument('--test-smtp', action='store_true',
                        help='测试SMTP连接后退出')
    parser.add_argument('--yes', '-y', action='store_true',
                        help='跳过确认，直接开始发送')
    parser.add_argument('--max-retries', type=int, default=3,
                        help='每封邮件的最大重试次数 (默认: 3)')

    args = parser.parse_args()

    # 切换到脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    print("=== 证书邮件发送工具 ===")
    print("功能：从Excel文件读取证书信息并发送邮件")
    print()

    # 加载环境变量
    env_path = os.path.join(script_dir, '..', '.env')
    from dotenv import load_dotenv
    load_dotenv(dotenv_path=env_path)

    # 测试SMTP连接
    if args.test_smtp:
        print("正在测试SMTP连接...")
        smtp_server = os.getenv('SMTP_SERVER')
        smtp_port = int(os.getenv('SMTP_PORT', '587'))
        sender_email = os.getenv('SENDER_EMAIL')
        sender_password = os.getenv('SENDER_PASSWORD')

        if not all([smtp_server, sender_email, sender_password]):
            print("错误：请先配置 .env 文件")
            return

        success, message = test_smtp_connection(
            smtp_server, smtp_port, sender_email, sender_password
        )
        print(message)
        return

    # 检查Excel文件
    excel_file = '证书信息确认表.xlsx'
    if not os.path.exists(excel_file):
        print(f"错误：找不到Excel文件 {excel_file}")
        print("请先运行 parse_certificates.py 生成Excel文件")
        return

    # 显示配置信息
    print("准备发送证书：")
    print(f"Excel文件：{excel_file}")
    print(f"证书文件夹：捐赠证书")
    if args.dry_run:
        print("模式：模拟发送（不会实际发送邮件）")
    print()

    # 确认发送
    if not args.yes and not args.dry_run:
        confirm = input("确认开始发送吗？(y/N): ").strip().lower()
        if confirm != 'y':
            print("已取消发送")
            return

    try:
        # 创建发送器
        sender = ExcelCertificateSender(
            dry_run=args.dry_run,
            max_retries=args.max_retries
        )

        # 发送所有证书
        sender.send_all_certificates()

    except Exception as e:
        print(f"发生错误：{e}")
        import traceback
        traceback.print_exc()
        return

    print("\n=== 全部完成 ===")

if __name__ == "__main__":
    main()