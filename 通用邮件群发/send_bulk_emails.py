#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通用附件群发邮件工具
从Excel文件读取收件人列表，发送带附件的邮件
支持3个自定义变量、双附件、自动重试、详细日志等功能
"""

import os
import smtplib
import time
import pandas as pd
import logging
import re
import mimetypes
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
from email.utils import make_msgid, formatdate
from pathlib import Path
from dotenv import load_dotenv
from datetime import datetime
import sys

# 导入邮件模板配置
from prompts import SENDER_NAME, EMAIL_SUBJECT, EMAIL_BODY

# 加载环境变量（从上级目录读取.env文件）
script_dir = os.path.dirname(os.path.abspath(__file__))
env_path = os.path.join(script_dir, '..', '.env')
load_dotenv(dotenv_path=env_path)

# 配置日志
log_file = os.path.join(script_dir, 'email_sender.log')
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


class EmailValidator:
    """
    邮件地址验证工具类
    """

    @staticmethod
    def is_valid_email(email):
        """
        验证邮件地址格式是否有效
        """
        if not email or not isinstance(email, str):
            return False

        # 基本的邮件格式验证正则表达式
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email.strip()) is not None


class BulkEmailSender:
    """
    通用附件群发邮件的类
    支持：
    - 3个自定义变量（var1, var2, var3）
    - 双附件支持
    - 邮件地址验证
    - 自动重试机制
    - SMTP连接池
    - 详细日志记录
    - Dry-run模式
    """

    def __init__(self):
        # 邮箱配置
        self.smtp_server = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
        self.smtp_port = int(os.getenv('SMTP_PORT', '587'))
        self.sender_email: str = os.getenv('SENDER_EMAIL') or ''
        self.sender_password: str = os.getenv('SENDER_PASSWORD') or ''

        # 发件人姓名（从prompts.py导入）
        self.sender_name = SENDER_NAME

        # 邮件模板（从prompts.py导入）
        self.email_subject_template = EMAIL_SUBJECT
        self.email_body_template = EMAIL_BODY

        # 发送配置
        self.delay_between_emails = int(os.getenv('DELAY_BETWEEN_EMAILS', '5'))
        self.emails_per_batch = int(os.getenv('EMAILS_PER_BATCH', '10'))

        # 可靠性配置（硬编码，不需要用户配置）
        self.max_retry_attempts = 3  # 失败重试次数
        self.retry_delay = 10  # 重试间隔时间（秒）

        # Dry-run模式（测试模式，不实际发送邮件）
        self.dry_run = os.getenv('DRY_RUN', 'false').lower() == 'true'

        # 文件配置
        self.excel_file = '发送列表.xlsx'
        self.attachments_folder = 'attachments'

        # SMTP连接池（用于复用连接）
        self.smtp_connection = None
        self.connection_email_count = 0
        self.max_emails_per_connection = 50  # 每个连接最多发送50封邮件

        # 验证必要配置
        if not self.sender_email or not self.sender_password:
            raise ValueError("请设置 SENDER_EMAIL 和 SENDER_PASSWORD 环境变量")

        if self.sender_name == "您的姓名":
            logger.warning("请在 prompts.py 中设置 SENDER_NAME")

        if self.dry_run:
            logger.info("=== DRY RUN 模式已启用 - 不会实际发送邮件 ===")

    def validate_data_before_send(self, recipients):
        """
        发送前验证所有数据
        """
        logger.info("正在进行发送前验证...")
        errors = []
        warnings = []

        for i, recipient in enumerate(recipients):
            row_num = i + 1

            # 验证邮箱格式
            if not EmailValidator.is_valid_email(recipient['email']):
                errors.append(f"第{row_num}行：邮箱格式无效 '{recipient['email']}'")

            # 检查附件文件是否存在
            for att_num in [1, 2]:
                att_key = f'attachment{att_num}'
                if recipient.get(att_key) and recipient[att_key].strip():
                    att_path = os.path.join(self.attachments_folder, recipient[att_key].strip())
                    if not os.path.exists(att_path):
                        warnings.append(f"第{row_num}行：附件文件不存在 '{att_path}'")

            # 检查是否至少有一个变量有值（可选检查）
            if not any([recipient.get('var1'), recipient.get('var2'), recipient.get('var3')]):
                warnings.append(f"第{row_num}行：所有变量(var1, var2, var3)都为空")

        # 显示错误和警告
        if errors:
            logger.error("发现以下错误：")
            for error in errors:
                logger.error(f"  - {error}")
            return False

        if warnings:
            logger.warning("发现以下警告：")
            for warning in warnings:
                logger.warning(f"  - {warning}")

            if not self.dry_run:
                confirm = input("\n存在警告，是否继续发送？(y/N): ").strip().lower()
                if confirm != 'y':
                    logger.info("用户取消发送")
                    return False

        logger.info("验证通过！")
        return True

    def load_recipients_from_excel(self):
        """
        从Excel文件加载收件人信息
        新格式：收件邮箱 | var1 | var2 | var3 | 附件名称1 | 附件名称2 | 发送情况
        """
        try:
            # 读取Excel文件
            df = pd.read_excel(self.excel_file)

            logger.info(f"从 {self.excel_file} 读取数据...")
            logger.info(f"Excel列名: {list(df.columns)}")

            # 检查必要的列
            required_columns = ['收件邮箱', 'var1', 'var2', 'var3', '附件名称1', '附件名称2', '发送情况']
            missing_columns = [col for col in required_columns if col not in df.columns]

            if missing_columns:
                logger.error(f"Excel文件缺少以下列：{', '.join(missing_columns)}")
                logger.error(f"要求的列格式：{' | '.join(required_columns)}")
                return []

            # 过滤出需要发送的邮件
            # 状态为 0 或空的才发送（1 表示已发送）
            recipients = []
            for index, row in df.iterrows():
                status = row.get('发送情况', 0)

                # 将状态转换为数字，处理各种可能的输入
                if pd.isna(status) or status == '' or status == '0':  # type: ignore[arg-type]
                    status_num = 0
                elif status == 1 or status == '1':
                    status_num = 1
                else:
                    # 其他情况视为未发送
                    status_num = 0

                # 只发送状态为 0 的邮件
                if status_num == 0:
                    # 安全地获取值，如果为NaN则转为空字符串
                    def safe_get(key: str) -> str:
                        val = row.get(key, '')
                        return '' if pd.isna(val) else str(val).strip()  # type: ignore[arg-type]

                    recipients.append({
                        'email': safe_get('收件邮箱'),
                        'var1': safe_get('var1'),
                        'var2': safe_get('var2'),
                        'var3': safe_get('var3'),
                        'attachment1': safe_get('附件名称1'),
                        'attachment2': safe_get('附件名称2'),
                        'row_index': index
                    })

            logger.info(f"找到 {len(recipients)} 条待发送记录")
            return recipients

        except FileNotFoundError:
            logger.error(f"找不到Excel文件 {self.excel_file}")
            logger.error("请先创建发送列表Excel文件")
            return []
        except Exception as e:
            logger.error(f"读取Excel文件时出错：{e}")
            import traceback
            logger.error(traceback.format_exc())
            return []

    def format_email_content(self, template, var1='', var2='', var3=''):
        """
        替换邮件模板中的变量
        支持：{var1}, {var2}, {var3}, {sender_name}
        """
        content = template.replace('{var1}', var1 if var1 else '')
        content = content.replace('{var2}', var2 if var2 else '')
        content = content.replace('{var3}', var3 if var3 else '')
        content = content.replace('{sender_name}', self.sender_name)
        return content

    def get_or_create_smtp_connection(self):
        """
        获取或创建SMTP连接（连接池）
        """
        # 如果没有连接或连接已发送太多邮件，则重新创建
        if self.smtp_connection is None or self.connection_email_count >= self.max_emails_per_connection:
            # 关闭旧连接
            if self.smtp_connection is not None:
                try:
                    self.smtp_connection.quit()
                    logger.debug("关闭旧的SMTP连接")
                except Exception:  # pylint: disable=broad-except
                    pass

            # 创建新连接
            try:
                logger.debug(f"连接到SMTP服务器 {self.smtp_server}:{self.smtp_port}")
                self.smtp_connection = smtplib.SMTP(self.smtp_server, self.smtp_port, timeout=30)
                self.smtp_connection.starttls()
                # 确保邮箱和密码已配置
                if not self.sender_email or not self.sender_password:
                    raise ValueError("SENDER_EMAIL 和 SENDER_PASSWORD 未配置")
                self.smtp_connection.login(self.sender_email, self.sender_password)
                self.connection_email_count = 0
                logger.info("SMTP连接已建立")
            except Exception as e:
                logger.error(f"建立SMTP连接失败：{e}")
                self.smtp_connection = None
                raise

        return self.smtp_connection

    def close_smtp_connection(self):
        """
        关闭SMTP连接
        """
        if self.smtp_connection is not None:
            try:
                self.smtp_connection.quit()
                logger.debug("SMTP连接已关闭")
            except Exception:  # pylint: disable=broad-except
                pass
            self.smtp_connection = None
            self.connection_email_count = 0

    def attach_file(self, msg, attachment_filename):
        """
        向邮件添加附件
        """
        if not attachment_filename or not attachment_filename.strip():
            return True

        attachment_path = os.path.join(self.attachments_folder, attachment_filename.strip())

        if not os.path.exists(attachment_path):
            logger.warning(f"附件文件不存在：{attachment_path}")
            return False

        try:
            with open(attachment_path, 'rb') as attachment:
                # 根据文件扩展名设置正确的MIME类型
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
                part.add_header(
                    'Content-Disposition',
                    'attachment',
                    filename=('utf-8', '', filename)
                )
                msg.attach(part)
                logger.debug(f"成功添加附件：{filename}")
                return True
        except Exception as e:
            logger.error(f"添加附件失败 {attachment_path}：{e}")
            return False

    def send_email_with_attachments(self, recipient_email, var1, var2, var3,
                                    attachment1, attachment2, retry_count=0):
        """
        发送单封带附件的邮件（支持双附件）
        """
        try:
            # Dry-run模式：不实际发送
            if self.dry_run:
                logger.info(f"[DRY RUN] 模拟发送邮件到：{recipient_email}")
                logger.info(f"  变量：var1={var1}, var2={var2}, var3={var3}")
                logger.info(f"  附件：{attachment1}, {attachment2}")
                time.sleep(0.1)  # 模拟发送延迟
                return True

            # 创建邮件对象
            msg = MIMEMultipart()
            msg['From'] = self.sender_email if self.sender_email else ''
            msg['To'] = recipient_email

            # 替换邮件标题中的变量
            subject = self.format_email_content(self.email_subject_template, var1, var2, var3)
            msg['Subject'] = subject

            # 添加必要的邮件头信息
            msg['Message-ID'] = make_msgid()
            msg['Date'] = formatdate(localtime=True)
            msg['MIME-Version'] = '1.0'

            # 替换邮件正文中的变量
            body = self.format_email_content(self.email_body_template, var1, var2, var3)

            # 添加邮件正文
            msg.attach(MIMEText(body, 'plain', 'utf-8'))

            # 添加附件1
            att1_success = self.attach_file(msg, attachment1)

            # 添加附件2
            att2_success = self.attach_file(msg, attachment2)

            # 如果有附件但都加载失败，记录警告但继续发送
            if (attachment1 or attachment2) and not (att1_success or att2_success):
                logger.warning(f"所有附件都加载失败，将发送无附件邮件")

            # 获取SMTP连接并发送邮件
            server = self.get_or_create_smtp_connection()
            text = msg.as_string()
            # 确保发件人邮箱已配置
            sender = self.sender_email if self.sender_email else ''
            server.sendmail(sender, recipient_email, text)
            self.connection_email_count += 1

            logger.info(f"✓ 发送成功：{recipient_email}")
            return True

        except Exception as e:
            logger.error(f"发送邮件失败（尝试 {retry_count + 1}/{self.max_retry_attempts}）：{e}")

            # 如果是连接错误，关闭连接以便下次重新建立
            if isinstance(e, (smtplib.SMTPException, ConnectionError, TimeoutError)):
                self.close_smtp_connection()

            # 自动重试
            if retry_count < self.max_retry_attempts - 1:
                logger.info(f"等待 {self.retry_delay} 秒后重试...")
                time.sleep(self.retry_delay)
                return self.send_email_with_attachments(
                    recipient_email, var1, var2, var3,
                    attachment1, attachment2, retry_count + 1
                )

            return False

    def send_all_emails(self):
        """
        发送所有邮件
        """
        # 从Excel文件加载收件人信息
        recipients = self.load_recipients_from_excel()

        if not recipients:
            logger.warning("没有需要发送的邮件")
            return

        # 发送前验证
        if not self.validate_data_before_send(recipients):
            logger.error("验证失败，取消发送")
            return

        total_recipients = len(recipients)
        logger.info(f"准备发送 {total_recipients} 封邮件")

        success_count = 0
        fail_count = 0

        # 用于记录实际发送结果，用于后续更新状态
        send_results = []

        start_time = datetime.now()

        for index, recipient in enumerate(recipients):
            logger.info(f"\n[{index + 1}/{total_recipients}] 发送给：{recipient['email']}")
            logger.info(f"  var1={recipient['var1']}, var2={recipient['var2']}, var3={recipient['var3']}")

            # 发送邮件
            send_success = self.send_email_with_attachments(
                recipient['email'],
                recipient['var1'],
                recipient['var2'],
                recipient['var3'],
                recipient['attachment1'],
                recipient['attachment2']
            )

            if send_success:
                success_count += 1
                send_results.append({'row_index': recipient['row_index'], 'status': 1})
            else:
                fail_count += 1
                send_results.append({'row_index': recipient['row_index'], 'status': 0})

            # 批量发送间隔
            if (index + 1) % self.emails_per_batch == 0 and index < total_recipients - 1:
                logger.info(f"批量发送 {self.emails_per_batch} 封邮件完成，等待 {self.delay_between_emails} 秒...")
                time.sleep(self.delay_between_emails)
            elif index < total_recipients - 1:
                # 每封邮件之间的短暂延迟
                time.sleep(1)

        # 关闭SMTP连接
        self.close_smtp_connection()

        # 计算总耗时
        end_time = datetime.now()
        duration = end_time - start_time

        # 发送统计
        logger.info(f"\n=== 发送完成 ===")
        logger.info(f"总计：{total_recipients}")
        logger.info(f"成功：{success_count}")
        logger.info(f"失败：{fail_count}")
        logger.info(f"耗时：{duration}")
        logger.info(f"日志文件：{log_file}")

        # 更新Excel文件状态
        if not self.dry_run:
            self.update_excel_status(send_results)

    def update_excel_status(self, send_results):
        """
        更新Excel文件中的发送状态
        """
        try:
            # 读取原始Excel文件
            df = pd.read_excel(self.excel_file)

            # 确保有发送情况列
            if '发送情况' not in df.columns:
                df['发送情况'] = ''

            # 更新状态列
            for result in send_results:
                df.at[result['row_index'], '发送情况'] = result['status']

            # 保存回原文件
            df.to_excel(self.excel_file, index=False)
            logger.info(f"状态已更新到：{self.excel_file}")

        except Exception as e:
            logger.error(f"更新Excel状态失败：{e}")


def main():
    """
    主函数
    """
    # 切换到脚本所在目录，确保相对路径正确
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    print("=" * 50)
    print("通用附件群发邮件工具")
    print("=" * 50)
    print("功能：")
    print("  - 支持3个自定义变量（var1, var2, var3）")
    print("  - 支持双附件")
    print("  - 自动重试失败的邮件")
    print("  - 详细日志记录")
    print("  - 邮件地址验证")
    print("=" * 50)
    print()

    # 检查Excel文件是否存在
    excel_file = '发送列表.xlsx'
    if not os.path.exists(excel_file):
        logger.error(f"找不到Excel文件 {excel_file}")
        logger.error("请先创建发送列表Excel文件")
        logger.error("要求的列格式：收件邮箱 | var1 | var2 | var3 | 附件名称1 | 附件名称2 | 发送情况")
        logger.error("发送情况列的值：0=未发送，1=已发送")
        return

    # 检查附件文件夹
    attachments_folder = 'attachments'
    if not os.path.exists(attachments_folder):
        logger.warning(f"attachments 文件夹不存在，将创建")
        os.makedirs(attachments_folder)

    # 显示配置信息
    logger.info("当前配置：")
    logger.info(f"  Excel文件：{excel_file}")
    logger.info(f"  附件文件夹：{attachments_folder}")
    logger.info(f"  邮件模板：prompts.py")
    logger.info(f"  日志文件：email_sender.log")

    dry_run = os.getenv('DRY_RUN', 'false').lower() == 'true'
    if dry_run:
        logger.info(f"  模式：DRY RUN（测试模式，不实际发送）")
    print()

    # 确认发送
    if not dry_run:
        confirm = input("确认开始发送吗？(y/N): ").strip().lower()
        if confirm != 'y':
            logger.info("已取消发送")
            return

    try:
        # 创建发送器
        sender = BulkEmailSender()

        # 发送所有邮件
        sender.send_all_emails()

    except Exception as e:
        logger.error(f"发生错误：{e}")
        import traceback
        logger.error(traceback.format_exc())
        return

    print("\n=== 全部完成 ===")


if __name__ == "__main__":
    main()
