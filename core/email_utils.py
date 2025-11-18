#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
邮件发送工具模块
提供SMTP连接管理、重试机制、日志记录等通用功能
"""

import smtplib
import time
import logging
import re
from functools import wraps
from typing import Optional, Callable, Any
from datetime import datetime
import os

class SMTPConnectionPool:
    """
    SMTP连接池管理器
    复用SMTP连接，提高发送效率和稳定性
    """

    def __init__(self, smtp_server: str, smtp_port: int,
                 sender_email: str, sender_password: str,
                 max_emails_per_connection: int = 50):
        """
        初始化SMTP连接池

        Args:
            smtp_server: SMTP服务器地址
            smtp_port: SMTP端口
            sender_email: 发件人邮箱
            sender_password: 邮箱密码/授权码
            max_emails_per_connection: 每个连接最多发送的邮件数
        """
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.sender_email = sender_email
        self.sender_password = sender_password
        self.max_emails_per_connection = max_emails_per_connection

        self.server: Optional[smtplib.SMTP] = None
        self.emails_sent_in_current_connection = 0
        self.logger = logging.getLogger(__name__)

    def connect(self) -> smtplib.SMTP:
        """建立SMTP连接"""
        try:
            self.logger.info(f"正在连接到SMTP服务器 {self.smtp_server}:{self.smtp_port}")
            server = smtplib.SMTP(self.smtp_server, self.smtp_port, timeout=30)
            server.starttls()
            server.login(self.sender_email, self.sender_password)
            self.logger.info("SMTP连接成功")
            return server
        except Exception as e:
            self.logger.error(f"SMTP连接失败: {e}")
            raise

    def get_connection(self) -> smtplib.SMTP:
        """
        获取可用的SMTP连接
        如果连接不存在或已发送过多邮件，则重新建立连接
        """
        if (self.server is None or
            self.emails_sent_in_current_connection >= self.max_emails_per_connection):
            # 关闭旧连接
            if self.server is not None:
                try:
                    self.server.quit()
                    self.logger.info("已关闭旧的SMTP连接")
                except:
                    pass

            # 建立新连接
            self.server = self.connect()
            self.emails_sent_in_current_connection = 0

        return self.server

    def send_email(self, msg: Any, recipient_email: str) -> bool:
        """
        通过连接池发送邮件

        Args:
            msg: 邮件消息对象
            recipient_email: 收件人邮箱

        Returns:
            bool: 发送是否成功
        """
        try:
            server = self.get_connection()
            text = msg.as_string()
            server.sendmail(self.sender_email, recipient_email, text)
            self.emails_sent_in_current_connection += 1
            self.logger.info(f"邮件发送成功: {recipient_email}")
            return True
        except Exception as e:
            self.logger.error(f"邮件发送失败: {recipient_email}, 错误: {e}")
            # 连接可能已失效，重置连接
            self.server = None
            raise

    def close(self):
        """关闭SMTP连接"""
        if self.server is not None:
            try:
                self.server.quit()
                self.logger.info("SMTP连接已关闭")
            except:
                pass
            finally:
                self.server = None

    def __enter__(self):
        """支持with语句"""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """支持with语句，自动关闭连接"""
        self.close()


def retry_on_failure(max_retries: int = 3, delay: int = 2, backoff: int = 2):
    """
    失败重试装饰器

    Args:
        max_retries: 最大重试次数
        delay: 初始延迟秒数
        backoff: 延迟倍数（指数退避）
    """
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            logger = logging.getLogger(__name__)
            current_delay = delay

            for attempt in range(max_retries + 1):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    if attempt == max_retries:
                        logger.error(f"{func.__name__} 失败，已重试 {max_retries} 次: {e}")
                        raise

                    logger.warning(
                        f"{func.__name__} 失败 (尝试 {attempt + 1}/{max_retries + 1}), "
                        f"{current_delay}秒后重试: {e}"
                    )
                    time.sleep(current_delay)
                    current_delay *= backoff

        return wrapper
    return decorator


def validate_email(email: str) -> bool:
    """
    验证邮箱地址格式

    Args:
        email: 邮箱地址

    Returns:
        bool: 是否有效
    """
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None


def setup_logger(name: str = 'email_sender', log_dir: str = 'logs') -> logging.Logger:
    """
    配置日志记录器

    Args:
        name: 日志记录器名称
        log_dir: 日志文件目录

    Returns:
        logging.Logger: 配置好的日志记录器
    """
    # 创建日志目录
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    # 创建日志记录器
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    # 避免重复添加handler
    if logger.handlers:
        return logger

    # 文件处理器 - 详细日志
    log_file = os.path.join(log_dir, f'{name}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    file_handler.setFormatter(file_formatter)

    # 控制台处理器 - 简洁输出
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_formatter = logging.Formatter('%(levelname)s - %(message)s')
    console_handler.setFormatter(console_formatter)

    # 添加处理器
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger


def test_smtp_connection(smtp_server: str, smtp_port: int,
                        sender_email: str, sender_password: str) -> tuple[bool, str]:
    """
    测试SMTP连接

    Args:
        smtp_server: SMTP服务器地址
        smtp_port: SMTP端口
        sender_email: 发件人邮箱
        sender_password: 邮箱密码/授权码

    Returns:
        tuple[bool, str]: (是否成功, 消息)
    """
    try:
        server = smtplib.SMTP(smtp_server, smtp_port, timeout=10)
        server.starttls()
        server.login(sender_email, sender_password)
        server.quit()
        return True, "SMTP连接测试成功"
    except smtplib.SMTPAuthenticationError:
        return False, "认证失败：请检查邮箱地址和密码/授权码"
    except smtplib.SMTPException as e:
        return False, f"SMTP错误：{e}"
    except Exception as e:
        return False, f"连接失败：{e}"


def format_time_remaining(seconds: float) -> str:
    """
    格式化剩余时间

    Args:
        seconds: 剩余秒数

    Returns:
        str: 格式化的时间字符串
    """
    if seconds < 60:
        return f"{int(seconds)}秒"
    elif seconds < 3600:
        minutes = int(seconds / 60)
        secs = int(seconds % 60)
        return f"{minutes}分{secs}秒"
    else:
        hours = int(seconds / 3600)
        minutes = int((seconds % 3600) / 60)
        return f"{hours}小时{minutes}分"
