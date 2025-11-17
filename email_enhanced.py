"""
邮件发送增强模块 - 提供HTML邮件、改进的头部、智能重试等功能
"""

import time
import random
import logging
import smtplib
from typing import Optional, Tuple, Dict
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import make_msgid, formatdate

try:
    from premailer import transform
    PREMAILER_AVAILABLE = True
except ImportError:
    PREMAILER_AVAILABLE = False
    logging.warning("premailer未安装,HTML CSS内联功能不可用")

try:
    import html2text
    HTML2TEXT_AVAILABLE = True
except ImportError:
    HTML2TEXT_AVAILABLE = False
    logging.warning("html2text未安装,HTML转纯文本功能不可用")


class SMTPErrorClassifier:
    """SMTP错误分类器 - 区分永久性和临时性错误"""

    # 永久性错误码(不应重试)
    PERMANENT_ERRORS = {
        501,  # 语法错误
        502,  # 命令未实现
        503,  # 错误的命令序列
        504,  # 命令参数未实现
        521,  # 服务器不接受邮件
        530,  # 需要认证
        535,  # 认证失败
        550,  # 邮箱不存在
        551,  # 用户不是本地的
        552,  # 邮箱存储空间不足
        553,  # 邮箱名不允许
        554,  # 交易失败
    }

    # 临时性错误码(可以重试)
    TEMPORARY_ERRORS = {
        421,  # 服务不可用
        450,  # 邮箱不可用(临时)
        451,  # 本地错误
        452,  # 系统存储不足
        455,  # 服务器无法接收参数
    }

    # 速率限制错误码
    RATE_LIMIT_ERRORS = {
        421,  # 速率限制
        450,  # 临时拒绝
        451,  # 本地错误/速率限制
        452,  # 邮件队列满
        454,  # 临时认证失败
    }

    @classmethod
    def classify_error(cls, exception: Exception) -> Tuple[str, bool, int]:
        """
        分类SMTP错误

        Args:
            exception: SMTP异常

        Returns:
            (错误类型, 是否应该重试, 建议等待时间)
            错误类型: 'permanent', 'temporary', 'rate_limit', 'connection', 'authentication', 'unknown'
        """
        error_str = str(exception).lower()

        # SMTP错误响应格式: (code, 'message')
        if isinstance(exception, smtplib.SMTPResponseException):
            code = exception.smtp_code

            # 认证错误
            if code in (530, 535):
                return 'authentication', False, 0

            # 永久性错误
            if code in cls.PERMANENT_ERRORS:
                return 'permanent', False, 0

            # 速率限制
            if code in cls.RATE_LIMIT_ERRORS or 'rate limit' in error_str or 'too many' in error_str:
                return 'rate_limit', True, 60  # 建议等待60秒

            # 临时性错误
            if code in cls.TEMPORARY_ERRORS:
                return 'temporary', True, 10

        # 连接错误
        if isinstance(exception, (smtplib.SMTPConnectError, smtplib.SMTPServerDisconnected,
                                   ConnectionError, TimeoutError, OSError)):
            return 'connection', True, 15

        # 认证错误
        if isinstance(exception, smtplib.SMTPAuthenticationError):
            return 'authentication', False, 0

        # 检查错误消息
        if 'authentication' in error_str or 'login' in error_str:
            return 'authentication', False, 0

        if 'connection' in error_str or 'timeout' in error_str:
            return 'connection', True, 15

        if 'rate' in error_str or 'quota' in error_str or 'limit' in error_str:
            return 'rate_limit', True, 60

        if 'mailbox' in error_str or 'recipient' in error_str or 'not found' in error_str:
            return 'permanent', False, 0

        # 默认: 未知错误,可以重试
        return 'unknown', True, 10

    @classmethod
    def get_error_message(cls, error_type: str) -> str:
        """获取用户友好的错误信息"""
        messages = {
            'permanent': '永久性错误(邮箱不存在或被拒绝)',
            'temporary': '临时性错误(服务器暂时不可用)',
            'rate_limit': '速率限制(发送过快,需要减速)',
            'connection': '连接错误(网络问题或服务器断开)',
            'authentication': '认证失败(用户名或密码错误)',
            'unknown': '未知错误'
        }
        return messages.get(error_type, '未知错误')


class ExponentialBackoff:
    """指数退避重试策略"""

    def __init__(self, base_delay: float = 10.0, max_delay: float = 300.0, jitter: bool = True):
        """
        初始化指数退避

        Args:
            base_delay: 基础延迟时间(秒)
            max_delay: 最大延迟时间(秒)
            jitter: 是否添加随机抖动(避免雷鸣群效应)
        """
        self.base_delay = base_delay
        self.max_delay = max_delay
        self.jitter = jitter

    def get_delay(self, attempt: int) -> float:
        """
        计算延迟时间

        Args:
            attempt: 当前尝试次数(从0开始)

        Returns:
            延迟时间(秒)
        """
        # 指数增长: base_delay * 2^attempt
        delay = min(self.base_delay * (2 ** attempt), self.max_delay)

        # 添加随机抖动(0-25%)
        if self.jitter:
            jitter_range = delay * 0.25
            delay += random.uniform(0, jitter_range)

        return delay


class EnhancedEmailBuilder:
    """增强的邮件构建器 - 支持HTML、改进的头部等"""

    def __init__(self, sender_email: str, sender_name: Optional[str] = None,
                 reply_to: Optional[str] = None, unsubscribe_email: Optional[str] = None):
        """
        初始化邮件构建器

        Args:
            sender_email: 发件人邮箱
            sender_name: 发件人姓名
            reply_to: 回复地址
            unsubscribe_email: 退订邮箱
        """
        self.sender_email = sender_email
        self.sender_name = sender_name
        self.reply_to = reply_to or sender_email
        self.unsubscribe_email = unsubscribe_email or sender_email

    def build_message(self, recipient_email: str, subject: str,
                      body_plain: str, body_html: Optional[str] = None,
                      extra_headers: Optional[Dict[str, str]] = None) -> MIMEMultipart:
        """
        构建邮件消息

        Args:
            recipient_email: 收件人邮箱
            subject: 主题
            body_plain: 纯文本正文
            body_html: HTML正文(可选)
            extra_headers: 额外的邮件头部

        Returns:
            构建好的邮件消息
        """
        # 创建multipart消息
        msg = MIMEMultipart('alternative')

        # 基本头部
        if self.sender_name:
            msg['From'] = f"{self.sender_name} <{self.sender_email}>"
        else:
            msg['From'] = self.sender_email

        msg['To'] = recipient_email
        msg['Subject'] = subject

        # 关键头部(提高送达率)
        msg['Message-ID'] = make_msgid(domain=self.sender_email.split('@')[-1])
        msg['Date'] = formatdate(localtime=True)
        msg['MIME-Version'] = '1.0'

        # Reply-To头部
        if self.reply_to and self.reply_to != self.sender_email:
            msg['Reply-To'] = self.reply_to

        # 退订头部(批量邮件必需)
        if self.unsubscribe_email:
            msg['List-Unsubscribe'] = f'<mailto:{self.unsubscribe_email}?subject=unsubscribe>'
            msg['List-Unsubscribe-Post'] = 'List-Unsubscribe=One-Click'

        # 批量邮件标识
        msg['Precedence'] = 'bulk'

        # 客户端标识
        msg['X-Mailer'] = 'Python Bulk Email Sender v2.0'

        # 添加自定义头部
        if extra_headers:
            for key, value in extra_headers.items():
                msg[key] = value

        # 添加纯文本正文
        part_plain = MIMEText(body_plain, 'plain', 'utf-8')
        msg.attach(part_plain)

        # 添加HTML正文(如果提供)
        if body_html:
            # 如果有premailer,内联CSS样式
            if PREMAILER_AVAILABLE:
                try:
                    body_html = transform(body_html)
                except Exception as e:
                    logging.warning(f"CSS内联失败: {e}")

            part_html = MIMEText(body_html, 'html', 'utf-8')
            msg.attach(part_html)

        return msg

    @staticmethod
    def html_to_plain(html: str) -> str:
        """
        将HTML转换为纯文本(作为备用)

        Args:
            html: HTML内容

        Returns:
            纯文本内容
        """
        if not HTML2TEXT_AVAILABLE:
            # 简单的清理
            import re
            text = re.sub('<[^<]+?>', '', html)  # 移除HTML标签
            text = text.replace('&nbsp;', ' ')
            text = text.replace('&amp;', '&')
            text = text.replace('&lt;', '<')
            text = text.replace('&gt;', '>')
            return text.strip()

        try:
            h = html2text.HTML2Text()
            h.ignore_links = False
            h.ignore_images = False
            h.body_width = 78
            return h.handle(html)
        except Exception as e:
            logging.warning(f"HTML转纯文本失败: {e}")
            return html


class SmartRetryHandler:
    """智能重试处理器"""

    def __init__(self, max_attempts: int = 3, base_delay: float = 10.0):
        """
        初始化重试处理器

        Args:
            max_attempts: 最大重试次数
            base_delay: 基础延迟时间(秒)
        """
        self.max_attempts = max_attempts
        self.backoff = ExponentialBackoff(base_delay=base_delay)
        self.classifier = SMTPErrorClassifier()

    def should_retry(self, exception: Exception, attempt: int) -> Tuple[bool, float, str]:
        """
        判断是否应该重试

        Args:
            exception: 发生的异常
            attempt: 当前尝试次数(从0开始)

        Returns:
            (是否重试, 等待时间, 错误类型描述)
        """
        # 达到最大重试次数
        if attempt >= self.max_attempts - 1:
            return False, 0, '已达最大重试次数'

        # 分类错误
        error_type, should_retry, suggested_delay = self.classifier.classify_error(exception)
        error_msg = self.classifier.get_error_message(error_type)

        if not should_retry:
            return False, 0, error_msg

        # 计算延迟时间(使用建议延迟或指数退避)
        if suggested_delay > 0:
            delay = suggested_delay
        else:
            delay = self.backoff.get_delay(attempt)

        # 速率限制错误需要更长的等待时间
        if error_type == 'rate_limit':
            delay = max(delay, 60)

        return True, delay, error_msg


class BounceHandler:
    """反弹邮件处理器"""

    @staticmethod
    def parse_smtp_response(exception: Exception) -> Optional[Dict[str, any]]:
        """
        解析SMTP响应,提取反弹信息

        Args:
            exception: SMTP异常

        Returns:
            反弹信息字典或None
        """
        if not isinstance(exception, smtplib.SMTPResponseException):
            return None

        code = exception.smtp_code
        message = str(exception.smtp_error) if hasattr(exception, 'smtp_error') else str(exception)

        result = {
            'code': code,
            'message': message,
            'is_bounce': False,
            'bounce_type': None,
            'reason': None
        }

        # 硬反弹(邮箱不存在等永久性错误)
        hard_bounce_codes = {550, 551, 552, 553, 554}
        if code in hard_bounce_codes:
            result['is_bounce'] = True
            result['bounce_type'] = 'hard'

            if code == 550:
                result['reason'] = '邮箱不存在或被拒绝'
            elif code == 551:
                result['reason'] = '用户不是本地用户'
            elif code == 552:
                result['reason'] = '邮箱存储空间已满'
            elif code == 553:
                result['reason'] = '邮箱名称不允许'
            elif code == 554:
                result['reason'] = '交易失败'

        # 软反弹(临时性错误)
        soft_bounce_codes = {421, 450, 451, 452}
        if code in soft_bounce_codes:
            result['is_bounce'] = True
            result['bounce_type'] = 'soft'

            if code == 421:
                result['reason'] = '服务不可用(临时)'
            elif code == 450:
                result['reason'] = '邮箱暂时不可用'
            elif code == 451:
                result['reason'] = '处理错误(临时)'
            elif code == 452:
                result['reason'] = '存储空间不足(临时)'

        return result


def add_unsubscribe_footer(body: str, unsubscribe_email: str, format_type: str = 'plain') -> str:
    """
    在邮件正文底部添加退订信息

    Args:
        body: 原始正文
        unsubscribe_email: 退订邮箱
        format_type: 格式类型('plain' 或 'html')

    Returns:
        添加退订信息后的正文
    """
    if format_type == 'html':
        footer = f"""
<hr style="margin-top: 30px; border: none; border-top: 1px solid #ccc;">
<p style="font-size: 12px; color: #666; text-align: center;">
    如果您不想再收到此类邮件,请回复邮件到
    <a href="mailto:{unsubscribe_email}?subject=unsubscribe">{unsubscribe_email}</a>
    并在主题中填写 "unsubscribe"
</p>
"""
        return body + footer
    else:
        footer = f"""

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
如果您不想再收到此类邮件,请回复邮件到 {unsubscribe_email}
并在主题中填写 "unsubscribe"
"""
        return body + footer


if __name__ == '__main__':
    # 测试代码
    print("邮件增强模块测试\n")

    # 测试错误分类
    classifier = SMTPErrorClassifier()

    test_errors = [
        (smtplib.SMTPResponseException(550, b'User not found'), '550 邮箱不存在'),
        (smtplib.SMTPResponseException(421, b'Rate limit exceeded'), '421 速率限制'),
        (smtplib.SMTPAuthenticationError(535, b'Authentication failed'), '535 认证失败'),
        (ConnectionError('Connection refused'), '连接错误'),
    ]

    print("错误分类测试:")
    for error, desc in test_errors:
        error_type, should_retry, delay = classifier.classify_error(error)
        print(f"  {desc}")
        print(f"    类型: {error_type}, 重试: {should_retry}, 延迟: {delay}秒")
        print(f"    描述: {classifier.get_error_message(error_type)}\n")

    # 测试指数退避
    print("\n指数退避测试:")
    backoff = ExponentialBackoff(base_delay=10.0, jitter=False)
    for i in range(5):
        delay = backoff.get_delay(i)
        print(f"  尝试 {i+1}: 等待 {delay:.1f} 秒")

    # 测试邮件构建
    print("\n邮件构建测试:")
    builder = EnhancedEmailBuilder(
        sender_email='test@example.com',
        sender_name='测试发件人',
        unsubscribe_email='unsubscribe@example.com'
    )

    msg = builder.build_message(
        recipient_email='recipient@example.com',
        subject='测试邮件',
        body_plain='这是纯文本内容',
        body_html='<h1>这是HTML内容</h1><p>测试段落</p>'
    )

    print(f"  From: {msg['From']}")
    print(f"  To: {msg['To']}")
    print(f"  Subject: {msg['Subject']}")
    print(f"  List-Unsubscribe: {msg.get('List-Unsubscribe', 'N/A')}")
    print(f"  Precedence: {msg.get('Precedence', 'N/A')}")
