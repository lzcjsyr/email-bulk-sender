"""
核心邮件发送模块

包含邮件发送、安全验证和工具函数
"""

from .email_enhanced import (
    SMTPErrorClassifier,
    ExponentialBackoff,
    EnhancedEmailBuilder,
    SmartRetryHandler,
    BounceHandler,
    add_unsubscribe_footer,
)

# TODO: Add email_utils exports when needed
# from .email_utils import ...

# TODO: Add email_security exports when needed
# from .email_security import ...

__version__ = "3.0.0"
__all__ = [
    "SMTPErrorClassifier",
    "ExponentialBackoff",
    "EnhancedEmailBuilder",
    "SmartRetryHandler",
    "BounceHandler",
    "add_unsubscribe_footer",
]
