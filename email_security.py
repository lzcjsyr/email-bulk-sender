"""
邮件安全模块 - 提供DKIM签名、SPF/DMARC验证、内容检查等功能
"""

import os
import re
import socket
import logging
from typing import Optional, Dict, List, Tuple
from email.message import EmailMessage

try:
    import dkim
    DKIM_AVAILABLE = True
except ImportError:
    DKIM_AVAILABLE = False
    logging.warning("dkimpy未安装,DKIM签名功能不可用。请运行: pip install dkimpy")

try:
    import dns.resolver
    DNS_AVAILABLE = True
except ImportError:
    DNS_AVAILABLE = False
    logging.warning("dnspython未安装,SPF/DMARC验证功能不可用。请运行: pip install dnspython")


class DKIMSigner:
    """DKIM邮件签名器"""

    def __init__(self, domain: str, selector: str, private_key_path: Optional[str] = None):
        """
        初始化DKIM签名器

        Args:
            domain: 发送域名 (例如: example.com)
            selector: DKIM选择器 (例如: default, mail)
            private_key_path: DKIM私钥文件路径
        """
        self.domain = domain
        self.selector = selector
        self.private_key = None

        if not DKIM_AVAILABLE:
            logging.warning("DKIM功能不可用: dkimpy未安装")
            return

        if private_key_path and os.path.exists(private_key_path):
            try:
                with open(private_key_path, 'rb') as f:
                    self.private_key = f.read()
                logging.info(f"DKIM私钥已加载: {private_key_path}")
            except Exception as e:
                logging.error(f"加载DKIM私钥失败: {e}")
        else:
            logging.warning(f"DKIM私钥文件不存在: {private_key_path}")

    def sign_message(self, message: bytes) -> bytes:
        """
        对邮件进行DKIM签名

        Args:
            message: 原始邮件内容(bytes)

        Returns:
            签名后的邮件内容(bytes)
        """
        if not DKIM_AVAILABLE or not self.private_key:
            logging.warning("跳过DKIM签名: 功能不可用或私钥未配置")
            return message

        try:
            signature = dkim.sign(
                message,
                self.selector.encode(),
                self.domain.encode(),
                self.private_key,
                include_headers=[b'from', b'to', b'subject', b'date', b'message-id']
            )

            # 将DKIM签名添加到邮件头部
            signed_message = signature + message
            logging.debug(f"DKIM签名成功: domain={self.domain}, selector={self.selector}")
            return signed_message

        except Exception as e:
            logging.error(f"DKIM签名失败: {e}")
            return message

    @staticmethod
    def generate_dns_record(selector: str, public_key: str) -> str:
        """
        生成DKIM DNS TXT记录

        Args:
            selector: DKIM选择器
            public_key: 公钥内容(PEM格式,去除头尾和换行)

        Returns:
            DNS记录字符串
        """
        # 清理公钥格式
        public_key_clean = public_key.replace('-----BEGIN PUBLIC KEY-----', '')
        public_key_clean = public_key_clean.replace('-----END PUBLIC KEY-----', '')
        public_key_clean = public_key_clean.replace('\n', '').replace('\r', '').strip()

        dns_record = f"{selector}._domainkey.yourdomain.com. IN TXT \"v=DKIM1; k=rsa; p={public_key_clean}\""
        return dns_record


class DNSValidator:
    """DNS记录验证器 - 用于检查SPF和DMARC配置"""

    def __init__(self):
        if not DNS_AVAILABLE:
            logging.warning("DNS验证功能不可用: dnspython未安装")
        self.resolver = dns.resolver.Resolver() if DNS_AVAILABLE else None
        # 配置超时
        if self.resolver:
            self.resolver.timeout = 5
            self.resolver.lifetime = 5

    def check_spf(self, domain: str) -> Tuple[bool, str]:
        """
        检查域名的SPF记录

        Args:
            domain: 要检查的域名

        Returns:
            (是否存在SPF记录, SPF记录内容或错误信息)
        """
        if not DNS_AVAILABLE:
            return False, "DNS功能不可用"

        try:
            answers = self.resolver.resolve(domain, 'TXT')
            spf_records = []

            for rdata in answers:
                txt_string = rdata.to_text().strip('"')
                if txt_string.startswith('v=spf1'):
                    spf_records.append(txt_string)

            if spf_records:
                return True, spf_records[0]
            else:
                return False, f"未找到SPF记录: {domain}"

        except dns.resolver.NXDOMAIN:
            return False, f"域名不存在: {domain}"
        except dns.resolver.NoAnswer:
            return False, f"未找到TXT记录: {domain}"
        except dns.resolver.Timeout:
            return False, f"DNS查询超时: {domain}"
        except Exception as e:
            return False, f"SPF检查失败: {str(e)}"

    def check_dmarc(self, domain: str) -> Tuple[bool, str]:
        """
        检查域名的DMARC记录

        Args:
            domain: 要检查的域名

        Returns:
            (是否存在DMARC记录, DMARC记录内容或错误信息)
        """
        if not DNS_AVAILABLE:
            return False, "DNS功能不可用"

        dmarc_domain = f"_dmarc.{domain}"

        try:
            answers = self.resolver.resolve(dmarc_domain, 'TXT')
            dmarc_records = []

            for rdata in answers:
                txt_string = rdata.to_text().strip('"')
                if txt_string.startswith('v=DMARC1'):
                    dmarc_records.append(txt_string)

            if dmarc_records:
                return True, dmarc_records[0]
            else:
                return False, f"未找到DMARC记录: {dmarc_domain}"

        except dns.resolver.NXDOMAIN:
            return False, f"DMARC域名不存在: {dmarc_domain}"
        except dns.resolver.NoAnswer:
            return False, f"未找到DMARC TXT记录: {dmarc_domain}"
        except dns.resolver.Timeout:
            return False, f"DNS查询超时: {dmarc_domain}"
        except Exception as e:
            return False, f"DMARC检查失败: {str(e)}"

    def check_mx(self, domain: str) -> Tuple[bool, List[str]]:
        """
        检查域名的MX记录

        Args:
            domain: 要检查的域名

        Returns:
            (是否存在MX记录, MX服务器列表或错误信息)
        """
        if not DNS_AVAILABLE:
            return False, ["DNS功能不可用"]

        try:
            answers = self.resolver.resolve(domain, 'MX')
            mx_records = [str(rdata.exchange).rstrip('.') for rdata in answers]
            return True, sorted(mx_records)

        except Exception as e:
            return False, [f"MX检查失败: {str(e)}"]


class ContentChecker:
    """邮件内容检查器 - 检测垃圾邮件特征"""

    # 中文垃圾邮件关键词
    SPAM_KEYWORDS_CN = [
        '点击领取', '立即领取', '免费领取', '恭喜中奖', '中奖通知',
        '紧急通知', '账户异常', '验证身份', '点击链接', '立即验证',
        '限时优惠', '仅限今日', '最后机会', '马上抢购', '秒杀',
        '赚钱', '兼职', '日赚', '月入', '暴富', '致富',
        '贷款', '无抵押', '零首付', '快速放款', '信用卡提额',
        '发票', '代开', '增值税', '刻章', '办证'
    ]

    # 英文垃圾邮件关键词
    SPAM_KEYWORDS_EN = [
        'free money', 'click here', 'act now', 'limited time',
        'congratulations', 'you won', 'winner', 'prize',
        'urgent', 'verify account', 'suspended account',
        'cheap', 'discount', 'lowest price', 'buy now',
        'earn money', 'work from home', 'make money fast',
        'loan', 'credit', 'debt', 'casino', 'viagra'
    ]

    # 可疑URL模式
    SUSPICIOUS_URL_PATTERNS = [
        r'bit\.ly', r'tinyurl\.com', r'goo\.gl',  # 短链接
        r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}',    # IP地址
        r'\.tk$', r'\.ml$', r'\.ga$', r'\.cf$'     # 免费域名
    ]

    def __init__(self):
        self.spam_keywords = self.SPAM_KEYWORDS_CN + self.SPAM_KEYWORDS_EN

    def check_spam_keywords(self, text: str) -> List[str]:
        """
        检查文本中的垃圾邮件关键词

        Args:
            text: 要检查的文本内容

        Returns:
            找到的垃圾邮件关键词列表
        """
        found_keywords = []
        text_lower = text.lower()

        for keyword in self.spam_keywords:
            if keyword.lower() in text_lower:
                found_keywords.append(keyword)

        return found_keywords

    def check_suspicious_urls(self, text: str) -> List[str]:
        """
        检查文本中的可疑URL

        Args:
            text: 要检查的文本内容

        Returns:
            找到的可疑URL列表
        """
        suspicious_urls = []

        # 提取所有URL
        url_pattern = r'https?://[^\s<>"]+|www\.[^\s<>"]+'
        urls = re.findall(url_pattern, text, re.IGNORECASE)

        for url in urls:
            for pattern in self.SUSPICIOUS_URL_PATTERNS:
                if re.search(pattern, url, re.IGNORECASE):
                    suspicious_urls.append(url)
                    break

        return suspicious_urls

    def check_content(self, subject: str, body: str) -> Dict[str, any]:
        """
        全面检查邮件内容

        Args:
            subject: 邮件主题
            body: 邮件正文

        Returns:
            检查结果字典
        """
        result = {
            'has_issues': False,
            'spam_keywords': [],
            'suspicious_urls': [],
            'warnings': []
        }

        # 检查垃圾邮件关键词
        spam_in_subject = self.check_spam_keywords(subject)
        spam_in_body = self.check_spam_keywords(body)
        result['spam_keywords'] = list(set(spam_in_subject + spam_in_body))

        # 检查可疑URL
        result['suspicious_urls'] = self.check_suspicious_urls(body)

        # 生成警告
        if result['spam_keywords']:
            result['has_issues'] = True
            result['warnings'].append(
                f"发现 {len(result['spam_keywords'])} 个垃圾邮件关键词: {', '.join(result['spam_keywords'][:5])}"
            )

        if result['suspicious_urls']:
            result['has_issues'] = True
            result['warnings'].append(
                f"发现 {len(result['suspicious_urls'])} 个可疑URL: {', '.join(result['suspicious_urls'][:3])}"
            )

        # 检查邮件长度
        if len(body) < 50:
            result['warnings'].append("邮件正文过短,可能被视为垃圾邮件")
            result['has_issues'] = True

        # 检查全大写
        if subject.isupper() and len(subject) > 10:
            result['warnings'].append("主题全部大写,可能被视为垃圾邮件")
            result['has_issues'] = True

        return result


class ReputationChecker:
    """发送者信誉检查器"""

    # 常见的DNSBL黑名单服务器
    DNSBL_SERVERS = [
        'zen.spamhaus.org',
        'bl.spamcop.net',
        'b.barracudacentral.org',
        'dnsbl.sorbs.net'
    ]

    def __init__(self):
        if not DNS_AVAILABLE:
            logging.warning("信誉检查功能不可用: dnspython未安装")
        self.resolver = dns.resolver.Resolver() if DNS_AVAILABLE else None
        if self.resolver:
            self.resolver.timeout = 3
            self.resolver.lifetime = 3

    def check_ip_blacklist(self, ip_address: str) -> Dict[str, any]:
        """
        检查IP是否在黑名单中

        Args:
            ip_address: 要检查的IP地址

        Returns:
            检查结果字典
        """
        if not DNS_AVAILABLE:
            return {
                'is_blacklisted': False,
                'blacklists': [],
                'error': 'DNS功能不可用'
            }

        result = {
            'is_blacklisted': False,
            'blacklists': [],
            'clean_lists': []
        }

        # 反转IP地址用于DNSBL查询
        reversed_ip = '.'.join(reversed(ip_address.split('.')))

        for dnsbl in self.DNSBL_SERVERS:
            query = f"{reversed_ip}.{dnsbl}"
            try:
                self.resolver.resolve(query, 'A')
                # 如果能解析成功,说明在黑名单中
                result['is_blacklisted'] = True
                result['blacklists'].append(dnsbl)
                logging.warning(f"IP {ip_address} 在黑名单中: {dnsbl}")
            except dns.resolver.NXDOMAIN:
                # NXDOMAIN表示不在黑名单中(正常)
                result['clean_lists'].append(dnsbl)
            except Exception as e:
                logging.debug(f"黑名单查询失败 {dnsbl}: {e}")

        return result

    def get_public_ip(self) -> Optional[str]:
        """
        获取本机的公网IP地址

        Returns:
            公网IP地址或None
        """
        try:
            # 方法1: 通过DNS查询
            import urllib.request
            ip = urllib.request.urlopen('https://api.ipify.org').read().decode('utf8')
            return ip
        except:
            pass

        try:
            # 方法2: 通过socket连接(不实际发送数据)
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
            s.close()
            # 注意: 这可能返回内网IP
            if not ip.startswith('192.168.') and not ip.startswith('10.'):
                return ip
        except:
            pass

        return None


def run_pre_send_checks(sender_email: str, subject: str, body: str, verbose: bool = True) -> Dict[str, any]:
    """
    运行发送前的全面安全检查

    Args:
        sender_email: 发件人邮箱
        subject: 邮件主题
        body: 邮件正文
        verbose: 是否显示详细信息

    Returns:
        检查结果字典
    """
    results = {
        'passed': True,
        'warnings': [],
        'errors': []
    }

    # 提取域名
    domain = sender_email.split('@')[-1] if '@' in sender_email else None

    if verbose:
        print("\n=== 邮件安全检查 ===\n")

    # 1. DNS记录检查
    if domain and DNS_AVAILABLE:
        validator = DNSValidator()

        # SPF检查
        spf_exists, spf_record = validator.check_spf(domain)
        if spf_exists:
            if verbose:
                print(f"✓ SPF记录存在: {spf_record[:60]}...")
        else:
            results['warnings'].append(f"SPF记录不存在或查询失败: {spf_record}")
            if verbose:
                print(f"⚠ {results['warnings'][-1]}")

        # DMARC检查
        dmarc_exists, dmarc_record = validator.check_dmarc(domain)
        if dmarc_exists:
            if verbose:
                print(f"✓ DMARC记录存在: {dmarc_record[:60]}...")
        else:
            results['warnings'].append(f"DMARC记录不存在或查询失败: {dmarc_record}")
            if verbose:
                print(f"⚠ {results['warnings'][-1]}")

    # 2. 内容检查
    checker = ContentChecker()
    content_result = checker.check_content(subject, body)

    if content_result['has_issues']:
        results['warnings'].extend(content_result['warnings'])
        if verbose:
            for warning in content_result['warnings']:
                print(f"⚠ {warning}")
    else:
        if verbose:
            print("✓ 内容检查通过,未发现明显垃圾邮件特征")

    # 3. IP信誉检查
    rep_checker = ReputationChecker()
    public_ip = rep_checker.get_public_ip()

    if public_ip and DNS_AVAILABLE:
        if verbose:
            print(f"\n检查IP信誉: {public_ip}")

        blacklist_result = rep_checker.check_ip_blacklist(public_ip)

        if blacklist_result['is_blacklisted']:
            error_msg = f"警告: IP {public_ip} 在以下黑名单中: {', '.join(blacklist_result['blacklists'])}"
            results['errors'].append(error_msg)
            results['passed'] = False
            if verbose:
                print(f"✗ {error_msg}")
        else:
            if verbose:
                print(f"✓ IP未在黑名单中 (检查了 {len(blacklist_result['clean_lists'])} 个黑名单)")

    # 汇总结果
    if verbose:
        print(f"\n{'='*50}")
        if results['passed'] and not results['warnings']:
            print("✓ 所有检查通过!")
        elif results['passed']:
            print(f"⚠ 检查通过,但有 {len(results['warnings'])} 个警告")
        else:
            print(f"✗ 检查未通过,有 {len(results['errors'])} 个错误")
        print(f"{'='*50}\n")

    return results


if __name__ == '__main__':
    # 测试代码
    print("邮件安全模块测试\n")

    # 测试DNS验证
    if DNS_AVAILABLE:
        validator = DNSValidator()
        print("测试SPF记录查询:")
        spf_ok, spf_record = validator.check_spf("gmail.com")
        print(f"  gmail.com SPF: {spf_ok} - {spf_record}\n")

        print("测试DMARC记录查询:")
        dmarc_ok, dmarc_record = validator.check_dmarc("gmail.com")
        print(f"  gmail.com DMARC: {dmarc_ok} - {dmarc_record}\n")

    # 测试内容检查
    checker = ContentChecker()
    result = checker.check_content(
        "恭喜中奖!点击领取",
        "您已中奖100万,请立即点击 http://bit.ly/xxx 领取奖金"
    )
    print(f"内容检查结果: {result}\n")

    # 测试IP信誉
    if DNS_AVAILABLE:
        rep_checker = ReputationChecker()
        public_ip = rep_checker.get_public_ip()
        if public_ip:
            print(f"本机公网IP: {public_ip}")
            bl_result = rep_checker.check_ip_blacklist(public_ip)
            print(f"黑名单检查: {bl_result}")
