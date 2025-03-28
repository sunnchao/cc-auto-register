from logger import info, error
import time
import re
import requests
import imaplib
import poplib
import email
from email.parser import Parser
from datetime import datetime

from config import (
    EMAIL_USERNAME,
    EMAIL_DOMAIN,
    EMAIL_PIN,
    EMAIL_VERIFICATION_RETRIES,
    EMAIL_VERIFICATION_WAIT,
    EMAIL_TYPE,
    EMAIL_PROXY_ADDRESS,
    EMAIL_PROXY_ENABLED,
    EMAIL_API,
    EMAIL_IMAP_SERVER,
    EMAIL_IMAP_PORT,
    EMAIL_IMAP_DIR,
    EMAIL_IMAP_PROTOCOL,
)


class EmailVerificationHandler:
    def __init__(self, username=None, domain=None, pin=None):
        self.email = EMAIL_TYPE
        self.username = username or EMAIL_USERNAME
        self.domain = domain or EMAIL_DOMAIN
        self.session = requests.Session()
        self.emailApi = EMAIL_API
        self.emailExtension = f"@{self.domain}"
        self.pin = pin or EMAIL_PIN
        if self.email == "tempemail":
            info(
                f"初始化邮箱验证器成功: {self.username}{self.emailExtension} pin: {self.pin}"
            )
        elif self.email == "zmail":
            info(
                f"初始化邮箱验证器成功: {self.emailApi}"
            )

    def get_verification_code(
            self, source_email=None, max_retries=None, wait_time=None
    ):
        """
        获取验证码，增加了重试机制

        Args:
            max_retries: 最大重试次数
            wait_time: 每次重试间隔时间(秒)

        Returns:
            str: 验证码或None
        """
        max_retries = max_retries or EMAIL_VERIFICATION_RETRIES
        wait_time = wait_time or EMAIL_VERIFICATION_WAIT
        info(f"开始获取邮箱验证码=>最大重试次数:{max_retries}, 等待时间:{wait_time}")
        for attempt in range(max_retries):
            try:
                info(f"当前EMail类型为： {EMAIL_TYPE}")
                if self.email == "tempemail":
                    code, mail_id = self.get_tempmail_email_code(source_email)
                elif self.email == "zmail":
                    code, mail_id = self.get_zmail_email_code(source_email)
                elif self.email == "imap":
                    code, mail_id = self.get_imap_email_code(source_email)
                if code:
                    info(f"成功获取验证码: {code}")
                    return code
                if attempt < max_retries - 1:
                    info(
                        f"未找到验证码，{wait_time}秒后重试 ({attempt + 1}/{max_retries})..."
                    )
                    time.sleep(wait_time)
            except Exception as e:
                error(f"获取验证码失败: {str(e)}")
                if attempt < max_retries - 1:
                    info(f"将在{wait_time}秒后重试...")
                    time.sleep(wait_time)

        return None

    # 手动输入验证码
    def get_tempmail_email_code(self, source_email=None):
        info("开始获取邮件列表")
        # 获取邮件列表
        mail_list_url = f"https://tempmail.plus/api/mails?email={self.username}{self.emailExtension}&limit=20&epin={self.pin}"
        try:
            mail_list_response = self.session.get(
                mail_list_url, timeout=10
            )  # 添加超时参数
            mail_list_data = mail_list_response.json()
            time.sleep(0.5)
            if not mail_list_data.get("result"):
                return None, None
        except requests.exceptions.Timeout:
            error("获取邮件列表超时")
            return None, None
        except requests.exceptions.ConnectionError:
            error("获取邮件列表连接错误")
            return None, None
        except Exception as e:
            error(f"获取邮件列表发生错误: {str(e)}")
            return None, None

        # 获取最新邮件的ID
        first_id = mail_list_data.get("first_id")
        if not first_id:
            return None, None
        info(f"开始获取邮件详情: {first_id}")
        # 获取具体邮件内容
        mail_detail_url = f"https://tempmail.plus/api/mails/{first_id}?email={self.username}{self.emailExtension}&epin={self.pin}"
        try:
            mail_detail_response = self.session.get(
                mail_detail_url, timeout=10
            )  # 添加超时参数
            mail_detail_data = mail_detail_response.json()
            time.sleep(0.5)
            if not mail_detail_data.get("result"):
                return None, None
        except requests.exceptions.Timeout:
            error("获取邮件详情超时")
            return None, None
        except requests.exceptions.ConnectionError:
            error("获取邮件详情连接错误")
            return None, None
        except Exception as e:
            error(f"获取邮件详情发生错误: {str(e)}")
            return None, None

        # 从邮件文本中提取6位数字验证码
        mail_text = mail_detail_data.get("text", "")

        # 如果提供了source_email，确保邮件内容中包含该邮箱地址
        if source_email and source_email.lower() not in mail_text.lower():
            error(f"邮件内容不包含指定的邮箱地址: {source_email}")
        else:
            info(f"邮件内容包含指定的邮箱地址: {source_email}")

        code_match = re.search(r"(?<![a-zA-Z@.])\b\d{6}\b", mail_text)

        if code_match:
            # 清理邮件
            self._cleanup_mail(first_id)
            return code_match.group(), first_id
        return None, None

    def _cleanup_mail(self, first_id):
        # 构造删除请求的URL和数据
        delete_url = "https://tempmail.plus/api/mails/"
        payload = {
            "email": f"{self.username}{self.emailExtension}",
            "first_id": first_id,
            "epin": self.pin,
        }

        # 最多尝试3次
        for _ in range(3):
            response = self.session.delete(delete_url, data=payload)
            try:
                result = response.json().get("result")
                if result is True:
                    return True
            except:
                pass

            # 如果失败,等待0.2秒后重试
            time.sleep(0.2)

        return False

    # 如果是zmail 需要先创建邮箱
    def create_zmail_email(account_info):
        # 如果邮箱类型是zmail 需要先创建邮箱
        session = requests.Session()
        if EMAIL_PROXY_ENABLED:
            proxy = {
                "http": f"{EMAIL_PROXY_ADDRESS}",
                "https": f"{EMAIL_PROXY_ADDRESS}",
            }
            session.proxies.update(proxy)
        # 创建临时邮箱URL
        create_url = f"{EMAIL_API}/api/mailboxes"
        username = account_info["email"].split("@")[0]
        # 生成临时邮箱地址
        payload = {
            "address": f"{username}",
            "expiresInHours": 24,
        }
        # 发送POST请求创建临时邮箱
        try:
            create_response = session.post(
                create_url, json=payload, timeout=100
            )  # 添加超时参数
            info(f"创建临时邮箱成功: {create_response.status_code}")
            create_data = create_response.json()
            info(f"创建临时邮箱返回数据: {create_data}")
            # 检查创建邮箱是否成功
            time.sleep(0.5)
            if create_data.get("success") is True or create_data.get('error') == '邮箱地址已存在':
                info(f"邮箱创建成功: {create_data}")
            else:
                error(f"邮箱创建失败: {create_data}")
                return None, None
        except requests.exceptions.Timeout:
            error("创建临时邮箱超时", create_url)
            return None, None
        info(f"创建临时邮箱成功: {create_data}, 返回值: {create_data}")

    # 获取zmail邮箱验证码
    def get_zmail_email_code(self, source_email=None):
        info("开始获取邮件列表")
        # 获取邮件列表
        username = source_email.split("@")[0]
        mail_list_url = f"{EMAIL_API}/api/mailboxes/{username}/emails"
        proxy = {
            "http": f"{EMAIL_PROXY_ADDRESS}",
            "https": f"{EMAIL_PROXY_ADDRESS}",
        }
        self.session.proxies.update(proxy)
        try:
            mail_list_response = self.session.get(
                mail_list_url, timeout=10000
            )  # 添加超时参数
            mail_list_data = mail_list_response.json()
            time.sleep(2)
            if not mail_list_data.get("emails"):
                return None, None
        except requests.exceptions.Timeout:
            error("获取邮件列表超时")
            return None, None
        except requests.exceptions.ConnectionError:
            error("获取邮件列表连接错误")
            return None, None
        except Exception as e:
            error(f"获取邮件列表发生错误: {str(e)}")
            return None, None

        # 获取最新邮件的ID、
        mail_detail_data_len = len(mail_list_data["emails"])
        if mail_detail_data_len == 0:
            return None, None
        mail_list_data = mail_list_data["emails"][0]
        # 获取最新邮件的ID
        mail_id = mail_list_data.get("id")
        if not mail_id:
            return None, None
        # 获取具体邮件内容
        mail_detail_url = f"{EMAIL_API}/api/emails/{mail_id}"
        returnData = ''
        try:
            mail_detail_response = self.session.get(
                mail_detail_url, timeout=10
            )  # 添加超时参数
            returnData = mail_detail_response.json()
            time.sleep(2)
        except requests.exceptions.Timeout:
            error("获取邮件详情超时")
            return None, None
        except requests.exceptions.ConnectionError:
            error("获取邮件详情连接错误")
            return None, None
        except Exception as e:
            error(f"获取邮件详情发生错误: {str(e)}")
            return None, None

        # 从邮件文本中提取6位数字验证码\
        mail_text = returnData.get("email")
        mail_text = mail_text.get("textContent")
        # 如果提供了source_email，确保邮件内容中包含该邮箱地址
        if source_email and source_email.lower() not in mail_text.lower():
            error(f"邮件内容不包含指定的邮箱地址: {source_email}")
        else:
            info(f"邮件内容包含指定的邮箱地址: {source_email}")

        code_match = re.search(r"(?<![a-zA-Z@.])\b\d{6}\b", mail_text)
        info(f"验证码匹配结果: {code_match}")
        # 如果找到验证码, 返回验证码和邮件ID
        if code_match:
            return code_match.group(), mail_id
        else:
            error("未找到验证码")
            return None, None

    # 获取 IMAP 的验证码
    def get_imap_email_code(self, source_email = None):
        # IMAP 配置
        imap_server = EMAIL_IMAP_SERVER
        imap_port = EMAIL_IMAP_PORT
        imap_dir = EMAIL_IMAP_DIR
        imap_protocol = EMAIL_IMAP_PROTOCOL

        imap_user = ""
        imap_pass = ""

        if imap_protocol == "IMAP" :
            verify_code, imap_id = self._get_mail_code_by_imap(imap_server, imap_port, imap_user, imap_pass, imap_dir)
            return verify_code, imap_id
        else:
            return None, None



    # 使用imap获取邮件
    def _get_mail_code_by_imap(self, imap_server, imap_port, imap_user, imap_pass, imap_dir, retry = 0):
        if retry > 0:
            time.sleep(3)
        if retry >= 20:
            raise Exception("获取验证码超时")
        try:
            # 连接到IMAP服务器
            mail = imaplib.IMAP4_SSL(imap_server, imap_port)
            mail.login(imap_user, imap_pass)
            search_by_date=False
            # 针对网易系邮箱，imap登录后需要附带联系信息，且后续邮件搜索逻辑更改为获取当天的未读邮件
            if imap_user.endswith(('@163.com', '@126.com', '@yeah.net')):
                imap_id = ("name", imap_user.split('@')[0], "contact", imap_user, "version", "1.0.0", "vendor", "imaplib")
                mail.xatom('ID', '("' + '" "'.join(imap_id) + '")')
                search_by_date=True
            mail.select(imap_dir)
            if search_by_date:
                date = datetime.now().strftime("%d-%b-%Y")
                status, messages = mail.search(None, f'ON {date} UNSEEN')
            else:
                status, messages = mail.search(None, 'ALL')
            if status != 'OK':
                return None

            mail_ids = messages[0].split()
            if not mail_ids:
                # 没有获取到，就在获取一次
                return self._get_mail_code_by_imap(retry=retry + 1)

            for mail_id in reversed(mail_ids):
                status, msg_data = mail.fetch(mail_id, '(RFC822)')
                if status != 'OK':
                    continue
                raw_email = msg_data[0][1]
                email_message = email.message_from_bytes(raw_email)

                # 如果是按日期搜索的邮件，需要进一步核对收件人地址是否对应
                if search_by_date and email_message['to'] !=self.account:
                    continue
                body = self._extract_imap_body(email_message)
                if body:
                    code_match = re.search(r"\b\d{6}\b", body)
                    if code_match:
                        code = code_match.group()
                        # 删除找到验证码的邮件
                        mail.store(mail_id, '+FLAGS', '\\Deleted')
                        mail.expunge()
                        mail.logout()
                        return code, imap_id
            # print("未找到验证码")
            mail.logout()
            return None
        except Exception as e:
            print(f"发生错误: {e}")
            return None

    def _extract_imap_body(self, email_message):
        # 提取邮件正文
        if email_message.is_multipart():
            for part in email_message.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                if content_type == "text/plain" and "attachment" not in content_disposition:
                    charset = part.get_content_charset() or 'utf-8'
                    try:
                        body = part.get_payload(decode=True).decode(charset, errors='ignore')
                        return body
                    except Exception as e:
                        error(f"解码邮件正文失败: {e}")
        else:
            content_type = email_message.get_content_type()
            if content_type == "text/plain":
                charset = email_message.get_content_charset() or 'utf-8'
                try:
                    body = email_message.get_payload(decode=True).decode(charset, errors='ignore')
                    return body
                except Exception as e:
                    error(f"解码邮件正文失败: {e}")
        return ""
