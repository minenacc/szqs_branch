import smtplib
from email.message import EmailMessage

SMTP_SERVER = 'smtp.qq.com'
SMTP_PORT = 465
EMAIL_SENDER = '605055322@qq.com'
EMAIL_PASSWORD = 'vxfnpttsxwtybaib'  # 在邮箱中申请的专用密码

def send_license_email(to_email, license_key, expire_date):
    subject = '您的授权密钥已生成'
    content = f"""亲爱的用户：

感谢您注册本软件，以下是您的授权信息：

- 授权邮箱：{to_email}
- 授权密钥：{license_key}
- 到期时间：{expire_date}

请妥善保管此密钥，用于软件激活验证。

"""

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = EMAIL_SENDER
    msg['To'] = to_email
    msg.set_content(content)

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
            smtp.send_message(msg)
        print(f"[✓] 邮件成功发送至 {to_email}")
        return True
    except smtplib.SMTPResponseException as e:
        # 容忍 (-1, b'\x00\x00\x00') 错误
        if e.smtp_code == -1 and e.smtp_error == b'\x00\x00\x00':
            print(f"[✓] SMTP返回非标准响应但邮件已发送至 {to_email}")
            return True
        print(f"[×] SMTP 错误: {str(e)}")
        return False
    except Exception as e:
        print(f"[×] 发送邮件失败: {str(e)}")
        return False