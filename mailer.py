"""
Gmail SMTP를 사용한 첨부파일 이메일 발송.

필요한 secrets:
  - GMAIL_SENDER: 발신자 Gmail 주소
  - GMAIL_APP_PASSWORD: Google 계정 앱 비밀번호 (16자리)
    발급: https://myaccount.google.com/apppasswords
    (2단계 인증이 활성화된 계정에서만 발급 가능)
"""

import os
import smtplib
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path


SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587


def send_with_attachment(
    recipients: list[str],
    subject: str,
    body: str,
    attachment_path: str | Path,
    sender_name: str = "금융시장브리프 자동발송",
    sender: str | None = None,
    password: str | None = None,
) -> None:
    """지정된 수신자에게 첨부파일 포함 메일 발송."""
    sender = sender or os.getenv("GMAIL_SENDER")
    password = password or os.getenv("GMAIL_APP_PASSWORD")
    if not sender or not password:
        raise RuntimeError(
            "Gmail 설정이 없습니다. Streamlit Secrets에 "
            "GMAIL_SENDER와 GMAIL_APP_PASSWORD를 추가하세요."
        )

    attachment_path = Path(attachment_path)
    if not attachment_path.exists():
        raise FileNotFoundError(f"첨부 파일을 찾을 수 없습니다: {attachment_path}")

    msg = EmailMessage()
    msg["From"] = formataddr((sender_name, sender))
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.set_content(body)

    with open(attachment_path, "rb") as f:
        data = f.read()
    msg.add_attachment(
        data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=attachment_path.name,
    )

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.login(sender, password)
        smtp.send_message(msg)
