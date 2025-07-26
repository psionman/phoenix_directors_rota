"""Send and or save emails."""

import smtplib
from email.mime.text import MIMEText
from smtplib import SMTPAuthenticationError

from psiutils.errors import ErrorMsg

from process import Director
from config import read_config, env


def send_emails(text: str, directors: dict[Director]) -> int | ErrorMsg:
    """Send emails to the directors."""
    emails_sent = 0
    for key, director in directors.items():
        if key and director. active:
            print(director.email)
            response = _create_email(text, director)
            if isinstance(response, ErrorMsg):
                return response
            emails_sent += 1
    return emails_sent


def _create_email(text: str, director: Director) -> str:
    config = read_config()
    try:
        _send_email(
            config.email_subject,
            text,
            director.email)
    except SMTPAuthenticationError:
        return ErrorMsg(
            header='Email error',
            message='Email authentication error.',
        )
    except TypeError:
        return ErrorMsg(
            header='Email error',
            message='Email setup error.',
        )
    return True


def _send_email(subject, body, recipient):
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = env['email_sender']
    msg['To'] = recipient
    # recipient = sender
    with smtplib.SMTP_SSL(env['smtp_server'], env['smtp_port']) as smtp_server:
        smtp_server.login(env['email_sender'], env['email_key'])
        smtp_server.sendmail(env['email_sender'], recipient, msg.as_string())
