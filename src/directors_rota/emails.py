"""Send and or save emails."""

import smtplib
from email.mime.text import MIMEText
from smtplib import SMTPAuthenticationError

from psiutils.constants import Status
from psiutils.errors import ErrorMsg

from directors_rota.process import Director
from directors_rota.config import read_config, env
from directors_rota import logger


def send_emails(text: str, directors: dict[Director]) -> int | ErrorMsg:
    """Send emails to the directors."""
    emails_sent = 0
    config = read_config()
    if not (env['email_sender']
            and env['email_key']
            and env['smtp_port']
            and env['smtp_server']):
        return Status.WARNING
    for key, director in directors.items():
        if key and director.active:
            response = _create_email(config.email_subject, text, director)
            if isinstance(response, ErrorMsg):
                return response
            emails_sent += 1
            if director.send_reminder:
                _create_reminder(config.email_reminder_dir, director)
    return Status.SUCCESS


def _create_email(subject: str, text: str, director: Director) -> str:
    # pylint: disable=no-member)
    try:
        _send_email(
            subject,
            text,
            director.email)
    except SMTPAuthenticationError:
        logger.error('Email authentication error.')
        return Status.ERROR
    except TypeError:
        logger.error('Email setup error.')
        return Status.ERROR
    return Status.SUCCESS


def _send_email(subject: str, body: str, recipient: str) -> None:
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = env['email_sender']
    msg['To'] = recipient
    # recipient = env['email_sender']
    with smtplib.SMTP_SSL(env['smtp_server'], env['smtp_port']) as smtp_server:
        smtp_server.login(env['email_sender'], env['email_key'])
        smtp_server.sendmail(env['email_sender'], recipient, msg.as_string())
    logger.info(f"Email sent to {recipient}")


def _create_reminder(path: str, director: Director) -> None:
    return {
        'date': '',
        'time': '06:00',
        'recipient': '',
        'subject': '',
        'body': '',
    }
