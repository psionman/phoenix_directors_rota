"""Config for Phoenix director's rota."""
import os
from psiconfig import TomlConfig
from pathlib import Path
from appdirs import user_data_dir
from dotenv import load_dotenv

from psiutils.known_paths import get_downloads_dir

from directors_rota.constants import CONFIG_PATH, APP_NAME, APP_AUTHOR, EMAIL_TEMPLATE

import directors_rota.text as text

load_dotenv()

email_key = os.getenv('EMAIL_KEY')
email_sender = os.getenv('EMAIL_SENDER')
smtp_server = os.getenv('SMTP_SERVER')
try:
    SMTP_PORT = int(os.getenv('SMTP_PORT'))
except TypeError:
    SMTP_PORT = 0

DEFAULT_CONFIG = {
    'workbook_dir': Path(get_downloads_dir()),
    'workbook_file_name': 'directors-rota.xlsx',
    'email_template': str(
        Path(user_data_dir(APP_NAME, APP_AUTHOR), EMAIL_TEMPLATE)),
    'main_sheet': 'Main',
    'directors_sheet': 'Directors',
    'initials_col': 0,
    'name_col': 1,
    'email_col': 2,
    'username_col': 3,
    'active_col': 4,
    'mon_date_col': 0,
    'wed_date_col': 3,
    'senders_email_address': email_sender,
    'smtp_server': smtp_server,
    'SMTP_PORT': SMTP_PORT,
    'email_password': email_key,
    'email_subject': f'Phoenix Bridge Club - BBO {text.DIRECTORS} rota',
    'send_emails': True,
    'geometry': {},
    'new_geometry': {},
}


def read_config() -> TomlConfig:
    """Return the config file."""
    toml_config = TomlConfig(path=CONFIG_PATH, defaults=DEFAULT_CONFIG)
    toml_config.check_defaults(toml_config.config)
    return toml_config


def save_config(changed_config: TomlConfig) -> TomlConfig | None:
    """Save the config file."""
    result = changed_config.save()
    if result != changed_config.STATUS_OK:
        return None
    toml_config = TomlConfig(CONFIG_PATH)
    return toml_config


def _get_env() -> dict:
    load_dotenv()
    try:
        SMTP_PORT = int(os.getenv('SMTP_PORT'))
    except TypeError:
        SMTP_PORT = 0

    return {
        'email_key': os.getenv('EMAIL_KEY'),
        'email_sender': os.getenv('EMAIL_SENDER'),
        'smtp_server': os.getenv('SMTP_SERVER'),
        'smtp_port': SMTP_PORT
    }


config = read_config()
env = _get_env()
