"""Constants for the Director's Rota app."""
from pathlib import Path
from appdirs import user_config_dir

from psiutils.known_paths import get_downloads_dir

import directors_rota.text as text

# Config
APP_NAME = 'phoenix_bbo_dir_rota'
APP_AUTHOR = 'phoenix'
EMAIL_TEMPLATE = Path('rota_email_template.txt')
CONFIG_PATH = Path(user_config_dir(APP_NAME, APP_AUTHOR), 'config.toml')

# App
APP_TITLE = f'{text.DIRECTORS} Rota'
ICON_FILE = Path( Path(__file__).parent, 'images', 'phoenix.png')
AUTHOR = 'Jeff Watkins'

# GUI
LARGE_FONT = ('Arial', 16)

# Files
DOWNLOADS_DIR = get_downloads_dir()
EMAIL_FILE_PREFIX = 'emails'

TXT_FILE_TYPES = (
    ('txt files', '*.txt'),
    ('All files', '*.*')
)

XLS_FILE_TYPES = (
    ('xlsx files', '*.xlsx'),
    ('All files', '*.*')
)

# Strings
MMYYYY = '%b %Y'

COL_MAXIMUM = 26

# Colours
ERROR_COLOUR = 'red'
MSG_COLOUR = 'green'
