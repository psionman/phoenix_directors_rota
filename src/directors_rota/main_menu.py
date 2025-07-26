import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import shutil

from psiutils.utilities import create_directories
from psiutils.menus import Menu, MenuItem

from constants import (
    APP_TITLE, DOWNLOADS_DIR, EMAIL_TEMPLATE, TXT_FILE_TYPES, AUTHOR)
from _version import __version__
import text

from config import config

from forms.frm_config import ConfigFrame

SPACES = ' '*20


class MainMenu():
    def __init__(self, parent, root):
        self.parent = parent
        self.root = root

    def create(self):
        menubar = tk.Menu()
        self.root['menu'] = menubar

        # File menu
        file_menu = Menu(menubar, self._file_menu_items())
        menubar.add_cascade(menu=file_menu, label='File')

        # Help menu
        help_menu = Menu(menubar, self._help_menu_items())
        menubar.add_cascade(menu=help_menu, label='Help')

    def _file_menu_items(self) -> list:
        return [
            MenuItem(f'Defaults{text.ELLIPSIS}', self.show_defaults),
            MenuItem(f'Copy template{text.ELLIPSIS}',
                     self.copy_template),
            MenuItem('Quit', self.dismiss),
        ]

    def _help_menu_items(self) -> list:
        return [
            # MenuItem(f'On line help{text.ELLIPSIS}', self._show_help),
            MenuItem(f'Data directory location{text.ELLIPSIS}',
                     self._show_data_directory),
            MenuItem(f'About{text.ELLIPSIS}', self.show_about),
        ]

    def copy_template(self) -> None:
        template_path = filedialog.askopenfilename(
            title='Copy email template',
            initialdir=DOWNLOADS_DIR,
            initialfile=str(EMAIL_TEMPLATE),
            filetypes=TXT_FILE_TYPES
        )
        if template_path:
            target_path = config.email_template
            create_directories(Path(target_path).parent)
            shutil.copyfile(
                template_path,
                target_path
                )
            messagebox.showinfo(
                'Copy email template',
                f'Email template copied to {target_path}',
                )

    def show_defaults(self):
        dlg = ConfigFrame(self)
        self.root.wait_window(dlg.root)

    def _show_data_directory(self) -> None:
        dir = f'Data directory: {Path(config.email_template).parent} {SPACES}'
        messagebox.showinfo(title='Data directory', message=dir)

    def show_about(self):
        about = f'Version: {__version__}\nAuthor: {AUTHOR}'
        messagebox.showinfo(title=f'{APP_TITLE} - About', message=about)

    def dismiss(self):
        self.root.destroy()
