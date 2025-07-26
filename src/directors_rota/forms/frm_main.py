
import os
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import datetime
from dateutil.relativedelta import relativedelta
from dateutil.parser import parse as date_parse

from psiutils.widgets import clickable_widget
from psiutils.buttons import ButtonFrame, IconButton
from psiutils.constants import PAD, LARGE_FONT
from psiutils.utilities import window_resize, geometry
import psiutils.text as psiText

from constants import MMYYYY, DOWNLOADS_DIR, XLS_FILE_TYPES
from config import config
from process import generate_rota, status as process_status
import text

from config import read_config

from forms.frm_email import EmailFrame
from main_menu import MainMenu

FRAME_TITLE = f'{text.DIRECTORS} Rota'


class MainFrame():
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.directors = []
        self.email = ''
        self.config = read_config()

        # Tk Vars
        workbook_path = Path(config.workbook_dir, config.workbook_file_name)
        self.workbook_path = tk.StringVar(value=workbook_path)
        self.email_template = tk.StringVar(value=config.email_template)
        self.rota_month = tk.StringVar(value='')

        self.show()

        self.period_starts = self._get_period_starts()
        self.selected_month = datetime.datetime(1, 1, 1)
        self._set_file_message()

    def show(self) -> None:
        root = self.root
        root.geometry(geometry(self.config, __file__))
        root.title(FRAME_TITLE)

        root.rowconfigure(0, weight=1)
        root.columnconfigure(0, weight=1)
        root.bind('<Control-q>', self.dismiss)
        root.bind('<Control-g>', self._generate_rota)
        root.bind('<Configure>',
                  lambda event, arg=None: window_resize(self, __file__))

        main_menu = MainMenu(self, root)
        main_menu.create()

        frame = ttk.Frame(root)
        frame.grid(row=0, column=0, sticky=tk.NSEW)
        frame.rowconfigure(1, weight=1)

        frame.columnconfigure(0, weight=1)

        header = ttk.Label(frame, text=FRAME_TITLE, font=LARGE_FONT)
        header.grid(row=0, column=0, columnspan=99, padx=PAD, pady=PAD)

        main_frame = self._main_frame(frame)
        main_frame.grid(row=1, column=0, sticky=tk.NSEW, padx=PAD, pady=PAD)

        self.button_frame = self._button_frame(frame)
        self.button_frame.grid(row=3, column=0, sticky=tk.EW, padx=PAD)

        sizegrip = ttk.Sizegrip(root)
        sizegrip.grid(sticky=tk.SE)

    def _main_frame(self, master: tk.Frame) -> tk.Frame:
        frame = ttk.Frame(master)

        frame.columnconfigure(1, weight=1)

        # Widgets for month file.
        row = 0
        label = ttk.Label(frame, text='Rota for month')
        label.grid(row=row, column=1, sticky=tk.SW, padx=PAD)

        row += 1
        button = IconButton(
            frame, psiText.PREVIOUS, 'previous', self._previous_month)
        button.grid(row=row, column=0, sticky=tk.E)

        month_entry = ttk.Entry(frame, textvariable=self.rota_month)
        month_entry.grid(row=row, column=1, sticky=tk.EW, padx=PAD, pady=PAD)

        button = IconButton(
            frame, psiText.NEXT, 'next', self._next_month)
        button.grid(row=row, column=2, sticky=tk.W, padx=PAD)

        # Workbook
        row += 1
        label = ttk.Label(frame, text="Director's rota workbook")
        label.grid(row=row, column=0, sticky=tk.E)

        workbook_file_name = ttk.Entry(frame,
                                       textvariable=self.workbook_path)
        workbook_file_name.grid(row=row, column=1, sticky=tk.EW,
                                padx=PAD, pady=PAD)
        self.workbook_path.trace_add("write", self._on_workbook_path_change)

        button = IconButton(
            frame, psiText.OPEN, 'open', self._get_workbook_path)
        button.grid(row=row, column=2, padx=PAD)

        return frame

    def _button_frame(self, master: tk.Frame) -> tk.Frame:
        frame = ButtonFrame(master, tk.HORIZONTAL)
        buttons = [
            frame.icon_button('build', True, self._generate_rota),
            frame.icon_button('close', False, self.dismiss)
        ]
        frame.buttons = buttons
        return frame

    def _get_period_starts(self) -> None:
        """Return a list of months."""
        now = datetime.datetime.now()
        start_date = datetime.datetime(now.year, now.month, 1)
        months = [start_date +
                  relativedelta(months=diff) for diff in range(-1, 4)]
        self.rota_month.set((start_date +
                             relativedelta(months=1)).strftime(MMYYYY))
        return months

    def _previous_month(self) -> None:
        """Get previous month."""
        selected_month = date_parse(f'1 {self.rota_month.get()}').date()
        self.selected_month = selected_month - relativedelta(months=1)
        self.rota_month.set((self.selected_month).strftime(MMYYYY))

    def _next_month(self) -> None:
        """Get next month."""
        selected_month = date_parse(f'1 {self.rota_month.get()}').date()
        self.selected_month = selected_month + relativedelta(months=1)
        self.rota_month.set((self.selected_month).strftime(MMYYYY))

    def _generate_rota(self, *args) -> None:
        """Show process screen."""
        if not Path(self.workbook_path.get()).is_file():
            messagebox.showerror(
                '', f'{self.workbook_path.get()} does not exist',
            )
            return

        selected_month = date_parse(f'1 {self.rota_month.get()}').date()
        (status, self.email, self.directors) = generate_rota(selected_month)

        if status == process_status['OK']:
            dlg = EmailFrame(self)
            self.root.wait_window(dlg.root)

        elif status == process_status['FILE_MISSING']:
            message = f'Workbook "{config.workbook_file_name}" does not exist'
            messagebox.showerror(title=FRAME_TITLE, message=message)
        elif status == process_status['SHEET_MISSING']:
            message = f'Sheet "{0}?????" does not exist'
            messagebox.showerror(title=FRAME_TITLE, message=message)
        self.root.destroy()

    def _get_workbook_path(self) -> None:
        """Set the workbook path"""
        initialdir = str(Path(self.workbook_path.get()).parent)
        if initialdir == '.':
            initialdir = DOWNLOADS_DIR
        workbook_file_name = filedialog.askopenfilename(
            title='Workbook',
            initialdir=initialdir,
            initialfile=str(Path(self.workbook_path.get()).name),
            filetypes=XLS_FILE_TYPES
        )

        if workbook_file_name:
            self.workbook_path.set(workbook_file_name)
            config.workbook_file_name = workbook_file_name

    def _on_workbook_path_change(self, *args) -> None:
        self._set_file_message()

    def _set_file_message(self) -> None:
        message = ''
        config_text = 'Click on Menu > Defaults to define.'
        email_template = os.path.isfile(config.email_template)
        directors_rota = os.path.isfile(self.workbook_path.get())
        if not email_template and not directors_rota:
            message = (f'{text.DIRECTORS} rota and email template not valid. '
                       f'{config_text}')
        elif not email_template and not directors_rota:
            message = f'Email template not valid. {config_text}'
        elif not email_template and not directors_rota:
            message = f'{text.DIRECTORS} rota not valid.'
        if message:
            self.button_frame.enable(False)

    def dismiss(self, *args) -> None:
        self.root.destroy()
