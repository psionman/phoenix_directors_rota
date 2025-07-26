import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from pathlib import Path

from psiutils.widgets import clickable_widget, separator_frame
from psiutils.buttons import ButtonFrame, IconButton
from psiutils.constants import PAD
from psiutils.utilities import window_resize, geometry
from psiutils import text as psiText

from constants import COL_MAXIMUM, APP_TITLE, TXT_FILE_TYPES
from config import read_config

import text

FRAME_TITLE = 'System defaults'


class ConfigFrame():
    def __init__(self, parent) -> None:
        self.root = tk.Toplevel(parent.root)
        self.parent = parent
        self.config = read_config()

        # Tk vars
        self.workbook_dir = tk.StringVar(value=self.config.workbook_dir)
        self.main_sheet = tk.StringVar(value=self.config.main_sheet)
        self.directors_sheet = tk.StringVar(value=self.config.directors_sheet)

        self.initials_col = tk.IntVar(value=self.config.initials_col)
        self.name_col = tk.IntVar(value=self.config.name_col)
        self.email_col = tk.IntVar(value=self.config.email_col)
        self.username_col = tk.IntVar(value=self.config.username_col)
        self.active_col = tk.IntVar(value=self.config.active_col)

        self.mon_date_col = tk.IntVar(value=self.config.mon_date_col)
        self.wed_date_col = tk.IntVar(value=self.config.wed_date_col)

        self.email_subject = tk.StringVar(value=self.config.email_subject)
        self.email_template = tk.StringVar(value=self.config.email_template)
        self.send_emails = tk.BooleanVar(value=self.config.send_emails)

        # Change check vars
        self.workbook_dir.trace_add('write', self._check_value_changed)
        self.main_sheet.trace_add('write', self._check_value_changed)
        self.directors_sheet.trace_add('write', self._check_value_changed)

        self.initials_col.trace_add('write', self._check_value_changed)
        self.name_col.trace_add('write', self._check_value_changed)
        self.email_col.trace_add('write', self._check_value_changed)
        self.username_col.trace_add('write', self._check_value_changed)
        self.active_col.trace_add('write', self._check_value_changed)

        self.mon_date_col.trace_add('write', self._check_value_changed)
        self.wed_date_col.trace_add('write', self._check_value_changed)

        self.email_subject.trace_add('write', self._check_value_changed)
        self.email_template.trace_add('write', self._check_value_changed)
        self.send_emails.trace_add('write', self._check_value_changed)

        self.show()

    def show(self) -> None:
        root = self.root
        root.geometry(geometry(self.config, __file__))
        root.title(FRAME_TITLE)
        root.transient(self.parent.root)

        root.bind('<Control-x>', self.dismiss)
        root.bind('<Control-s>', self._save_config)
        root.bind('<Configure>',
                  lambda event, arg=None: window_resize(self, __file__))

        root.rowconfigure(0, weight=1)
        root.columnconfigure(0, weight=1)

        main_frame = self._main_frame(root)
        main_frame.grid(row=0, column=0, sticky=tk.NSEW, padx=PAD, pady=PAD)

        sizegrip = ttk.Sizegrip(root)
        sizegrip.grid(sticky=tk.SE)

    def _main_frame(self, master: tk.Frame) -> tk.Frame:
        frame = ttk.Frame(master)
        frame.columnconfigure(2, weight=1)

        row = 0
        # header = ttk.Label(frame, text=FRAME_TITLE, font=LARGE_FONT)
        # header.grid(row=row, column=0, columnspan=4, pady=PAD)

        row += 1
        heading_frame = separator_frame(frame, 'Workbook')
        heading_frame.grid(row=row, column=0, columnspan=4, sticky=tk.EW)

        # Workbook directory
        row += 1
        label = ttk.Label(frame, text="Workbook directory")
        label.grid(row=row, column=0, sticky=tk.E)

        workbook_dir = ttk.Entry(frame, textvariable=self.workbook_dir)
        workbook_dir.grid(row=row, column=1, columnspan=2, sticky=tk.EW,
                          padx=PAD, pady=PAD)

        button = IconButton(
            frame, psiText.OPEN, 'open', self._get_workbook_dir)
        button.grid(row=row, column=3, sticky=tk.W, padx=PAD)

        row += 1
        # Email template
        label = ttk.Label(frame, text="Email template")
        label.grid(row=row, column=0, sticky=tk.E)

        email_template = ttk.Entry(frame,
                                   textvariable=self.email_template)
        email_template.grid(row=row, column=1, columnspan=2, sticky=tk.EW,
                            padx=PAD, pady=PAD)

        button = IconButton(
            frame, psiText.OPEN, 'open', self._get_email_template)
        button.grid(row=row, column=3, sticky=tk.W, padx=PAD)

        # Main sheet name
        row += 1
        label = ttk.Label(frame, text="Main sheet name")
        label.grid(row=row, column=0, sticky=tk.E)
        entry = ttk.Entry(frame, textvariable=self.main_sheet)
        entry.grid(row=row, column=1, sticky=tk.EW, padx=PAD)

        # Director's sheet name.
        row += 1
        label = ttk.Label(frame, text="Director sheet name")
        label.grid(row=row, column=0, sticky=tk.E, pady=PAD/2)
        entry = ttk.Entry(frame, textvariable=self.directors_sheet)
        entry.grid(row=row, column=1, sticky=tk.EW, padx=PAD)

        # Sheet column numbers.
        row += 1
        heading_frame = separator_frame(frame, 'Director tab')
        heading_frame.grid(row=row, column=0, columnspan=4, sticky=tk.EW)

        row += 1
        self._label_and_spinbox(frame, row, 'Initials column',
                                self.initials_col)

        row += 1
        self._label_and_spinbox(frame, row, 'Name column',
                                self.name_col)

        row += 1
        self._label_and_spinbox(frame, row, 'Email column',
                                self.email_col)

        row += 1
        self._label_and_spinbox(frame, row, 'Username column',
                                self.username_col)

        row += 1
        self._label_and_spinbox(frame, row, 'Active column',
                                self.active_col)

        row += 1
        heading_frame = separator_frame(frame, 'Rota tab')
        heading_frame.grid(row=row, column=0, columnspan=4, sticky=tk.EW)

        row += 1
        self._label_and_spinbox(frame, row, 'Monday date column',
                                self.mon_date_col)

        row += 1
        self._label_and_spinbox(frame, row, 'Wednesday date column',
                                self.wed_date_col)

        row += 1
        separator = separator_frame(frame, 'Email details')
        separator.grid(row=row, column=0, columnspan=3,
                       sticky=tk.EW, padx=PAD, pady=PAD)

        row += 1
        label = ttk.Label(frame, text='Email subject')
        label.grid(row=row, column=0, sticky=tk.E, padx=PAD, pady=PAD)

        entry = ttk.Entry(frame, textvariable=self.email_subject)
        entry.grid(row=row, column=1, columnspan=2, sticky=tk.EW)

        row += 1
        check_button = tk.Checkbutton(frame, text='Send emails',
                                      variable=self.send_emails)
        check_button.grid(row=row, column=1, sticky=tk.W)

        row += 1
        frame.rowconfigure(row, weight=1)
        row += 1
        self.button_frame = self._button_frame(frame)
        self.button_frame.grid(row=row, column=0, columnspan=4,
                               sticky=tk.EW, pady=PAD)

        return frame

    def _label_and_spinbox(self, frame, row, text, textvariable):
        label = ttk.Label(frame, text=text)
        label.grid(row=row, column=0, sticky=tk.E, pady=PAD/2)
        spinbox = ttk.Spinbox(
            frame,
            width=2,
            from_=0,
            to=COL_MAXIMUM,
            increment=1,
            textvariable=textvariable)
        spinbox.grid(row=row, column=1, sticky=tk.EW, padx=PAD)
        clickable_widget(spinbox)

    def _button_frame(self, master: tk.Frame) -> tk.Frame:
        frame = ButtonFrame(master, tk.HORIZONTAL)
        buttons = [
            frame.icon_button('save', True, self._save_config),
            frame.icon_button('exit', False, self.dismiss),
        ]
        frame.buttons = buttons
        frame.enable(False)
        return frame

    def _show_password(self, *args):
        if self.password_hide:
            self.password_hide = ''
            self.password_button.self.config(text=text.HIDE)
        else:
            self.password_hide = '*'
            self.password_button.self.config(text=text.SHOW)
        self.password_entry.self.config(show=self.password_hide)

    def _get_email_template(self):
        """Set email template path."""
        email_template = filedialog.askopenfilename(
            initialdir=Path(self.email_template.get()).parent,
            initialfile=self.email_template.get(),
            filetypes=TXT_FILE_TYPES,
            parent=self.root,
        )
        if email_template:
            self.email_template.set(email_template)

    def _get_workbook_dir(self) -> None:
        """Set workbook dir."""
        workbook_dir = filedialog.askdirectory(
            initialdir=Path(self.workbook_dir.get()),
            parent=self.root,
        )
        if workbook_dir:
            self.workbook_dir.set(workbook_dir)

    def _save_config(self, *args):
        """Save defaults to self.config."""
        result = self._config_save()
        if result == self.config.STATUS_OK:
            messagebox.showinfo(title=APP_TITLE, message='Defaults saved',
                                parent=self.root)
        else:
            lf = '\n'
            message = f'Defaults not saved{lf}{result}'
            messagebox.showerror(title=APP_TITLE, message=message,
                                 parent=self.root)
        self.dismiss()

    def _value_changed(self) -> bool:
        return (
            self.workbook_dir.get() != self.config.workbook_dir or
            self.main_sheet.get() != self.config.main_sheet or
            self.directors_sheet.get() != self.config.directors_sheet or

            self.initials_col.get() != self.config.initials_col or
            self.name_col.get() != self.config.name_col or
            self.email_col.get() != self.config.email_col or
            self.username_col.get() != self.config.username_col or
            self.active_col.get() != self.config.active_col or

            self.mon_date_col.get() != self.config.mon_date_col or
            self.wed_date_col.get() != self.config.wed_date_col or

            self.email_template.get() != self.config.email_template or
            self.email_subject.get() != self.config.email_subject or
            self.send_emails.get() != self.config.send_emails)

    def _check_value_changed(self, *args):
        enable = False
        if self._value_changed():
            enable = True
        self.button_frame.enable(enable)

    def _config_save(self):
        self.config.config['workbook_dir'] = self.workbook_dir.get()
        self.config.config['main_sheet'] = self.main_sheet.get()
        self.config.config['directors_sheet'] = self.directors_sheet.get()

        self.config.config['initials_col'] = self.initials_col.get()
        self.config.config['name_col'] = self.name_col.get()
        self.config.config['email_col'] = self.email_col.get()
        self.config.config['username_col'] = self.username_col.get()
        self.config.config['active_col'] = self.active_col.get()

        self.config.config['mon_date_col'] = self.mon_date_col.get()
        self.config.config['wed_date_col'] = self.wed_date_col.get()

        self.config.config['email_template'] = self.email_template.get()
        self.config.config['email_subject'] = self.email_subject.get()
        self.config.config['send_emails'] = self.send_emails.get()

        result = self.config.save()
        return result

    def dismiss(self, *args):
        self.root.destroy()
