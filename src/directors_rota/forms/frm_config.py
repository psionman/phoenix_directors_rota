import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from pathlib import Path

from psiutils.widgets import clickable_widget, separator_frame
from psiutils.buttons import ButtonFrame, IconButton
from psiutils.constants import PAD
from psiutils.utilities import window_resize, geometry

from directors_rota.constants import COL_MAXIMUM, APP_TITLE, TXT_FILE_TYPES
from directors_rota.config import read_config
from directors_rota import logger
from directors_rota.text import Text

txt = Text()

FRAME_TITLE = 'System defaults'

FIELDS = {
    "workbook_dir": tk.StringVar,
    "main_sheet": tk.StringVar,
    "directors_sheet": tk.StringVar,
    "initials_col": tk.IntVar,
    "name_col": tk.IntVar,
    "email_col": tk.IntVar,
    "username_col": tk.IntVar,
    "active_col": tk.IntVar,
    "mon_date_col": tk.IntVar,
    "wed_date_col": tk.IntVar,
    "email_subject": tk.StringVar,
    "email_template": tk.StringVar,
    "send_emails": tk.BooleanVar,
}


class ConfigFrame():
    workbook_dir: tk.StringVar
    main_sheet: tk.StringVar
    directors_sheet: tk.StringVar
    initials_col: tk.IntVar
    name_col: tk.IntVar
    email_col: tk.IntVar
    username_col: tk.IntVar
    active_col: tk.IntVar
    mon_date_col: tk.IntVar
    wed_date_col: tk.IntVar
    email_subject: tk.StringVar
    email_template: tk.StringVar
    send_emails: tk.BooleanVar

    def __init__(self, parent) -> None:
        # pylint: disable=no-member)
        self.root = tk.Toplevel(parent.root)
        self.parent = parent
        config = read_config()
        self.config = config

        for field, f_type in FIELDS.items():
            if f_type is tk.StringVar:
                setattr(self, field, self._stringvar(getattr(config, field)))
            elif f_type is tk.IntVar:
                setattr(self, field, self._intvar(getattr(config, field)))
            elif f_type is tk.BooleanVar:
                setattr(self, field, self._boolvar(getattr(config, field)))

        self.password_hide = ''
        self.button_frame = None
        self.password_button = None
        self.password_entry = None

        self._show()

    def _stringvar(self, value: str) -> tk.StringVar:
        stringvar = tk.StringVar(value=value)
        stringvar.trace_add('write', self._check_value_changed)
        return stringvar

    def _intvar(self, value: int) -> tk.IntVar:
        intvar = tk.IntVar(value=value)
        intvar.trace_add('write', self._check_value_changed)
        return intvar

    def _boolvar(self, value: bool) -> tk.BooleanVar:
        boolvar = tk.BooleanVar(value=value)
        boolvar.trace_add('write', self._check_value_changed)
        return boolvar

    def _show(self) -> None:
        root = self.root
        root.geometry(geometry(self.config, __file__))
        root.title(FRAME_TITLE)
        root.transient(self.parent.root)

        root.bind('<Control-x>', self._dismiss)
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
            frame, txt.OPEN, 'open', self._get_workbook_dir)
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
            frame, txt.OPEN, 'open', self._get_email_template)
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
            frame.icon_button('save', self._save_config, True),
            frame.icon_button('exit', self._dismiss),
        ]
        frame.buttons = buttons
        frame.enable(False)
        return frame

    def _show_password(self, *args):
        if self.password_hide:
            self.password_hide = ''
            self.password_button.self.config(text=txt.HIDE)
        else:
            self.password_hide = '*'
            self.password_button.self.config(text=txt.SHOW)
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
        self._dismiss()

    def _check_value_changed(self, *args):
        enable = bool(self._config_changes())
        self.button_frame.enable(enable)

    def _config_save(self):
        changes = {field: f'(old value={change[0]}, new_value={change[1]})'
                   for field, change in self._config_changes().items()}

        logger.info(
            "Config saved",
            changes=changes
        )

        for field in FIELDS:
            self.config.config[field] = getattr(self, field).get()
        return self.config.save()

    def _config_changes(self) -> dict:
        stored = self.config.config
        return {
            field: (stored[field], getattr(self, field).get())
            for field in FIELDS
            if stored[field] != getattr(self, field).get()
        }

    def _dismiss(self, *args):
        self.root.destroy()
