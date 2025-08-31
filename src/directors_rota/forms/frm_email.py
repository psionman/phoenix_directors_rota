
import tkinter as tk
from tkinter import ttk, messagebox
import clipboard

from psiutils.buttons import ButtonFrame
from psiutils.widgets import WaitCursor
from psiutils.constants import PAD
from psiutils.utilities import window_resize, geometry

from directors_rota.config import read_config
from directors_rota.emails import send_emails


class EmailFrame():
    def __init__(self, parent):
        self.root = tk.Toplevel(parent.root)
        self.parent = parent
        self.directors = parent.directors
        self.config = read_config()
        self.email_text = None

        # tk variables
        self.email = tk.StringVar(value=parent.email)
        self.send_emails = tk.BooleanVar(value=self.config.send_emails)

        self.show()

    def show(self) -> None:
        root = self.root
        root.geometry(geometry(self.config, __file__))
        root.title('Rota')

        root.bind('<Control-x>', self._dismiss)
        root.bind('<Control-s>', self._send_emails)
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

        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        self.email_text = tk.Text(frame, height=30)
        self.email_text.grid(row=0, column=0, sticky=tk.NSEW)
        self.email_text.insert('0.0', self.email.get())

        check_button = tk.Checkbutton(frame, text='Send emails',
                                      variable=self.send_emails)
        check_button.grid(row=1, column=0, sticky=tk.W)

        buttons = self._button_frame(frame)
        buttons.grid(row=9, column=0, columnspan=9,
                     sticky=tk.EW, pady=PAD)

        return frame

    def _button_frame(self, master: tk.Frame) -> tk.Frame:
        frame = ButtonFrame(master, tk.HORIZONTAL)
        buttons = [
            frame.icon_button('send', True, self._send_emails),
            frame.icon_button('exit', False, self._dismiss),
        ]
        frame.buttons = buttons
        return frame

    def _send_emails(self) -> None:
        text = self.email_text.get('1.0', 'end')
        clipboard.copy(text)
        if self.send_emails.get():
            with WaitCursor(self.root):
                response = send_emails(text, self.directors)
            if isinstance(response, int):
                messagebox.showinfo(
                    'Emails', f'{response} emails sent.', parent=self.root)
                self._dismiss()
                return
            messagebox.showerror(
                'Emails', 'Emails not sent.', parent=self.root)

    def _dismiss(self, event: object = None):
        self.root.destroy()
