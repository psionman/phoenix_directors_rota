
"""
A tkinter application to present the choices and config for Director's Rota.
"""
import sys
import tkinter as tk

from psiutils.widgets import get_styles
from psiutils.utilities import display_icon

from constants import ICON_FILE

from forms.frm_main import MainFrame
from module_caller import ModuleCaller


class Root():
    def __init__(self) -> None:
        """Create the app's root and loop."""
        self.root = tk.Tk()
        root = self.root
        display_icon(root, ICON_FILE, ignore_error=True)
        root.title('{APP_TITLE}')
        root.protocol("WM_DELETE_WINDOW", root.destroy)

        get_styles()

        dlg = None
        if len(sys.argv) > 1 and sys.argv[1] != 'main':
            module = sys.argv[1]
            dlg = ModuleCaller(self, module)
        if not dlg or dlg.invalid:
            MainFrame(self.root)

        root.mainloop()
