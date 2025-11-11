
"""
A tkinter application to present the choices and config for Director's Rota.
"""
import sys
import tkinter as tk

from psiutils.widgets import get_styles
from psiutils.utilities import display_icon, resource_path

from directors_rota.constants import ICON_FILE

from directors_rota.forms.frm_main import MainFrame
from directors_rota.module_caller import ModuleCaller


class Root():
    """Create the app's root and loop."""
    def __init__(self) -> None:
        self.root = tk.Tk()
        root = self.root
        icon_path = resource_path(__file__, ICON_FILE)
        display_icon(root, icon_path, ignore_error=True)
        # root.title(f'{APP_TITLE}')
        root.protocol("WM_DELETE_WINDOW", root.destroy)

        get_styles()

        dlg = None
        if len(sys.argv) > 1 and sys.argv[1] != 'main':
            module = sys.argv[1]
            dlg = ModuleCaller(self, module)
        if not dlg or dlg.invalid:
            MainFrame(self.root)

        root.mainloop()
