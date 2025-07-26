"""Module caller for Phoenix Director's Rota."""
from forms.frm_config import ConfigFrame


class ModuleCaller():
    def __init__(self, root, module) -> None:
        modules = {
            'config': self._config,
            }

        self.invalid = False
        if module == '-h':
            print(modules.keys())
            self.invalid = True
            return

        if module not in modules:
            print(f'Invalid function name: {module}')
            self.invalid = True
            return

        self.root = root.root
        modules[module]()
        self.root.destroy()

    def _config(self) -> None:
        dlg = ConfigFrame(self)
        self.root.wait_window(dlg.root)
