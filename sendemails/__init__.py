from .constants import (
    EMAIL_PATTERN, NAME_EMAIL_PATTERN, EMAIL_PROVIDERS,
    EMAIL_HEADER_KEYWORDS, NAME_HEADER_KEYWORDS, GRAPH_SCOPES,
    DeviceFlowCancelled,
)
from .utils import get_resource_path, ConfigManager, TemplateManager, SharedAPI


def main():
    import sys
    from PyQt6.QtWidgets import QApplication
    from qfluentwidgets import setTheme, Theme
    from sendemails.app import EmailSenderApp
    app = QApplication(sys.argv)
    setTheme(Theme.LIGHT)
    w = EmailSenderApp()
    w.show()
    sys.exit(app.exec())


# Lazy imports for UI-heavy modules to avoid Qt init at import time
def __getattr__(name):
    import importlib
    lazy = {
        'MicrosoftAuthWorker': 'sendemails.workers',
        'SmtpTestWorker': 'sendemails.workers',
        'EmailSendWorker': 'sendemails.workers',
        'WidgetBase': 'sendemails.interfaces',
        'AccountInterface': 'sendemails.interfaces',
        'ContentInterface': 'sendemails.interfaces',
        'SendInterface': 'sendemails.interfaces',
        'EmailSenderApp': 'sendemails.app',
    }
    if name in lazy:
        mod = importlib.import_module(lazy[name])
        return getattr(mod, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
