"""Точка входа GUI (запуск: python app/main_gui.py или run_gui.bat)."""
import sys
from pathlib import Path

# чтобы импорты вида `from core... / from ui...` работали из любого места
sys.path.insert(0, str(Path(__file__).resolve().parent))

from PySide6.QtWidgets import QApplication  # noqa: E402
from PySide6.QtGui import QIcon             # noqa: E402

from core.paths import ASSETS_DIR           # noqa: E402
from core.logger import init_logging, install_interceptor, app_log  # noqa: E402
from core.settings import load_settings     # noqa: E402
from core.patches import apply_patches      # noqa: E402
from ui.main_window import MainWindow, APP_TITLE  # noqa: E402


def main():
    init_logging()
    install_interceptor()          # print() фоновых потоков → логи экранов
    apply_patches()                # ask_choice/input больше не заблокируют GUI

    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    ico = ASSETS_DIR / "icon.ico"
    if ico.is_file():
        app.setWindowIcon(QIcon(str(ico)))

    s = load_settings()
    w = MainWindow(s)
    w.show()
    app_log("GUI запущен")
    rc = app.exec()
    app_log(f"GUI завершён, код {rc}")
    sys.exit(rc)


if __name__ == "__main__":
    main()
