"""Главное окно: боковое меню (QListWidget) + QStackedWidget, статус-строка,
меню «Файл/Справка», тема, горячие клавиши, мост для ask_choice."""
import json
import threading
from pathlib import Path

from PySide6.QtCore import Qt, Signal, Slot
from PySide6.QtGui import QKeySequence, QAction, QIcon
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QListWidget, QStackedWidget,
    QLabel, QStatusBar, QMessageBox, QDialog, QListWidgetItem,
    QApplication, QAbstractItemView,
)

from core.paths import APP_DIR, ASSETS_DIR, SETTINGS_PATH
from core.logger import app_log, open_in_explorer
from core.settings import save_settings
from core import patches

APP_TITLE = "Сравнение таблиц КМД — v1.0"
VERSION = "v1.0"

# (id, элемент меню) — по порядку вкладок
SCREENS = [
    ("parse",     "Парсинг PDF"),
    ("prepare",   "Подготовка отчётной"),
    ("compare",   "Сравнение"),
    ("transfer",  "Перенос из DEP"),
    ("cycle",     "Полный цикл"),
    ("settings",  "Настройки"),
]

DARK_QSS = """
QWidget { background: #232629; color: #eff0f1; }
QLineEdit, QPlainTextEdit, QComboBox, QListWidget, QDialogButtonBox {
    background: #31363b; border: 1px solid #4f5b66; border-radius: 4px; padding: 3px; }
QPushButton { background: #3daee9; color: #15171a; border: none;
    border-radius: 4px; padding: 6px 14px; font-weight: bold; }
QPushButton:hover { background: #6cc4f5; }
QPushButton:disabled { background: #4f5b66; color: #8b95a1; }
QPushButton[accent="true"] { background: #27ae60; color: white; }
QPushButton[accent="true"]:hover { background: #2ecc71; }
QProgressBar { background: #31363b; border: 1px solid #4f5b66;
    border-radius: 4px; text-align: center; }
QProgressBar::chunk { background: #3daee9; border-radius: 3px; }
QFrame#Card { background: #2a2e32; border: 1px solid #3a3f44; border-radius: 8px; }
QListWidget#Menu { background: #1d2023; border: none; outline: none; }
QListWidget#Menu::item { padding: 12px 16px; border-left: 3px solid transparent; }
QListWidget#Menu::item:selected { background: #31363b; border-left: 3px solid #3daee9; }
QCheckBox { color: #eff0f1; }
QLabel { color: #eff0f1; }
QMenuBar { background: #232629; }
QMenuBar::item:selected { background: #31363b; }
QMenu { background: #2a2e32; border: 1px solid #4f5b66; }
QMenu::item:selected { background: #3daee9; color: #15171a; }
QStatusBar { background: #1d2023; color: #9aa0a6; }
"""

LIGHT_QSS = """
QFrame#Card { background: #ffffff; border: 1px solid #d7dbe0; border-radius: 8px; }
QListWidget#Menu { background: #f2f4f6; border: none; outline: none; }
QListWidget#Menu::item { padding: 12px 16px; border-left: 3px solid transparent; }
QListWidget#Menu::item:selected { background: #e3ecf5; border-left: 3px solid #2d7dd2; }
QPushButton { border-radius: 4px; padding: 6px 14px; }
QPushButton[accent="true"] { background: #27ae60; color: white; border: none; }
QPushButton[accent="true"]:hover { background: #2ecc71; }
QStatusBar { color: #666; }
"""


class _ChoiceDialog(QDialog):
    """Модальный выбор варианта (резервный режим ask_choice)."""

    def __init__(self, prompt: str, options: list, parent=None):
        super().__init__(parent)
        self.setWindowTitle(APP_TITLE)
        self.selected = None
        lay = QHBoxLayout(self)
        from PySide6.QtWidgets import QVBoxLayout, QPushButton
        v = QVBoxLayout()
        lbl = QLabel(prompt)
        lbl.setWordWrap(True)
        v.addWidget(lbl)
        self.lst = QListWidget()
        for o in options:
            self.lst.addItem(str(o))
        self.lst.itemDoubleClicked.connect(self._accept_item)
        v.addWidget(self.lst)
        btns = QHBoxLayout()
        ok = QPushButton("Выбрать")
        ok.clicked.connect(lambda: self._accept_item(self.lst.currentItem()))
        no = QPushButton("Отмена")
        no.clicked.connect(self.reject)
        btns.addStretch(1)
        btns.addWidget(ok)
        btns.addWidget(no)
        v.addLayout(btns)
        lay.addLayout(v)
        self.resize(420, 380)

    def _accept_item(self, item):
        if item is not None:
            self.selected = item.text()
            self.accept()


class MainWindow(QMainWindow):
    _askRequested = Signal(str, str, object, object)  # prompt, options_json, event, box

    def __init__(self, settings: dict):
        super().__init__()
        self.settings = settings
        self._askRequested.connect(self._on_ask_requested)
        self.setWindowTitle(APP_TITLE)
        self.resize(1100, 750)
        self.setMinimumSize(900, 600)

        ico = ASSETS_DIR / "icon.ico"
        if ico.is_file():
            self.setWindowIcon(QIcon(str(ico)))

        # ---- layout: меню слева, экраны справа ----
        central = QWidget()
        h = QHBoxLayout(central)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(0)

        self.menu = QListWidget()
        self.menu.setObjectName("Menu")
        self.menu.setFixedWidth(220)
        self.menu.setSelectionMode(QAbstractItemView.SingleSelection)
        for sid, label in SCREENS:
            self.menu.addItem(QListWidgetItem(label))
        self.menu.currentRowChanged.connect(self._on_menu)
        h.addWidget(self.menu)

        self.stack = QStackedWidget()
        h.addWidget(self.stack, 1)
        self.setCentralWidget(central)

        # ---- экраны ----
        from ui.screens.screen_parse_pdf import ScreenParsePdf
        from ui.screens.screen_prepare import ScreenPrepare
        from ui.screens.screen_compare import ScreenCompare
        from ui.screens.screen_transfer import ScreenTransfer
        from ui.screens.screen_full_cycle import ScreenFullCycle
        from ui.screens.screen_settings import ScreenSettings

        self.screens = {}
        deps = {"settings": self.settings, "window": self}
        self.screens["parse"] = ScreenParsePdf(deps)
        self.screens["prepare"] = ScreenPrepare(deps)
        self.screens["compare"] = ScreenCompare(deps)
        self.screens["transfer"] = ScreenTransfer(deps)
        self.screens["cycle"] = ScreenFullCycle(deps)
        self.screens["settings"] = ScreenSettings(deps)
        for w in self.screens.values():
            self.stack.addWidget(w)

        # ---- меню ----
        mb = self.menuBar()
        m_file = mb.addMenu("&Файл")
        act_set = QAction("На&стройки", self)
        act_set.triggered.connect(lambda: self.menu.setCurrentRow(len(SCREENS) - 1))
        m_file.addAction(act_set)
        act_quit = QAction("&Выход", self)
        act_quit.setShortcut(QKeySequence("Ctrl+Q"))
        act_quit.triggered.connect(self.close)
        m_file.addAction(act_quit)
        m_help = mb.addMenu("&Справка")
        act_about = QAction("&О программе", self)
        act_about.triggered.connect(self._about)
        m_help.addAction(act_about)

        # ---- статус-строка ----
        sb = QStatusBar()
        self.setStatusBar(sb)
        self.status_state = QLabel("Готово")
        sb.addWidget(QLabel(f"{VERSION} · "))
        sb.addWidget(QLabel(f"settings.json: {SETTINGS_PATH} · "))
        sb.addPermanentWidget(self.status_state)

        # bridge для ask_choice из фоновых потоков
        patches.set_dialog_bridge(self._ask_from_worker)

        self.apply_theme(settings.get("theme", "system"))
        self.menu.setCurrentRow(0)

        geo = settings.get("window_geometry") or ""
        if geo:
            try:
                from PySide6.QtCore import QByteArray
                self.restoreGeometry(QByteArray.fromHex(geo.encode()))
            except Exception:
                pass

    # ---------- навигация ----------
    def _on_menu(self, row: int):
        if 0 <= row < self.stack.count():
            self.stack.setCurrentIndex(row)

    def goto_screen(self, sid: str):
        for i, (k, _) in enumerate(SCREENS):
            if k == sid:
                self.menu.setCurrentRow(i)
                return

    def set_status(self, text: str):
        self.status_state.setText(text)

    # ---------- тема ----------
    def apply_theme(self, theme: str):
        app = QApplication.instance()
        if theme == "dark":
            app.setStyle("Fusion")
            app.setStyleSheet(DARK_QSS)
        elif theme == "light":
            app.setStyle("Fusion")
            app.setStyleSheet(LIGHT_QSS)
        else:  # system
            app.setStyle("windowsvista" if Path().exists() and __import__("sys").platform == "win32" else "Fusion")
            app.setStyleSheet("")

    # ---------- ask_choice bridge ----------
    def _ask_from_worker(self, prompt: str, options: list):
        """Вызывается из фонового потока: блокированно показывает диалог в GUI-потоке.

        Python-объекты нельзя передать через Q_ARG, а invokeMethod с
        BlockingQueuedConnection не вызывает python-слоты. Используем сигнал
        (queued connection) + threading.Event (ТЗ 6.2)."""
        import json
        ev = threading.Event()
        box = {}
        self._askRequested.emit(prompt, json.dumps([str(o) for o in options], ensure_ascii=False), ev, box)
        if not ev.wait(timeout=3600):   # 1 час на ответ пользователя
            return None
        return box.get("answer")

    @Slot(str, str, object, object)
    def _on_ask_requested(self, prompt, options_json, ev, box):
        options = json.loads(options_json)
        dlg = _ChoiceDialog(prompt, options, self)
        if dlg.exec() and dlg.selected is not None:
            box["answer"] = dlg.selected
        ev.set()

    # ---------- прочее ----------
    def _about(self):
        QMessageBox.about(
            self, "О программе",
            f"<b>{APP_TITLE}</b><br><br>"
            "Графическая обёртка над консольным пайплайном excel-compare:<br>"
            "парсинг PDF → подготовка отчётной → сравнение КМД → перенос из DEP.<br><br>"
            "Бизнес-логика находится в src/, GUI — в app/.")

    def open_path(self, path):
        open_in_explorer(path)

    def closeEvent(self, ev):
        from PySide6.QtCore import QByteArray
        self.settings["window_geometry"] = bytes(QByteArray(self.saveGeometry()).toHex()).decode()
        for s in self.screens.values():
            s.collect_settings(self.settings)
        try:
            save_settings(self.settings)
        except OSError as e:
            app_log(f"Не удалось сохранить настройки: {e}")
        super().closeEvent(ev)
