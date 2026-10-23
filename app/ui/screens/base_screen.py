"""Базовый экран: карточка с формой + кнопка «Выполнить» + прогресс + лог.

Потомки реализуют:
  card_header() -> (заголовок, описание)
  build_form(card)   — поля экрана в Card
  validate() -> str  — текст ошибки или "" (всё ок)
  start_task() -> (fn, args, kwargs, done_cb)  — что запускать в фоне
  on_success(result) — пост-обработка в GUI-потоке
Кнопка «Открыть результат» появляется после успешного выполнения.
"""
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QScrollArea, QPushButton,
    QProgressBar, QLabel, QMessageBox, QLineEdit, QFrame,
)

from core.logger import install_interceptor
from core.worker import Worker
from ui.widgets import Card, LogPanel


class BaseScreen(QWidget):
    action_label = "Выполнить"
    action_tip = "Запустить операцию (Ctrl+Enter)"

    def __init__(self, deps: dict):
        super().__init__()
        self.deps = deps
        self.settings: dict = deps["settings"]
        self.window = deps["window"]
        self._worker: Worker | None = None
        self._result_path: Path | None = None

        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 16, 16, 16)
        outer.setSpacing(10)

        self.card = Card(*self.card_header())
        sa = QScrollArea()
        sa.setWidgetResizable(True)
        sa.setFrameShape(QFrame.NoFrame)
        holder = QWidget()
        hv = QVBoxLayout(holder)
        hv.setContentsMargins(0, 0, 0, 0)
        self.build_form(self.card)
        hv.addWidget(self.card)
        hv.addStretch(1)
        sa.setWidget(holder)
        outer.addWidget(sa, 2)

        brow = QHBoxLayout()
        self.btn_run = QPushButton(self.action_label)
        self.btn_run.setProperty("accent", True)
        self.btn_run.setToolTip(self.action_tip)
        self.btn_run.clicked.connect(self.on_run)
        brow.addWidget(self.btn_run)
        self.btn_open = QPushButton("Открыть результат")
        self.btn_open.setToolTip("Открыть созданный файл стандартным приложением")
        self.btn_open.setVisible(False)
        self.btn_open.clicked.connect(self._open_result)
        brow.addWidget(self.btn_open)
        brow.addStretch(1)
        outer.addLayout(brow)

        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setRange(0, 0)   # indeterminate
        self.progress.setVisible(False)
        outer.addWidget(self.progress)

        self.log = LogPanel()
        outer.addWidget(self.log, 1)

        sc = QShortcut(QKeySequence("Ctrl+O"), self)
        sc.setContext(Qt.WidgetWithChildrenShortcut)
        sc.activated.connect(self.focus_first_input)
        sc2 = QShortcut(QKeySequence("Ctrl+Return"), self)
        sc2.setContext(Qt.WidgetWithChildrenShortcut)
        sc2.activated.connect(self.on_run)

    # -------- переопределяется потомками --------
    def card_header(self) -> tuple[str, str]:
        return ("Экран", "")

    def build_form(self, card: Card):
        pass

    def validate(self) -> str:
        return ""

    def start_task(self):
        raise NotImplementedError

    def on_success(self, result):
        pass

    def collect_settings(self, s: dict):
        pass

    def focus_first_input(self):
        le = self.findChild(QLineEdit)
        if le:
            le.setFocus()

    # -------- общий механизм запуска --------
    def on_run(self):
        if self._worker is not None and self._worker.isRunning():
            return
        err = self.validate()
        if err:
            self.log.append_line(f"⚠️ {err}")
            QMessageBox.warning(self, "Проверьте поля", err)
            return
        fn, args, kwargs, done_cb = self.start_task()
        self.btn_run.setEnabled(False)
        self.btn_open.setVisible(False)
        self.progress.setVisible(True)
        self._result_path = None
        self.window.set_status(f"{self.card_header()[0]}: выполняется…")
        w = Worker(install_interceptor(), fn, *args, **kwargs)
        w.log.connect(self.log.append_line)
        w.finished_with_result.connect(lambda r: self._done(w, r, done_cb))
        w.failed.connect(lambda tb: self._fail(w, tb))
        self._worker = w
        w.start()

    def _done(self, w, result, done_cb):
        self._worker = None
        self.progress.setVisible(False)
        self.btn_run.setEnabled(True)
        self.window.set_status("Готово")
        try:
            if done_cb:
                done_cb(result)
            self.on_success(result)
        except Exception as e:
            self.log.append_line(f"⚠️ Ошибка обработки результата: {e}")

    def _fail(self, w, tb):
        self._worker = None
        self.progress.setVisible(False)
        self.btn_run.setEnabled(True)
        self.window.set_status("Ошибка — см. лог")
        short = tb.strip().splitlines()[-1] if tb.strip() else "Неизвестная ошибка"
        hint = ""
        if "PermissionError" in tb:
            hint = "\n\n💡 Закройте файл в Excel и попробуйте снова."
        elif "FileNotFoundError" in tb:
            hint = "\n\n💡 Проверьте, что файл существует и путь указан верно."
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Critical)
        box.setWindowTitle("Ошибка выполнения")
        box.setText(short + hint)
        box.setDetailedText(tb)
        box.exec()

    def show_result_button(self, path: Path):
        self._result_path = Path(path)
        self.log.append_line(f"✅ Результат: {self._result_path}")
        self.btn_open.setVisible(True)

    def _open_result(self):
        if self._result_path:
            self.window.open_path(self._result_path)
