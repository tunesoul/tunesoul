"""Переиспользуемые виджеты: строка «файл + Обзор…», лог-панель, карточка экрана."""
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QWidget, QHBoxLayout, QVBoxLayout, QLineEdit, QPushButton,
    QFileDialog, QPlainTextEdit, QLabel, QFrame,
)

FILE_FILTERS = {
    ".pdf": "PDF-документы (*.pdf)",
    ".xlsm": "Книги Excel с макросами (*.xlsm)",
    ".xlsb": "Двоичные книги Excel (*.xlsb)",
    ".xlsx": "Книги Excel (*.xlsx)",
}


class FileSelect(QWidget):
    """QLineEdit + кнопка «Обзор…» с фильтром по расширению."""

    def __init__(self, ext: str, placeholder: str = "", parent: QWidget | None = None):
        super().__init__(parent)
        self.ext = ext
        self.edit = QLineEdit()
        self.edit.setPlaceholderText(placeholder or f"Путь к файлу {ext}")
        self.edit.setToolTip(f"Файл {ext}. Можно вставить путь и нажать Enter.")
        self.btn = QPushButton("Обзор…")
        self.btn.setToolTip(f"Выбрать файл {ext}")
        self.btn.clicked.connect(self.browse)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(self.edit, 1)
        lay.addWidget(self.btn)

    def browse(self):
        start = self.text() or str(Path.home())
        flt = FILE_FILTERS.get(self.ext, f"Файлы (*{self.ext})")
        path, _ = QFileDialog.getOpenFileName(self, f"Выберите {self.ext}", start, flt)
        if path:
            self.edit.setText(str(Path(path).resolve()))

    def text(self) -> str:
        return self.edit.text().strip()

    def set_text(self, value: str):
        self.edit.setText(value or "")

    def validate(self) -> tuple[bool, str]:
        p = self.text()
        if not p:
            return False, f"Не выбран файл {self.ext}"
        if not p.lower().endswith(self.ext):
            return False, f"Ожидается файл {self.ext}: {p}"
        if not Path(p).is_file():
            return False, f"Файл не найден: {p}"
        return True, ""


class LogPanel(QPlainTextEdit):
    """read-only лог в реальном времени, моноширинный, автоскролл."""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setReadOnly(True)
        f = QFont("Consolas")
        f.setStyleHint(QFont.Monospace)
        f.setPointSize(9)
        self.setFont(f)
        self.setMaximumBlockCount(5000)

    def append_line(self, line: str):
        self.appendPlainText(line)
        sb = self.verticalScrollBar()
        sb.setValue(sb.maximum())

    def clear_log(self):
        self.setPlainText("")


class Card(QFrame):
    """Блок с заголовком экрана и описанием."""

    def __init__(self, title: str, subtitle: str = "", parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("Card")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 10, 14, 12)
        t = QLabel(title)
        tf = QFont(); tf.setPointSize(13); tf.setBold(True)
        t.setFont(tf)
        lay.addWidget(t)
        if subtitle:
            s = QLabel(subtitle)
            s.setWordWrap(True)
            s.setStyleSheet("color: gray;")
            lay.addWidget(s)
        self.body = QVBoxLayout()
        lay.addLayout(self.body)

    def add(self, widget: QWidget):
        self.body.addWidget(widget)


def hline() -> QFrame:
    ln = QFrame()
    ln.setFrameShape(QFrame.HLine)
    ln.setFrameShadow(QFrame.Sunken)
    return ln
