"""Экран «Настройки» (ТЗ 5.7): папки input/output, тема, автоочистка, сброс/сохранение."""
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QCheckBox, QFileDialog,
    QPushButton, QLineEdit, QComboBox, QMessageBox, QFormLayout,
)

from ui.widgets import Card
from core.settings import DEFAULTS, save_settings


class ScreenSettings(QWidget):
    def __init__(self, deps: dict):
        super().__init__()
        self.settings = deps["settings"]
        self.window = deps["window"]

        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 16, 16, 16)
        card = Card("Настройки", "Сохраняются в settings.json в корне проекта.")
        form = QFormLayout()
        form.setLabelAlignment(__import__("PySide6.QtCore", fromlist=["Qt"]).Qt.AlignLeft)

        self.ed_in = QLineEdit(self.settings.get("input_dir", ""))
        self.ed_in.setToolTip("Папка с входными файлами (.pdf/.xlsm/.xlsb/.xlsx)")
        form.addRow("Папка input/:", self._row_with_browse(self.ed_in, is_dir=True))

        self.ed_out = QLineEdit(self.settings.get("output_dir", ""))
        self.ed_out.setToolTip("Папка для результатов")
        form.addRow("Папка output/:", self._row_with_browse(self.ed_out, is_dir=True))

        self.cb_theme = QComboBox()
        self.cb_theme.addItems(["system", "light", "dark"])
        t = self.settings.get("theme", "system")
        self.cb_theme.setCurrentText(t if t in ("system", "light", "dark") else "system")
        self.cb_theme.setToolTip("Тема оформления (применяется сразу)")
        self.cb_theme.currentTextChanged.connect(self._on_theme)
        form.addRow("Тема:", self.cb_theme)

        self.chk_cleanup = QCheckBox("Автоочистка промежуточных файлов после полного цикла")
        self.chk_cleanup.setChecked(bool(self.settings.get("cleanup_after_cycle", False)))
        self.chk_cleanup.setToolTip("Удалять kmd_*.xlsx, report_prepared.xlsx, comparison.xlsx после успешного цикла")
        form.addRow("", self.chk_cleanup)

        card.body.addLayout(form)

        btns = QHBoxLayout()
        b_save = QPushButton("Сохранить")
        b_save.setProperty("accent", True)
        b_save.setToolTip("Записать настройки в settings.json")
        b_save.clicked.connect(self._save)
        b_reset = QPushButton("Сбросить настройки")
        b_reset.setToolTip("Вернуть значения по умолчанию (в полях; нажмите «Сохранить»)")
        b_reset.clicked.connect(self._reset)
        btns.addWidget(b_save)
        btns.addWidget(b_reset)
        btns.addStretch(1)
        card.body.addLayout(btns)

        outer.addWidget(card)
        outer.addStretch(1)

    def _row_with_browse(self, edit: QLineEdit, is_dir: bool) -> QWidget:
        w = QWidget()
        h = QHBoxLayout(w)
        h.setContentsMargins(0, 0, 0, 0)
        h.addWidget(edit, 1)
        b = QPushButton("Обзор…")
        b.setToolTip("Выбрать папку")
        b.clicked.connect(lambda: self._browse(edit, is_dir))
        h.addWidget(b)
        return w

    @staticmethod
    def _browse(edit: QLineEdit, is_dir: bool):
        if is_dir:
            p = QFileDialog.getExistingDirectory(edit.parent(), "Выберите папку", edit.text() or str(Path.home()))
        else:
            p, _ = QFileDialog.getOpenFileName(edit.parent(), "Выберите файл", edit.text() or str(Path.home()))
        if p:
            edit.setText(p)

    def _on_theme(self, theme: str):
        self.settings["theme"] = theme
        self.window.apply_theme(theme)

    def _save(self):
        self.collect_settings(self.settings)
        try:
            save_settings(self.settings)
            QMessageBox.information(self, "Настройки", "Настройки сохранены в settings.json")
        except OSError as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить:\n{e}")

    def _reset(self):
        self.ed_in.setText(DEFAULTS["input_dir"])
        self.ed_out.setText(DEFAULTS["output_dir"])
        self.cb_theme.setCurrentText(DEFAULTS["theme"])
        self.chk_cleanup.setChecked(DEFAULTS["cleanup_after_cycle"])

    def collect_settings(self, s: dict):
        s["input_dir"] = self.ed_in.text().strip() or DEFAULTS["input_dir"]
        s["output_dir"] = self.ed_out.text().strip() or DEFAULTS["output_dir"]
        s["theme"] = self.cb_theme.currentText()
        s["cleanup_after_cycle"] = self.chk_cleanup.isChecked()
