"""Экран «Парсинг PDF» (ТЗ §5.2, ШАГ 3).

Поля: PDF-файл (FileSelect, автоподстановка из settings["last_pdf"]), кнопка
«Парсить», indeterminate-прогресс, живой лог, «Открыть результат».
Бизнес-логика не дублируется: вызывается src/pdf_to_kmd.parse_pdf через импорт.
"""
from pathlib import Path

from PySide6.QtWidgets import QLabel

from core.settings import save_settings
from core.src_loader import load_src
from ui.screens.base_screen import BaseScreen
from ui.widgets import Card, FileSelect


class ScreenParsePdf(BaseScreen):
    action_label = "Парсить"
    action_tip = "Распознать таблицы КМД из PDF в Excel (Ctrl+Enter)"

    def card_header(self) -> tuple[str, str]:
        return ("Парсинг PDF", "Прошмандовка (PDF) → output/kmd_*.xlsx")

    def build_form(self, card: Card):
        self.fs_pdf = FileSelect(".pdf", "Путь к PDF-файлу прошмандовки")
        last = self.settings.get("last_pdf", "")
        if last and Path(last).is_file():
            self.fs_pdf.set_text(last)
        card.add(QLabel("PDF-файл:"))
        card.add(self.fs_pdf)

    def validate(self) -> str:
        ok, err = self.fs_pdf.validate()
        return "" if ok else err

    def start_task(self):
        pdf_path = Path(self.fs_pdf.text())
        mod = load_src("pdf_to_kmd")   # тяжёлый импорт — тоже в фоне
        return (mod.parse_pdf, (pdf_path,), {}, None)

    def on_success(self, result):
        self.settings["last_pdf"] = self.fs_pdf.text()
        save_settings(self.settings)
        self.show_result_button(Path(result))
