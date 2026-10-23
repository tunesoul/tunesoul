"""Заглушка до ШАГА 3 (экран будет реализован согласно ТЗ)."""
from ui.screens.base_screen import BaseScreen
from ui.widgets import Card


class ScreenParsePdf(BaseScreen):
    def card_header(self):
        return ("screen_parse_pdf", "Будет реализован на ШАГЕ 3")

    def start_task(self):
        raise NotImplementedError
