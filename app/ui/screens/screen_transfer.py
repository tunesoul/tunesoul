"""Заглушка до ШАГА 6 (экран будет реализован согласно ТЗ)."""
from ui.screens.base_screen import BaseScreen
from ui.widgets import Card


class ScreenTransfer(BaseScreen):
    def card_header(self):
        return ("screen_transfer", "Будет реализован на ШАГЕ 6")

    def start_task(self):
        raise NotImplementedError
