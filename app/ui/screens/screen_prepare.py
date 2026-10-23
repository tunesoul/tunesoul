"""Заглушка до ШАГА 4 (экран будет реализован согласно ТЗ)."""
from ui.screens.base_screen import BaseScreen
from ui.widgets import Card


class ScreenPrepare(BaseScreen):
    def card_header(self):
        return ("screen_prepare", "Будет реализован на ШАГЕ 4")

    def start_task(self):
        raise NotImplementedError
