"""Заглушка до ШАГА 5 (экран будет реализован согласно ТЗ)."""
from ui.screens.base_screen import BaseScreen
from ui.widgets import Card


class ScreenCompare(BaseScreen):
    def card_header(self):
        return ("screen_compare", "Будет реализован на ШАГЕ 5")

    def start_task(self):
        raise NotImplementedError
