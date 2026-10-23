"""Заглушка до ШАГА 7 (экран будет реализован согласно ТЗ)."""
from ui.screens.base_screen import BaseScreen
from ui.widgets import Card


class ScreenFullCycle(BaseScreen):
    def card_header(self):
        return ("screen_full_cycle", "Будет реализован на ШАГЕ 7")

    def start_task(self):
        raise NotImplementedError
