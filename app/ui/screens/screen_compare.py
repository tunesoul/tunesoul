"""Экран «Сравнение» (ТЗ §5.4, ШАГ 5).

Поля: kmd_*.xlsx и report_prepared.xlsx с автопоиском в output/.
Вызывается src/compare_kmd.compare — бизнес-логика не дублируется.
"""
from pathlib import Path

from PySide6.QtWidgets import QLabel, QPushButton

from core.paths import OUTPUT_DIR
from core.settings import save_settings
from core.src_loader import load_src
from ui.screens.base_screen import BaseScreen
from ui.widgets import Card, FileSelect


def _find_latest(pattern: str) -> str:
    cands = sorted(OUTPUT_DIR.glob(pattern), key=lambda p: p.stat().st_mtime,
                   reverse=True)
    return str(cands[0]) if cands else ""


class ScreenCompare(BaseScreen):
    action_label = "Сравнить"
    action_tip = "Прошмандовка vs подготовленная отчётная → output/comparison.xlsx (Ctrl+Enter)"

    def card_header(self) -> tuple[str, str]:
        return ("Сравнение", "kmd_*.xlsx + report_prepared.xlsx → comparison.xlsx")

    def build_form(self, card: Card):
        self.fs_kmd = FileSelect(".xlsx", "Прошмандовка (output/kmd_*.xlsx)")
        self.fs_rep = FileSelect(".xlsx", "Отчётная подготовленная (report_prepared.xlsx)")
        btn = QPushButton("Автопоиск последних файлов в output/")
        btn.clicked.connect(self._autofind)

        card.add(QLabel("Прошмандовка (kmd_*.xlsx):"))
        card.add(self.fs_kmd)
        card.add(QLabel("Отчётная подготовленная:"))
        card.add(self.fs_rep)
        card.add(btn)

        # автоподстановка: сохранённые пути, иначе свежие файлы из output/
        self.fs_kmd.set_text(self.settings.get("last_xlsx", "") or _find_latest("kmd_*.xlsx"))
        self.fs_rep.set_text(_find_latest("report_prepared.xlsx"))

    def _autofind(self):
        k, r = _find_latest("kmd_*.xlsx"), _find_latest("report_prepared.xlsx")
        if k:
            self.fs_kmd.set_text(k)
        if r:
            self.fs_rep.set_text(r)
        self.log.append_line(f"🔎 Автопоиск: kmd={'есть' if k else 'нет'}, "
                             f"report_prepared={'есть' if r else 'нет'}")

    def validate(self) -> str:
        for fs in (self.fs_kmd, self.fs_rep):
            ok, err = fs.validate()
            if not ok:
                return err
        return ""

    def start_task(self):
        mod = load_src("compare_kmd")
        return (mod.compare, (Path(self.fs_kmd.text()), Path(self.fs_rep.text())), {}, None)

    def on_success(self, result):
        self.settings["last_xlsx"] = self.fs_kmd.text()
        save_settings(self.settings)
        self.show_result_button(Path(result))
