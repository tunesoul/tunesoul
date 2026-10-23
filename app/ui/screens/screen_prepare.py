"""Экран «Подготовка отчётной» (ТЗ §5.3, ШАГ 4).

Лист/КМ/DP читаются в GUI-потоке через openpyxl/pandas и заполняют QComboBox;
выбранные значения передаются в prepare_report через предзаполненную очередь
patches.set_prefilled (ask_choice смонки-патчен). Значения КМ/DP берутся из
модуля src/prepare_report (его же функции find_header_row/find_columns/COL_KM/
COL_DP) — логика не дублируется и гарантированно совпадает с этапом выполнения.
"""
from pathlib import Path

from PySide6.QtWidgets import QLabel, QComboBox, QPushButton

from core.patches import set_prefilled
from core.settings import save_settings
from core.src_loader import load_src
from ui.screens.base_screen import BaseScreen
from ui.widgets import Card, FileSelect


class ScreenPrepare(BaseScreen):
    action_label = "Подготовить"
    action_tip = "Сгруппировать отчётную таблицу → output/report_prepared.xlsx (Ctrl+Enter)"

    def card_header(self) -> tuple[str, str]:
        return ("Подготовка отчётной", ".xlsm отчётная → report_prepared.xlsx")

    def build_form(self, card: Card):
        self.fs_xlsm = FileSelect(".xlsm", "Путь к отчётному .xlsm")
        last = self.settings.get("last_xlsm", "")
        if last and Path(last).is_file():
            self.fs_xlsm.set_text(last)
        self.combo_sheet = QComboBox()
        self.combo_km = QComboBox()
        self.combo_dp = QComboBox()
        self.btn_refresh = QPushButton("Обновить списки (лист / КМ / DP)")
        self.btn_refresh.clicked.connect(self._load_sheets)

        for lbl, w in (("Файл .xlsm:", self.fs_xlsm), ("Лист:", self.combo_sheet),
                       ("КМ:", self.combo_km), ("DP:", self.combo_dp),
                       ("", self.btn_refresh)):
            if lbl:
                card.add(QLabel(lbl))
            card.add(w)

        self.fs_xlsm.edit.textChanged.connect(lambda _: self._clear_combos())
        self.combo_sheet.currentIndexChanged.connect(self._load_km_dp)

    # -------- автозаполнение комбобоксов (в GUI-потоке) --------
    def _clear_combos(self):
        for c in (self.combo_sheet, self.combo_km, self.combo_dp):
            c.blockSignals(True)
            c.clear()
            c.blockSignals(False)

    def _read_only(self):
        """(sheets, km_vals, dp_vals, err) — теми же функциями, что и prepare_report."""
        try:
            from openpyxl import load_workbook
            import pandas as pd
            mod = load_src("prepare_report")   # только чтение констант/функций
            p = Path(self.fs_xlsm.text())
            wb = load_workbook(p, read_only=True, data_only=True)
            sheets = list(wb.sheetnames)
            wb.close()
            sheet = self.combo_sheet.currentText() or (sheets[0] if sheets else "")
            raw = pd.read_excel(p, sheet_name=sheet, header=None,
                                engine="openpyxl", dtype=object)
            hr = mod.find_header_row(raw)
            if hr is None:
                return sheets, [], [], ("Строка шапки не найдена (нужны ячейки "
                                        "'Марка' и 'Вес' в первых строках листа)")
            km_vals = sorted({mod.clean_str(v) for v in raw.iloc[hr + 1:, mod.COL_KM]
                              if mod.clean_str(v)})
            dp_vals = sorted({mod.clean_str(v) for v in raw.iloc[hr + 1:, mod.COL_DP]
                              if mod.clean_str(v)})
            return sheets, km_vals, dp_vals, ""
        except Exception as e:
            return [], [], [], f"{type(e).__name__}: {e}"

    def _load_sheets(self):
        if not self.fs_xlsm.text():
            self.log.append_line("⚠️ Сначала выберите .xlsm-файл")
            return
        sheets, _, _, err = self._read_only()
        if err:
            self.log.append_line(f"⚠️ Листы не прочитаны: {err}")
            return
        self.combo_sheet.blockSignals(True)
        self.combo_sheet.clear()
        self.combo_sheet.addItems(sheets)
        saved = self.settings.get("last_sheet", "")
        if saved in sheets:
            self.combo_sheet.setCurrentText(saved)
        self.combo_sheet.blockSignals(False)
        self.log.append_line(f"📄 Листов: {len(sheets)}")
        self._load_km_dp()

    def _load_km_dp(self):
        if not self.fs_xlsm.text() or not self.combo_sheet.currentText():
            return
        self.btn_refresh.setEnabled(False)
        try:
            _, km_vals, dp_vals, err = self._read_only()
        finally:
            self.btn_refresh.setEnabled(True)
        if err:
            self.log.append_line(f"⚠️ КМ/DP не прочитаны: {err}")
            return
        for combo, vals, key in ((self.combo_km, km_vals, "last_km"),
                                 (self.combo_dp, dp_vals, "last_dp")):
            cur = self.settings.get(key, "")
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(vals)
            if cur in vals:
                combo.setCurrentText(cur)
            combo.blockSignals(False)
        self.log.append_line(f"🔎 Уникальных КМ: {len(km_vals)}, DP: {len(dp_vals)}")

    def validate(self) -> str:
        ok, err = self.fs_xlsm.validate()
        if not ok:
            return err
        if not self.combo_sheet.currentText():
            return "Не выбран лист — нажмите «Обновить списки»"
        if not self.combo_km.currentText() or not self.combo_dp.currentText():
            return "Не выбраны КМ/DP — нажмите «Обновить списки»"
        return ""

    def start_task(self):
        mod = load_src("prepare_report")
        # выбранные в GUI значения уходят в ask_choice через предзаполненную очередь
        set_prefilled([("рабочий лист", self.combo_sheet.currentText()),
                       ("Выберите КМ", self.combo_km.currentText()),
                       ("Выберите DP", self.combo_dp.currentText())])
        return (mod.prepare_report, (Path(self.fs_xlsm.text()),), {}, None)

    def on_success(self, result):
        self.settings.update(last_xlsm=self.fs_xlsm.text(),
                             last_sheet=self.combo_sheet.currentText(),
                             last_km=self.combo_km.currentText(),
                             last_dp=self.combo_dp.currentText())
        save_settings(self.settings)
        self.show_result_button(Path(result))
