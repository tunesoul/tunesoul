"""Экран «Перенос из DEP» (ТЗ §5.5, ШАГ 6).

transfer_dep возвращает Workbook в памяти; итоговый файл собирается через
build_final_xlsx, импортируемый из src/main.py (не копируется!).
Листы/КМ/DP читаются теми же функциями и константами src-модулей.
"""
from pathlib import Path

from PySide6.QtWidgets import QLabel, QComboBox, QPushButton

from core.patches import set_prefilled
from core.settings import save_settings
from core.src_loader import load_src
from ui.screens.base_screen import BaseScreen
from ui.widgets import Card, FileSelect


class ScreenTransfer(BaseScreen):
    action_label = "Перенести"
    action_tip = "DEP (.xlsb) → листы актуальной отчётной + итоговый файл (Ctrl+Enter)"

    def card_header(self) -> tuple[str, str]:
        return ("Перенос из DEP", ".xlsb (DEP) + .xlsm (наша) → output/<наша>_актуал.xlsx")

    def build_form(self, card: Card):
        self.fs_xlsb = FileSelect(".xlsb", "Путь к DEP-файлу (.xlsb)")
        self.fs_xlsm = FileSelect(".xlsm", "Путь к нашей отчётной (.xlsm)")
        last_xlsb = self.settings.get("last_xlsb", "")
        if last_xlsb and Path(last_xlsb).is_file():
            self.fs_xlsb.set_text(last_xlsb)
        last_xlsm = self.settings.get("last_xlsm", "")
        if last_xlsm and Path(last_xlsm).is_file():
            self.fs_xlsm.set_text(last_xlsm)

        self.combo_dep_sheet = QComboBox()
        self.combo_rep_sheet = QComboBox()
        self.combo_km = QComboBox()
        self.combo_dp = QComboBox()
        self.btn_refresh = QPushButton("Обновить списки (листы / КМ / DP)")
        self.btn_refresh.clicked.connect(self._load_lists)

        for lbl, w in (("DEP-файл (.xlsb):", self.fs_xlsb),
                       ("Наша отчётная (.xlsm):", self.fs_xlsm),
                       ("Лист DEP:", self.combo_dep_sheet),
                       ("Лист отчётной:", self.combo_rep_sheet),
                       ("КМ (из DEP):", self.combo_km),
                       ("DP (из DEP):", self.combo_dp),
                       ("", self.btn_refresh)):
            if lbl:
                card.add(QLabel(lbl))
            card.add(w)

    # -------- чтение вариантов в GUI-потоке (те же функции, что и при выполнении) --------
    def _load_lists(self):
        ok1, err1 = self.fs_xlsb.validate()
        ok2, err2 = self.fs_xlsm.validate()
        if not (ok1 and ok2):
            self.log.append_line(f"⚠️ {err1 or err2}")
            return
        try:
            import pyxlsb, openpyxl, pandas as pd
            td = load_src("transfer_dep")
            with pyxlsb.open_workbook(Path(self.fs_xlsb.text())) as wb:
                dep_sheets = list(wb.sheets)
            wb2 = openpyxl.load_workbook(Path(self.fs_xlsm.text()), read_only=True,
                                         data_only=True)
            rep_sheets = list(wb2.sheetnames)
            wb2.close()
            dep_raw = pd.read_excel(self.fs_xlsb.text(), sheet_name=dep_sheets[0],
                                    header=None, engine="pyxlsb", dtype=object)
            dep_data = dep_raw.iloc[2:]
            km_vals = sorted({td.clean_str(v) for v in dep_data.iloc[:, td.DEP_COL_KM]
                              if td.clean_str(v)})
            dp_vals = sorted({td.clean_str(v) for v in dep_data.iloc[:, td.DEP_COL_DP]
                              if td.clean_str(v)})
        except Exception as e:
            self.log.append_line(f"⚠️ Не удалось прочитать списки: {type(e).__name__}: {e}")
            return
        self._fill(self.combo_dep_sheet, dep_sheets)
        self._fill(self.combo_rep_sheet, rep_sheets)
        self._fill(self.combo_km, km_vals)
        self._fill(self.combo_dp, dp_vals)
        self.log.append_line(f"📄 Листов DEP: {len(dep_sheets)}, листов нашей: {len(rep_sheets)}, "
                             f"КМ: {len(km_vals)}, DP: {len(dp_vals)}")

    @staticmethod
    def _fill(combo: QComboBox, values: list):
        combo.blockSignals(True)
        combo.clear()
        combo.addItems([str(v) for v in values])
        combo.blockSignals(False)

    def validate(self) -> str:
        for fs in (self.fs_xlsb, self.fs_xlsm):
            ok, err = fs.validate()
            if not ok:
                return err
        for c, name in ((self.combo_dep_sheet, "лист DEP"), (self.combo_rep_sheet,
                          "лист отчётной"), (self.combo_km, "КМ"), (self.combo_dp, "DP")):
            if not c.currentText():
                return f"Не выбран {name} — нажмите «Обновить списки»"
        return ""

    def start_task(self):
        td = load_src("transfer_dep")
        mn = load_src("main")   # build_final_xlsx берём из src/main.py (по ТЗ — не копировать!)
        xlsb, xlsm = Path(self.fs_xlsb.text()), Path(self.fs_xlsm.text())
        set_prefilled([("лист DEP", self.combo_dep_sheet.currentText()),
                       ("лист отчётной", self.combo_rep_sheet.currentText()),
                       ("Выберите КМ", self.combo_km.currentText()),
                       ("Выберите DP", self.combo_dp.currentText())])

        def run():
            wb, rep_sheet, n_changes = td.transfer_dep(xlsb, xlsm)
            if wb is None:
                raise RuntimeError("transfer_dep вернул пустой результат")
            print(f"\n--- СБОРКА ИТОГОВОГО ФАЙЛА ---")
            return mn.build_final_xlsx(
                xlsm_path=xlsm, main_sheet_name=rep_sheet, wb_from_transfer=wb,
                kmd_path=None, report_prepared_path=None, comparison_path=None)

        return (run, (), {}, None)

    def on_success(self, result):
        self.settings["last_xlsb"] = self.fs_xlsb.text()
        save_settings(self.settings)
        self.show_result_button(Path(result))
