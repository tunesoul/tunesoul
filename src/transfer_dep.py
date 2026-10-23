r"""
Перенос данных из DEP-таблицы (.xlsb) в нашу отчётную (.xlsm).

Возвращает openpyxl.Workbook (в памяти), НЕ сохраняет файл.
Сохранение делает main.py в конце полного цикла.

Запуск: python src/transfer_dep.py (сохраняет как отдельный файл для отладки)
"""
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import pyxlsb

PROJECT_DIR = Path(__file__).parent.parent
INPUT_DIR = PROJECT_DIR / "input"
OUTPUT_DIR = PROJECT_DIR / "output"


# ---------- Индексы столбцов (0-based) ----------
# В DEP (.xlsb)
DEP_COL_KM = 1
DEP_COL_DP = 3
DEP_COL_MARK = 4
DEP_COL_WEIGHT = 5
DEP_COL_IGS = 7
DEP_COL_DATE_IN = 8
DEP_COL_DATE_ACC = 9
DEP_COL_RFI = 11
DEP_COL_RFI_DATE = 12
DEP_COL_SDANO_DATE = 13
DEP_COL_GUID = 22

# В нашей (.xlsm)
REP_COL_KM = 1
REP_COL_DP = 3
REP_COL_MARK = 6
REP_COL_WEIGHT = 7
REP_COL_RFI = 10
REP_COL_RFI_DATE = 11
REP_COL_SDANO_DATE = 12
REP_COL_IGS = 27
REP_COL_IGS_DATE_IN = 28
REP_COL_IGS_DATE_ACC = 29
REP_COL_GUID = 39
REP_COL_GUID_DEP = 40

# ---------- Служебные константы ----------
# FIX 1.5: заголовок таблицы — вторая строка Excel (было захардкожено row=2 в коде).
HEADER_ROW = 2
# FIX 1.1: данные в .xlsm начинаются со 3-й строки Excel.
# После .iloc[2:] индекс первой строки данных = 2, значит Excel row = idx + 1.
DATA_ROW_OFFSET = 1

DATE_FORMAT = "DD.MM.YYYY"

YELLOW_FILL = PatternFill("solid", fgColor="FFFF00")


# ============================================================
# УТИЛИТЫ
# ============================================================
def clean_str(val):
    if val is None:
        return ""
    if isinstance(val, float) and pd.isna(val):
        return ""
    s = str(val)
    s = s.replace("\u00a0", " ").replace("\t", " ")
    s = s.replace("\u200b", "").replace("\ufeff", "")
    return s.strip()


def _isna(x):
    """FIX 1.4: универсальная проверка на NA (работает с pd.NA, pd.NaT, np.nan, None)."""
    if x is None:
        return True
    try:
        return bool(pd.isna(x))
    except (TypeError, ValueError):
        return False


def parse_weight(val):
    if _isna(val):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = clean_str(val)
    if not s:
        return None
    s = s.replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def excel_date_to_datetime(n):
    """FIX 1.2: поддержка строк + проверка диапазона."""
    if _isna(n):
        return None
    if isinstance(n, datetime):
        return n
    if isinstance(n, pd.Timestamp):
        return n.to_pydatetime()
    if isinstance(n, (int, float)):
        # Диапазон, в котором имеет смысл Excel-серийная дата.
        # ~1899-12-31 .. 2173 — отсекаем заведомо мусорные значения.
        if not (1 < float(n) < 100000):
            return None
        try:
            return datetime(1899, 12, 30) + timedelta(days=float(n))
        except (ValueError, OverflowError):
            return None
    if isinstance(n, str):
        s = n.strip()
        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%y"):
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                continue
        return None
    return None


def norm_guid(s):
    """FIX 1.3: нормализация GUID для сравнения ({ABC} == {abc} == ABC)."""
    s = clean_str(s)
    if s.startswith("{") and s.endswith("}"):
        s = s[1:-1]
    return s.strip().upper()


def ask_choice(prompt, options):
    """FIX 1.8: не зависаем, если options пуст."""
    if not options:
        raise RuntimeError(f"Нет вариантов для выбора: {prompt}")
    print(f"\n{prompt}")
    for i, opt in enumerate(options, 1):
        print(f"  [{i}] {opt}")
    while True:
        raw = input("Введите номер: ").strip()
        if raw.isdigit():
            idx = int(raw) - 1
            if 0 <= idx < len(options):
                return options[idx]
        print("  ⚠️ Некорректный ввод, попробуйте ещё раз.")


def find_xlsb():
    xlsbs = sorted(INPUT_DIR.glob("*.xlsb"))
    if len(xlsbs) != 1:
        print(f"❌ XLSB-файлов должно быть ровно 1, найдено {len(xlsbs)}")
        return None
    return xlsbs[0]


def find_xlsm():
    xlsms = sorted(INPUT_DIR.glob("*.xlsm"))
    if len(xlsms) != 1:
        print(f"❌ XLSM-файлов должно быть ровно 1, найдено {len(xlsms)}")
        return None
    return xlsms[0]


# ============================================================
# ОСНОВНАЯ ЛОГИКА
# ============================================================
def transfer_dep(xlsb_path: Path, xlsm_path: Path):
    """
    Переносит данные из DEP в нашу.
    Возвращает (openpyxl.Workbook, rep_sheet_name, changes_count).
    Workbook — в памяти, НЕ сохранён.
    """

    # ---- 1. Список листов DEP ----
    with pyxlsb.open_workbook(xlsb_path) as wb:
        dep_sheets = wb.sheets

    print(f"\n📂 DEP-файл: {xlsb_path.name}")
    print(f"   Листов: {len(dep_sheets)}")
    for i, s in enumerate(dep_sheets):
        print(f"     [{i}] {s}")

    dep_sheet = ask_choice("📄 Выберите лист DEP:", dep_sheets)

    # ---- 2. Список листов нашей ----
    wb_rep = openpyxl.load_workbook(xlsm_path, read_only=True, data_only=True)
    rep_sheets = wb_rep.sheetnames
    wb_rep.close()

    print(f"\n📂 Наша отчётная: {xlsm_path.name}")
    rep_sheet = ask_choice("📄 Выберите лист отчётной:", rep_sheets)

    # ---- 3. Читаем DEP ----
    print(f"\n⏳ Читаю DEP (лист «{dep_sheet}»)...")
    dep_raw = pd.read_excel(xlsb_path, sheet_name=dep_sheet, header=None,
                            engine="pyxlsb", dtype=object)
    print(f"   Строк: {len(dep_raw)}, столбцов: {len(dep_raw.columns)}")

    dep_data = dep_raw.iloc[2:].copy()
    dep_data.columns = [f"dep_{i}" for i in range(dep_data.shape[1])]

    # ---- 4. Читаем нашу ----
    print(f"\n⏳ Читаю отчётную (лист «{rep_sheet}»)...")
    rep_raw = pd.read_excel(xlsm_path, sheet_name=rep_sheet, header=None,
                            engine="openpyxl", dtype=object)
    print(f"   Строк: {len(rep_raw)}, столбцов: {len(rep_raw.columns)}")

    rep_data = rep_raw.iloc[2:].copy()
    rep_data.columns = [f"rep_{i}" for i in range(rep_data.shape[1])]

    # ---- 5. Спрашиваем КМ и DP ----
    dep_km = sorted({clean_str(v) for v in dep_data.iloc[:, DEP_COL_KM] if clean_str(v)})
    dep_dp = sorted({clean_str(v) for v in dep_data.iloc[:, DEP_COL_DP] if clean_str(v)})

    print(f"\n🔎 В DEP найдено КМ: {dep_km}")
    print(f"🔎 В DEP найдено DP: {dep_dp}")

    km = ask_choice("🏗️  Выберите КМ:", dep_km)
    dp = ask_choice("🏗️  Выберите DP:", dep_dp)

    print(f"\n✅ Выбрано: КМ = «{km}», DP = «{dp}»")

    # ---- 6. Фильтруем (FIX 1.1: БЕЗ reset_index, сохраняем исходные метки) ----
    def filter_df(df, col_km, col_dp):
        km_vals = df.iloc[:, col_km].map(clean_str)
        dp_vals = df.iloc[:, col_dp].map(clean_str)
        return df[(km_vals == km) & (dp_vals == dp)].copy()

    dep_f = filter_df(dep_data, DEP_COL_KM, DEP_COL_DP)
    rep_f = filter_df(rep_data, REP_COL_KM, REP_COL_DP)

    print(f"\n📊 DEP после фильтра: {len(dep_f)}")
    print(f"📊 Наша после фильтра: {len(rep_f)}")

    if len(dep_f) == 0 or len(rep_f) == 0:
        print("❌ Одна из таблиц пуста после фильтрации. Прекращаю.")
        return None, None, 0

    # FIX 1.9: проверяем, что выбранный лист соответствует выбранным КМ/DP
    if len(rep_f) != len(rep_data):
        first_idx = rep_f.index[0]
        print(f"   ℹ️  Первая подходящая строка отчётной — Excel row {first_idx + 1} "
              f"(индекс данных {first_idx})")

    # ---- 7. Карта DEP: (mark, weight) -> [список меток индекса] ----
    dep_groups = {}
    for idx, row in dep_f.iterrows():
        mark = clean_str(row.iloc[DEP_COL_MARK])
        weight = parse_weight(row.iloc[DEP_COL_WEIGHT])
        key = (mark, weight)
        dep_groups.setdefault(key, []).append(idx)

    # ---- 8. Сопоставление (порядковое, как в ТЗ) ----
    changes = []
    rep_groups_used = {}
    transfer_results = []

    for rep_idx, rep_row in rep_f.iterrows():
        rep_mark = clean_str(rep_row.iloc[REP_COL_MARK])
        rep_weight = parse_weight(rep_row.iloc[REP_COL_WEIGHT])
        key = (rep_mark, rep_weight)

        used = rep_groups_used.get(key, 0)
        dep_indices = dep_groups.get(key, [])

        if used >= len(dep_indices):
            transfer_results.append({
                "rep_idx": rep_idx,
                "status": "not_found",
                "dep_idx": None,
            })
            continue

        dep_idx = dep_indices[used]
        rep_groups_used[key] = used + 1

        dep_row = dep_f.loc[dep_idx]

        dep_igs = dep_row.iloc[DEP_COL_IGS]
        dep_date_in = excel_date_to_datetime(dep_row.iloc[DEP_COL_DATE_IN])
        dep_date_acc = excel_date_to_datetime(dep_row.iloc[DEP_COL_DATE_ACC])
        dep_rfi = dep_row.iloc[DEP_COL_RFI]
        dep_rfi_date = excel_date_to_datetime(dep_row.iloc[DEP_COL_RFI_DATE])
        dep_sdano_date = excel_date_to_datetime(dep_row.iloc[DEP_COL_SDANO_DATE])
        dep_guid = clean_str(dep_row.iloc[DEP_COL_GUID])

        rep_guid = clean_str(rep_row.iloc[REP_COL_GUID])

        transfer_results.append({
            "rep_idx": rep_idx,
            "status": "ok",
            "dep_idx": dep_idx,
            "dep_igs": dep_igs,
            "dep_date_in": dep_date_in,
            "dep_date_acc": dep_date_acc,
            "dep_rfi": dep_rfi,
            "dep_rfi_date": dep_rfi_date,
            "dep_sdano_date": dep_sdano_date,
            "dep_guid": dep_guid,
            "rep_guid": rep_guid,
            # FIX 1.3: сравниваем нормализованные GUID
            "guid_match": norm_guid(dep_guid) == norm_guid(rep_guid),
        })

    # ---- 9. Открываем нашу книгу ----
    print("\n⏳ Открываю нашу книгу...")
    wb = openpyxl.load_workbook(xlsm_path)
    ws = wb[rep_sheet]

    # FIX 1.5: заголовок пишем в константную строку
    ws.cell(row=HEADER_ROW, column=REP_COL_GUID_DEP + 1, value="Уникальный код (DEP)")

    # FIX 1.6: захватываем max_column ДО цикла закраски,
    # чтобы не читать его при каждой итерации и не зависеть от новых ячеек.
    max_col = ws.max_column

    n_ok = 0
    n_not_found = 0
    n_guid_diff = 0

    for tr in transfer_results:
        rep_idx = tr["rep_idx"]
        # FIX 1.1: правильный номер строки Excel
        excel_row = rep_idx + DATA_ROW_OFFSET

        if tr["status"] == "not_found":
            for col_idx in range(1, max_col + 2):
                ws.cell(row=excel_row, column=col_idx).fill = YELLOW_FILL
            n_not_found += 1
            changes.append({
                "Строка": excel_row,
                "Марка": clean_str(rep_f.loc[rep_idx, f"rep_{REP_COL_MARK}"]),
                "Вес": rep_f.loc[rep_idx, f"rep_{REP_COL_WEIGHT}"],
                "Действие": "Не найдено в DEP",
                "Подробности": "",
            })
            continue

        # --- Числовые/текстовые поля ---
        if not _isna(tr["dep_igs"]):
            ws.cell(row=excel_row, column=REP_COL_IGS + 1, value=tr["dep_igs"])
        if not _isna(tr["dep_rfi"]):
            ws.cell(row=excel_row, column=REP_COL_RFI + 1, value=tr["dep_rfi"])

        # --- Даты: пишем + формат FIX 1.7 ---
        if tr["dep_date_in"] is not None:
            c = ws.cell(row=excel_row, column=REP_COL_IGS_DATE_IN + 1, value=tr["dep_date_in"])
            c.number_format = DATE_FORMAT
        if tr["dep_date_acc"] is not None:
            c = ws.cell(row=excel_row, column=REP_COL_IGS_DATE_ACC + 1, value=tr["dep_date_acc"])
            c.number_format = DATE_FORMAT
        if tr["dep_rfi_date"] is not None:
            c = ws.cell(row=excel_row, column=REP_COL_RFI_DATE + 1, value=tr["dep_rfi_date"])
            c.number_format = DATE_FORMAT
        if tr["dep_sdano_date"] is not None:
            c = ws.cell(row=excel_row, column=REP_COL_SDANO_DATE + 1, value=tr["dep_sdano_date"])
            c.number_format = DATE_FORMAT

        n_ok += 1

        if not tr["guid_match"]:
            if tr["dep_guid"]:
                ws.cell(row=excel_row, column=REP_COL_GUID_DEP + 1, value=tr["dep_guid"])
            ws.cell(row=excel_row, column=REP_COL_GUID_DEP + 1).fill = YELLOW_FILL
            n_guid_diff += 1
            changes.append({
                "Строка": excel_row,
                "Марка": clean_str(rep_f.loc[rep_idx, f"rep_{REP_COL_MARK}"]),
                "Вес": rep_f.loc[rep_idx, f"rep_{REP_COL_WEIGHT}"],
                "Действие": "GUID отличается",
                "Подробности": f"Наш: {tr['rep_guid']} / DEP: {tr['dep_guid']}",
            })

    print(f"\n📊 Результаты переноса:")
    print(f"   ✅ Перенесено: {n_ok}")
    print(f"   ⚠️  Не найдено в DEP: {n_not_found}")
    print(f"   ⚠️  GUID отличается: {n_guid_diff}")

    # ---- 10. Лист "Изменения" ----
    if changes:
        ws_changes = wb.create_sheet("Изменения")
        ws_changes.append(["Строка", "Марка", "Вес", "Действие", "Подробности"])
        for ch in changes:
            ws_changes.append([ch["Строка"], ch["Марка"], ch["Вес"], ch["Действие"], ch["Подробности"]])
        print(f"\n📋 Лист «Изменения»: {len(changes)} записей")

    # НЕ сохраняем — возвращаем в память
    return wb, rep_sheet, len(changes)


# ============================================================
# ЗАПУСК ОТДЕЛЬНО (для отладки)
# ============================================================
def main():
    print("=" * 70)
    print("ПЕРЕНОС ДАННЫХ ИЗ DEP В НАШУ ОТЧЁТНУЮ")
    print("=" * 70)

    xlsb = find_xlsb()
    if xlsb is None:
        return
    xlsm = find_xlsm()
    if xlsm is None:
        return

    wb, sheet, changes = transfer_dep(xlsb, xlsm)
    if wb is None:
        return
    out = OUTPUT_DIR / f"{xlsm.stem}_актуал_ОТЛАДКА.xlsx"
    wb.save(out)
    print(f"\n💾 Сохранено (отладка): {out}")


if __name__ == "__main__":
    main()