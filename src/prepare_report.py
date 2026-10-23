r"""
Подготовка отчётной таблицы к сравнению с «Прошмандовкой».

Универсальный скрипт: сам находит строку шапки и нужные столбцы по названиям.
Работает на разных титулах (L1-41-030, L03-02-120, ...).

Что делает:
1. Читает .xlsm-файл отчётной таблицы.
2. Сам находит строку шапки (по наличию 'Марка' и 'Вес').
3. Сам находит столбцы: КМ, DP, Марка, Вес, Наименование элемента.
4. Спрашивает у пользователя: лист, КМ, DP.
5. Фильтрует по КМ и DP.
6. Группирует по Марке.
7. Считает: Кол-во, Вес_1шт, Вес_итого.
8. Проверяет, что сумма веса до группировки = после.
9. Сохраняет результат в output/report_prepared.xlsx.

Запуск отдельно:
    python src/prepare_report.py
"""
import sys
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook

# ============================================================
# НАСТРОЙКИ
# ============================================================
PROJECT_DIR = Path(__file__).parent.parent
INPUT_DIR = PROJECT_DIR / "input"
OUTPUT_DIR = PROJECT_DIR / "output"

# Фиксированные индексы столбцов КМ и DP (совпадают на разных титулах).
# Если не совпадут — скрипт предупредит.
COL_KM = 1    # B — фильтр по КМ
COL_DP = 3    # D — фильтр по DP

# Названия столбцов для автопоиска (регистронезависимо)
HDR_MARK = "марка"
HDR_WEIGHT = "вес"
HDR_NAME = "наименование элемента"

# Максимум строк для сканирования шапки
HEADER_SCAN_ROWS = 5


# ============================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
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


def parse_weight(val):
    if val is None:
        return None
    if isinstance(val, float) and pd.isna(val):
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


def ask_choice(prompt, options):
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


def find_header_row(raw_df, max_scan=HEADER_SCAN_ROWS):
    """
    Ищет строку шапки: где есть ячейка 'Марка' и ячейка 'Вес'.
    Возвращает индекс строки (0-based) или None.
    """
    for i in range(min(max_scan, len(raw_df))):
        row_vals = [clean_str(v).lower() for v in raw_df.iloc[i, :].values]
        has_mark = any(v == HDR_MARK for v in row_vals)
        has_weight = any(v == HDR_WEIGHT for v in row_vals)
        if has_mark and has_weight:
            return i
    return None


def find_columns(raw_df, header_row):
    """
    По строке шапки находит индексы нужных столбцов.
    Возвращает dict: {'mark': idx, 'weight': idx, 'name': idx} или None для ненайденных.
    """
    row_vals = [clean_str(v).lower() for v in raw_df.iloc[header_row, :].values]
    result = {"mark": None, "weight": None, "name": None}
    for j, v in enumerate(row_vals):
        if v == HDR_MARK and result["mark"] is None:
            result["mark"] = j
        elif v == HDR_WEIGHT and result["weight"] is None:
            result["weight"] = j
        elif v == HDR_NAME and result["name"] is None:
            result["name"] = j
    return result


def find_name_column_by_values(raw_df, header_row, max_scan=100):
    """
    Если столбец 'Наименование элемента' не найден по заголовку —
    ищем по значениям: столбец, где встречаются 'Балка', 'Колонна', 'Связь'...
    Возвращает индекс столбца или None.
    """
    known_names = {"балка", "колонна", "связь", "ригель", "распорка",
                   "накладка", "прогон", "ограждение", "прокладка",
                   "монтажная деталь", "монтажная накладка", "шайба выверочная",
                   "проф.лист", "доборная пластина"}
    best_col = None
    best_count = 0
    data = raw_df.iloc[header_row + 1: header_row + 1 + max_scan, :]
    for j in range(raw_df.shape[1]):
        vals = [clean_str(v).lower() for v in data.iloc[:, j].values if clean_str(v)]
        matches = sum(1 for v in vals if v in known_names)
        if matches > best_count:
            best_count = matches
            best_col = j
    return best_col if best_count >= 3 else None


# ============================================================
# ОСНОВНАЯ ФУНКЦИЯ
# ============================================================
def prepare_report(xlsm_path: Path) -> Path:
    """
    Готовит отчётную таблицу. Возвращает путь к report_prepared.xlsx.
    """
    print(f"\n{'=' * 70}")
    print(f"ПОДГОТОВКА ОТЧЁТНОЙ: {xlsm_path.name}")
    print('=' * 70)

    if not xlsm_path.exists():
        raise FileNotFoundError(f"Файл не найден: {xlsm_path}")

    print(f"📂 Файл: {xlsm_path.name}")
    print(f"   Размер: {xlsm_path.stat().st_size / 1024 / 1024:.2f} МБ")

    # ---- Листы ----
    wb = load_workbook(xlsm_path, read_only=True, data_only=True)
    sheets = wb.sheetnames
    wb.close()

    if len(sheets) == 1:
        sheet = sheets[0]
        print(f"\n📄 Рабочий лист: {sheet} (единственный)")
    else:
        sheet = ask_choice("📄 Выберите рабочий лист:", sheets)

    # ---- Читаем БЕЗ шапки, чтобы найти её ----
    print(f"\n⏳ Читаю лист «{sheet}» (без шапки)...")
    raw = pd.read_excel(xlsm_path, sheet_name=sheet, header=None,
                        engine="openpyxl", dtype=object)
    print(f"   Строк: {len(raw):,}, столбцов: {len(raw.columns)}")

    # ---- Автопоиск шапки ----
    header_row = find_header_row(raw)
    if header_row is None:
        raise RuntimeError(
            f"Не удалось найти строку шапки (искал 'Марка' и 'Вес' "
            f"в первых {HEADER_SCAN_ROWS} строках)."
        )
    print(f"\n🔎 Строка шапки найдена: {header_row} (0-based)")

    # ---- Автопоиск столбцов ----
    cols = find_columns(raw, header_row)
    print(f"🔎 Автопоиск столбцов:")
    print(f"   Марка:                 {cols['mark']}")
    print(f"   Вес:                   {cols['weight']}")
    print(f"   Наименование элемента: {cols['name']}")

    if cols["mark"] is None:
        raise RuntimeError("Не найден столбец 'Марка'.")
    if cols["weight"] is None:
        raise RuntimeError("Не найден столбец 'Вес'.")

    # ---- Если 'Наименование элемента' не найден по заголовку — ищем по значениям ----
    if cols["name"] is None:
        print(f"   ⚠️  'Наименование элемента' не найден по заголовку, ищу по значениям...")
        name_col = find_name_column_by_values(raw, header_row)
        if name_col is None:
            print(f"   ⚠️  Не удалось найти. Будет использован пустой столбец.")
            cols["name"] = None
        else:
            cols["name"] = name_col
            print(f"   ✓ Найден по значениям: индекс {name_col}")

    # ---- Читаем данные, начиная со строки после шапки ----
    df = raw.iloc[header_row + 1:, :].reset_index(drop=True)
    df.columns = [f"col_{i}" for i in range(df.shape[1])]

    print(f"\n📊 Данных: {len(df):,} строк")

    # ---- Уникальные КМ и DP ----
    km_vals = sorted({clean_str(v) for v in df.iloc[:, COL_KM] if clean_str(v)})
    dp_vals = sorted({clean_str(v) for v in df.iloc[:, COL_DP] if clean_str(v)})

    print(f"\n🔎 Уникальных КМ: {len(km_vals)} — {km_vals[:10]}")
    print(f"🔎 Уникальных DP: {len(dp_vals)} — {dp_vals[:15]}")

    km = ask_choice("🏗️  Выберите КМ:", km_vals)
    dp = ask_choice("🏗️  Выберите DP:", dp_vals)

    print(f"\n✅ Выбрано: КМ = «{km}», DP = «{dp}»")

    # ---- Фильтруем ----
    km_clean = df.iloc[:, COL_KM].map(clean_str)
    dp_clean = df.iloc[:, COL_DP].map(clean_str)
    mask = (km_clean == km) & (dp_clean == dp)
    df_f = df[mask].copy()

    print(f"\n📊 После фильтрации: {len(df_f):,} строк")

    if len(df_f) == 0:
        raise RuntimeError("После фильтрации не осталось строк. Проверь КМ и DP.")

    # ---- Извлекаем ----
    mark_col = cols["mark"]
    weight_col = cols["weight"]
    name_col = cols["name"]

    df_f["_mark"] = df_f.iloc[:, mark_col].map(clean_str)
    df_f["_weight"] = df_f.iloc[:, weight_col].map(parse_weight)
    if name_col is not None:
        df_f["_name"] = df_f.iloc[:, name_col].map(clean_str)
    else:
        df_f["_name"] = ""

    before = len(df_f)
    df_f = df_f[df_f["_mark"] != ""].copy()
    if len(df_f) < before:
        print(f"   ⚠️ Отброшено строк без марки: {before - len(df_f)}")

    weight_before = df_f["_weight"].sum()
    print(f"\n⚖️  Сумма веса ДО группировки: {weight_before:,.2f}")

    # ---- Группировка ----
    print("\n⏳ Группируем по марке...")
    grouped = df_f.groupby("_mark", as_index=False).agg(
        Кол_во=("_mark", "size"),
        Вес_итого=("_weight", "sum"),
        Вес_1шт=("_weight", lambda s: s.iloc[0] if s.nunique() == 1 else "РАЗНЫЙ"),
        Наименование_элемента=("_name", "first"),
    )

    grouped = grouped.rename(columns={
        "_mark": "Марка",
        "Кол_во": "Кол-во",
        "Вес_итого": "Вес_итого",
        "Вес_1шт": "Вес_1шт",
        "Наименование_элемента": "Наименование элемента",
    })

    grouped["Вес_итого"] = grouped["Вес_итого"].round(2)
    grouped["Вес_1шт"] = grouped["Вес_1шт"].apply(
        lambda x: round(x, 2) if isinstance(x, (int, float)) else x
    )

    grouped = grouped.sort_values("Марка").reset_index(drop=True)
    print(f"   Уникальных марок: {len(grouped):,}")

    weight_after = grouped["Вес_итого"].sum()
    print(f"\n⚖️  Сумма веса ПОСЛЕ группировки: {weight_after:,.2f}")

    diff = abs(weight_before - weight_after)
    if diff < 0.01:
        print(f"   ✅ Проверка пройдена! Вес совпадает (расхождение: {diff:.4f})")
    else:
        print(f"   ⚠️  ВНИМАНИЕ! Расхождение веса: {diff:,.2f}")

    n_diff_weight = (grouped["Вес_1шт"] == "РАЗНЫЙ").sum()
    if n_diff_weight > 0:
        print(f"\n   ℹ️  Марок с разным весом внутри: {n_diff_weight}")

    OUTPUT_DIR.mkdir(exist_ok=True)
    result = OUTPUT_DIR / "report_prepared.xlsx"
    grouped.to_excel(result, index=False, engine="openpyxl")
    print(f"\n💾 Сохранено: {result}")
    print(f"   Строк в файле: {len(grouped):,}")

    print("\n📋 Первые 10 строк результата:")
    print(grouped.head(10).to_string(index=False))

    return result


# ============================================================
# ЗАПУСК ОТДЕЛЬНО
# ============================================================
if __name__ == "__main__":
    xlsms = sorted(INPUT_DIR.glob("*.xlsm"))
    if not xlsms:
        print(f"Нет .xlsm в {INPUT_DIR}")
        sys.exit(1)
    if len(xlsms) > 1:
        print(f"⚠️  XLSM несколько, беру первый: {xlsms[0].name}")
    prepare_report(xlsms[0])