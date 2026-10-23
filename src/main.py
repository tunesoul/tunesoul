r"""
Главный скрипт: меню + полный пайплайн сравнения таблиц.

Меню:
  [1] Парсить PDF → Excel (с проверкой качества)
  [2] Подготовить отчётную таблицу
  [3] Сравнить прошмандовку с отчётной
  [4] Полный цикл (1 → 2 → 3 → перенос DEP → итоговый файл)
  [5] Перенести данные из DEP (.xlsb) в нашу отчётную (.xlsm) → итоговый файл
  [6] Полный цикл без парсинга (готовый .xlsx → сравнение → перенос DEP → итоговый файл)
  [0] Выход

Итоговый файл: <наша>_актуал.xlsx — один файл с листами:
  1. Отчётная (актуал) — с перенесёнными данными
  2. Сравнение (группировка) — сгруппированная отчётная
  3. Прошмандовка — распарсенный PDF
  4. Изменения — лог изменений при переносе
  5. Сравнение — расхождения

Запуск: python src/main.py
"""
import re
import shutil
import sys
from pathlib import Path

import pandas as pd
import openpyxl
import pdfplumber

# Импортируем наши модули
sys.path.insert(0, str(Path(__file__).parent))
from pdf_to_kmd import parse_pdf
from prepare_report import prepare_report
from compare_kmd import compare
from transfer_dep import transfer_dep

# ============================================================
# ПУТИ
# ============================================================
BASE_DIR = Path(__file__).parent.parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"


# ============================================================
# АВТООПРЕДЕЛЕНИЕ ФАЙЛОВ
# ============================================================
def find_files(require_pdf=False, require_xlsx=False, require_xlsm=False, require_xlsb=False):
    pdfs = sorted(INPUT_DIR.glob("*.pdf"))
    xlsxs = sorted(INPUT_DIR.glob("*.xlsx"))
    xlsbs = sorted(INPUT_DIR.glob("*.xlsb"))
    xlsms = sorted(INPUT_DIR.glob("*.xlsm"))

    result = {
        "pdf": pdfs[0] if len(pdfs) == 1 else None,
        "xlsx": xlsxs[0] if len(xlsxs) == 1 else None,
        "xlsb": xlsbs[0] if len(xlsbs) == 1 else None,
        "xlsm": xlsms[0] if len(xlsms) == 1 else None,
    }

    errors = []
    if require_pdf and len(pdfs) != 1:
        errors.append(f"PDF-файлов должно быть ровно 1, найдено {len(pdfs)}")
    if require_xlsx and len(xlsxs) != 1:
        errors.append(f"XLSX-файлов должно быть ровно 1, найдено {len(xlsxs)}")
    if require_xlsm and len(xlsms) != 1:
        errors.append(f"XLSM-файлов должно быть ровно 1, найдено {len(xlsms)}")
    if require_xlsb and len(xlsbs) != 1:
        errors.append(f"XLSB-файлов должно быть ровно 1, найдено {len(xlsbs)}")

    if errors:
        print("❌ Проблемы с файлами в input/:")
        for e in errors:
            print(f"   - {e}")
        return None

    return result


def find_xlsb_input():
    xlsbs = sorted(INPUT_DIR.glob("*.xlsb"))
    return xlsbs[0] if len(xlsbs) == 1 else None


def find_xlsm_input():
    xlsms = sorted(INPUT_DIR.glob("*.xlsm"))
    return xlsms[0] if len(xlsms) == 1 else None


def find_kmd_file():
    kmds = sorted(OUTPUT_DIR.glob("kmd_*.xlsx"))
    if not kmds:
        return None
    return kmds[0]


def find_report_file():
    f = OUTPUT_DIR / "report_prepared.xlsx"
    return f if f.exists() else None


def detect_dp_from_xlsx(xlsx_path: Path) -> str:
    df = pd.read_excel(xlsx_path)
    if "DP" not in df.columns:
        raise RuntimeError(f"В {xlsx_path.name} нет столбца 'DP'")
    dp_values = df["DP"].dropna().astype(str).str.strip()
    dp_values = dp_values[dp_values != ""]
    unique_dps = sorted(set(dp_values))
    if not unique_dps:
        raise RuntimeError(f"Не удалось определить DP")
    return unique_dps[0]


# ============================================================
# СБОРКА ИТОГОВОГО ФАЙЛА
# ============================================================
def add_sheet_from_xlsx(wb_dst, xlsx_path, sheet_name):
    """Читает xlsx и добавляет его данные как новый лист в wb_dst."""
    df = pd.read_excel(xlsx_path, header=None, dtype=object)
    ws = wb_dst.create_sheet(sheet_name)
    for r_idx, row in df.iterrows():
        for c_idx, val in enumerate(row):
            if pd.notna(val):
                ws.cell(row=r_idx + 1, column=c_idx + 1, value=val)


def build_final_xlsx(
    xlsm_path: Path,
    main_sheet_name: str,
    wb_from_transfer,  # openpyxl.Workbook или None
    kmd_path,          # Path или None
    report_prepared_path,  # Path или None
    comparison_path,       # Path или None
) -> Path:
    """
    Собирает итоговый файл <наша>_актуал.xlsx со всеми листами.
    """
    final_path = OUTPUT_DIR / f"{xlsm_path.stem}_актуал.xlsx"

    # ---- 1. Определяем основу ----
    if wb_from_transfer is not None:
        wb = wb_from_transfer
    else:
        # Без переноса — берём исходную .xlsm
        wb = openpyxl.load_workbook(xlsm_path)
        # Оставляем только рабочий лист
        for s in list(wb.sheetnames):
            if s != main_sheet_name:
                del wb[s]
        if main_sheet_name in wb.sheetnames:
            wb[main_sheet_name].title = "Отчётная (актуал)"

    # ---- 2. Если в wb есть лишние листы (кроме рабочего и Изменения) — удаляем ----
    keep = {"Отчётная (актуал)", main_sheet_name, "Изменения"}
    for s in list(wb.sheetnames):
        if s not in keep:
            del wb[s]

    # Переименовываем рабочий лист, если он ещё называется по-старому
    if main_sheet_name in wb.sheetnames and "Отчётная (актуал)" not in wb.sheetnames:
        wb[main_sheet_name].title = "Отчётная (актуал)"

    # ---- 3. Добавляем "Сравнение (группировка)" ----
    if report_prepared_path and report_prepared_path.exists():
        add_sheet_from_xlsx(wb, report_prepared_path, "Сравнение (группировка)")
        print(f"   ✅ Добавлен лист: Сравнение (группировка)")

    # ---- 4. Добавляем "Прошмандовка" ----
    if kmd_path and kmd_path.exists():
        add_sheet_from_xlsx(wb, kmd_path, "Прошмандовка")
        print(f"   ✅ Добавлен лист: Прошмандовка")

    # ---- 5. Добавляем "Сравнение" ----
    if comparison_path and comparison_path.exists():
        add_sheet_from_xlsx(wb, comparison_path, "Сравнение")
        print(f"   ✅ Добавлен лист: Сравнение")

    # ---- 6. Сохраняем ----
    wb.save(final_path)
    print(f"\n💾 Итоговый файл: {final_path}")
    return final_path


# ============================================================
# ПРОВЕРКА PDF
# ============================================================
def check_pdf_is_image(pdf_path: Path) -> bool:
    if not pdf_path or not pdf_path.exists():
        return False
    total_words = 0
    total_marks = 0
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(x_tolerance=1.5, y_tolerance=2.0)
            total_words += len(words)
            for w in words:
                text = w["text"]
                if re.fullmatch(r"DP\d+", text, re.IGNORECASE):
                    total_marks += 1
                elif text.startswith("L") and "-" in text and len(text) > 10:
                    total_marks += 1
    print(f"   📊 Слов в PDF: {total_words}, маркеров: {total_marks}")
    return total_marks == 0


def print_recommendation(pdf_path: Path):
    print()
    print("─" * 70)
    print("РЕКОМЕНДАЦИЯ:")
    print("─" * 70)
    print("1. Откройте PDF и проверьте, есть ли там таблица с марками.")
    print("2. Если PDF — картинка, конвертируйте его в Excel через OCR:")
    print("   👉 https://tools.pdf24.org/en/pdf-to-excel")
    print("3. Скачанный Excel положите в папку input/ с именем:")
    if pdf_path:
        print(f"   👉 {pdf_path.stem}.xlsx")
    print("4. Запустите пункт [6] — «Полный цикл без парсинга».")
    print("─" * 70)


# ============================================================
# ДЕЙСТВИЯ
# ============================================================
def action_parse_pdf():
    """[1] Парсить PDF → Excel + проверка качества."""
    print("\n" + "=" * 70)
    print("ДЕЙСТВИЕ: ПАРСИНГ PDF")
    print("=" * 70)
    files = find_files(require_pdf=True)
    if files is None:
        return None
    pdf_path = files["pdf"]
    print(f"\n⏳ Предварительная проверка PDF...")
    if check_pdf_is_image(pdf_path):
        print("\n⚠️  PDF выглядит как КАРТИНКА.")
        print_recommendation(pdf_path)
        return None
    try:
        kmd_path = parse_pdf(pdf_path)
    except Exception as e:
        print(f"\n❌ Ошибка парсинга PDF: {e}")
        print_recommendation(pdf_path)
        return None
    return kmd_path


def action_prepare_report():
    """[2] Подготовить отчётную."""
    print("\n" + "=" * 70)
    print("ДЕЙСТВИЕ: ПОДГОТОВКА ОТЧЁТНОЙ")
    print("=" * 70)
    files = find_files(require_xlsm=True)
    if files is None:
        return None
    try:
        return prepare_report(files["xlsm"])
    except Exception as e:
        print(f"\n❌ Ошибка подготовки: {e}")
        return None


def action_compare():
    """[3] Сравнить прошмандовку с отчётной."""
    print("\n" + "=" * 70)
    print("ДЕЙСТВИЕ: СРАВНЕНИЕ")
    print("=" * 70)
    kmd_path = find_kmd_file()
    report_path = find_report_file()
    if kmd_path is None:
        print("❌ Не найден kmd_*.xlsx. Запустите [1].")
        return None
    if report_path is None:
        print("❌ Не найден report_prepared.xlsx. Запустите [2].")
        return None
    try:
        return compare(kmd_path, report_path)
    except Exception as e:
        print(f"\n❌ Ошибка сравнения: {e}")
        return None


def action_full_cycle():
    """[4] Полный цикл с парсингом PDF."""
    print("\n" + "=" * 70)
    print("ДЕЙСТВИЕ: ПОЛНЫЙ ЦИКЛ (с парсингом PDF)")
    print("=" * 70)

    files = find_files(require_pdf=True, require_xlsm=True)
    if files is None:
        return None

    pdf_path = files["pdf"]
    xlsm_path = files["xlsm"]

    # ---- ШАГ 1: парсинг ----
    print(f"\n--- ШАГ 1/4: ПАРСИНГ PDF ---")
    print(f"⏳ Проверка PDF...")
    if check_pdf_is_image(pdf_path):
        print("\n⚠️  PDF — КАРТИНКА.")
        print_recommendation(pdf_path)
        return None
    try:
        kmd_path = parse_pdf(pdf_path)
    except Exception as e:
        print(f"\n❌ Ошибка парсинга: {e}")
        return None

    # ---- ШАГ 2: подготовка отчётной ----
    print(f"\n--- ШАГ 2/4: ПОДГОТОВКА ОТЧЁТНОЙ ---")
    try:
        report_path = prepare_report(xlsm_path)
    except Exception as e:
        print(f"\n❌ Ошибка подготовки: {e}")
        return None

    # ---- ШАГ 3: сравнение ----
    print(f"\n--- ШАГ 3/4: СРАВНЕНИЕ ---")
    try:
        comparison_path = compare(kmd_path, report_path)
    except Exception as e:
        print(f"\n❌ Ошибка сравнения: {e}")
        comparison_path = None

    # ---- ШАГ 4: перенос из DEP ----
    print(f"\n--- ШАГ 4/4: ПЕРЕНОС ДАННЫХ ИЗ DEP ---")
    xlsb_path = find_xlsb_input()
    wb_from_transfer = None
    main_sheet_name = None

    if xlsb_path is None:
        print("⚠️  Нет .xlsb (DEP) в input/. Пропускаю перенос.")
    else:
        try:
            wb_from_transfer, main_sheet_name, _ = transfer_dep(xlsb_path, xlsm_path)
        except Exception as e:
            print(f"\n❌ Ошибка переноса: {e}")
            import traceback
            traceback.print_exc()

    # Если переноса не было — нужно спросить лист вручную
    if main_sheet_name is None:
        wb_tmp = openpyxl.load_workbook(xlsm_path, read_only=True)
        sheets = wb_tmp.sheetnames
        wb_tmp.close()
        print(f"\n📄 Лист отчётной: {sheets}")
        main_sheet_name = sheets[0]  # по умолчанию первый

    # ---- СБОРКА ИТОГОВОГО ФАЙЛА ----
    print(f"\n--- СБОРКА ИТОГОВОГО ФАЙЛА ---")
    try:
        final_path = build_final_xlsx(
            xlsm_path=xlsm_path,
            main_sheet_name=main_sheet_name,
            wb_from_transfer=wb_from_transfer,
            kmd_path=kmd_path,
            report_prepared_path=report_path,
            comparison_path=comparison_path,
        )
    except Exception as e:
        print(f"\n❌ Ошибка сборки: {e}")
        import traceback
        traceback.print_exc()
        return None

    # ---- Удаляем промежуточные файлы ----
    cleanup_intermediate()

    print(f"\n{'=' * 70}")
    print("✅ ПОЛНЫЙ ЦИКЛ ЗАВЕРШЁН!")
    print(f"📄 Итоговый файл: {final_path}")
    print('=' * 70)
    return final_path


def action_prepare_dep():
    """[5] Только перенос из DEP."""
    print("\n" + "=" * 70)
    print("ДЕЙСТВИЕ: ПЕРЕНОС ДАННЫХ ИЗ DEP")
    print("=" * 70)

    xlsb = find_xlsb_input()
    if xlsb is None:
        print("❌ Не найден .xlsb (DEP) в input/")
        return None
    xlsm = find_xlsm_input()
    if xlsm is None:
        print("❌ Не найден .xlsm (отчётная) в input/")
        return None

    try:
        wb, main_sheet_name, _ = transfer_dep(xlsb, xlsm)
    except Exception as e:
        print(f"\n❌ Ошибка переноса: {e}")
        import traceback
        traceback.print_exc()
        return None

    if wb is None:
        return None

    # Собираем итоговый файл (только с листами, которые доступны)
    print(f"\n--- СБОРКА ИТОГОВОГО ФАЙЛА ---")
    final_path = build_final_xlsx(
        xlsm_path=xlsm,
        main_sheet_name=main_sheet_name,
        wb_from_transfer=wb,
        kmd_path=None,             # нет парсинга
        report_prepared_path=None, # нет группировки
        comparison_path=None,      # нет сравнения
    )

    print(f"\n{'=' * 70}")
    print("✅ ПЕРЕНОС ЗАВЕРШЁН!")
    print(f"📄 Итоговый файл: {final_path}")
    print('=' * 70)
    return final_path


def action_full_cycle_no_parse():
    """[6] Полный цикл без парсинга."""
    print("\n" + "=" * 70)
    print("ДЕЙСТВИЕ: ПОЛНЫЙ ЦИКЛ БЕЗ ПАРСИНГА")
    print("=" * 70)

    files = find_files(require_xlsx=True, require_xlsm=True)
    if files is None:
        return None

    xlsm_path = files["xlsm"]

    # ---- ШАГ 1: копирование готового xlsx ----
    print(f"\n--- ШАГ 1/4: КОПИРОВАНИЕ ГОТОВОЙ ПРОШМАНДОВКИ ---")
    print(f"📂 Читаю: {files['xlsx'].name}")
    try:
        dp = detect_dp_from_xlsx(files["xlsx"])
    except Exception as e:
        print(f"\n❌ Ошибка чтения xlsx: {e}")
        return None

    target = OUTPUT_DIR / f"kmd_{dp.lower()}.xlsx"
    shutil.copy2(files["xlsx"], target)
    print(f"   ✅ Скопировано: {target}")
    kmd_path = target

    # ---- ШАГ 2: подготовка отчётной ----
    print(f"\n--- ШАГ 2/4: ПОДГОТОВКА ОТЧЁТНОЙ ---")
    try:
        report_path = prepare_report(xlsm_path)
    except Exception as e:
        print(f"\n❌ Ошибка подготовки: {e}")
        return None

    # ---- ШАГ 3: сравнение ----
    print(f"\n--- ШАГ 3/4: СРАВНЕНИЕ ---")
    try:
        comparison_path = compare(kmd_path, report_path)
    except Exception as e:
        print(f"\n❌ Ошибка сравнения: {e}")
        comparison_path = None

    # ---- ШАГ 4: перенос из DEP ----
    print(f"\n--- ШАГ 4/4: ПЕРЕНОС ИЗ DEP ---")
    xlsb_path = find_xlsb_input()
    wb_from_transfer = None
    main_sheet_name = None

    if xlsb_path is None:
        print("⚠️  Нет .xlsb. Пропускаю перенос.")
    else:
        try:
            wb_from_transfer, main_sheet_name, _ = transfer_dep(xlsb_path, xlsm_path)
        except Exception as e:
            print(f"\n❌ Ошибка переноса: {e}")
            import traceback
            traceback.print_exc()

    if main_sheet_name is None:
        wb_tmp = openpyxl.load_workbook(xlsm_path, read_only=True)
        sheets = wb_tmp.sheetnames
        wb_tmp.close()
        main_sheet_name = sheets[0]

    # ---- СБОРКА ----
    print(f"\n--- СБОРКА ИТОГОВОГО ФАЙЛА ---")
    final_path = build_final_xlsx(
        xlsm_path=xlsm_path,
        main_sheet_name=main_sheet_name,
        wb_from_transfer=wb_from_transfer,
        kmd_path=kmd_path,
        report_prepared_path=report_path,
        comparison_path=comparison_path,
    )

    cleanup_intermediate()

    print(f"\n{'=' * 70}")
    print("✅ ПОЛНЫЙ ЦИКЛ (БЕЗ ПАРСИНГА) ЗАВЕРШЁН!")
    print(f"📄 Итоговый файл: {final_path}")
    print('=' * 70)
    return final_path


def cleanup_intermediate():
    """Удаляет промежуточные файлы (kmd_*.xlsx, report_prepared.xlsx, comparison.xlsx)."""
    patterns = ["kmd_*.xlsx", "report_prepared.xlsx", "comparison.xlsx"]
    for pattern in patterns:
        for f in OUTPUT_DIR.glob(pattern):
            try:
                f.unlink()
                print(f"   🗑️  Удалён: {f.name}")
            except Exception as e:
                print(f"   ⚠️  Не удалось удалить {f.name}: {e}")


# ============================================================
# МЕНЮ
# ============================================================
def print_menu():
    print()
    print("=" * 70)
    print("СРАВНЕНИЕ ТАБЛИЦ")
    print("=" * 70)

    pdfs = sorted(INPUT_DIR.glob("*.pdf"))
    xlsxs = sorted(INPUT_DIR.glob("*.xlsx"))
    xlsbs = sorted(INPUT_DIR.glob("*.xlsb"))
    xlsms = sorted(INPUT_DIR.glob("*.xlsm"))

    print("Состояние input/:")
    print(f"  PDF:  {'✅ ' + pdfs[0].name if len(pdfs) == 1 else ('❌ нет' if not pdfs else f'⚠️ {len(pdfs)} шт.')}")
    print(f"  XLSX: {'✅ ' + xlsxs[0].name if len(xlsxs) == 1 else ('❌ нет' if not xlsxs else f'⚠️ {len(xlsxs)} шт.')}")
    print(f"  XLSM: {'✅ ' + xlsms[0].name if len(xlsms) == 1 else ('❌ нет' if not xlsms else f'⚠️ {len(xlsms)} шт.')}")
    print(f"  XLSB: {'✅ ' + xlsbs[0].name if len(xlsbs) == 1 else ('❌ нет' if not xlsbs else f'⚠️ {len(xlsbs)} шт.')}")

    finals = sorted(OUTPUT_DIR.glob("*_актуал.xlsx"))
    if finals:
        print("\nСостояние output/:")
        for f in finals:
            print(f"  📄 {f.name}")

    print()
    print("[1] Парсить PDF → Excel (с проверкой качества)")
    print("[2] Подготовить отчётную таблицу")
    print("[3] Сравнить прошмандовку с отчётной")
    print("[4] Полный цикл (1 → 2 → 3 → перенос DEP → итоговый файл)")
    print("[5] Перенести данные из DEP → итоговый файл")
    print("[6] Полный цикл БЕЗ парсинга (готовый .xlsx → сравнение → перенос → итоговый файл)")
    print("[0] Выход")
    print()


def main():
    actions = {
        "1": action_parse_pdf,
        "2": action_prepare_report,
        "3": action_compare,
        "4": action_full_cycle,
        "5": action_prepare_dep,
        "6": action_full_cycle_no_parse,
    }

    while True:
        print_menu()
        choice = input("Выберите действие: ").strip()

        if choice == "0":
            print("Выход.")
            break

        action = actions.get(choice)
        if action is None:
            print("⚠️  Некорректный ввод.")
            continue

        try:
            action()
        except KeyboardInterrupt:
            print("\n⚠️  Прервано.")
        except Exception as e:
            print(f"\n❌ Непредвиденная ошибка: {e}")
            import traceback
            traceback.print_exc()

        input("\nНажмите Enter, чтобы вернуться в меню...")


if __name__ == "__main__":
    main()