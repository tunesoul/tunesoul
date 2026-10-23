r"""
Парсер Прошмандовки (PDF) в Excel на чистом pdfplumber.
Поддерживает кириллицу, отсекает штампы, ищет шапку как по цифрам, так и по тексту.

Запуск: python src/pdf_to_kmd.py
"""
import re
import sys
from pathlib import Path

import pdfplumber
import pandas as pd

# ---------- Пути ----------
PROJECT_DIR = Path(__file__).parent.parent
INPUT_DIR = PROJECT_DIR / "input"
OUTPUT_DIR = PROJECT_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

# ---------- Параметры ----------
Y_TOLERANCE = 3.0
HEADER_SCAN_ROWS = 30
BOTTOM_CUT_RATIO = 0.99

# ---------- Номера столбцов ----------
COL_NUM = 1
COL_DP = 2
COL_MARK = 4
COL_NAME = 5
COL_QTY = 11
COL_W1 = 16
COL_WTOTAL = 17

# ---------- Словарь исправлений названий элементов ----------
NAME_FIXES = {
    "Св зь": "Связь",
    "Свзь": "Связь",
    "Св зь вертикальная": "Связь вертикальная",
    "Свзь вертикальная": "Связь вертикальная",
    "Св зь горизонтальная": "Связь горизонтальная",
    "Свзь горизонтальная": "Связь горизонтальная",
    "Монтажна деталь": "Монтажная деталь",
    "Монтажна деталь ": "Монтажная деталь",
    "Монтажны элемент": "Монтажный элемент",
    "Монтажна накладка": "Монтажная накладка",
    "Доборна пластина": "Доборная пластина",
    "Доборны элемент": "Доборный элемент",
    "Св ÿзь": "Связь",
    "Монтажна ÿ деталь": "Монтажная деталь",
}

# ======================================================================
# УТИЛИТЫ
# ======================================================================
def group_words_into_lines(words, tolerance=Y_TOLERANCE):
    if not words:
        return []
    sorted_words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines = []
    current_line = [sorted_words[0]]
    current_top = sorted_words[0]["top"]
    for w in sorted_words[1:]:
        if abs(w["top"] - current_top) <= tolerance:
            current_line.append(w)
        else:
            lines.append(sorted(current_line, key=lambda x: x["x0"]))
            current_line = [w]
            current_top = w["top"]
    if current_line:
        lines.append(sorted(current_line, key=lambda x: x["x0"]))
    return lines

def fix_encoding(text):
    if text is None:
        return None
    if not isinstance(text, str):
        text = str(text)
    try:
        fixed = text.encode("latin-1").decode("cp1251")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return text
    if re.search(r"[А-Яа-яЁё]", fixed):
        return fixed
    return text

def fix_name(text):
    if not text:
        return text
    t = text.strip()
    if t in NAME_FIXES:
        return NAME_FIXES[t]
    return text

def parse_number(val):
    if val is None:
        return None
    s = str(val).strip().replace(" ", "")
    if not s:
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None

def extract_dp(val):
    if not val:
        return None
    m = re.search(r"DP\d+", str(val), re.IGNORECASE)
    return m.group(0) if m else None

def extract_mark(val):
    if not val:
        return None
    for part in str(val).split():
        if part.startswith("L") and "-" in part:
            return part
    return None

# ======================================================================
# ПОИСК ШАПКИ (2 ВАРИАНТА)
# ======================================================================
def find_columns_row(lines):
    """Стандартный поиск: строка с цифрами 1, 2, 3, 4, 5."""
    for line_idx, line in enumerate(lines[:HEADER_SCAN_ROWS]):
        texts = [w["text"].strip() for w in line]
        expected = ["1", "2", "3", "4", "5"]
        found = 0
        for t in texts:
            if found < len(expected) and t == expected[found]:
                found += 1
        if found >= 4:
            return line_idx
    return None

def find_columns_row_fallback(lines):
    """Запасной поиск: строка с текстовыми заголовками."""
    for line_idx, line in enumerate(lines[:HEADER_SCAN_ROWS]):
        texts = [w["text"].strip() for w in line]
        if "Марка" in texts and "Кол-во," in texts and "ИТОГО" in texts:
            virtual_line = []
            itogo_count = 0
            for i, w in enumerate(line):
                t = w["text"].strip()
                if t == "№":
                    virtual_line.append({"text": "1", "x0": w["x0"]})
                elif t == "DP":
                    virtual_line.append({"text": "2", "x0": w["x0"]})
                elif t == "Марка":
                    virtual_line.append({"text": "4", "x0": w["x0"]})
                elif t == "Название":
                    virtual_line.append({"text": "5", "x0": w["x0"]})
                elif t.startswith("Кол-во"):
                    virtual_line.append({"text": "11", "x0": w["x0"]})
                elif t == "Вес":
                    next_texts = [x["text"] for x in line[i+1:i+3]]
                    if "ИТОГО" in next_texts:
                        itogo_count += 1
                        if itogo_count == 1:
                            virtual_line.append({"text": "16", "x0": w["x0"]})
                        elif itogo_count == 2:
                            virtual_line.append({"text": "17", "x0": w["x0"]})
            if len(virtual_line) >= 7:
                return line_idx, virtual_line
    return None, None

def build_column_ranges(columns_line):
    """Строит диапазоны столбцов по строке с номерами."""
    nums = []
    for w in columns_line:
        t = w["text"].strip()
        if t.isdigit():
            nums.append((int(t), w["x0"]))
    nums.sort()

    if len(nums) < 2:
        return {}

    boundaries = []
    for i in range(len(nums) - 1):
        mid = (nums[i][1] + nums[i + 1][1]) / 2
        boundaries.append(mid)

    shifts = {(3, 4): -30, (4, 5): +10}

    for i in range(len(boundaries)):
        left_num = nums[i][0]
        right_num = nums[i + 1][0]
        shift = shifts.get((left_num, right_num), 0)
        boundaries[i] += shift

    ranges = {}
    for i, (num, x0) in enumerate(nums):
        if i == 0:
            x_min = x0 - 30
        else:
            x_min = boundaries[i - 1]
        if i == len(nums) - 1:
            x_max = x0 + 60
        else:
            x_max = boundaries[i]
        ranges[num] = (x_min, x_max)

    return ranges

def shift_ranges(ranges, delta):
    return {num: (x_min + delta, x_max + delta) for num, (x_min, x_max) in ranges.items()}

def find_dp_anchor_x(words, top_min=None):
    candidates = []
    for w in words:
        text = w["text"].strip()
        if re.fullmatch(r"DP\d+", text, re.IGNORECASE):
            if top_min is None or w["top"] >= top_min:
                candidates.append(w)
    if not candidates:
        return None
    candidates.sort(key=lambda w: (w["top"], w["x0"]))
    return candidates[0]["x0"]

def assign_word_to_column(word, ranges):
    x0 = word["x0"]
    for num, (x_min, x_max) in ranges.items():
        if x_min <= x0 < x_max:
            return num
    return None

def extract_row_values(line_words, ranges):
    buckets = {}
    for w in line_words:
        num = assign_word_to_column(w, ranges)
        if num is not None:
            buckets.setdefault(num, []).append(w["text"])

    result = {}
    for col_num in [COL_NUM, COL_DP, COL_MARK, COL_NAME, COL_QTY, COL_W1, COL_WTOTAL]:
        parts = buckets.get(col_num, [])
        if not parts:
            result[col_num] = None
            continue
        text = " ".join(parts).strip()
        text = text if text else None
        if text is not None:
            text = fix_encoding(text)
        if col_num == COL_NAME and text is not None:
            text = fix_name(text)
        result[col_num] = text

    return result

# ======================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ======================================================================
def parse_pdf(pdf_path: Path) -> Path:
    print(f"\n{'=' * 70}")
    print(f"ПАРСИНГ PDF: {pdf_path.name}")
    print('=' * 70)

    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF не найден: {pdf_path}")

    all_rows = []
    prev_ranges = None
    ref_dp_x = None

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        print(f"Всего страниц: {total_pages}")

        for page_num in range(1, total_pages + 1):
            page = pdf.pages[page_num - 1]
            words = page.extract_words(x_tolerance=1.5, y_tolerance=2.0)

            if not words:
                print(f"  Страница {page_num}: пусто, пропускаю")
                continue

            lines = group_words_into_lines(words)
            cols_line_idx = find_columns_row(lines)
            virtual_columns_line = None

            # Если нет цифр, ищем текстовые заголовки
            if cols_line_idx is None:
                cols_line_idx, virtual_columns_line = find_columns_row_fallback(lines)

            y_max = page.height * BOTTOM_CUT_RATIO

            if cols_line_idx is not None:
                if virtual_columns_line is not None:
                    ranges = build_column_ranges(virtual_columns_line)
                else:
                    ranges = build_column_ranges(lines[cols_line_idx])

                if not ranges:
                    print(f"  Страница {page_num}: не удалось построить диапазоны, пропускаю")
                    continue

                data_lines = lines[cols_line_idx:]
                data_lines = [l for l in data_lines if l[0]["top"] < y_max]

                header_bottom_top = lines[cols_line_idx][-1]["bottom"]
                anchor_x = find_dp_anchor_x(words, top_min=header_bottom_top)
                if anchor_x is not None:
                    ref_dp_x = anchor_x
                prev_ranges = ranges
                
                if virtual_columns_line is not None:
                    print(f"  Страница {page_num}: диапазоны обновлены (по тексту), ref_dp_x={ref_dp_x}")
                else:
                    print(f"  Страница {page_num}: диапазоны обновлены (по цифрам), ref_dp_x={ref_dp_x}")
            elif prev_ranges is not None:
                anchor_x = find_dp_anchor_x(words)
                if anchor_x is None:
                    print(f"  Страница {page_num}: шапки нет, якорь 'DP' не найден, пропускаю")
                    continue
                if ref_dp_x is None:
                    delta = 0
                else:
                    delta = anchor_x - ref_dp_x
                ranges = shift_ranges(prev_ranges, delta)
                data_lines = [l for l in lines if l[0]["top"] < y_max]
                print(f"  Страница {page_num}: шапки нет, якорь DP на x={anchor_x:.1f}, delta={delta:+.1f}")
            else:
                print(f"  Страница {page_num}: нет ни шапки, ни диапазонов, пропускаю")
                continue

            page_rows = 0
            for line_words in data_lines:
                texts_lower = [w["text"].strip().lower() for w in line_words]
                if any(t.startswith("итого") for t in texts_lower):
                    continue

                values = extract_row_values(line_words, ranges)

                dp_val = extract_dp(values.get(COL_DP))
                mark_val = extract_mark(values.get(COL_MARK))

                if dp_val is None or mark_val is None:
                    continue

                row = {
                    "№": values.get(COL_NUM),
                    "DP": dp_val,
                    "Марка": mark_val,
                    "Название элемента": values.get(COL_NAME),
                    "Кол-во (шт)": parse_number(values.get(COL_QTY)),
                    "Вес 1 шт": parse_number(values.get(COL_W1)),
                    "Вес всех": parse_number(values.get(COL_WTOTAL)),
                    "_page": page_num,
                }
                all_rows.append(row)
                page_rows += 1

            print(f"  Страница {page_num}: {page_rows} строк")

    print(f"\nВсего строк: {len(all_rows)}")

    if not all_rows:
        raise RuntimeError("НЕТ ДАННЫХ — проверьте логику поиска шапки.")

    df = pd.DataFrame(all_rows)

    dp_values = df["DP"].dropna().astype(str).str.strip()
    dp_values = dp_values[dp_values != ""]
    unique_dps = sorted(set(dp_values))

    if not unique_dps:
        raise RuntimeError("Не удалось определить DP из данных.")

    if len(unique_dps) > 1:
        print(f"ℹ️  В PDF найдено несколько DP: {unique_dps}. Сохраняю все в один файл, имя — по первому: {unique_dps[0]}")

    dp = unique_dps[0]
    xlsx_path = OUTPUT_DIR / f"kmd_{dp.lower()}.xlsx"

    df.to_excel(xlsx_path, index=False)
    print(f"\n💾 Сохранено: {xlsx_path}")
    print(f"   Строк: {len(df)}")

    print()
    print("=" * 60)
    print("САМОПРОВЕРКА")
    print("=" * 60)
    print(f"Итого Кол-во (шт):   {df['Кол-во (шт)'].sum():,.2f}")
    print(f"Итого Вес 1 шт:      {df['Вес 1 шт'].sum():,.2f}")
    print(f"Итого Вес всех:      {df['Вес всех'].sum():,.2f}")

    return xlsx_path

if __name__ == "__main__":
    pdfs = sorted(INPUT_DIR.glob("*.pdf"))
    if not pdfs:
        print(f"Нет PDF в {INPUT_DIR}")
        sys.exit(1)
    if len(pdfs) > 1:
        print(f"⚠️  PDF несколько, беру первый: {pdfs[0].name}")
    parse_pdf(pdfs[0])