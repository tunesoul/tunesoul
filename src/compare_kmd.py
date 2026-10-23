r"""
Сравнение «Прошмандовки» (актуальная ревизия КМД) с подготовленной отчётной таблицей.

Вход:
  - output/kmd_dp6.xlsx          — прошмандовка (эталон)
  - output/report_prepared.xlsx  — наша отчётная (после группировки)

Выход:
  - output/comparison.xlsx       — отчёт с изменениями

Запуск отдельно:
    python src/compare_kmd.py
"""
import sys
from pathlib import Path
import pandas as pd

# ============================================================
# НАСТРОЙКИ
# ============================================================
PROJECT_DIR = Path(__file__).parent.parent
OUTPUT_DIR = PROJECT_DIR / "output"

KMD_COL_MARK = "Марка"
KMD_COL_NAME = "Название элемента"
KMD_COL_QTY = "Кол-во (шт)"
KMD_COL_W1 = "Вес 1 шт"
KMD_COL_WTOTAL = "Вес всех"

REP_COL_MARK = "Марка"
REP_COL_NAME = "Наименование элемента"
REP_COL_QTY = "Кол-во"
REP_COL_W1 = "Вес_1шт"
REP_COL_WTOTAL = "Вес_итого"

EPS = 0.0


# ============================================================
# ВСПОМОГАТЕЛЬНЫЕ
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


def norm_name(val):
    return clean_str(val).lower().replace(" ", "")


def parse_num(val):
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


def num_diff(a, b):
    if a is None and b is None:
        return False
    if a is None or b is None:
        return True
    return abs(a - b) > EPS


def fmt_num(x):
    if x is None:
        return ""
    if isinstance(x, float) and pd.isna(x):
        return ""
    return round(float(x), 2)


def _isna(x):
    return x is None or (isinstance(x, float) and pd.isna(x))


# ============================================================
# ОСНОВНАЯ ФУНКЦИЯ
# ============================================================
def compare(kmd_path: Path, report_path: Path) -> Path:
    """
    Сравнивает прошмандовку с отчётной.
    Возвращает путь к comparison.xlsx.
    """
    print(f"\n{'=' * 70}")
    print("СРАВНЕНИЕ «ПРОШМАНДОВКИ» С ОТЧЁТНОЙ")
    print('=' * 70)

    if not kmd_path.exists():
        raise FileNotFoundError(f"Прошмандовка не найдена: {kmd_path}")
    if not report_path.exists():
        raise FileNotFoundError(f"Отчётная не найдена: {report_path}")

    print(f"📂 {kmd_path.name}  ({kmd_path.stat().st_size / 1024:.1f} КБ)")
    print(f"📂 {report_path.name}  ({report_path.stat().st_size / 1024:.1f} КБ)")

    print("\n⏳ Читаем прошмандовку...")
    kmd = pd.read_excel(kmd_path, dtype=object)
    print(f"   Строк: {len(kmd):,}")

    print("\n⏳ Читаем отчётную...")
    rep = pd.read_excel(report_path, dtype=object)
    print(f"   Строк: {len(rep):,}")

    needed_kmd = [KMD_COL_MARK, KMD_COL_NAME, KMD_COL_QTY, KMD_COL_W1, KMD_COL_WTOTAL]
    needed_rep = [REP_COL_MARK, REP_COL_NAME, REP_COL_QTY, REP_COL_W1, REP_COL_WTOTAL]

    missing_kmd = [c for c in needed_kmd if c not in kmd.columns]
    missing_rep = [c for c in needed_rep if c not in rep.columns]
    if missing_kmd:
        raise RuntimeError(f"В прошмандовке нет столбцов: {missing_kmd}")
    if missing_rep:
        raise RuntimeError(f"В отчётной нет столбцов: {missing_rep}")

    kmd_n = pd.DataFrame({
        "mark": kmd[KMD_COL_MARK].map(clean_str),
        "name": kmd[KMD_COL_NAME].map(clean_str),
        "name_norm": kmd[KMD_COL_NAME].map(norm_name),
        "qty": kmd[KMD_COL_QTY].map(parse_num),
        "w1": kmd[KMD_COL_W1].map(parse_num),
        "wt": kmd[KMD_COL_WTOTAL].map(parse_num),
    })
    kmd_n = kmd_n[kmd_n["mark"] != ""].copy()

    rep_n = pd.DataFrame({
        "mark": rep[REP_COL_MARK].map(clean_str),
        "name": rep[REP_COL_NAME].map(clean_str),
        "name_norm": rep[REP_COL_NAME].map(norm_name),
        "qty": rep[REP_COL_QTY].map(parse_num),
        "w1": rep[REP_COL_W1].map(parse_num),
        "wt": rep[REP_COL_WTOTAL].map(parse_num),
    })
    rep_n = rep_n[rep_n["mark"] != ""].copy()

    kmd_n = kmd_n.groupby("mark", as_index=False).agg(
        name=("name", "first"),
        name_norm=("name_norm", "first"),
        qty=("qty", "sum"),
        w1=("w1", "max"),
        wt=("wt", "sum"),
    )
    rep_n = rep_n.groupby("mark", as_index=False).agg(
        name=("name", "first"),
        name_norm=("name_norm", "first"),
        qty=("qty", "sum"),
        w1=("w1", "max"),
        wt=("wt", "sum"),
    )

    print(f"\n📊 Уникальных марок в прошмандовке: {len(kmd_n):,}")
    print(f"📊 Уникальных марок в отчётной:    {len(rep_n):,}")

    print("\n⏳ Сопоставляем по марке...")
    merged = kmd_n.merge(
        rep_n, on="mark", how="outer",
        suffixes=("_kmd", "_rep"), indicator=True,
    )

    rows = []
    for _, r in merged.iterrows():
        mark = r["mark"]
        only_kmd = r["_merge"] == "left_only"
        only_rep = r["_merge"] == "right_only"

        name_kmd = r.get("name_kmd")
        name_rep = r.get("name_rep")
        name_kmd = "" if _isna(name_kmd) else name_kmd
        name_rep = "" if _isna(name_rep) else name_rep
        name = name_kmd if name_kmd else name_rep

        qty_kmd = r.get("qty_kmd"); qty_kmd = None if _isna(qty_kmd) else qty_kmd
        qty_rep = r.get("qty_rep"); qty_rep = None if _isna(qty_rep) else qty_rep
        w1_kmd = r.get("w1_kmd"); w1_kmd = None if _isna(w1_kmd) else w1_kmd
        w1_rep = r.get("w1_rep"); w1_rep = None if _isna(w1_rep) else w1_rep
        wt_kmd = r.get("wt_kmd"); wt_kmd = None if _isna(wt_kmd) else wt_kmd
        wt_rep = r.get("wt_rep"); wt_rep = None if _isna(wt_rep) else wt_rep

        name_norm_kmd = r.get("name_norm_kmd"); name_norm_kmd = "" if _isna(name_norm_kmd) else name_norm_kmd
        name_norm_rep = r.get("name_norm_rep"); name_norm_rep = "" if _isna(name_norm_rep) else name_norm_rep

        changes = []
        if only_kmd:
            changes.append("Добавлено")
        elif only_rep:
            changes.append("Удалено")
        else:
            if num_diff(qty_kmd, qty_rep):
                changes.append("Изменилось количество")
            if num_diff(w1_kmd, w1_rep):
                changes.append("Изменился вес 1 шт")
            if num_diff(wt_kmd, wt_rep):
                changes.append("Изменился вес итого")
            if name_norm_kmd and name_norm_rep and name_norm_kmd != name_norm_rep:
                changes.append("Изменилось наименование")

        if not changes:
            continue

        d_qty = (qty_kmd - qty_rep) if (qty_kmd is not None and qty_rep is not None) else None
        d_w1 = (w1_kmd - w1_rep) if (w1_kmd is not None and w1_rep is not None) else None
        d_wt = (wt_kmd - wt_rep) if (wt_kmd is not None and wt_rep is not None) else None

        rows.append({
            "Марка": mark,
            "Наименование": name,
            "Кол-во (прошм.)": fmt_num(qty_kmd),
            "Кол-во (наше)": fmt_num(qty_rep),
            "Δ Кол-во": fmt_num(d_qty),
            "Вес 1 шт (прошм.)": fmt_num(w1_kmd),
            "Вес 1 шт (наше)": fmt_num(w1_rep),
            "Δ Вес 1 шт": fmt_num(d_w1),
            "Вес итого (прошм.)": fmt_num(wt_kmd),
            "Вес итого (наше)": fmt_num(wt_rep),
            "Δ Вес итого": fmt_num(d_wt),
            "Наименование (прошм.)": name_kmd,
            "Наименование (наше)": name_rep,
            "Тип изменения": "; ".join(changes),
        })

    result_path = OUTPUT_DIR / "comparison.xlsx"

    if not rows:
        print("\n" + "=" * 70)
        print("✅ ИЗМЕНЕНИЙ НЕТ. Отчётная таблица актуальна.")
        print("=" * 70)
        pd.DataFrame([{
            "Марка": "", "Наименование": "",
            "Кол-во (прошм.)": "", "Кол-во (наше)": "", "Δ Кол-во": "",
            "Вес 1 шт (прошм.)": "", "Вес 1 шт (наше)": "", "Δ Вес 1 шт": "",
            "Вес итого (прошм.)": "", "Вес итого (наше)": "", "Δ Вес итого": "",
            "Наименование (прошм.)": "", "Наименование (наше)": "",
            "Тип изменения": "Изменений нет. Отчётная таблица актуальна.",
        }]).to_excel(result_path, index=False, engine="openpyxl")
        print(f"💾 Сохранено: {result_path}")
        return result_path

    df_res = pd.DataFrame(rows)
    order = {
        "Добавлено": 1, "Удалено": 2,
        "Изменилось количество": 3, "Изменился вес 1 шт": 4,
        "Изменился вес итого": 5, "Изменилось наименование": 6,
    }
    def sort_key(t):
        for k, v in order.items():
            if k in str(t):
                return v
        return 99
    df_res["_sort"] = df_res["Тип изменения"].map(sort_key)
    df_res = df_res.sort_values(["_sort", "Марка"]).drop(columns="_sort").reset_index(drop=True)

    print(f"\n📊 Найдено изменений: {len(df_res):,}")
    print("\n   По типам:")
    for t, cnt in df_res["Тип изменения"].value_counts().items():
        print(f"     {t}: {cnt}")

    df_res.to_excel(result_path, index=False, engine="openpyxl")
    print(f"\n💾 Сохранено: {result_path}")

    print("\n📋 Первые 20 строк результата:")
    with pd.option_context("display.max_columns", None, "display.width", 250):
        print(df_res.head(20).to_string(index=False))

    return result_path


# ============================================================
# ЗАПУСК ОТДЕЛЬНО
# ============================================================
if __name__ == "__main__":
    kmds = sorted(OUTPUT_DIR.glob("kmd_*.xlsx"))
    if not kmds:
        print("Прошмандовка (kmd_*.xlsx) не найдена в output/")
        sys.exit(1)
    compare(kmds[0], OUTPUT_DIR / "report_prepared.xlsx")