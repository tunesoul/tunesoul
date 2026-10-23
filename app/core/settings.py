"""Загрузка/сохранение settings.json. Пути по умолчанию — относительные от проекта."""
import json

from core.paths import SETTINGS_PATH, INPUT_DIR, OUTPUT_DIR

DEFAULTS = {
    "input_dir": str(INPUT_DIR),
    "output_dir": str(OUTPUT_DIR),
    "theme": "system",           # light | dark | system
    "last_pdf": "",
    "last_xlsm": "",
    "last_xlsb": "",
    "last_xlsx": "",
    "last_sheet": "",
    "last_km": "",
    "last_dp": "",
    "cleanup_after_cycle": False,
    "window_geometry": "",
}


def load_settings() -> dict:
    """Читает settings.json; если файла нет или он битый — создаёт с дефолтами."""
    s = dict(DEFAULTS)
    try:
        with open(SETTINGS_PATH, encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            s.update({k: data[k] for k in DEFAULTS if k in data})
    except FileNotFoundError:
        save_settings(s)
    except (json.JSONDecodeError, OSError):
        pass  # битый файл — работаем на дефолтах, перезапишем при сохранении
    return s


def save_settings(s: dict) -> None:
    tmp = SETTINGS_PATH.with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(s, f, ensure_ascii=False, indent=2)
    tmp.replace(SETTINGS_PATH)
