"""Пути проекта: всё относительно расположения приложения (работа и в dev, и в .exe)."""
import sys
from pathlib import Path

if getattr(sys, "frozen", False):  # собрано PyInstaller
    APP_DIR = Path(sys.executable).resolve().parent          # app.exe лежит в app/
    PROJECT_DIR = APP_DIR.parent                              # корень excel-compare
else:
    APP_DIR = Path(__file__).resolve().parent.parent          # .../excel-compare/app
    PROJECT_DIR = APP_DIR.parent                              # .../excel-compare

SRC_DIR = PROJECT_DIR / "src"
INPUT_DIR = PROJECT_DIR / "input"
OUTPUT_DIR = PROJECT_DIR / "output"
LOGS_DIR = APP_DIR / "logs"
ASSETS_DIR = APP_DIR / "assets"
SETTINGS_PATH = PROJECT_DIR / "settings.json"
