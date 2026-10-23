"""Импорт модулей консольного пайплайна из src/ (по ТЗ — только импорт, без subprocess).

src-модули самодостаточны; их PROJECT_DIR = Path(__file__).parent.parent остаётся
корректным и при импорте из GUI. Импортируем один раз (sys.modules кэширует).
"""
import sys

from core.paths import SRC_DIR


def load_src(name: str):
    """Вернуть модуль src/<name>.py по имени (например, 'pdf_to_kmd')."""
    if str(SRC_DIR) not in sys.path:
        sys.path.insert(0, str(SRC_DIR))
    return __import__(name)
