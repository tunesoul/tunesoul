"""Monkey-patching ask_choice/input из src-модулей (ТЗ, разделы 4 и 6).

src/ НЕ трогаем: подменяем ссылки на функции в импортированных модулях.
Режимы gui_ask_choice:
  1) предзаполненный ответ из очереди (по подстроке prompt);
  2) модальный диалог в GUI-потоке (blocking queued call);
  3) если диалог показать невозможно — options[0] + предупреждение в лог.
"""
import builtins
import sys
import threading
from pathlib import Path

from core.paths import SRC_DIR

if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

# очередь предзаполненных ответов: [(подстрока_промпта, ответ), ...]
_prefilled: list[tuple[str, object]] = []
_lock = threading.Lock()

# мост, установленный MainWindow: callable(prompt: str, options: list) -> answer | None
_dialog_bridge = None


def set_prefilled(pairs: list[tuple[str, object]]) -> None:
    with _lock:
        _prefilled.clear()
        _prefilled.extend(pairs)


def clear_prefilled() -> None:
    with _lock:
        _prefilled.clear()


def set_dialog_bridge(fn) -> None:
    global _dialog_bridge
    _dialog_bridge = fn


def gui_ask_choice(prompt: str, options: list):
    if not options:
        raise RuntimeError(f"Нет вариантов для выбора: {prompt}")

    with _lock:
        for i, (pat, ans) in enumerate(_prefilled):
            if pat and pat in prompt:
                _prefilled.pop(i)
                print(f"Выбрано автоматически: {ans}")
                return ans

    if _dialog_bridge is not None:
        res = _dialog_bridge(prompt, list(options))
        if res is not None:
            return res

    print(f"⚠️ ask_choice вызван без GUI — взят первый вариант: {options[0]}")
    return options[0]


def gui_input(prompt: str = "") -> str:
    """Страховка от прямого input() в src-модулях."""
    print(f"⚠️ Перехвачен прямой input({prompt!r}) — возвращена пустая строка.")
    return ""


_patched = False


def apply_patches() -> None:
    """Подменяет ask_choice в prepare_report/transfer_dep и builtins.input."""
    global _patched
    if _patched:
        return
    import prepare_report
    import transfer_dep

    prepare_report.ask_choice = gui_ask_choice
    transfer_dep.ask_choice = gui_ask_choice
    builtins.input = gui_input
    _patched = True
