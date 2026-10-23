"""Логгер приложения + перехват stdout/stderr.

StdoutInterceptor подменяет sys.stdout внутри фонового потока (thread-local):
всё, что старые модули печатают через print(), попадает в сигнал log Worker'а
и в файл app/logs/app.log. Окно консоли не открывается вообще — UnicodeEncodeError
исключён по построению.
"""
import sys
import threading
from datetime import datetime
from pathlib import Path

from core.paths import LOGS_DIR

_log_file: Path | None = None
_lock = threading.Lock()


def init_logging() -> Path:
    """Создаёт app/logs/ и файл лога на сегодня. Возвращает путь к логу."""
    global _log_file
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    _log_file = LOGS_DIR / f"app_{datetime.now():%Y-%m-%d}.log"
    return _log_file


def app_log(msg: str) -> None:
    line = f"[{datetime.now():%H:%M:%S}] {msg}"
    if _log_file is not None:
        with _lock:
            try:
                with open(_log_file, "a", encoding="utf-8") as f:
                    f.write(line + "\n")
            except OSError:
                pass


class StdoutInterceptor:
    """thread-local подмена sys.stdout: print() фоновых потоков → callback;
    print() из GUI-потока → в original_stdout (не теряем вывод)."""

    def __init__(self, original):
        self._original = original
        self._local = threading.local()

    def set_callback(self, cb):
        self._local.cb = cb

    def clear_callback(self):
        self._local.cb = None

    def write(self, text):
        cb = getattr(self._local, "cb", None)
        if cb is not None:
            if text and text.strip():
                cb(text.rstrip("\n"))
        elif self._original is not None:
            try:
                self._original.write(text)
            except (UnicodeEncodeError, ValueError):
                pass
        return len(text)

    def flush(self):
        try:
            if self._original is not None:
                self._original.flush()
        except ValueError:
            pass


_interceptor: StdoutInterceptor | None = None


def install_interceptor():
    """Вызывается один раз при старте GUI. Возвращает объект-перехватчик."""
    global _interceptor
    if _interceptor is None:
        _interceptor = StdoutInterceptor(sys.__stdout__ or sys.stderr)
        sys.stdout = _interceptor
        sys.stderr = _interceptor
    return _interceptor


def open_in_explorer(path: str | Path) -> None:
    """Открыть файл/папку стандартным приложением (Windows) или аналогом."""
    import subprocess
    path = Path(path)
    try:
        if sys.platform == "win32":
            import os
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except OSError as e:
        app_log(f"Не удалось открыть {path}: {e}")
