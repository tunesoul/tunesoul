"""Worker(QThread): фоновое выполнение callable с сигналами log/finished/failed.

print() внутри fn попадает в сигнал log (через StdoutInterceptor, thread-local).
traceback при ошибке — в сигнал failed и в app/logs/app.log.
"""
import traceback

from PySide6.QtCore import QThread, Signal

from core.logger import StdoutInterceptor, app_log


class Worker(QThread):
    log = Signal(str)
    finished_with_result = Signal(object)
    failed = Signal(str)  # полный traceback

    def __init__(self, interceptor: StdoutInterceptor, fn, *args, **kwargs):
        super().__init__()
        self._interceptor = interceptor
        self._fn = fn
        self._args = args
        self._kwargs = kwargs

    def run(self):
        self._interceptor.set_callback(self.log.emit)
        try:
            result = self._fn(*self._args, **self._kwargs)
        except Exception:
            tb = traceback.format_exc()
            app_log(f"Ошибка в фоновой задаче:\n{tb}")
            self.failed.emit(tb)
        else:
            self.finished_with_result.emit(result)
        finally:
            self._interceptor.clear_callback()
