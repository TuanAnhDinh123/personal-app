"""Chạy một hàm nặng ở luồng nền, trả kết quả về luồng giao diện qua signal.

Dùng khi KHÔNG cần hộp thoại tiến trình (cập nhật trạng thái tại chỗ):

    self._task = Task(lambda emit: nang_viec(progress=lambda *a: emit(a)), parent)
    self._task.signals.progress.connect(...)
    self._task.signals.finished.connect(...)
    self._task.signals.failed.connect(...)
    self._task.start()

Giữ tham chiếu (self._task) để QThread không bị thu hồi giữa chừng.
"""
from PySide6.QtCore import QObject, QThread, Signal


class _TaskSignals(QObject):
    progress = Signal(object)
    finished = Signal(object)
    failed = Signal(str)


class Task(QThread):
    def __init__(self, fn, parent=None):
        super().__init__(parent)
        self._fn = fn
        self.signals = _TaskSignals()

    def run(self):
        try:
            result = self._fn(self.signals.progress.emit)
            self.signals.finished.emit(result)
        except Exception as exc:
            self.signals.failed.emit(str(exc))
