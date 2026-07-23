"""Hộp thoại tiến trình + luồng nền dùng chung (QThread + signals).

Mẫu chuẩn cho các tool chạy lâu (quét CV AI, quét CV regex, PDF→Text…):
chạy công việc nặng trong luồng nền, đẩy cập nhật (log/status/tiến độ) về luồng
giao diện an toàn qua signal. Giao diện: bo góc, thanh tiến trình, ô nhật ký.

Dùng:
    dlg = ProgressDialog(parent, "Đang xử lý…", total=len(items),
                         subtitle="Quét 10 CV")
    dlg.start(job, on_finish)     # job(ctx) chạy ở luồng nền; on_finish(dlg, kq)
Trong job dùng: ctx.log(str) · ctx.status(str) · ctx.step(n=1) · ctx.cancelled
"""
from PySide6.QtCore import QObject, Qt, QThread, Signal
from PySide6.QtWidgets import (
    QFrame, QHBoxLayout, QLabel, QProgressBar, QTextEdit, QVBoxLayout,
)

from app_qt import widgets


class _Signals(QObject):
    log = Signal(str)
    status = Signal(str)
    step = Signal(int)
    finished = Signal(object)
    failed = Signal(str)


class _Worker(QThread):
    """Chạy `job(ctx)` trong luồng nền; ctx chính là worker này."""

    def __init__(self, job, parent=None):
        super().__init__(parent)
        self._job = job
        self._cancel = False
        self.signals = _Signals()

    # --- API cho job dùng (đẩy cập nhật về luồng giao diện) ---
    def log(self, text):
        self.signals.log.emit(str(text))

    def status(self, text):
        self.signals.status.emit(str(text))

    def step(self, n=1):
        self.signals.step.emit(int(n))

    @property
    def cancelled(self):
        return self._cancel

    def cancel(self):
        self._cancel = True

    def run(self):
        try:
            result = self._job(self)
            self.signals.finished.emit(result)
        except Exception as exc:
            self.signals.failed.emit(str(exc))


class ProgressDialog(QFrame):
    """Cửa sổ tiến trình độc lập (frameless) — giữ tham chiếu để không bị thu hồi."""

    _alive = set()   # giữ tham chiếu các dialog đang chạy

    def __init__(self, parent, title, total=0, subtitle=""):
        # dùng QFrame làm cửa sổ riêng (Tool) để bo góc + shadow như dialog khác
        super().__init__(None, Qt.FramelessWindowHint | Qt.Dialog)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self._on_finish = None
        self._worker = None
        self._drag = None

        shell = QVBoxLayout(self)
        shell.setContentsMargins(24, 24, 24, 24)
        card = QFrame(self)
        card.setObjectName("Dialog")
        card.setMinimumWidth(560)
        widgets.add_shadow(card, blur=48, dy=12, alpha=70)
        shell.addWidget(card)

        lay = QVBoxLayout(card)
        lay.setContentsMargins(22, 20, 22, 18)   # padding thẻ→nội dung chuẩn
        lay.setSpacing(10)

        t = QLabel(title)
        t.setObjectName("DialogTitle")
        lay.addWidget(t)
        self._status = QLabel(subtitle or "Chuẩn bị…")
        self._status.setObjectName("DialogMsg")
        self._status.setWordWrap(True)
        lay.addWidget(self._status)

        self._bar = QProgressBar()
        self._bar.setMaximum(max(1, total))
        self._bar.setValue(0)
        self._bar.setTextVisible(False)
        self._bar.setFixedHeight(8)
        lay.addWidget(self._bar)

        self._log = QTextEdit()
        self._log.setReadOnly(True)
        self._log.setMinimumHeight(200)
        lay.addWidget(self._log, 1)

        foot = QHBoxLayout()
        foot.addStretch(1)
        self._btn = widgets.button(card, "Hủy", variant="neutral", icon="x",
                                   command=self._on_button)
        foot.addWidget(self._btn)
        lay.addLayout(foot)

        self.resize(600, 460)
        if parent is not None:
            self._center_on(parent)

    def _center_on(self, parent):
        try:
            g = parent.window().frameGeometry()
            self.move(g.center().x() - 300, g.center().y() - 230)
        except Exception:
            pass

    # --------------------------------------------------------------- chạy job
    def start(self, job, on_finish=None):
        self._on_finish = on_finish
        self._worker = _Worker(job, self)
        s = self._worker.signals
        s.log.connect(self._append_log)
        s.status.connect(self._status.setText)
        s.step.connect(self._advance)
        s.finished.connect(self._finished)
        s.failed.connect(self._failed)
        ProgressDialog._alive.add(self)
        self.show()
        self.raise_()
        self._worker.start()

    def _append_log(self, line):
        self._log.append(line)

    def _advance(self, n):
        self._bar.setValue(self._bar.value() + n)

    def _on_button(self):
        if self._worker is not None and self._worker.isRunning():
            self._worker.cancel()
            self._status.setText("Đang hủy…")
        else:
            self.close()

    def _finished(self, result):
        self._done()
        if self._on_finish is not None:
            self._on_finish(self, result)

    def _failed(self, msg):
        self._done()
        self._append_log(f"⚠ Lỗi: {msg}")
        self._status.setText("Đã dừng do lỗi.")

    def _done(self):
        self._btn.setText("Đóng")

    def set_final_status(self, text):
        self._status.setText(text)

    def log(self, text):
        self._append_log(str(text))

    # kéo di chuyển
    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag = e.globalPosition().toPoint() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, e):
        if self._drag is not None and e.buttons() & Qt.LeftButton:
            self.move(e.globalPosition().toPoint() - self._drag)

    def mouseReleaseEvent(self, e):
        self._drag = None

    def closeEvent(self, e):
        if self._worker is not None and self._worker.isRunning():
            self._worker.cancel()
            self._worker.wait(3000)
        ProgressDialog._alive.discard(self)
        super().closeEvent(e)
