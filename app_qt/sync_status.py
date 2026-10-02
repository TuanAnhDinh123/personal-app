"""Ô trạng thái đồng bộ DB ở sidebar: "Synced 09:15" + nút Refresh + nhãn Read-only.

Chỉ là lớp vỏ của `app.core.shared_db`: đọc `shared_db.status()`, nghe thay đổi
qua `shared_db.add_listener` (có thể đến từ luồng phụ → chuyển về luồng giao
diện bằng Signal), bấm Refresh thì chạy `ensure_fresh()` ở luồng nền. Tính năng
tắt (chưa cấu hình thư mục dùng chung) thì cả ô ẩn đi.
"""
from PySide6.QtCore import QObject, QSize, Qt, QTimer, Signal
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QPushButton, QVBoxLayout

from app.core import shared_db
from app_qt import dialogs, theme, widgets
from app_qt.components.task import Task

# Nhịp kiểm tra bản chủ mới khi app đang mở (ngoài các lượt từ get_connection).
CHECK_INTERVAL_MS = 3 * 60 * 1000

_DOT = {
    "synced":  theme.PALETTE["--success"],
    "pending": theme.PALETTE["--warning"],
    "offline": theme.PALETTE["--sidebar-muted"],
    "blocked": theme.PALETTE["--danger"],
    "error":   theme.PALETTE["--danger"],
}


def _hhmm(dt):
    return f"{dt:%H:%M}" if dt else ""


def status_text(st: shared_db.SyncStatus) -> str:
    """Dòng chữ ngắn cho ô trạng thái (tiếng Anh)."""
    if st.state == "synced":
        return f"Synced {_hhmm(st.synced_at)}".strip()
    if st.state == "pending":
        return "Update waiting…"
    if st.state == "offline":
        return (f"Offline · synced {_hhmm(st.synced_at)}" if st.synced_at
                else "Offline")
    if st.state == "blocked":
        return "Sync paused"
    if st.state == "error":
        return "Sync error"
    return ""


class _Bridge(QObject):
    """Chuyển thông báo của shared_db (bất kỳ luồng nào) về luồng giao diện."""
    changed = Signal(object, bool)


class SyncStatusBar(QFrame):
    """Ô trạng thái đồng bộ. Phát `db_replaced` khi file .db local vừa bị thay
    bằng bản kéo về — cửa sổ chính nghe để tải lại các trang đang mở."""

    db_replaced = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("SidebarSync")
        self._task = None
        self._refreshing = False
        self._warned = set()       # thông báo "blocked" đã hiện trong phiên này

        outer = QVBoxLayout(self)
        outer.setContentsMargins(26, 6, 12, 4)   # thẳng mép trái ô hồ sơ
        outer.setSpacing(6)

        row = QHBoxLayout()
        row.setSpacing(8)
        self._dot = QLabel(self)
        self._dot.setFixedSize(8, 8)
        row.addWidget(self._dot, 0, Qt.AlignVCenter)
        self._text = QLabel(self)
        self._text.setObjectName("SyncText")
        row.addWidget(self._text, 1)
        self._btn = QPushButton(self)
        self._btn.setObjectName("SyncRefresh")
        self._btn.setToolTip("Check the shared database now")
        self._btn.setCursor(Qt.PointingHandCursor)
        self._btn.setIcon(widgets.svg_icon("refresh", theme.PALETTE["--sidebar-text"], 14))
        self._btn.setIconSize(QSize(14, 14))
        self._btn.setText("Refresh")
        self._btn.clicked.connect(lambda *_: self.refresh_now())
        row.addWidget(self._btn)
        outer.addLayout(row)

        # Nhãn Read-only đứng riêng một dòng, nền màu cảnh báo: máy chỉ đọc phải
        # NHÌN LÀ BIẾT, không phải di chuột mới thấy.
        self._ro = QLabel("READ-ONLY  ·  changes can't be saved", self)
        self._ro.setObjectName("SyncReadOnly")
        self._ro.setToolTip("This computer only reads the shared database. "
                            "Turn it off in Settings → Shared database.")
        outer.addWidget(self._ro, 0, Qt.AlignLeft)

        self._bridge = _Bridge(self)
        self._bridge.changed.connect(self._on_changed)
        shared_db.add_listener(lambda st, replaced: self._bridge.changed.emit(st, replaced))

        self._timer = QTimer(self)
        self._timer.setInterval(CHECK_INTERVAL_MS)
        self._timer.timeout.connect(lambda: shared_db.check_in_background("timer"))
        self._timer.start()

        self._apply(shared_db.status())

    # ------------------------------------------------------------ hiển thị
    def _apply(self, st):
        self.setVisible(st.state != "off")
        self._ro.setVisible(st.state != "off" and st.read_only)
        self._text.setText(status_text(st))
        self._text.setToolTip(st.message)
        self._dot.setStyleSheet(
            f"background: {_DOT.get(st.state, theme.PALETTE['--sidebar-muted'])};"
            "border-radius: 4px;")

    def _on_changed(self, st, replaced):
        self._apply(st)
        if replaced:
            self.db_replaced.emit()
        # Bị chặn là việc cần NGƯỜI quyết định → báo một lần mỗi phiên cho mỗi
        # lý do, không đợi người dùng tình cờ di chuột vào ô trạng thái. Đang
        # Refresh thì để `_on_refreshed` báo, tránh hiện hai hộp thoại.
        if (st.state == "blocked" and not self._refreshing
                and st.message not in self._warned):
            self._warned.add(st.message)
            dialogs.warning(self.window(), "Shared database", st.message)

    # ------------------------------------------------------------ Refresh
    def refresh_now(self):
        if self._task is not None and self._task.isRunning():
            return
        self._refreshing = True
        self._btn.setEnabled(False)
        self._text.setText("Checking…")
        # Ổ mạng có thể treo hàng chục giây → chạy nền, không đứng giao diện.
        self._task = Task(lambda _emit: shared_db.ensure_fresh("manual"), self)
        self._task.signals.finished.connect(self._on_refreshed)
        self._task.start()

    def _on_refreshed(self, result):
        self._refreshing = False
        self._btn.setEnabled(True)
        st = result.status
        self._apply(st)
        # Người dùng vừa CHỦ ĐỘNG bấm → trạng thái xấu thì nói rõ luôn (kể cả
        # thông báo "blocked" đã hiện trước đó).
        if st.state in ("blocked", "error", "offline", "pending"):
            self._warned.add(st.message)
            show = dialogs.error if st.state == "error" else dialogs.warning
            show(self.window(), "Shared database", st.message)
