"""Nút "Columns" — dropdown tích chọn cột nào được hiển thị trong DataTable.

Dùng cho bảng có RẤT NHIỀU cột (vd nhân viên): mặc định chỉ hiện vài cột chính,
người dùng tự bật thêm cột cần xem. Menu KHÔNG tự đóng sau mỗi lần tích để tích
được nhiều cột liền một lúc.

Cách dùng:
    picker = ColumnPicker(table, default_keys=DEFAULT_COLS, on_change=save_cfg)
    picker.set_keys(saved_cols or DEFAULT_COLS)   # nạp cấu hình đã lưu (nếu có)
    toolbar.addWidget(picker)
"""
from PySide6.QtCore import Qt
from PySide6.QtGui import QAction
from PySide6.QtWidgets import QMenu, QPushButton

from app_qt import theme, widgets


class _StayOpenMenu(QMenu):
    """QMenu không đóng khi bấm vào một mục CHECKABLE (tích được nhiều mục).

    Mặc định Qt đóng popup ngay sau khi trigger action → mỗi lần bật/tắt một cột
    lại phải mở lại menu. Ở đây tự trigger rồi chặn không cho lớp cha xử lý tiếp.
    """

    def mouseReleaseEvent(self, e):
        act = self.activeAction()
        if act is not None and act.isEnabled() and act.isCheckable():
            act.trigger()
            e.accept()
            return
        super().mouseReleaseEvent(e)

    def keyPressEvent(self, e):
        act = self.activeAction()
        if (e.key() == Qt.Key_Space and act is not None
                and act.isEnabled() and act.isCheckable()):
            act.trigger()
            e.accept()
            return
        super().keyPressEvent(e)


class ColumnPicker(QPushButton):
    """Nút mở dropdown chọn cột hiển thị cho một `DataTable`.

    • `table`        – DataTable cần điều khiển (đọc danh sách cột từ nó).
    • `default_keys` – các cột hiện mặc định; None/rỗng = hiện tất cả.
    • `on_change`    – callback(list_keys) mỗi khi người dùng đổi lựa chọn (để
                       tool tự lưu cấu hình xuống đĩa nếu muốn).
    • `min_keys`     – số cột tối thiểu phải còn hiện (không cho ẩn hết bảng).
    """

    def __init__(self, table, default_keys=None, on_change=None, parent=None,
                 label="Columns", min_keys=1):
        super().__init__(label, parent)
        self._table = table
        self._columns = table.data_columns()
        self._defaults = list(default_keys or [k for k, _ in self._columns])
        self._on_change = on_change
        self._min_keys = max(1, min_keys)

        self.setProperty("variant", "neutral")
        self.setCursor(Qt.PointingHandCursor)
        self.setIcon(widgets.svg_icon("table", theme.TEXT, 16))

        self._menu = _StayOpenMenu(self)
        self._menu.setObjectName("ColumnMenu")   # QSS: chừa chỗ cho ô tick
        self._acts = {}
        for key, title in self._columns:
            act = QAction(title or key, self._menu)
            act.setCheckable(True)
            act.triggered.connect(lambda _checked=False, k=key: self._on_toggle(k))
            self._menu.addAction(act)
            self._acts[key] = act
        self._menu.addSeparator()
        act_all = QAction("Show all", self._menu)
        act_all.triggered.connect(lambda: self.set_keys(None))
        self._menu.addAction(act_all)
        act_def = QAction("Reset to default", self._menu)
        act_def.triggered.connect(lambda: self.set_keys(self._defaults))
        self._menu.addAction(act_def)
        self.setMenu(self._menu)

        self.set_keys(self._defaults, notify=False)

    # -------------------------------------------------------------------- API
    def keys(self):
        """Các cột đang hiện (thứ tự theo khai báo cột của bảng)."""
        return [k for k, _ in self._columns if self._acts[k].isChecked()]

    def set_keys(self, keys, notify=True):
        """Đặt lại tập cột hiện. None/rỗng → hiện tất cả.

        Bỏ qua key lạ (cấu hình cũ còn cột đã xóa) và tự quay về mặc định nếu
        không còn key nào khớp, để bảng không bao giờ trống cột.
        """
        valid = {k for k, _ in self._columns}
        wanted = {k for k in (keys or ()) if k in valid}
        if not wanted:
            wanted = valid if not keys else {k for k in self._defaults if k in valid}
        for key, act in self._acts.items():
            act.setChecked(key in wanted)
        self._sync(notify=notify)

    # ---------------------------------------------------------------- nội bộ
    def _on_toggle(self, key):
        # Giữ lại tối thiểu `min_keys` cột: bỏ tick cột cuối cùng thì tick lại.
        if len(self.keys()) < self._min_keys:
            self._acts[key].setChecked(True)
            return
        self._sync()

    def _sync(self, notify=True):
        keys = self.keys()
        self._table.set_visible_keys(keys)
        self.setText(f"Columns ({len(keys)}/{len(self._columns)})")
        if notify and callable(self._on_change):
            self._on_change(keys)
