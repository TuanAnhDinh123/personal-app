"""Nút "Columns" — mở modal tích chọn cột nào được hiển thị trong DataTable.

Dùng cho bảng có RẤT NHIỀU cột (vd nhân viên ~90 cột): mặc định chỉ hiện vài cột
chính, người dùng tự bật thêm cột cần xem. Danh sách cột được GOM NHÓM theo chủ
đề (Identity / Contact / Education…) để dễ tìm, mỗi nhóm xếp thành nhiều cột
checkbox trong một modal cuộn được.

Cách dùng:
    picker = ColumnPicker(table, default_keys=DEFAULT_COLS, groups=GROUPS,
                          on_change=save_cfg)
    picker.set_keys(saved_cols or DEFAULT_COLS)   # nạp cấu hình đã lưu (nếu có)
    toolbar.addWidget(picker)

`groups` = [(tên nhóm, [khóa cột…])…]. Khóa lạ bị bỏ qua; cột không nằm trong
nhóm nào được dồn vào nhóm "Other" ở cuối nên không bao giờ bị mất.
"""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox, QGridLayout, QHBoxLayout, QLabel, QPushButton, QVBoxLayout,
    QWidget,
)

from app_qt import dialogs, theme, widgets
from app_qt.components.modal import ModalDialog

# Số checkbox trên MỘT HÀNG trong modal (modal cỡ "lg" đủ rộng cho 3 cột).
_CHECKS_PER_ROW = 3


class _ColumnPickerDialog(ModalDialog):
    """Modal tích chọn cột, chia theo nhóm. `.run()` → list khóa đã chọn / None."""

    def __init__(self, parent, groups, checked, defaults, min_keys=1):
        super().__init__(parent, "lg")
        self._defaults = defaults
        self._min_keys = min_keys
        self._boxes = {}          # khóa cột → QCheckBox
        self._result = None
        self._total = sum(len(cols) for _title, cols in groups)

        card, lay = self.build_shell("Choose columns")

        desc = QLabel("Tick the columns you want to see in the table.")
        desc.setObjectName("DialogMsg")
        desc.setWordWrap(True)
        lay.addWidget(desc)

        body = QWidget()
        col = QVBoxLayout(body)
        col.setContentsMargins(0, 0, 8, 0)
        col.setSpacing(2)
        for title, cols in groups:
            self._build_group(body, col, title, cols)
        col.addStretch(1)

        sa = widgets.scroll_area(body)
        # Chỉ cuộn DỌC: lưới checkbox tự co theo bề rộng modal, không đẩy ra
        # thanh cuộn ngang (nhìn rối + nuốt mất một hàng nội dung).
        sa.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        lay.addWidget(sa, 1)
        self.set_grow_region(sa)

        self._count_lbl = QLabel("")
        self._count_lbl.setObjectName("Hint")
        lay.addWidget(self._count_lbl)

        foot = QHBoxLayout()
        foot.setContentsMargins(0, 6, 0, 0)
        foot.addWidget(widgets.button(card, "Apply", variant="success", icon="check",
                                      command=self._apply))
        foot.addWidget(widgets.button(card, "Cancel", variant="neutral", icon="x",
                                      command=self.reject))
        foot.addStretch(1)
        foot.addWidget(widgets.button(card, "Show all", variant="neutral",
                                      command=lambda: self._set_all(True)))
        foot.addWidget(widgets.button(card, "Reset to default", variant="neutral",
                                      icon="refresh", command=self._reset))
        lay.addLayout(foot)

        # Tick sẵn theo tập cột đang hiện.
        for key, box in self._boxes.items():
            box.setChecked(key in set(checked))
        self._sync_count()

    def _build_group(self, body, col, title, cols):
        """Một nhóm = tiêu đề + lưới checkbox `_CHECKS_PER_ROW` cột."""
        if title:
            widgets.section_label(body, title)   # tự thêm vào layout của `body`

        grid = QGridLayout()
        grid.setContentsMargins(0, 4, 0, 6)
        grid.setHorizontalSpacing(14)
        grid.setVerticalSpacing(6)
        for i in range(_CHECKS_PER_ROW):
            grid.setColumnStretch(i, 1)
        for i, (key, col_title) in enumerate(cols):
            box = QCheckBox(col_title or key, body)
            box.setCursor(Qt.PointingHandCursor)
            box.setToolTip(col_title or key)   # tiêu đề dài bị cắt → xem ở tooltip
            box.toggled.connect(self._sync_count)
            grid.addWidget(box, i // _CHECKS_PER_ROW, i % _CHECKS_PER_ROW)
            self._boxes[key] = box
        col.addLayout(grid)

    # -------------------------------------------------------------- hành động
    def _set_all(self, checked):
        for box in self._boxes.values():
            box.setChecked(checked)

    def _reset(self):
        wanted = set(self._defaults)
        for key, box in self._boxes.items():
            box.setChecked(key in wanted)

    def _keys(self):
        return [k for k, box in self._boxes.items() if box.isChecked()]

    def _sync_count(self):
        self._count_lbl.setText(
            f"{len(self._keys())} of {self._total} columns selected")

    def _apply(self):
        keys = self._keys()
        if len(keys) < self._min_keys:
            dialogs.warning(self, "Too few columns",
                            f"Please keep at least {self._min_keys} column(s) visible.")
            return
        self._result = keys
        self.accept()

    def run(self):
        self.exec()
        return self._result


class ColumnPicker(QPushButton):
    """Nút mở modal chọn cột hiển thị cho một `DataTable`.

    • `table`        – DataTable cần điều khiển (đọc danh sách cột từ nó).
    • `default_keys` – các cột hiện mặc định; None/rỗng = hiện tất cả.
    • `groups`       – [(tên nhóm, [khóa cột…])…] để gom nhóm trong modal;
                       None = một danh sách phẳng theo thứ tự cột của bảng.
    • `on_change`    – callback(list_keys) mỗi khi người dùng đổi lựa chọn (để
                       tool tự lưu cấu hình xuống đĩa nếu muốn).
    • `min_keys`     – số cột tối thiểu phải còn hiện (không cho ẩn hết bảng).
    """

    def __init__(self, table, default_keys=None, on_change=None, parent=None,
                 label="Columns", min_keys=1, groups=None):
        super().__init__(label, parent)
        self._table = table
        self._columns = table.data_columns()
        self._defaults = list(default_keys or [k for k, _ in self._columns])
        self._on_change = on_change
        self._min_keys = max(1, min_keys)
        self._groups = self._resolve_groups(groups)
        self._keys = []

        self.setProperty("variant", "neutral")
        self.setCursor(Qt.PointingHandCursor)
        self.setIcon(widgets.svg_icon("table", theme.TEXT, 16))
        self.clicked.connect(self._open)

        self.set_keys(self._defaults, notify=False)

    # -------------------------------------------------------------------- API
    def keys(self):
        """Các cột đang hiện (thứ tự theo khai báo cột của bảng)."""
        return list(self._keys)

    def set_keys(self, keys, notify=True):
        """Đặt lại tập cột hiện. None/rỗng → hiện tất cả.

        Bỏ qua key lạ (cấu hình cũ còn cột đã xóa) và tự quay về mặc định nếu
        không còn key nào khớp, để bảng không bao giờ trống cột.
        """
        valid = {k for k, _ in self._columns}
        wanted = {k for k in (keys or ()) if k in valid}
        if not wanted:
            wanted = valid if not keys else {k for k in self._defaults if k in valid}
        # Giữ THỨ TỰ khai báo cột của bảng cho mọi nơi dùng lại (lưu config…).
        self._keys = [k for k, _ in self._columns if k in wanted]
        self._sync(notify=notify)

    # ---------------------------------------------------------------- nội bộ
    def _resolve_groups(self, groups):
        """[(tên nhóm, [(khóa, tiêu đề)…])…] — bỏ khóa lạ, dồn cột lẻ vào "Other"."""
        titles = dict(self._columns)
        if not groups:
            return [(None, list(self._columns))]
        out, used = [], set()
        for name, keys in groups:
            cols = [(k, titles[k]) for k in keys if k in titles and k not in used]
            used.update(k for k, _ in cols)
            if cols:
                out.append((name, cols))
        rest = [(k, t) for k, t in self._columns if k not in used]
        if rest:
            out.append(("Other", rest))
        return out

    def _open(self):
        keys = _ColumnPickerDialog(self.window(), self._groups, self._keys,
                                   self._defaults, self._min_keys).run()
        if keys is not None:
            self.set_keys(keys)

    def _sync(self, notify=True):
        self._table.set_visible_keys(self._keys)
        self.setText(f"Columns ({len(self._keys)}/{len(self._columns)})")
        if notify and callable(self._on_change):
            self._on_change(list(self._keys))
