"""Panel CRUD dùng chung — toolbar (Thêm/Sửa/Xóa[/Nhân bản]/Refresh) + bảng.

Bản Qt của _MasterTab. Nhận một `spec` mô tả bảng + form (xem _master_specs
trong tool candidate_db). Dùng cho các trang master: Bộ phận / Vị trí / Khóa học…

Spec có khóa "duplicate" (callable(id) → id mới) thì toolbar thêm nút Duplicate:
nhân bản dòng đang chọn rồi mở luôn form sửa bản sao.

Spec có "link_keys" (tập khóa cột) + "on_link" (callable(panel, row, key)) thì
các cột đó hiển thị dạng link bấm được — vd cột "JD file" ở trang Positions.
"""
from PySide6.QtWidgets import QHBoxLayout, QVBoxLayout, QWidget

from app_qt import dialogs, widgets
from app_qt.components.form_dialog import FormDialog
from app_qt.components.table import DataTable


class CrudTablePanel(QWidget):
    def __init__(self, spec, on_change=None, parent=None):
        super().__init__(parent)
        self.spec = spec
        self.on_change = on_change

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(10)

        bar = QHBoxLayout()
        bar.setSpacing(6)
        bar.addWidget(widgets.button(self, "Add", variant="success", icon="plus",
                                     command=self._add))
        bar.addWidget(widgets.button(self, "Edit", variant="info", icon="pencil",
                                     command=self._edit))
        bar.addWidget(widgets.button(self, "Delete", variant="danger", icon="trash",
                                     command=self._delete))
        # Nút Duplicate chỉ hiện với spec có khai báo hàm nhân bản (vd mẫu mail).
        if self.spec.get("duplicate"):
            bar.addWidget(widgets.button(self, "Duplicate", variant="warning",
                                         icon="copy", command=self._duplicate))
        bar.addStretch(1)
        bar.addWidget(widgets.button(self, "Reload", variant="neutral", icon="refresh",
                                     command=self.reload))
        lay.addLayout(bar)

        self.table = DataTable(
            self.spec["columns"], pk=self.spec["pk"],
            on_double=lambda _id: self._edit_id(_id),
            link_keys=self.spec.get("link_keys"),
            on_link=self._link_clicked if self.spec.get("on_link") else None)
        lay.addWidget(self.table, 1)

        self.reload()

    def reload(self):
        self.table.set_rows(self.spec["list_fn"]())

    def _link_clicked(self, row, key):
        """Bấm ô link trong bảng → gọi handler của spec; panel truyền theo để
        handler mở dialog đúng cửa sổ cha và nạp lại bảng khi cần."""
        self.spec["on_link"](self, row, key)

    def _changed(self):
        self.reload()
        if self.on_change:
            self.on_change()

    def _selected(self):
        rid = self.table.selected_id()
        if rid is None:
            dialogs.info(self, "Nothing selected", "Please select a row.")
        return rid

    def _add(self):
        FormDialog(
            self, "Add " + self.spec["title"], self.spec["form"], None,
            on_save=lambda data: (self.spec["insert"](data), self._changed()),
            size=self.spec.get("modal_size", "sm")
        ).run()

    def _edit(self):
        rid = self._selected()
        if rid is not None:
            self._edit_id(rid)

    def _edit_id(self, rid):
        current = self.spec["get"](rid)
        FormDialog(
            self, "Edit " + self.spec["title"], self.spec["form"], current,
            on_save=lambda data: (self.spec["update"](rid, data), self._changed()),
            size=self.spec.get("modal_size", "sm")
        ).run()

    def _duplicate(self):
        """Nhân bản dòng đang chọn → mở luôn form sửa bản sao vừa tạo."""
        rid = self._selected()
        if rid is None:
            return
        new_id = self.spec["duplicate"](rid)
        if not new_id:
            dialogs.error(self, "Duplicate failed",
                          f'Could not duplicate {self.spec["title"]} #{rid}.')
            return
        self._changed()
        self._edit_id(new_id)

    def _delete(self):
        rid = self._selected()
        if rid is None:
            return
        if dialogs.confirm(self, "Confirm delete",
                           f'Delete {self.spec["title"]} #{rid}?',
                           ok_label="Delete"):
            self.spec["delete"](rid)
            self._changed()
