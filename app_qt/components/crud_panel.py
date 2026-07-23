"""Panel CRUD dùng chung — toolbar (Thêm/Sửa/Xóa/Refresh) + bảng.

Bản Qt của _MasterTab. Nhận một `spec` mô tả bảng + form (xem _master_specs
trong tool candidate_db). Dùng cho 3 trang master: Bộ phận / Vị trí / JD.
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
        bar.addWidget(widgets.button(self, "Thêm", variant="success", icon="plus",
                                     command=self._add))
        bar.addWidget(widgets.button(self, "Sửa", variant="info", icon="pencil",
                                     command=self._edit))
        bar.addWidget(widgets.button(self, "Xóa", variant="danger", icon="trash",
                                     command=self._delete))
        bar.addStretch(1)
        bar.addWidget(widgets.button(self, "Tải lại", variant="neutral", icon="refresh",
                                     command=self.reload))
        lay.addLayout(bar)

        self.table = DataTable(
            self.spec["columns"], pk=self.spec["pk"],
            on_double=lambda _id: self._edit_id(_id))
        lay.addWidget(self.table, 1)

        self.reload()

    def reload(self):
        self.table.set_rows(self.spec["list_fn"]())

    def _changed(self):
        self.reload()
        if self.on_change:
            self.on_change()

    def _selected(self):
        rid = self.table.selected_id()
        if rid is None:
            dialogs.info(self, "Chưa chọn", "Vui lòng chọn một dòng.")
        return rid

    def _add(self):
        FormDialog(
            self, "Thêm " + self.spec["title"], self.spec["form"], None,
            on_save=lambda data: (self.spec["insert"](data), self._changed())
        ).run()

    def _edit(self):
        rid = self._selected()
        if rid is not None:
            self._edit_id(rid)

    def _edit_id(self, rid):
        current = self.spec["get"](rid)
        FormDialog(
            self, "Sửa " + self.spec["title"], self.spec["form"], current,
            on_save=lambda data: (self.spec["update"](rid, data), self._changed())
        ).run()

    def _delete(self):
        rid = self._selected()
        if rid is None:
            return
        if dialogs.confirm(self, "Xác nhận xóa",
                           f'Xóa {self.spec["title"]} #{rid}?',
                           ok_label="Xóa"):
            self.spec["delete"](rid)
            self._changed()
