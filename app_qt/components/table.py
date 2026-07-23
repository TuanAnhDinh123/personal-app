"""Bảng dữ liệu dùng chung — QTableView + model đơn giản trên list các dòng.

Dùng lại cho mọi tool cần hiển thị bảng (ứng viên, master data, danh sách trùng…).
Mỗi dòng là sqlite3.Row hoặc dict; cột khai báo bằng list (key, title, width, align).
"""
from PySide6.QtCore import QAbstractTableModel, Qt
from PySide6.QtGui import QColor, QFont, QKeySequence
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QHeaderView, QMenu, QStyle,
    QStyledItemDelegate, QTableView,
)

from app_qt import theme

_ALIGN = {
    "w": Qt.AlignLeft | Qt.AlignVCenter,
    "center": Qt.AlignCenter,
    "e": Qt.AlignRight | Qt.AlignVCenter,
}


def _cell(row, key):
    """Đọc row[key] an toàn cho cả sqlite3.Row lẫn dict."""
    try:
        val = row[key]
    except (KeyError, IndexError):
        return ""
    return "" if val is None else val


class DictTableModel(QAbstractTableModel):
    def __init__(self, columns, rows=None, pk=None, link_keys=None):
        super().__init__()
        # columns: list of (key, title, width, align[, formatter])
        # formatter (tùy chọn): callable(value) -> str để đổi cách hiển thị ô.
        self.columns = columns
        self.rows = list(rows or [])
        self.pk = pk
        self.link_keys = set(link_keys or ())

    def rowCount(self, _parent=None):
        return len(self.rows)

    def columnCount(self, _parent=None):
        return len(self.columns)

    def _text(self, index):
        col = self.columns[index.column()]
        val = _cell(self.rows[index.row()], col[0])
        fmt = col[4] if len(col) > 4 else None
        return fmt(val) if fmt else str(val)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        col = self.columns[index.column()]
        if role == Qt.DisplayRole:
            return self._text(index)
        if role == Qt.TextAlignmentRole:
            align = col[3] if len(col) > 3 else "w"
            return _ALIGN.get(align, _ALIGN["w"])
        if col[0] in self.link_keys and self._text(index):
            if role == Qt.ForegroundRole:
                return QColor(theme.PALETTE["--info"])
            if role == Qt.FontRole:
                font = QFont()
                font.setUnderline(True)
                return font
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.columns[section][1]
        return None

    def set_rows(self, rows):
        self.beginResetModel()
        self.rows = list(rows or [])
        self.endResetModel()

    def row_at(self, r):
        return self.rows[r] if 0 <= r < len(self.rows) else None

    def id_at(self, r):
        row = self.row_at(r)
        return None if row is None or self.pk is None else row[self.pk]


class _RowHoverDelegate(QStyledItemDelegate):
    """Tô nền cả DÒNG khi rê chuột (QSS ::item:hover chỉ tô 1 ô)."""

    def __init__(self, view):
        super().__init__(view)
        self._view = view

    def paint(self, painter, option, index):
        if (index.row() == self._view.hover_row
                and not (option.state & QStyle.State_Selected)):
            painter.fillRect(option.rect, QColor(theme.PALETTE["--row-hover"]))
        # bỏ trạng thái hover mặc định của style (tô lẻ từng ô) cho gọn
        option.state &= ~QStyle.State_MouseOver
        super().paint(painter, option, index)


class DataTable(QTableView):
    """QTableView đã cấu hình sẵn: chọn/hover theo cả hàng, ẩn số dòng."""

    def __init__(self, columns, pk=None, stretch_key=None, on_double=None,
                 link_keys=None, on_link=None, parent=None):
        super().__init__(parent)
        self.hover_row = -1
        self._model = DictTableModel(columns, pk=pk, link_keys=link_keys)
        self.setModel(self._model)
        self._link_cols = {i for i, c in enumerate(columns)
                           if c[0] in self._model.link_keys}
        self._on_link = on_link
        self.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.setSelectionMode(QAbstractItemView.SingleSelection)
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.setAlternatingRowColors(False)
        self.setItemDelegate(_RowHoverDelegate(self))
        self.setMouseTracking(True)
        self.verticalHeader().setVisible(False)
        self.verticalHeader().setDefaultSectionSize(36)
        self.horizontalHeader().setHighlightSections(False)
        self.setShowGrid(False)
        self.setWordWrap(False)

        header = self.horizontalHeader()
        for i, col in enumerate(columns):
            width = col[2] if len(col) > 2 else 120
            self.setColumnWidth(i, width)
            key = col[0]
            if stretch_key and key == stretch_key:
                header.setSectionResizeMode(i, QHeaderView.Stretch)
            else:
                header.setSectionResizeMode(i, QHeaderView.Interactive)

        if on_double:
            self.doubleClicked.connect(lambda idx: on_double(self._model.id_at(idx.row())))
        if self._on_link:
            self.clicked.connect(self._maybe_link)

        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._context_menu)

    def set_rows(self, rows):
        self._model.set_rows(rows)

    def selected_id(self):
        idxs = self.selectionModel().selectedRows()
        if not idxs:
            return None
        return self._model.id_at(idxs[0].row())

    def _set_hover(self, row):
        if row != self.hover_row:
            self.hover_row = row
            self.viewport().update()

    def _is_link_cell(self, index):
        return (index.isValid() and index.column() in self._link_cols
                and bool(self._model._text(index)))

    def _maybe_link(self, index):
        if self._is_link_cell(index):
            row = self._model.row_at(index.row())
            key = self._model.columns[index.column()][0]
            self._on_link(row, key)

    # -------------------------------------------------------------- sao chép
    @staticmethod
    def _copy(value):
        QApplication.clipboard().setText(str(value))

    def keyPressEvent(self, e):
        if e.matches(QKeySequence.Copy):
            idx = self.currentIndex()
            if idx.isValid():
                self._copy(self._model._text(idx))
            return
        super().keyPressEvent(e)

    def _context_menu(self, pos):
        idx = self.indexAt(pos)
        if not idx.isValid():
            return
        menu = QMenu(self)
        act = menu.addAction("Copy", lambda: self._copy(self._model._text(idx)))
        act.setShortcut(QKeySequence.Copy)
        menu.exec(self.viewport().mapToGlobal(pos))

    def mouseMoveEvent(self, e):
        idx = self.indexAt(e.position().toPoint())
        self._set_hover(idx.row())
        self.viewport().setCursor(
            Qt.PointingHandCursor if self._is_link_cell(idx) else Qt.ArrowCursor)
        super().mouseMoveEvent(e)

    def leaveEvent(self, e):
        self._set_hover(-1)
        super().leaveEvent(e)
