"""Bảng dữ liệu dùng chung — QTableView + model đơn giản trên list các dòng.

Dùng lại cho mọi tool cần hiển thị bảng (ứng viên, master data, danh sách trùng…).
Mỗi dòng là sqlite3.Row hoặc dict; cột khai báo bằng list (key, title, width, align).
"""
from PySide6.QtCore import QAbstractTableModel, QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QFont, QKeySequence, QPainter, QPen
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

# Khóa cột "chọn" (checkbox). DataTable tự chèn cột này ở đầu khi checkable=True.
_CHECK_KEY = "__check__"

# ── Kích thước cột checkbox (px) — chỉnh tùy ý ──
CHECK_COL_WIDTH = 30    # bề rộng cả cột checkbox
CHECK_BOX_INSET = 8     # lề trái của ô tick trong cột (để không dính sát biên)


def _draw_checkbox(painter, rect, checked, partial=False):
    """Vẽ 1 ô checkbox theo phong cách app (viền bo góc; tick trắng trên nền nhấn).

    partial=True → trạng thái 'một phần' (dấu gạch ngang) cho ô check-all ở header.
    """
    # Canh TRÁI với lề cố định 8px (bằng padding cột khác) thay vì canh giữa:
    # thu hẹp cột không đẩy checkbox ra sát biên, mà kéo cột kế (ID) lại gần.
    size = 16
    cy = rect.center().y()
    cx = rect.left() + CHECK_BOX_INSET + size / 2
    box = QRectF(cx - size / 2, cy - size / 2, size, size)
    painter.save()
    painter.setRenderHint(QPainter.Antialiasing, True)
    if checked or partial:
        painter.setBrush(QColor(theme.PALETTE["--accent"]))
        painter.setPen(QPen(QColor(theme.PALETTE["--accent"]), 1))
        painter.drawRoundedRect(box, 4, 4)
        pen = QPen(QColor("#ffffff"))
        pen.setWidthF(2.0)
        pen.setCapStyle(Qt.RoundCap)
        pen.setJoinStyle(Qt.RoundJoin)
        painter.setPen(pen)
        if partial and not checked:
            painter.drawLine(QPointF(cx - 4, cy), QPointF(cx + 4, cy))
        else:
            painter.drawLine(QPointF(cx - 3.5, cy + 0.5), QPointF(cx - 1, cy + 3))
            painter.drawLine(QPointF(cx - 1, cy + 3), QPointF(cx + 4, cy - 3))
    else:
        painter.setBrush(QColor(theme.PALETTE["--input-bg"]))
        painter.setPen(QPen(QColor(theme.PALETTE["--border-strong"]), 1))
        painter.drawRoundedRect(box, 4, 4)
    painter.restore()


def _cell(row, key):
    """Đọc row[key] an toàn cho cả sqlite3.Row lẫn dict."""
    try:
        val = row[key]
    except (KeyError, IndexError):
        return ""
    return "" if val is None else val


def _is_number(s):
    try:
        float(s)
        return True
    except (TypeError, ValueError):
        return False


class DictTableModel(QAbstractTableModel):
    def __init__(self, columns, rows=None, pk=None, link_keys=None):
        super().__init__()
        # columns: list of (key, title, width, align[, formatter])
        # formatter (tùy chọn): callable(value) -> str để đổi cách hiển thị ô.
        self.columns = columns
        self.rows = list(rows or [])
        self.pk = pk
        self.link_keys = set(link_keys or ())
        self.checked = set()   # tập PK của các dòng đang được tick

    def rowCount(self, _parent=None):
        return len(self.rows)

    def columnCount(self, _parent=None):
        return len(self.columns)

    def _text(self, index):
        col = self.columns[index.column()]
        if col[0] == _CHECK_KEY:
            return ""
        val = _cell(self.rows[index.row()], col[0])
        fmt = col[4] if len(col) > 4 else None
        return fmt(val) if fmt else str(val)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        col = self.columns[index.column()]
        if col[0] == _CHECK_KEY:
            # Không có chữ; delegate tự vẽ checkbox. Chỉ trả canh giữa.
            if role == Qt.TextAlignmentRole:
                return _ALIGN["center"]
            return None
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
        self.checked.clear()          # dữ liệu mới → bỏ mọi tick cũ
        self.endResetModel()

    def sort(self, column, order=Qt.AscendingOrder):
        """Sắp xếp theo cột `column`. Cột toàn số → sắp theo SỐ; còn lại theo
        chữ (không phân biệt hoa/thường). Ô trống coi như nhỏ nhất."""
        if not self.rows or not (0 <= column < len(self.columns)):
            return
        key = self.columns[column][0]
        if key == _CHECK_KEY:         # cột chọn không sắp xếp
            return
        vals = [str(_cell(r, key)).strip() for r in self.rows]
        numeric = any(vals) and all(_is_number(v) for v in vals if v)

        def sort_key(row):
            v = str(_cell(row, key)).strip()
            if numeric:
                return float(v) if v else float("-inf")
            return v.lower()

        self.layoutAboutToBeChanged.emit()
        self.rows.sort(key=sort_key, reverse=(order == Qt.DescendingOrder))
        self.layoutChanged.emit()

    def row_at(self, r):
        return self.rows[r] if 0 <= r < len(self.rows) else None

    def id_at(self, r):
        row = self.row_at(r)
        return None if row is None or self.pk is None else row[self.pk]

    # ----------------------------------------------------------- checkbox chọn
    def is_row_checked(self, r):
        pk = self.id_at(r)
        return pk is not None and pk in self.checked

    def toggle_row(self, r):
        pk = self.id_at(r)
        if pk is None:
            return
        self.checked.discard(pk) if pk in self.checked else self.checked.add(pk)
        self.dataChanged.emit(self.index(r, 0),
                              self.index(r, self.columnCount() - 1))

    def set_all_checked(self, value):
        if value:
            self.checked = {pk for r in range(len(self.rows))
                            if (pk := self.id_at(r)) is not None}
        else:
            self.checked = set()
        if self.rows:
            self.dataChanged.emit(self.index(0, 0),
                                  self.index(len(self.rows) - 1,
                                             self.columnCount() - 1))

    def checked_rows(self):
        return [self.rows[r] for r in range(len(self.rows)) if self.is_row_checked(r)]


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
        model = index.model()
        if model.columns[index.column()][0] == _CHECK_KEY:
            _draw_checkbox(painter, option.rect, model.is_row_checked(index.row()))


class _CheckHeader(QHeaderView):
    """Header ngang có ô 'chọn tất cả' ở cột checkbox (3 trạng thái: none/some/all)."""

    def __init__(self, table):
        super().__init__(Qt.Horizontal, table)
        self._table = table
        self.setSectionsClickable(True)
        self.setHighlightSections(False)

    def paintSection(self, painter, rect, logical_index):
        super().paintSection(painter, rect, logical_index)
        if logical_index == self._table._check_col:
            m = self._table._model
            total = len(m.rows)
            n = sum(1 for r in range(total) if m.is_row_checked(r))
            _draw_checkbox(painter, rect, checked=(total > 0 and n == total),
                           partial=(0 < n < total))

    def mousePressEvent(self, e):
        if self.logicalIndexAt(e.position().toPoint()) == self._table._check_col:
            self._table._toggle_all()      # bấm header cột chọn → đảo chọn tất cả
            return                          # KHÔNG phát sectionClicked (tránh sort)
        super().mousePressEvent(e)


class DataTable(QTableView):
    """QTableView đã cấu hình sẵn: chọn/hover theo cả hàng, ẩn số dòng."""

    def __init__(self, columns, pk=None, stretch_key=None, on_double=None,
                 link_keys=None, on_link=None, checkable=False, check_width=None,
                 parent=None):
        super().__init__(parent)
        self.hover_row = -1
        # checkable → chèn cột checkbox ở đầu (cần pk để nhớ dòng nào được tick).
        self._checkable = bool(checkable)
        self._check_col = 0 if self._checkable else -1
        if self._checkable:
            # rộng cột = CHECK_COL_WIDTH (mặc định) hoặc `check_width` truyền vào;
            # checkbox canh trái theo CHECK_BOX_INSET (xem _draw_checkbox).
            cw = CHECK_COL_WIDTH if check_width is None else check_width
            columns = [(_CHECK_KEY, "", cw, "center")] + list(columns)
        self._model = DictTableModel(columns, pk=pk, link_keys=link_keys)
        self.setModel(self._model)
        if self._checkable:
            self.setHorizontalHeader(_CheckHeader(self))
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
        self.setShowGrid(False)
        self.setWordWrap(False)

        header = self.horizontalHeader()
        header.setHighlightSections(False)
        # Mặc định QHeaderView có minimumSectionSize ~30px (theo style) → đặt
        # width nhỏ hơn sẽ bị kẹp lại. Hạ min để cột checkbox thu gọn được.
        header.setMinimumSectionSize(24)
        for i, col in enumerate(columns):
            width = col[2] if len(col) > 2 else 120
            self.setColumnWidth(i, width)
            key = col[0]
            if key == _CHECK_KEY:
                header.setSectionResizeMode(i, QHeaderView.Fixed)
            elif stretch_key and key == stretch_key:
                header.setSectionResizeMode(i, QHeaderView.Stretch)
            else:
                header.setSectionResizeMode(i, QHeaderView.Interactive)

        # Sắp xếp theo cột khi bấm tiêu đề (bấm lại đảo chiều). Mặc định chưa
        # sort → giữ thứ tự do truy vấn trả về (mới nhất trước).
        self._sort_col = -1
        self._sort_order = Qt.AscendingOrder
        header.setSectionsClickable(True)
        header.sectionClicked.connect(self._on_header_clicked)

        if on_double:
            self.doubleClicked.connect(lambda idx: on_double(self._model.id_at(idx.row())))
        if self._on_link or self._checkable:
            self.clicked.connect(self._on_cell_clicked)

        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._context_menu)

    def set_rows(self, rows):
        self._model.set_rows(rows)
        if self._sort_col >= 0:            # giữ nguyên sort người dùng đã chọn
            self._model.sort(self._sort_col, self._sort_order)
        self._refresh_header_check()

    # ------------------------------------------------------------- checkbox chọn
    def checked_rows(self):
        """Danh sách dòng (dict/Row) đang được tick."""
        return self._model.checked_rows()

    def _on_cell_clicked(self, index):
        if index.column() == self._check_col:
            self._model.toggle_row(index.row())
            self._refresh_header_check()
            return
        self._maybe_link(index)

    def _toggle_all(self):
        rows = range(self._model.rowCount())
        all_on = bool(rows) and all(self._model.is_row_checked(r) for r in rows)
        self._model.set_all_checked(not all_on)
        self._refresh_header_check()

    def _refresh_header_check(self):
        if self._check_col >= 0:
            self.horizontalHeader().updateSection(self._check_col)

    def _on_header_clicked(self, col):
        if col == self._check_col:        # cột chọn: đã xử lý toggle-all ở header
            return
        order = (Qt.DescendingOrder
                 if self._sort_col == col and self._sort_order == Qt.AscendingOrder
                 else Qt.AscendingOrder)
        self._sort_col, self._sort_order = col, order
        header = self.horizontalHeader()
        header.setSortIndicatorShown(True)
        header.setSortIndicator(col, order)
        self._model.sort(col, order)

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
