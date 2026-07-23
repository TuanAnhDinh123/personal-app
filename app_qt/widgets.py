"""Thư viện widget dựng sẵn (PySide6) — bản thay thế cho app/ui/widgets.py.

Giữ NGUYÊN tên hàm (file_row, text_row, text_area, dropdown, checkbox,
section_label, hint, button…) và giao diện `.get()/.set()` của giá trị trả về,
để khi port từng tool sang chỉ phải sửa tối thiểu.

Khác biệt với bản Tk: mỗi hàm nhận `parent` là QWidget CÓ SẴN layout dạng cột
(QVBoxLayout). Widget tự thêm mình vào layout đó — y như .pack() trước đây.
"""
import os

from PySide6.QtCore import QSize, Qt
from PySide6.QtGui import QColor, QIcon, QImage, QPainter, QPixmap
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QFileDialog, QFrame, QGraphicsDropShadowEffect,
    QHBoxLayout, QLabel, QLineEdit, QPushButton, QTextEdit, QVBoxLayout, QWidget,
)

from app_qt import theme


# ----- Icon line (SVG) tô màu theo yêu cầu (crisp, thay cho emoji nhòe) -----
_icon_cache = {}


def svg_pixmap(name, color, px=18):
    """Render icon line SVG (assets/icons/<name>.svg) rồi tô thành `color`.

    Đơn sắc: vẽ hình rồi dùng SourceIn phủ màu → icon sắc nét ở mọi kích cỡ.
    """
    key = (name, color, px)
    if key in _icon_cache:
        return _icon_cache[key]
    path = os.path.join(os.path.dirname(__file__), "assets", "icons", f"{name}.svg")
    if not os.path.isfile(path):
        return QPixmap()
    r = QSvgRenderer(path)
    dpr = 2
    img = QImage(px * dpr, px * dpr, QImage.Format_ARGB32_Premultiplied)
    img.fill(Qt.transparent)
    p = QPainter(img)
    r.render(p)
    p.setCompositionMode(QPainter.CompositionMode_SourceIn)
    p.fillRect(img.rect(), QColor(color))
    p.end()
    pm = QPixmap.fromImage(img)
    pm.setDevicePixelRatio(dpr)
    _icon_cache[key] = pm
    return pm


def svg_icon(name, color, px=18):
    return QIcon(svg_pixmap(name, color, px))


def add_shadow(widget, blur=28, dy=6, alpha=38):
    """Đổ bóng mềm cho thẻ (QSS không hỗ trợ box-shadow → dùng hiệu ứng Qt)."""
    eff = QGraphicsDropShadowEffect(widget)
    eff.setBlurRadius(blur)
    eff.setXOffset(0)
    eff.setYOffset(dy)
    eff.setColor(QColor(31, 41, 55, alpha))   # xám xanh, độ mờ nhẹ
    widget.setGraphicsEffect(eff)
    return widget


# ---------------------------------------------------------------- tiện ích
def _col(parent):
    """Trả về layout cột của `parent`; tạo nếu chưa có."""
    lay = parent.layout()
    if lay is None:
        lay = QVBoxLayout(parent)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)
    return lay


def _field_block(parent, label_text):
    """Khối 1 trường: nhãn nhỏ ở trên + vùng chứa input bên dưới. Trả về block."""
    block = QWidget(parent)
    v = QVBoxLayout(block)
    v.setContentsMargins(0, 6, 0, 6)
    v.setSpacing(5)
    lbl = QLabel(label_text, block)
    lbl.setObjectName("FieldLabel")
    v.addWidget(lbl)
    _col(parent).addWidget(block)
    return block, v


# ---------------------------------------------------------------- value holders
class StringValue:
    """Bọc QLineEdit cho giống tk.StringVar: .get() / .set()."""

    def __init__(self, line_edit: QLineEdit):
        self._w = line_edit

    def get(self):
        return self._w.text()

    def set(self, value):
        self._w.setText(str(value) if value is not None else "")

    @property
    def widget(self):
        return self._w


class TextValue:
    """Bọc QTextEdit cho giống ô nhập nhiều dòng."""

    def __init__(self, text_edit: QTextEdit):
        self._w = text_edit

    def get(self):
        return self._w.toPlainText()

    def set(self, value):
        self._w.setPlainText(str(value) if value is not None else "")

    def get_html(self):
        return self._w.toHtml()

    @property
    def widget(self):
        return self._w


class BoolValue:
    """Bọc QCheckBox cho giống tk.BooleanVar."""

    def __init__(self, checkbox: QCheckBox):
        self._w = checkbox

    def get(self):
        return self._w.isChecked()

    def set(self, value):
        self._w.setChecked(bool(value))

    @property
    def widget(self):
        return self._w


class ChoiceValue:
    """Bọc QComboBox."""

    def __init__(self, combo: QComboBox):
        self._w = combo

    def get(self):
        return self._w.currentText()

    def set(self, value):
        i = self._w.findText(str(value))
        if i >= 0:
            self._w.setCurrentIndex(i)

    @property
    def widget(self):
        return self._w


# ---------------------------------------------------------------- nhãn / ghi chú
def section_label(parent, text):
    lbl = QLabel(text, parent)
    lbl.setObjectName("SectionLabel")
    lbl.setContentsMargins(0, 10, 0, 4)
    _col(parent).addWidget(lbl)
    return lbl


def hint(parent, text):
    lbl = QLabel(text, parent)
    lbl.setObjectName("Hint")
    lbl.setWordWrap(True)
    lbl.setContentsMargins(0, 2, 0, 2)
    _col(parent).addWidget(lbl)
    return lbl


# ---------------------------------------------------------------- ô nhập
def text_row(parent, label, placeholder=""):
    """Ô nhập chữ 1 dòng. Trả về StringValue (.get()/.set())."""
    block, v = _field_block(parent, label)
    edit = QLineEdit(block)
    if placeholder:
        edit.setPlaceholderText(placeholder)
    v.addWidget(edit)
    return StringValue(edit)


def digit_entry(parent, label, placeholder=""):
    """Ô nhập chỉ cho phép chữ số."""
    from PySide6.QtGui import QIntValidator
    block, v = _field_block(parent, label)
    edit = QLineEdit(block)
    edit.setValidator(QIntValidator(0, 10_000_000, edit))
    if placeholder:
        edit.setPlaceholderText(placeholder)
    v.addWidget(edit)
    return StringValue(edit)


def text_area(parent, label, value="", height=8):
    """Ô nhập chữ nhiều dòng. Trả về TextValue."""
    block, v = _field_block(parent, label)
    edit = QTextEdit(block)
    edit.setAcceptRichText(False)
    edit.setFixedHeight(max(60, height * 20))
    if value:
        edit.setPlainText(value)
    v.addWidget(edit)
    return TextValue(edit)


def dropdown(parent, label, options):
    """Danh sách chọn (combobox). Trả về ChoiceValue."""
    block, v = _field_block(parent, label)
    combo = QComboBox(block)
    combo.addItems([str(o) for o in options])
    v.addWidget(combo)
    return ChoiceValue(combo)


def checkbox(parent, label, checked=True):
    """Công tắc bật/tắt. Trả về BoolValue."""
    cb = QCheckBox(label, parent)
    cb.setChecked(checked)
    _col(parent).addWidget(cb)
    return BoolValue(cb)


def file_row(parent, label, mode="file"):
    """Ô chọn file/thư mục: nhãn + ô đường dẫn + nút Chọn. Trả về StringValue.

    mode: "file" (mở file) | "folder" (chọn thư mục) | "save" (lưu file .xlsx).
    """
    block, v = _field_block(parent, label)
    row = QHBoxLayout()
    row.setSpacing(8)
    edit = QLineEdit(block)
    row.addWidget(edit, 1)

    def browse():
        if mode == "folder":
            path = QFileDialog.getExistingDirectory(parent, "Chọn thư mục")
        elif mode == "save":
            path, _ = QFileDialog.getSaveFileName(
                parent, "Lưu file", "", "Excel (*.xlsx)")
        else:
            path, _ = QFileDialog.getOpenFileName(parent, "Chọn file")
        if path:
            edit.setText(path)

    btn = button(block, "Chọn…", variant="neutral", command=browse)
    row.addWidget(btn)
    v.addLayout(row)
    return StringValue(edit)


def export_target_row(parent, label):
    """Chọn ĐÍCH xuất: nút Thư mục hoặc File Excel. Trả về StringValue."""
    block, v = _field_block(parent, label)
    row = QHBoxLayout()
    row.setSpacing(8)
    edit = QLineEdit(block)
    row.addWidget(edit, 1)

    def pick_folder():
        p = QFileDialog.getExistingDirectory(parent, "Chọn thư mục")
        if p:
            edit.setText(p)

    def pick_file():
        p, _ = QFileDialog.getOpenFileName(parent, "Chọn file Excel", "", "Excel (*.xlsx)")
        if p:
            edit.setText(p)

    row.addWidget(button(block, "Thư mục", variant="neutral", command=pick_folder))
    row.addWidget(button(block, "File Excel", variant="neutral", command=pick_file))
    v.addLayout(row)
    return StringValue(edit)


# ---------------------------------------------------------------- nút
# emoji nhỏ đứng trước chữ, thay cho icon vector vẽ tay ở bản Tk.
def button(parent, text="", variant="primary", icon=None, command=None):
    """Nút chuẩn của app. variant ∈ primary/success/danger/warning/info/neutral.
    icon: tên icon line trong assets/icons (vẽ sắc nét, tô theo màu chữ nút)."""
    btn = QPushButton(text, parent)
    btn.setProperty("variant", variant)
    btn.setCursor(Qt.PointingHandCursor)
    if icon:
        fg = theme.TEXT if variant == "neutral" else "#ffffff"
        btn.setIcon(svg_icon(icon, fg, 16))
        btn.setIconSize(QSize(16, 16))
    if command is not None:
        btn.clicked.connect(lambda *_: command())
    return btn


# ---------------------------------------------------------------- chip icon
def icon_badge(parent, emoji, color, size=44):
    """Chip vuông bo góc: nền tông nhạt của `color` + icon line màu `color`.

    `emoji` là icon cũ của tool (vd "🤖"); tự tra sang icon line qua icons.name_for.
    """
    from app_qt.icons import name_for
    lbl = QLabel(parent)
    lbl.setFixedSize(size, size)
    lbl.setAlignment(Qt.AlignCenter)
    lbl.setPixmap(svg_pixmap(name_for(emoji), color, int(size * 0.52)))
    r, g, b = _hex_to_rgb(color)
    lbl.setStyleSheet(
        f"background: rgba({r},{g},{b},0.14); border-radius: {int(size * 0.28)}px;")
    return lbl


def _hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


# ---------------------------------------------------------------- vùng cuộn
def scroll_area(child: QWidget):
    """Bọc `child` trong vùng cuộn dọc (nội dung co giãn theo bề ngang).

    Viewport của QScrollArea mặc định tự vẽ nền xám hệ thống (#efefef) → phải
    ép trong suốt để nó ăn theo nền cha (một nền duy nhất cho toàn trang).
    """
    from PySide6.QtWidgets import QScrollArea
    sa = QScrollArea()
    sa.setWidgetResizable(True)
    sa.setFrameShape(QFrame.NoFrame)
    sa.setStyleSheet(
        "QScrollArea { background: transparent; border: none; }"
        "QScrollArea > QWidget { background: transparent; }")
    sa.viewport().setAutoFillBackground(False)
    sa.setWidget(child)
    # setWidget bật autoFillBackground của child (nền Window #efefef) → tắt đi để
    # child trong suốt, ăn theo một nền duy nhất của trang.
    child.setAutoFillBackground(False)
    return sa
