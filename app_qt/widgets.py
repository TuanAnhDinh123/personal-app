"""Thư viện widget dựng sẵn (PySide6) — bản thay thế cho app/ui/widgets.py.

Giữ NGUYÊN tên hàm (file_row, text_row, text_area, dropdown, checkbox,
section_label, hint, button…) và giao diện `.get()/.set()` của giá trị trả về,
để khi port từng tool sang chỉ phải sửa tối thiểu.

Khác biệt với bản Tk: mỗi hàm nhận `parent` là QWidget CÓ SẴN layout dạng cột
(QVBoxLayout). Widget tự thêm mình vào layout đó — y như .pack() trước đây.
"""
import os

from PySide6.QtCore import QRectF, QSize, Qt, Signal
from PySide6.QtGui import QColor, QIcon, QImage, QPainter, QPen, QPixmap
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QFileDialog, QFrame,
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


def add_shadow(widget, blur=48, dy=12, alpha=70):
    """No-op: bóng đổ đã được gỡ khỏi toàn bộ thẻ/hộp thoại. Giữ hàm (và tham số)
    để các nơi gọi cũ không phải đổi; lề chừa quanh thẻ vẫn giữ nguyên bố cục."""
    return widget


# Trước đây là vùng đệm quanh thẻ để chứa bóng; bóng đã gỡ nên = 0. Giữ hằng số
# (và các chỗ '+ CARD_PAD') để không phải sửa loạt công thức căn lề: giờ mép
# NHÌN THẤY của thẻ = mép widget, nên cộng thêm 0.
#
# LƯU Ý QUAN TRỌNG: padding thẻ→nội dung = contentsMargins của LAYOUT con, KHÔNG
# phải pad. Qt để layout con GHI ĐÈ contentsMargins của widget, nên hai lề không
# cộng dồn. Viền thẻ vẽ tại mép widget (pad=0) → lề layout con hiện đúng là
# padding thấy được. (Khi pad>0 như thẻ Trang chủ: padding = lề layout − pad.)
CARD_PAD = 0


class Card(QFrame):
    """Thẻ trắng bo góc, viền 1px, TỰ VẼ bằng QPainter (không QSS #Card).

    `pad` = khoảng CHỪA NGOÀI viền thẻ (mặc định 0). Trang chủ dùng pad>0 để tạo
    khe hở giữa các thẻ trong lưới. Viền vẽ thụt vào `pad`; padding bên trong do
    contentsMargins của layout con quyết định (Qt cho layout con ghi đè lề widget
    → hai lề không cộng dồn, xem ghi chú ở CARD_PAD).
    """

    def __init__(self, parent=None, *, pad=CARD_PAD, radius=16, dy=6,
                 alpha=7, tint=(15, 23, 42), fill=None, border=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.NoFrame)
        self.setAttribute(Qt.WA_StyledBackground, False)
        self._pad = pad
        self._radius = radius
        self._dy = dy
        self._alpha = alpha               # độ mờ MỖI lớp bóng (cộng dồn khi chồng)
        self._spread = max(1, pad - dy)   # số lớp; +dy để lớp dưới cùng không lố pad
        self._tint = tint
        self._fill = QColor(fill or theme.CARD_BG)
        # Không còn bóng → viền phải rõ hơn để thẻ tách khỏi nền (nếu không thẻ
        # trắng trên nền xám rất nhạt gần như 'chìm', trông vỡ bố cục).
        self._border = QColor(border or theme.BORDER_STRONG)
        super().setContentsMargins(pad, pad, pad, pad)

    def _card_rect(self):
        r = self.rect()
        return QRectF(r.adjusted(self._pad, self._pad, -self._pad, -self._pad))

    def _colors(self):
        """(fill, border) — tách riêng để lớp con đổi màu theo trạng thái."""
        return self._fill, self._border

    def paintEvent(self, _e):
        r = self._card_rect()
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)
        p.setPen(Qt.NoPen)
        # Bóng đã gỡ — chỉ vẽ nền thẻ. Vẫn chừa CARD_PAD quanh thẻ để mọi chỗ
        # căn theo mép thẻ (dùng CARD_PAD) giữ nguyên bố cục.
        # Nền thẻ + viền 1px
        fill, border = self._colors()
        p.setBrush(fill)
        pen = QPen(border)
        pen.setWidthF(1.0)
        p.setPen(pen)
        p.drawRoundedRect(r, self._radius, self._radius)


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
        v = str(value) if value is not None else ""
        i = self._w.findText(v)
        if i >= 0:
            self._w.setCurrentIndex(i)
        elif self._w.isEditable():
            # Combobox cho gõ tay (vd ô Model): giữ giá trị dù chưa có trong list.
            self._w.setCurrentText(v)

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


class _FetchCombo(QComboBox):
    """Combobox tự TẢI danh sách lựa chọn khi người dùng bấm mở.

    `fetch_fn()` chạy ở luồng nền (tránh treo UI), trả về list[str] hoặc ném lỗi.
    Cho phép gõ tay (editable) để vẫn nhập được giá trị tùy ý. Trạng thái tải /
    lỗi hiển thị qua nhãn gắn bằng `bind_status`.
    """

    # Kết quả từ luồng nền → phát tín hiệu để cập nhật UI ở luồng chính (an toàn).
    _loaded = Signal(list)
    _failed = Signal(str)

    def __init__(self, fetch_fn, parent=None):
        super().__init__(parent)
        self.setEditable(True)
        self.setInsertPolicy(QComboBox.NoInsert)
        self._fetch_fn = fetch_fn
        self._loading = False
        self._fetched_once = False
        self._status = None
        self._loaded.connect(self._on_loaded)
        self._failed.connect(self._on_failed)

    def bind_status(self, label):
        self._status = label

    def reset(self):
        """Cho phép tải lại danh sách ở lần mở kế tiếp (vd khi API key đổi)."""
        self._fetched_once = False

    def showPopup(self):
        if not self._loading and not self._fetched_once:
            self._start_fetch()
        super().showPopup()

    def _start_fetch(self):
        import threading
        self._loading = True
        self._set_status("Đang tải danh sách model…")

        def work():
            try:
                models = list(self._fetch_fn() or [])
                self._loaded.emit(models)
            except Exception as exc:  # noqa: BLE001 — báo mọi lỗi lại cho UI
                self._failed.emit(str(exc))

        threading.Thread(target=work, daemon=True).start()

    def _on_loaded(self, models):
        self._loading = False
        self._fetched_once = True
        current = self.currentText()
        self.blockSignals(True)
        self.clear()
        self.addItems(models)
        if current:
            i = self.findText(current)
            self.setCurrentIndex(i) if i >= 0 else self.setEditText(current)
        self.blockSignals(False)
        self._set_status("" if models else "Không có model khả dụng cho key này.")
        # Người dùng vẫn đang mở danh sách → mở lại để thấy list vừa tải.
        if self.view().isVisible():
            self.hidePopup()
            self.showPopup()

    def _on_failed(self, msg):
        self._loading = False
        self._set_status(f"Không tải được danh sách model: {msg}")
        if self.view().isVisible():
            self.hidePopup()

    def _set_status(self, text):
        if self._status is None:
            return
        self._status.setText(text)
        self._status.setVisible(bool(text))


def model_select_row(parent, label, fetch_fn):
    """Ô chọn Model: combobox bấm vào sẽ tự gọi `fetch_fn()` lấy danh sách model.

    fetch_fn() trả về list[str] (chạy nền). Cho gõ tay nếu muốn model ngoài list.
    Trả về ChoiceValue (.get()/.set()) như dropdown; `.widget.reset()` để buộc
    tải lại ở lần mở sau.
    """
    block, v = _field_block(parent, label)
    combo = _FetchCombo(fetch_fn, block)
    v.addWidget(combo)
    status = QLabel("", block)
    status.setObjectName("Hint")
    status.setWordWrap(True)
    status.hide()
    v.addWidget(status)
    combo.bind_status(status)
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
