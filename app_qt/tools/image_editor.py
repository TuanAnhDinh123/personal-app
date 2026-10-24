"""AI Image Editor: sửa ảnh bằng cách mô tả thay đổi bằng lời (Gemini).

Màn hình cố ý tối giản cho người không rành: một ô gõ yêu cầu (kèm chip prompt
mẫu), ảnh gốc bên trái, ảnh kết quả bên phải để so trước/sau, nút tải về.
Logic gọi AI nằm ở app.core.image_edit.

Không có mục ở sidebar (`show_in_sidebar = False`) — chỉ mở từ thẻ ở Trang
chủ, để sidebar gọn cho các việc hằng ngày.

Ảnh vào: bấm khung trái để chọn file, kéo-thả file/ảnh vào khung, hoặc Ctrl+V
ảnh đang có trong clipboard (vd vừa chụp màn hình). Ảnh JPEG/PNG/WEBP nhỏ hơn
MAX_UPLOAD_BYTES gửi NGUYÊN byte gốc cho khỏi giảm chất lượng; định dạng khác,
ảnh quá lớn hay ảnh điện thoại có cờ xoay EXIF thì chuyển sang PNG/JPEG đúng
chiều trước khi gửi.

Lời gọi AI chạy ở luồng nền (Task). Bấm Cancel chỉ BỎ QUA kết quả của lượt đó
(urllib không ngắt ngang được một request đang chờ) — mỗi lượt có mã riêng, kết
quả của lượt đã hủy về sau bị lờ đi.
"""
import os
import re
import time

from PySide6.QtCore import (
    QBuffer, QByteArray, QEvent, QIODevice, QObject, QPoint, QRect, QRectF,
    QSize, QStandardPaths, Qt, QTimer, Signal,
)
from PySide6.QtGui import (
    QColor, QImage, QImageIOHandler, QImageReader, QKeySequence, QPainter, QPen,
    QPixmap, QShortcut, QTextCursor,
)
from PySide6.QtWidgets import (
    QApplication, QFileDialog, QFrame, QHBoxLayout, QLabel, QLayout,
    QSizePolicy, QStackedLayout, QVBoxLayout, QWidget,
)

from app.core import image_edit, settings
from app_qt import theme, widgets
from app_qt.base_tool import BaseTool
from app_qt.components.task import Task

# Cạnh dài tối đa khi phải chuyển định dạng ảnh trước khi gửi (model cũng chỉ
# xử lý tới cỡ 4K, gửi lớn hơn chỉ tốn băng thông).
_MAX_EDGE = 4096

_OPEN_FILTER = "Images (*.png *.jpg *.jpeg *.webp *.bmp *.gif *.tif *.tiff)"
_IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif", ".tif", ".tiff"}

# Chỗ còn để trống trong prompt mẫu, vd "[object]".
_PLACEHOLDER = re.compile(r"\[[^\]\n]+\]")

# Dòng chữ đổi dần trong lúc chờ (mỗi dòng ~6 giây, dừng ở dòng cuối) để người
# dùng thấy app vẫn đang chạy chứ không treo.
_LOADING_LINES = [
    "Sending your image to the AI…",
    "The AI is studying your image…",
    "Applying your edit…",
    "Adding the finishing touches…",
    "Still working — bigger images take a little longer…",
]
_LINE_SECONDS = 6


# ------------------------------------------------------------ xử lý ảnh
def _read_image(data):
    """bytes → (QImage đã xoay đúng chiều theo EXIF, có phải xoay không)."""
    buf = QBuffer()
    buf.setData(QByteArray(data))
    buf.open(QIODevice.ReadOnly)
    reader = QImageReader(buf)
    reader.setAutoTransform(True)
    image = reader.read()
    rotated = reader.transformation() != QImageIOHandler.TransformationNone
    return image, rotated


def _save_bytes(image, fmt, quality=-1):
    buf = QBuffer()
    buf.open(QIODevice.WriteOnly)
    image.save(buf, fmt, quality)
    return bytes(buf.data())


def _encode(image):
    """QImage → (bytes, mime) vừa trần dung lượng gửi đi.

    Ảnh có kênh trong suốt giữ PNG nếu vừa; không thì (hoặc PNG quá nặng) lót
    nền trắng rồi nén JPEG, hạ dần chất lượng tới khi vừa.
    """
    if max(image.width(), image.height()) > _MAX_EDGE:
        image = image.scaled(_MAX_EDGE, _MAX_EDGE, Qt.KeepAspectRatio,
                             Qt.SmoothTransformation)
    if image.hasAlphaChannel():
        data = _save_bytes(image, "PNG")
        if len(data) <= image_edit.MAX_UPLOAD_BYTES:
            return data, "image/png"
        flat = QImage(image.size(), QImage.Format_RGB32)
        flat.fill(Qt.white)
        p = QPainter(flat)
        p.drawImage(0, 0, image)
        p.end()
        image = flat
    for quality in (92, 85, 75, 60):
        data = _save_bytes(image, "JPEG", quality)
        if len(data) <= image_edit.MAX_UPLOAD_BYTES:
            break
    return data, "image/jpeg"


def _prepare_upload(data, ext):
    """Byte file ảnh → (byte gửi đi, mime, QImage để xem). Ném ValueError nếu
    không đọc được ảnh."""
    image, rotated = _read_image(data)
    if image.isNull():
        raise ValueError("unreadable image")
    mime = image_edit.UPLOAD_MIME.get(ext)
    if mime and not rotated and len(data) <= image_edit.MAX_UPLOAD_BYTES:
        return data, mime, image
    payload, mime = _encode(image)
    return payload, mime, image


def _local_image_path(mime_data):
    """File ảnh đầu tiên trong dữ liệu kéo-thả / clipboard (nếu có)."""
    if not mime_data.hasUrls():
        return ""
    for url in mime_data.urls():
        path = url.toLocalFile()
        if path and os.path.splitext(path)[1].lower() in _IMAGE_EXTS:
            return path
    return ""


# ------------------------------------------------------------ widget nhỏ
class _EventHook(QObject):
    """Event filter gọi `fn(obj, event)`; fn trả True để nuốt sự kiện."""

    def __init__(self, fn, parent):
        super().__init__(parent)
        self._fn = fn

    def eventFilter(self, obj, event):
        return bool(self._fn(obj, event))


class _FlowLayout(QLayout):
    """Xếp widget theo hàng, hết chỗ thì tự xuống dòng (cho dải chip prompt mẫu
    — cửa sổ hẹp không bị tràn ngang)."""

    def __init__(self, parent=None, spacing=8):
        super().__init__(parent)
        self._items = []
        self._gap = spacing
        self.setContentsMargins(0, 0, 0, 0)

    def addItem(self, item):
        self._items.append(item)

    def count(self):
        return len(self._items)

    def itemAt(self, index):
        return self._items[index] if 0 <= index < len(self._items) else None

    def takeAt(self, index):
        return self._items.pop(index) if 0 <= index < len(self._items) else None

    def expandingDirections(self):
        return Qt.Orientation(0)

    def hasHeightForWidth(self):
        return True

    def heightForWidth(self, width):
        return self._arrange(QRect(0, 0, width, 0), apply=False)

    def setGeometry(self, rect):
        super().setGeometry(rect)
        self._arrange(rect, apply=True)

    def sizeHint(self):
        return self.minimumSize()

    def minimumSize(self):
        size = QSize()
        for item in self._items:
            size = size.expandedTo(item.minimumSize())
        return size

    def _arrange(self, rect, apply):
        x, y, line_h = rect.x(), rect.y(), 0
        for item in self._items:
            hint = item.sizeHint()
            if x + hint.width() > rect.right() + 1 and line_h > 0:
                x = rect.x()
                y += line_h + self._gap
                line_h = 0
            if apply:
                item.setGeometry(QRect(QPoint(x, y), hint))
            x += hint.width() + self._gap
            line_h = max(line_h, hint.height())
        return y + line_h - rect.y()


class _Spinner(QWidget):
    """Vòng xoay tự vẽ (vòng nền nhạt + cung màu nhấn quay đều)."""

    def __init__(self, size=46, parent=None):
        super().__init__(parent)
        self.setFixedSize(size, size)
        self._angle = 0
        self._timer = QTimer(self)
        self._timer.setInterval(16)
        self._timer.timeout.connect(self._tick)

    def start(self):
        self._timer.start()

    def stop(self):
        self._timer.stop()

    def _tick(self):
        self._angle = (self._angle + 6) % 360
        self.update()

    def paintEvent(self, _e):
        width = 4
        rect = QRectF(width / 2 + 1, width / 2 + 1,
                      self.width() - width - 2, self.height() - width - 2)
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)
        pen = QPen(QColor(theme.PALETTE["--accent-soft-hi"]), width)
        p.setPen(pen)
        p.drawEllipse(rect)
        pen.setColor(QColor(theme.ACCENT))
        pen.setCapStyle(Qt.RoundCap)
        p.setPen(pen)
        p.drawArc(rect, int((90 - self._angle) * 16), -110 * 16)


class _ImageView(QLabel):
    """Hiện ảnh vừa khung (giữ tỉ lệ), tự vẽ lại khi khung đổi kích thước."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAlignment(Qt.AlignCenter)
        # Ignored: pixmap to không được đẩy khung giãn ra theo ảnh.
        self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)
        self.setStyleSheet("background: transparent;")
        self._pm = QPixmap()

    def set_image(self, image):
        self._pm = QPixmap.fromImage(image)
        self._redraw()

    def clear_image(self):
        self._pm = QPixmap()
        self.clear()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._redraw()

    def _redraw(self):
        if self._pm.isNull():
            return
        box = self.size() - QSize(16, 16)
        if box.width() < 10 or box.height() < 10:
            return
        pm = self._pm
        if pm.width() > box.width() or pm.height() > box.height():
            pm = pm.scaled(box, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.setPixmap(pm)


class _Stage(QFrame):
    """Khung chứa ảnh (nền xám, bo góc), các 'mặt' xếp chồng bằng QStackedLayout.

    `droppable=True` → nhận kéo-thả file/ảnh và bấm vào khi còn trống.
    """
    file_dropped = Signal(str)
    image_dropped = Signal(QImage)
    clicked = Signal()

    def __init__(self, droppable=False, parent=None):
        super().__init__(parent)
        self.setObjectName("ImageStage")
        self.setMinimumHeight(160)
        self.stack = QStackedLayout(self)
        self._droppable = droppable
        self._empty = True
        self.accepting = droppable
        self.setAcceptDrops(droppable)

    def _set_prop(self, name, value):
        self.setProperty(name, value)
        self.style().unpolish(self)
        self.style().polish(self)

    def set_empty(self, empty):
        self._empty = empty
        if self._droppable:
            self._set_prop("drop", "true" if empty else "false")
            self.setCursor(Qt.PointingHandCursor if empty else Qt.ArrowCursor)

    def _wanted(self, mime_data):
        return self.accepting and (mime_data.hasImage()
                                   or bool(_local_image_path(mime_data)))

    def dragEnterEvent(self, e):
        if self._wanted(e.mimeData()):
            e.acceptProposedAction()
            self._set_prop("dragOver", "true")

    def dragLeaveEvent(self, e):
        self._set_prop("dragOver", "false")
        super().dragLeaveEvent(e)

    def dropEvent(self, e):
        self._set_prop("dragOver", "false")
        md = e.mimeData()
        if not self._wanted(md):
            return
        e.acceptProposedAction()
        path = _local_image_path(md)
        if path:
            self.file_dropped.emit(path)
        else:
            data = md.imageData()
            if isinstance(data, QPixmap):
                data = data.toImage()
            self.image_dropped.emit(data if isinstance(data, QImage) else QImage())

    def mousePressEvent(self, e):
        if (self._droppable and self._empty and self.accepting
                and e.button() == Qt.LeftButton):
            self.clicked.emit()
        super().mousePressEvent(e)


def _empty_face(icon, title, lines):
    """Mặt 'chưa có ảnh': icon mờ + tiêu đề + vài dòng hướng dẫn, canh giữa."""
    face = QWidget()
    face.setStyleSheet("background: transparent;")
    lay = QVBoxLayout(face)
    lay.setSpacing(6)
    lay.addStretch(1)
    ic = QLabel()
    ic.setAlignment(Qt.AlignCenter)
    ic.setPixmap(widgets.svg_pixmap(icon, theme.PALETTE["--text-faint"], 40))
    lay.addWidget(ic)
    lay.addSpacing(4)
    t = QLabel(title)
    t.setObjectName("StageTitle")
    t.setAlignment(Qt.AlignCenter)
    lay.addWidget(t)
    for line in lines:
        s = QLabel(line)
        s.setObjectName("StageSub")
        s.setAlignment(Qt.AlignCenter)
        s.setWordWrap(True)
        lay.addWidget(s)
    lay.addStretch(1)
    return face


class _Pane(widgets.Card):
    """Thẻ một bên (Original / Result): tiêu đề + thông tin ảnh, khung ảnh,
    hàng nút bên dưới."""

    def __init__(self, title, stage, parent=None):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 14, 16, 14)
        lay.setSpacing(10)
        head = QHBoxLayout()
        head.setSpacing(10)
        t = QLabel(title)
        t.setObjectName("SectionLabel")
        head.addWidget(t)
        self.meta = QLabel("")
        self.meta.setObjectName("StageMeta")
        # Dòng thông tin dài (tên file) không được ép thẻ nở ngang.
        self.meta.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        head.addWidget(self.meta, 1)
        lay.addLayout(head)
        lay.addWidget(stage, 1)
        self.note = QLabel("")
        self.note.setObjectName("Hint")
        self.note.setWordWrap(True)
        self.note.hide()
        lay.addWidget(self.note)
        self.footer = QHBoxLayout()
        self.footer.setSpacing(8)
        lay.addLayout(self.footer)
        # Hai thẻ luôn chia đôi bề ngang (so trước/sau cho công bằng) — không
        # để hàng nút dài hơn đẩy một bên nở ra, ép cả trang rộng quá cửa sổ.
        self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)


# ------------------------------------------------------------ tool
class ImageEditorTool(BaseTool):
    name = "AI Image Editor"
    description = ("Edit a photo by describing the change — remove objects, "
                   "swap backgrounds, fix colours.")
    icon = "🖼"
    category = "Files & Documents"
    order = 40
    fills_height = True
    show_in_sidebar = False

    def build(self, parent=None):
        self._src = None        # dict: data, mime, name (không đuôi), folder
        self._result = None     # image_edit.EditResult
        self._busy = False
        self._req_id = 0        # mã lượt sửa hiện tại; lượt cũ hơn bị bỏ qua
        self._tasks = set()     # giữ tham chiếu QThread tới khi chạy xong
        self._started = 0.0
        self._retry_note = ""
        self._last_dir = ""

        page = QWidget(parent)
        self._page = page
        outer = QVBoxLayout(page)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(14)
        self._build_controls(outer)
        self._build_panes(outer)
        self._install_shortcuts()

        self._tick_timer = QTimer(page)
        self._tick_timer.setInterval(1000)
        self._tick_timer.timeout.connect(self._tick)
        # Model/độ phân giải đổi ở Settings → cập nhật dòng thông tin mỗi lần
        # trang hiện lại.
        page.installEventFilter(_EventHook(
            lambda _o, e: e.type() == QEvent.Show and self._show_model(), page))
        self._show_model()
        self._refresh()
        return page

    def build_body(self, parent):
        pass

    # ------------------------------------------------------------ dựng giao diện
    def _build_controls(self, outer):
        card = widgets.Card()
        lay = QVBoxLayout(card)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(8)

        label = QLabel("What should change?")
        label.setObjectName("FieldLabel")
        lay.addWidget(label)

        row = QHBoxLayout()
        row.setSpacing(12)
        self._prompt = widgets.TextEdit()
        self._prompt.setAcceptRichText(False)
        self._prompt.setPlaceholderText(
            "Describe the change, e.g. “Remove the person on the left and fill "
            "in the background naturally.” Any language works.")
        self._prompt.set_visible_rows(3)
        self._prompt.installEventFilter(_EventHook(self._prompt_keys, self._prompt))
        row.addWidget(self._prompt, 1)

        side = QVBoxLayout()
        side.setSpacing(6)
        side.addStretch(1)
        self._apply_btn = widgets.button(card, "Apply edit", variant="primary",
                                         icon="sparkles", command=self._apply)
        self._apply_btn.setToolTip("Ctrl+Enter")
        side.addWidget(self._apply_btn)
        self._model_lbl = QLabel("")
        self._model_lbl.setObjectName("StageMeta")
        self._model_lbl.setAlignment(Qt.AlignCenter)
        self._model_lbl.setToolTip("Change the model and resolution in Settings.")
        side.addWidget(self._model_lbl)
        row.addLayout(side)
        lay.addLayout(row)

        chips_row = QHBoxLayout()
        chips_row.setSpacing(10)
        quick = QLabel("Quick start:")
        quick.setObjectName("FieldLabel")
        chips_row.addWidget(quick, 0, Qt.AlignTop)
        holder = QWidget()
        flow = _FlowLayout(holder, spacing=6)
        for label_text, template in image_edit.PROMPT_TEMPLATES:
            chip = widgets.button(holder, label_text, variant="")
            chip.setObjectName("PromptChip")
            chip.setToolTip(template)
            chip.clicked.connect(lambda *_, t=template: self._use_template(t))
            flow.addWidget(chip)
        chips_row.addWidget(holder, 1)
        lay.addLayout(chips_row)

        self._status = QLabel(
            "Tip: say what to change AND what to keep, and be specific (“the red "
            "car on the left”, not “the car”). Not quite right? Apply again for a "
            "new attempt, or refine the wording. Each attempt is billed as one image.")
        self._status.setObjectName("Hint")
        self._status.setWordWrap(True)
        lay.addWidget(self._status)
        outer.addWidget(card)

    def _build_panes(self, outer):
        row = QHBoxLayout()
        row.setContentsMargins(widgets.CARD_PAD, 0, widgets.CARD_PAD, 0)
        row.setSpacing(16)

        # --- trái: ảnh gốc ---
        self._src_stage = _Stage(droppable=True)
        self._src_stage.stack.addWidget(_empty_face(
            "image", "Drop an image here",
            ["or click to browse · paste with Ctrl+V", "PNG, JPG, WEBP"]))
        self._src_view = _ImageView()
        self._src_stage.stack.addWidget(self._src_view)
        self._src_stage.clicked.connect(self._browse)
        self._src_stage.file_dropped.connect(self._load_path)
        self._src_stage.image_dropped.connect(
            lambda img: self._load_qimage(img, "dropped-image"))
        self._src_pane = _Pane("Original", self._src_stage)
        self._open_btn = widgets.button(self._src_pane, "Open image…",
                                        variant="neutral", icon="folder",
                                        command=self._browse)
        self._open_btn.setToolTip("Ctrl+O")
        self._clear_btn = widgets.button(self._src_pane, "Clear", variant="neutral",
                                         icon="x", command=self._clear)
        self._src_pane.footer.addWidget(self._open_btn)
        self._src_pane.footer.addWidget(self._clear_btn)
        self._src_pane.footer.addStretch(1)
        row.addWidget(self._src_pane, 1)

        # --- phải: kết quả ---
        self._res_stage = _Stage()
        self._res_stage.stack.addWidget(_empty_face(
            "sparkles", "Your edited image will appear here",
            ["Open an image, describe the change, then click Apply edit."]))
        self._res_view = _ImageView()
        self._res_stage.stack.addWidget(self._res_view)
        self._res_stage.stack.addWidget(self._build_loading())
        self._res_pane = _Pane("Result", self._res_stage)
        self._save_btn = widgets.button(self._res_pane, "Download", variant="success",
                                        icon="download", command=self._save)
        self._save_btn.setToolTip("Ctrl+S")
        # Chỉ icon: ba nút phải vừa nửa bề ngang khi cửa sổ thu nhỏ hết cỡ.
        self._copy_btn = widgets.button(self._res_pane, "", variant="neutral",
                                        icon="copy", command=self._copy)
        self._copy_btn.setToolTip("Copy to clipboard")
        self._copy_btn.setFixedWidth(40)
        self._again_btn = widgets.button(self._res_pane, "Keep editing",
                                         variant="neutral", icon="pencil",
                                         command=self._continue_from_result)
        self._again_btn.setToolTip("Use the result as the new original, so the "
                                   "next edit builds on it.")
        for btn in (self._save_btn, self._copy_btn, self._again_btn):
            self._res_pane.footer.addWidget(btn)
        self._res_pane.footer.addStretch(1)
        row.addWidget(self._res_pane, 1)

        outer.addLayout(row, 1)

    def _build_loading(self):
        face = QWidget()
        face.setStyleSheet("background: transparent;")
        lay = QVBoxLayout(face)
        lay.setSpacing(8)
        lay.addStretch(1)
        self._spinner = _Spinner()
        lay.addWidget(self._spinner, 0, Qt.AlignCenter)
        lay.addSpacing(6)
        self._load_title = QLabel("Applying your edit…")
        self._load_title.setObjectName("StageTitle")
        self._load_title.setAlignment(Qt.AlignCenter)
        lay.addWidget(self._load_title)
        self._load_sub = QLabel("")
        self._load_sub.setObjectName("StageSub")
        self._load_sub.setAlignment(Qt.AlignCenter)
        self._load_sub.setWordWrap(True)
        lay.addWidget(self._load_sub)
        self._load_time = QLabel("")
        self._load_time.setObjectName("StageMeta")
        self._load_time.setAlignment(Qt.AlignCenter)
        lay.addWidget(self._load_time)
        lay.addSpacing(6)
        cancel = widgets.button(face, "Cancel", variant="neutral", icon="x",
                                command=self._cancel)
        lay.addWidget(cancel, 0, Qt.AlignCenter)
        lay.addStretch(1)
        return face

    def _install_shortcuts(self):
        for key, fn in ((QKeySequence.Paste, self._paste),
                        (QKeySequence.Open, self._browse),
                        (QKeySequence.Save, self._save)):
            sc = QShortcut(QKeySequence(key), self._page)
            sc.setContext(Qt.WidgetWithChildrenShortcut)
            sc.activated.connect(fn)

    # ------------------------------------------------------------ trạng thái
    def _show_model(self):
        cfg = settings.load()
        model = (cfg.get("image_model") or "").strip() or settings.DEFAULTS["image_model"]
        size = cfg.get("image_size") or settings.DEFAULTS["image_size"]
        if not image_edit.supports_image_size(model):
            size = "1K"
        self._model_lbl.setText(f"{model} · {size}")

    def _refresh(self):
        has_src = self._src is not None
        has_res = self._result is not None
        busy = self._busy
        self._apply_btn.setEnabled(has_src and not busy)
        self._open_btn.setEnabled(not busy)
        self._clear_btn.setEnabled(has_src and not busy)
        for btn in (self._save_btn, self._copy_btn, self._again_btn):
            btn.setEnabled(has_res and not busy)

        self._src_stage.accepting = not busy
        self._src_stage.set_empty(not has_src)
        self._src_stage.stack.setCurrentIndex(1 if has_src else 0)
        self._res_stage.stack.setCurrentIndex(2 if busy else (1 if has_res else 0))

    def _set_busy(self, busy):
        self._busy = busy
        if busy:
            self._started = time.monotonic()
            self._retry_note = ""
            self._tick()
            self._spinner.start()
            self._tick_timer.start()
        else:
            self._spinner.stop()
            self._tick_timer.stop()
        self._refresh()

    def _tick(self):
        elapsed = int(time.monotonic() - self._started)
        self._load_time.setText(f"{elapsed}s elapsed · usually 10–40s")
        line = _LOADING_LINES[min(elapsed // _LINE_SECONDS, len(_LOADING_LINES) - 1)]
        self._load_sub.setText(self._retry_note or line)

    # ------------------------------------------------------------ nạp ảnh gốc
    def _browse(self):
        if self._busy:
            return
        folder = self._last_dir or QStandardPaths.writableLocation(
            QStandardPaths.PicturesLocation)
        path, _ = QFileDialog.getOpenFileName(self._page, "Open image", folder,
                                              _OPEN_FILTER)
        if path:
            self._load_path(path)

    def _load_path(self, path):
        try:
            with open(path, "rb") as fh:
                data = fh.read()
        except OSError as exc:
            self.error("Couldn't open the image", str(exc))
            return
        folder = os.path.dirname(path)
        self._last_dir = folder
        name = os.path.splitext(os.path.basename(path))[0]
        self._load_bytes(data, os.path.splitext(path)[1].lower(), name, folder)

    def _load_qimage(self, image, name):
        if image.isNull():
            self.error("Couldn't read the image",
                       "The pasted or dropped content isn't an image.")
            return
        data, mime = _encode(image)
        self._set_source(data, mime, image, name, self._last_dir)

    def _load_bytes(self, data, ext, name, folder):
        try:
            payload, mime, image = _prepare_upload(data, ext)
        except ValueError:
            self.error("Couldn't read the image",
                       "This file isn't an image the app can open. "
                       "Try a PNG, JPG or WEBP file.")
            return
        self._set_source(payload, mime, image, name, folder)

    def _set_source(self, data, mime, image, name, folder):
        self._src = {"data": data, "mime": mime, "name": name, "folder": folder}
        self._src_view.set_image(image)
        self._src_pane.meta.setText(f"{name} · {image.width()} × {image.height()}")
        # Kết quả cũ thuộc về ảnh trước → bỏ, tránh tải nhầm.
        self._set_result(None)
        self._refresh()
        self._prompt.setFocus()

    def _paste(self):
        if self._busy:
            return
        md = QApplication.clipboard().mimeData()
        path = _local_image_path(md)
        if path:
            self._load_path(path)
        elif md.hasImage():
            self._load_qimage(QApplication.clipboard().image(), "pasted-image")

    def _clear(self):
        self._src = None
        self._src_view.clear_image()
        self._src_pane.meta.setText("")
        self._set_result(None)
        self._refresh()

    # ------------------------------------------------------------ prompt
    def _prompt_keys(self, _obj, event):
        """Ctrl+Enter = Apply edit; Ctrl+V khi clipboard là ẢNH (không có chữ)
        thì nạp làm ảnh gốc thay vì dán vào ô chữ."""
        if event.type() != QEvent.KeyPress:
            return False
        if (event.key() in (Qt.Key_Return, Qt.Key_Enter)
                and event.modifiers() & Qt.ControlModifier):
            self._apply()
            return True
        if event.matches(QKeySequence.Paste):
            md = QApplication.clipboard().mimeData()
            if not md.hasText() and (md.hasImage() or _local_image_path(md)):
                self._paste()
                return True
        return False

    def _use_template(self, template):
        self._prompt.setPlainText(template)
        self._prompt.setFocus()
        self._select_placeholder()

    def _select_placeholder(self):
        """Bôi đen chỗ [trống] đầu tiên để người dùng gõ đè luôn."""
        match = _PLACEHOLDER.search(self._prompt.toPlainText())
        cursor = self._prompt.textCursor()
        if match:
            cursor.setPosition(match.start())
            cursor.setPosition(match.end(), QTextCursor.KeepAnchor)
        else:
            cursor.movePosition(QTextCursor.End)
        self._prompt.setTextCursor(cursor)
        return bool(match)

    # ------------------------------------------------------------ gọi AI
    def _apply(self):
        if self._busy:
            return
        if self._src is None:
            self.error("No image yet", "Open, drop or paste an image first.")
            return
        request = self._prompt.toPlainText().strip()
        if not request:
            self.error("Nothing to do yet",
                       "Describe the change you want, or pick one of the "
                       "Quick start examples.")
            self._prompt.setFocus()
            return
        if _PLACEHOLDER.search(request):
            self.error("Fill in the blanks",
                       "Your request still has a part in [square brackets]. "
                       "Replace it with what you mean, e.g. [object] → "
                       "“the man in the blue shirt”.")
            self._prompt.setFocus()
            self._select_placeholder()
            return

        cfg = settings.load()
        api_key = (cfg.get("api_key") or "").strip()
        if not api_key:
            self.error("Gemini API key missing",
                       "Image editing uses your Gemini API key. Add it in Settings "
                       "first.")
            return
        model = (cfg.get("image_model") or "").strip() or settings.DEFAULTS["image_model"]
        size = cfg.get("image_size") or settings.DEFAULTS["image_size"]
        self._show_model()

        self._req_id += 1
        rid = self._req_id
        src = self._src

        def job(emit):
            return image_edit.edit_image(
                api_key, model, src["data"], src["mime"], request, size,
                on_retry=lambda attempt, wait, _reason: emit(wait),
                should_cancel=lambda: rid != self._req_id)

        task = Task(job, self._page)
        task.signals.progress.connect(lambda wait, r=rid: self._on_retry(r, wait))
        task.signals.finished.connect(
            lambda res, r=rid, m=model: self._on_done(r, m, res))
        task.signals.failed.connect(lambda msg, r=rid: self._on_failed(r, msg))
        task.finished.connect(lambda t=task: self._tasks.discard(t))
        self._tasks.add(task)
        self._set_busy(True)
        task.start()

    def _on_retry(self, rid, wait):
        if rid == self._req_id and self._busy:
            self._retry_note = (f"The AI service is busy — trying again in "
                                f"{wait}s…")
            self._tick()

    def _on_done(self, rid, model, result):
        if rid != self._req_id:
            return      # lượt đã hủy
        elapsed = int(time.monotonic() - self._started)
        self._set_busy(False)
        image, _ = _read_image(result.image)
        if image.isNull():
            self.error("Couldn't show the result",
                       "The AI returned an image the app can't display. "
                       "Please try again.")
            return
        self._set_result(result, image,
                         f"{image.width()} × {image.height()} · {model} · {elapsed}s")
        self._status.setText("Done. Compare it with the original, then Download — "
                             "or describe another change.")

    def _on_failed(self, rid, msg):
        if rid != self._req_id:
            return
        self._set_busy(False)
        self.error("Couldn't edit the image", msg)

    def _cancel(self):
        if not self._busy:
            return
        self._req_id += 1       # kết quả của lượt đang chạy sẽ bị bỏ qua
        self._set_busy(False)
        self._status.setText("Cancelled. If the AI had already started, that "
                             "attempt may still be billed.")

    def _set_result(self, result, image=None, meta=""):
        self._result = result
        if result is None:
            self._res_view.clear_image()
            self._res_pane.meta.setText("")
            self._res_pane.note.hide()
        else:
            self._res_view.set_image(image)
            self._res_pane.meta.setText(meta)
            self._res_pane.note.setText(f"AI note: {result.note}" if result.note else "")
            self._res_pane.note.setVisible(bool(result.note))
        self._refresh()

    # ------------------------------------------------------------ kết quả
    def _result_name(self):
        name = self._src["name"] if self._src else "image"
        return name if name.endswith("_edited") else f"{name}_edited"

    def _save(self):
        if self._result is None or self._busy:
            return
        ext = image_edit.extension_for(self._result.mime_type)
        folder = (self._src or {}).get("folder") or self._last_dir or \
            QStandardPaths.writableLocation(QStandardPaths.PicturesLocation)
        path, _ = QFileDialog.getSaveFileName(
            self._page, "Save edited image",
            os.path.join(folder, self._result_name() + ext), f"Image (*{ext})")
        if not path:
            return
        if not os.path.splitext(path)[1]:
            path += ext
        try:
            with open(path, "wb") as fh:
                fh.write(self._result.image)
        except OSError as exc:
            self.error("Couldn't save the image", str(exc))
            return
        self._last_dir = os.path.dirname(path)
        self._status.setText(f"Saved: {path}")

    def _copy(self):
        if self._result is None:
            return
        image, _ = _read_image(self._result.image)
        QApplication.clipboard().setImage(image)
        self._status.setText("Result copied — paste it into Word, Outlook, Teams…")

    def _continue_from_result(self):
        if self._result is None or self._busy:
            return
        result = self._result
        # Qua _prepare_upload như ảnh mở từ file: kết quả 4K có thể vượt trần
        # dung lượng gửi đi.
        try:
            data, mime, image = _prepare_upload(
                result.image, image_edit.extension_for(result.mime_type))
        except ValueError:
            self.error("Couldn't reuse the result",
                       "The result can't be read back as an image.")
            return
        folder = (self._src or {}).get("folder", "")
        self._set_source(data, mime, image, self._result_name(), folder)
        self._prompt.clear()
        self._status.setText("The result is now the original. Describe the next "
                             "change.")
