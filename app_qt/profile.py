"""HỒ SƠ NGƯỜI DÙNG — tên hiển thị + ảnh đại diện (bảng `users`).

KHÔNG phải tính năng bảo mật: không mật khẩu, không phân quyền, không chặn gì
cả. Mỗi máy tự nhận ra mình bằng TÀI KHOẢN WINDOWS (`cv_repository.
windows_login()`), tra ra `user_id` rồi dùng id đó cho `created_by`/`updated_by`
của mọi bảng nghiệp vụ. Hồ sơ ở đây chỉ để giao diện gọi đúng tên người dùng và
để nhìn ra dòng nào của ai khi hai người dùng chung một file .db.

ẢNH LƯU BẰNG BYTES, KHÔNG PHẢI ĐƯỜNG DẪN. File .db được copy qua lại giữa hai
máy, mà đường dẫn ảnh là đường dẫn CỦA MỘT MÁY — lưu đường dẫn thì máy kia mở
lên chỉ thấy ô trống. Ảnh được thu nhỏ về {px}×{px} PNG trước khi ghi nên mỗi hồ
sơ chỉ tốn vài chục KB.

Xử lý ảnh nằm ở đây (tầng giao diện) chứ không ở `app/core`: nó dùng QImage của
Qt, mà `app/core` là logic thuần Python không dính giao diện.
"""
from PySide6.QtCore import QBuffer, QByteArray, QEvent, QIODevice, Qt
from PySide6.QtGui import (
    QColor, QCursor, QImage, QPainter, QPainterPath, QPen, QPixmap,
)
from PySide6.QtWidgets import (
    QFileDialog, QFrame, QHBoxLayout, QLabel, QVBoxLayout,
)

from app.core import cv_repository as repo
from app_qt import dialogs, theme, widgets
from app_qt.components.dialog_base import build_dialog_shell

# Cạnh ảnh đại diện lưu xuống DB (px). Đủ nét cho mọi chỗ app hiển thị (to nhất
# là ô xem trước 96px trong modal sửa hồ sơ) mà vẫn chỉ ~10–30 KB một hồ sơ.
AVATAR_PX = 128

_IMAGE_FILTER = "Images (*.png *.jpg *.jpeg *.bmp *.gif *.webp)"

# Hàm cần gọi lại khi hồ sơ đổi (sidebar vẽ lại avatar ngay sau khi bấm Save).
_watchers = []


def watch(callback) -> None:
    """Đăng ký hàm chạy lại mỗi khi hồ sơ người dùng được lưu."""
    _watchers.append(callback)


def _notify() -> None:
    for callback in list(_watchers):
        try:
            callback()
        except RuntimeError:
            # Widget đã bị Qt hủy (trang đóng rồi) → bỏ khỏi danh sách.
            _watchers.remove(callback)


# ───────────────────────────── ảnh đại diện ──────────────────────────────

def encode_avatar(path: str) -> bytes:
    """Đọc một file ảnh → PNG vuông {AVATAR_PX}×{AVATAR_PX} dạng bytes.

    Cắt giữa theo cạnh ngắn trước khi thu nhỏ, nên ảnh chân dung hay ảnh ngang
    đều ra khuôn tròn cân đối chứ không bị bóp méo. Ném ValueError nếu file
    không phải ảnh đọc được.
    """
    image = QImage(path)
    if image.isNull():
        raise ValueError("This file is not an image the app can read.")
    side = min(image.width(), image.height())
    square = image.copy((image.width() - side) // 2, (image.height() - side) // 2,
                        side, side)
    small = square.scaled(AVATAR_PX, AVATAR_PX, Qt.IgnoreAspectRatio,
                          Qt.SmoothTransformation)
    # QBuffer chỉ GIỮ CON TRỎ tới QByteArray chứ không sao chép nó, nên ô nhớ
    # đích phải nằm trong một biến sống lâu hơn buffer. Truyền thẳng
    # `QBuffer(QByteArray())` thì mảng tạm bị thu hồi ngay khi hàm dựng trả về,
    # và lần ghi kế tiếp là truy cập bộ nhớ đã giải phóng — app chết ở tầng C,
    # không có traceback Python và không kịp hiện lỗi gì.
    store = QByteArray()
    buffer = QBuffer(store)
    buffer.open(QIODevice.WriteOnly)
    small.save(buffer, "PNG")
    buffer.close()
    return bytes(store)


def _initials(name: str) -> str:
    """1–2 chữ cái đầu để vẽ khi hồ sơ chưa có ảnh."""
    parts = [p for p in str(name or "").replace(".", " ").split() if p]
    if not parts:
        return "?"
    if len(parts) == 1:
        return parts[0][:2].upper()
    return (parts[0][0] + parts[-1][0]).upper()


def avatar_pixmap(data, name: str, px: int) -> QPixmap:
    """Ảnh đại diện hình TRÒN cỡ `px`.

    Chưa có ảnh thì vẽ vòng tròn màu nhấn kèm chữ cái đầu của tên — luôn có gì
    đó để nhìn, không bao giờ là một ô trống.
    """
    pm = QPixmap(px, px)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing, True)
    path = QPainterPath()
    path.addEllipse(0, 0, px, px)
    p.setClipPath(path)
    if data:
        image = QImage.fromData(QByteArray(bytes(data)))
        if not image.isNull():
            p.drawPixmap(0, 0, QPixmap.fromImage(image.scaled(
                px, px, Qt.IgnoreAspectRatio, Qt.SmoothTransformation)))
            p.end()
            return pm
    p.setPen(Qt.NoPen)
    p.setBrush(QColor(theme.ACCENT))
    p.drawEllipse(0, 0, px, px)
    p.setPen(QPen(QColor("#ffffff")))
    font = p.font()
    font.setPixelSize(max(10, int(px * 0.4)))
    font.setBold(True)
    p.setFont(font)
    p.drawText(pm.rect(), Qt.AlignCenter, _initials(name))
    p.end()
    return pm


def _edit_badge(px: int) -> QPixmap:
    """Chấm tròn TRẮNG + bút chì màu nhấn, đè lên góc dưới ảnh đại diện.

    Nền trắng viền tối đọc được trên mọi thứ nằm dưới nó: ảnh sáng, ảnh tối, và
    cả khuôn chữ cái đầu — khuôn đó tô đúng màu nhấn, nên chấm mà cũng màu nhấn
    thì tan biến vào nền.
    """
    pm = QPixmap(px, px)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing, True)
    p.setPen(QPen(QColor(theme.PALETTE["--sidebar-bg"]), 1.5))
    p.setBrush(QColor("#ffffff"))
    p.drawEllipse(1, 1, px - 2, px - 2)
    icon = widgets.svg_pixmap("pencil", theme.ACCENT, round(px * 0.6))
    # svg_pixmap trả pixmap dpr=2 → lấy cỡ LOGIC để canh giữa cho đúng.
    side = icon.width() / icon.devicePixelRatio()
    p.drawPixmap(round((px - side) / 2), round((px - side) / 2), icon)
    p.end()
    return pm


def display_name(row=None) -> str:
    """Tên hiển thị của một dòng `users` (lùi về tài khoản Windows nếu trống)."""
    if row is None:
        return repo.windows_login() or "Unknown user"
    name = (row["display_name"] or "").strip()
    return name or (row["windows_login"] or "") or f"#{row['user_id']}"


def current_user():
    """Dòng `users` của người đang chạy app (None nếu chưa đọc được).

    Dựng DB trước khi đọc: hồ sơ hiện ở SIDEBAR nên có thể được hỏi trước cả
    khi người dùng mở tool nào — lúc đó bảng `users` còn chưa tồn tại. DB hỏng
    hay đang bị khóa thì trả None và giao diện lùi về tên tài khoản Windows,
    không được làm app không mở lên được.
    """
    try:
        repo.init_db()
        return repo.current_user()
    except Exception:
        return None


def avatar_bytes(row) -> bytes:
    """Ảnh đại diện của một dòng `users` (b"" nếu chưa đặt)."""
    if row is None or not row["avatar"]:
        return b""
    return bytes(row["avatar"])


# ───────────────────────── modal sửa hồ sơ người dùng ────────────────────

# Cạnh ô xem trước trong modal — to hơn hẳn ô 34px ở sidebar để soi được ảnh
# vừa chọn trước khi lưu.
PREVIEW_PX = 96


def edit_profile(parent) -> None:
    """Mở modal *Edit profile*: đổi ảnh đại diện + tên hiển thị rồi lưu.

    ĐỨNG RIÊNG, không nằm trong màn hình Settings: hồ sơ ghi xuống bảng `users`
    của file .db chứ không phải `config.json` như mọi ô ở Settings, và đường vào
    tự nhiên của nó là bấm thẳng vào ô hồ sơ ở sidebar. Chỉ có đúng hai thứ sửa
    được (ảnh + tên) nên modal có nút Save riêng, bấm mới ghi.
    """
    user = current_user()
    if user is None:
        dialogs.error(parent, "Profile unavailable",
                      "The app could not read your profile from the database.")
        return

    dlg, card, lay = build_dialog_shell(parent, "Edit profile")
    # Ảnh đang chọn giữ trong dict để các hàm con bên dưới gán lại được.
    state = {"avatar": avatar_bytes(user)}

    # Ảnh bên trái, hai nút ảnh xếp dọc bên phải: mọi khối trong modal (ảnh, ô
    # tên, nút Save) cùng một mép trái, không trộn căn giữa với căn trái.
    head = QHBoxLayout()
    head.setSpacing(16)
    preview = QLabel(card)
    preview.setFixedSize(PREVIEW_PX, PREVIEW_PX)
    head.addWidget(preview, 0, Qt.AlignVCenter)

    def draw():
        preview.setPixmap(avatar_pixmap(state["avatar"], name.get(), PREVIEW_PX))

    def pick():
        path, _ = QFileDialog.getOpenFileName(
            dlg, "Choose profile photo", "", _IMAGE_FILTER)
        if not path:
            return
        try:
            state["avatar"] = encode_avatar(path)
        except ValueError as exc:
            dialogs.error(dlg, "Could not read the image", str(exc))
            return
        draw()

    def clear():
        state["avatar"] = b""
        draw()

    photos = QVBoxLayout()
    photos.setSpacing(8)
    photos.addStretch(1)
    photos.addWidget(widgets.button(card, "Choose photo…", variant="neutral",
                                    icon="folder", command=pick))
    photos.addWidget(widgets.button(card, "Remove photo", variant="neutral",
                                    icon="trash", command=clear))
    photos.addStretch(1)
    head.addLayout(photos)
    head.addStretch(1)
    lay.addLayout(head)

    name = widgets.text_row(card, "Display name")
    name.set(display_name(user))
    name.widget.textChanged.connect(lambda *_: draw())
    # Con trỏ nằm sẵn ở ô tên: nếu để mặc định, viền focus rơi vào nút
    # "Choose photo…" trông như nút đang được chọn.
    name.widget.setFocus()
    draw()

    def save():
        try:
            repo.update_user(user["user_id"], {
                "display_name": name.get().strip(),
                # b"" = người dùng vừa gỡ ảnh → ghi NULL, không ghi chuỗi rỗng.
                "avatar": state["avatar"] or None,
            })
        except Exception as exc:   # noqa: BLE001 — DB khóa/hỏng: báo tại chỗ
            dialogs.error(dlg, "Could not save", str(exc))
            return
        dlg.accept()
        _notify()

    foot = QHBoxLayout()
    foot.addWidget(widgets.button(card, "Save", variant="primary", icon="save",
                                  command=save))
    foot.addWidget(widgets.button(card, "Cancel", variant="neutral", icon="x",
                                  command=dlg.reject))
    foot.addStretch(1)
    lay.addLayout(foot)
    dlg.exec()


# ──────────────────────────── ô hồ sơ ở sidebar ──────────────────────────

class SidebarProfile(QFrame):
    """Ô hồ sơ ở ĐÁY sidebar: ảnh tròn + tên hiển thị.

    Chỉ MỘT dòng chữ — tên hiển thị. Mặc định tên đó đã là tài khoản Windows
    (xem `display_name`), nên ghi thêm một dòng tài khoản nữa là lặp lại đúng
    cùng một chữ.

    CHỈ ẢNH bấm được (mở modal sửa hồ sơ), phần tên thì không: vùng bấm trùng
    khít với thứ người dùng nhắm tới, và ô này không có nền/viền riêng nên một
    vùng bấm rộng bằng cả hàng là vùng bấm vô hình. Bù lại, rê chuột vào **ảnh
    hay tên** đều hiện chấm bút chì ở góc dưới ảnh — chỉ dấu duy nhất cho biết
    chỗ này bấm được.
    """

    AVATAR = 34
    BADGE = 15

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("SidebarProfile")

        lay = QHBoxLayout(self)
        # Lề trái = margin + padding của #NavItem (12 + 14) để ảnh thẳng hàng
        # với cột icon của menu; lề trên = lề dưới để ô nằm cân giữa dải trống
        # còn lại ở đáy sidebar.
        lay.setContentsMargins(26, 10, 12, 10)
        lay.setSpacing(10)
        self._avatar = QLabel(self)
        self._avatar.setFixedSize(self.AVATAR, self.AVATAR)
        self._avatar.setCursor(Qt.PointingHandCursor)
        self._avatar.mousePressEvent = self._avatar_clicked
        lay.addWidget(self._avatar)

        # Chấm bút chì nằm TRONG ảnh (con của nhãn ảnh) nên bám theo ảnh mà
        # không cần tính lại vị trí mỗi lần layout đổi. Trong suốt với chuột để
        # bấm trúng chấm vẫn tính là bấm vào ảnh.
        self._badge = QLabel(self._avatar)
        self._badge.setFixedSize(self.BADGE, self.BADGE)
        self._badge.setPixmap(_edit_badge(self.BADGE))
        self._badge.move(self.AVATAR - self.BADGE, self.AVATAR - self.BADGE)
        self._badge.setAttribute(Qt.WA_TransparentForMouseEvents)
        self._badge.hide()

        self._name = QLabel(self)
        self._name.setObjectName("ProfileName")
        lay.addWidget(self._name)
        lay.addStretch(1)

        for w in (self, self._avatar, self._name):
            w.installEventFilter(self)

        self.refresh()
        watch(self.refresh)

    def eventFilter(self, obj, event):
        """Hiện/ẩn chấm bút chì theo việc trỏ chuột có nằm trong ô hồ sơ không.

        Qt gửi Leave cho widget CHA ngay khi trỏ chuột chạy sang widget con, nên
        chỉ nghe Enter/Leave của từng widget thì chấm sẽ nhấp nháy lúc rê từ tên
        sang ảnh. Mỗi lần có Enter/Leave ở bất kỳ widget nào trong ô thì hỏi lại
        vị trí trỏ THẬT và quyết định một lần.
        """
        if event.type() in (QEvent.Enter, QEvent.Leave):
            self._badge.setVisible(
                self.rect().contains(self.mapFromGlobal(QCursor.pos())))
        return False

    def refresh(self):
        """Đọc lại hồ sơ từ DB và vẽ lại (gọi sau mỗi lần lưu ở modal)."""
        user = current_user()
        name = display_name(user)
        self._avatar.setPixmap(avatar_pixmap(avatar_bytes(user), name, self.AVATAR))
        self._name.setText(name)

    def _avatar_clicked(self, event):
        # Modal treo vào CỬA SỔ chứ không vào ô này: luật QSS "#Sidebar QWidget"
        # (nền tối) thắng "#Dialog" về độ ưu tiên, nên hộp thoại nào nhận widget
        # trong sidebar làm cha cũng bị nhuộm tối cả modal.
        if event.button() == Qt.LeftButton:
            edit_profile(self.window())
