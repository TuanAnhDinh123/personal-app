"""Kích thước & khung modal DÙNG CHUNG cho toàn app.

► SỬA KÍCH THƯỚC MODAL Ở ĐÂY — chỉ một chỗ duy nhất.

Có đúng 3 cỡ: sm / md / lg. Mọi modal trong app phải tham chiếu một trong ba cỡ
này (truyền `size="sm"|"md"|"lg"`), không đặt số rộng/cao rời rạc nữa.

    sm = cỡ gốc (như hộp "Sửa vị trí").
    md = rộng hơn sm 1.25×, cao hơn 1.5×.
    lg = rộng hơn sm 1.5×,  cao hơn 2×.

`ModalDialog` là lớp nền chung: frameless, nền trong suốt, kéo di chuyển được, và
CANH GIỮA MÀN HÌNH VẬT LÝ chứa cửa sổ — kẹp kích thước theo chiều cao màn hình
thật (không phải theo viewport/kích thước cửa sổ app), nên modal không bị co
lại khi cửa sổ app nhỏ hoặc bị lệch theo sidebar.
"""
from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QApplication, QDialog, QFrame, QHBoxLayout, QLabel, QVBoxLayout,
)

from app_qt import widgets

# Lề ngoài shell (mỗi phía) để lộ bóng — trừ ra khi kẹp modal vào viewport.
SHELL_MARGIN = 24

# ── Cỡ modal (rộng × cao, px) ───────────────────────────────────────────────
# Đổi _BASE_W / _BASE_H là ba cỡ tự giãn theo hệ số bên dưới.
_BASE_W, _BASE_H = 520, 480

MODAL_SM = (_BASE_W, _BASE_H)                                   # 520 × 480
MODAL_MD = (round(_BASE_W * 1.5), round(_BASE_H * 1.5))        # 650 × 720
MODAL_LG = (round(_BASE_W * 2),  round(_BASE_H * 1.75))          # 780 × 960

SIZES = {"sm": MODAL_SM, "md": MODAL_MD, "lg": MODAL_LG}


def dims(size):
    """(rộng, cao) theo tên cỡ; mặc định về sm nếu tên lạ."""
    return SIZES.get(size, MODAL_SM)


class ModalDialog(QDialog):
    """Lớp nền cho MỌI modal.

    Tham số:
        size ∈ 'sm' | 'md' | 'lg' — quyết định bề rộng chuẩn của thẻ và chiều cao
        chuẩn của vùng nội dung (dùng qua `self.modal_w` / `self.modal_h`).
    """

    def __init__(self, parent, size="sm"):
        super().__init__(parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setModal(True)
        self.modal_w, self.modal_h = dims(size)
        self._drag = None
        # Thẻ nội dung để kẹp bề rộng theo viewport (đặt trong build_shell / lớp con).
        self._card = None
        self._card_fixed_width = True
        # Vùng co giãn (scroll area / bảng) sẽ cao = modal_h, tự co khi viewport thấp.
        self._grow = None
        # "chrome" (phần cao ngoài vùng co giãn) đo 1 lần rồi cache — xem lý do
        # trong _clamp_to.
        self._chrome = None

    def set_grow_region(self, widget):
        """Đăng ký vùng co giãn chính của modal (thường là scroll area / bảng).

        Chiều cao của nó = `modal_h` khi màn hình đủ chỗ → modal LUÔN cao đúng cỡ
        (đổi MODAL_MD/… là cao lên thật), và tự co lại khi viewport thấp để không
        tràn (max-height = 100vh).
        """
        self._grow = widget

    # ---------------------------------------------------------- khung shell + card
    def build_shell(self, title, *, fixed_width=True, closable=True, spacing=12):
        """Dựng lề-bóng + thẻ #Dialog (rộng theo cỡ) + header (tiêu đề + nút ✕).

        Trả về (card, content_layout) để bên gọi thêm nội dung + nút.
        """
        shell = QVBoxLayout(self)
        shell.setContentsMargins(24, 24, 24, 24)   # lề ngoài để lộ bóng

        card = QFrame(self)
        card.setObjectName("Dialog")
        if fixed_width:
            card.setFixedWidth(self.modal_w)
        else:
            card.setMaximumWidth(self.modal_w)
        widgets.add_shadow(card, blur=48, dy=12, alpha=70)
        shell.addWidget(card)
        self._card = card
        self._card_fixed_width = fixed_width

        lay = QVBoxLayout(card)
        lay.setContentsMargins(22, 20, 22, 18)   # padding thẻ→nội dung chuẩn
        lay.setSpacing(spacing)

        if title is not None:
            head = QHBoxLayout()
            t = QLabel(title)
            t.setObjectName("DialogTitle")
            head.addWidget(t, 1)
            if closable:
                close = QLabel("✕")
                close.setObjectName("DialogClose")
                close.setFixedSize(24, 24)
                close.setAlignment(Qt.AlignCenter)
                close.setCursor(Qt.PointingHandCursor)
                close.mousePressEvent = lambda _e: self.reject()
                head.addWidget(close, 0, Qt.AlignTop)
            lay.addLayout(head)

        return card, lay

    # ---------------------------------------------- kẹp viewport + canh giữa nội dung
    def showEvent(self, e):
        super().showEvent(e)
        self._fit()
        # Chạy lại sau khi layout ổn định (kích thước cuối) để kẹp + canh chuẩn.
        QTimer.singleShot(0, self._fit)

    def _fit(self):
        area = viewport_rect(self.parent())
        self._clamp_to(area)
        self._center_in(area)

    def _clamp_to(self, area):
        """Không cho modal vượt viewport: max-width = 100vw, max-height = 100vh."""
        if area is None:
            return
        max_w = max(240, area.width() - 2 * SHELL_MARGIN)
        w = self.modal_w
        if self._card is not None:
            w = min(self.modal_w, max_w)
            # setFixed → giữ đúng bề rộng thiết kế; kẹp xuống khi màn hình hẹp.
            if self._card_fixed_width:
                self._card.setFixedWidth(w)
            else:
                self._card.setMaximumWidth(w)
        # Chiều cao: vùng co giãn cao đúng = modal_h, nhưng kẹp lại theo màn hình.
        # chrome = phần cao ngoài vùng co giãn (header/footer/label/lề) — đo 1 LẦN
        # DUY NHẤT (lúc chưa co giãn gì) rồi cache lại. Không được đo lại mỗi lần
        # bằng self.height() - self._grow.height(): setFixedHeight() làm
        # self._grow co lại NGAY, còn self.height() (cửa sổ ngoài) chỉ đuổi kịp ở
        # vòng lặp sự kiện sau → đo lại kiểu đó sẽ cho chrome tụt về ~0 mỗi lần
        # gọi tiếp theo, khiến vùng co giãn cứ phồng to dần mỗi lần _fit() chạy.
        new_h = None
        if self._grow is not None:
            if self._chrome is None:
                self._chrome = max(0, self.height() - self._grow.height())
            avail_h = area.height() - 2 * SHELL_MARGIN - self._chrome
            new_h = max(160, min(self.modal_h, avail_h))
            self._grow.setFixedHeight(new_h)
        # Trần cho cả dialog (kể cả lề bóng) → không bao giờ vượt màn hình.
        self.setMaximumSize(area.width(), area.height())
        # Co lại kích thước THẬT ngay bằng resize() trực tiếp với số đã tính —
        # không dùng sizeHint()/chờ Qt tự co: setFixedHeight() ở trên đã co
        # self._grow ngay, nhưng self (cửa sổ ngoài) chỉ tự co theo layout ở
        # vòng lặp sự kiện SAU lời gọi này. Nếu không resize() ngay bằng đúng số
        # đã tính, _center_in() bên dưới sẽ canh giữa theo kích thước cũ (còn to)
        # — rồi khi Qt co cửa sổ thật sự lại (muộn hơn), nó chỉ co từ mép dưới,
        # không canh giữa lại, khiến modal tràn xuống dưới & lệch khỏi giữa màn hình.
        total_w = w + 2 * SHELL_MARGIN
        total_h = (self._chrome + new_h) if new_h is not None else self.height()
        self.resize(total_w, total_h)

    def _center_in(self, area):
        if area is None:
            return
        g = self.frameGeometry()
        g.moveCenter(area.center())
        self.move(g.topLeft())

    # ---------------------------------------------------------- kéo di chuyển
    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag = e.globalPosition().toPoint() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, e):
        if self._drag is not None and e.buttons() & Qt.LeftButton:
            self.move(e.globalPosition().toPoint() - self._drag)

    def mouseReleaseEvent(self, e):
        self._drag = None


def viewport_rect(ref):
    """"Viewport" để canh giữa + kẹp kích thước modal.

    Luôn là vùng khả dụng của MÀN HÌNH VẬT LÝ chứa cửa sổ (100vw/100vh = trọn
    màn hình thật) — modal kẹp theo chiều cao màn hình thật và canh giữa MÀN
    HÌNH, không phải giữa vùng nội dung app (nên không co lại khi cửa sổ nhỏ
    hoặc lệch theo sidebar).
    """
    win = ref.window() if ref is not None else None
    screen = (win.screen() if win is not None else None) or QApplication.primaryScreen()
    return screen.availableGeometry() if screen is not None else None
