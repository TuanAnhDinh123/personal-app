"""Cửa sổ chính (PySide6): sidebar điều hướng + vùng nội dung dạng stack."""
import os
import sys

from PySide6.QtCore import QEvent, QObject, Qt, QSize, QTimer, Signal
from PySide6.QtGui import QColor, QCursor, QIcon
from PySide6.QtWidgets import (
    QApplication, QButtonGroup, QFrame, QGridLayout, QHBoxLayout, QLabel,
    QMainWindow, QPushButton, QStackedWidget, QVBoxLayout, QWidget,
)

from app_qt import icons, settings_page, theme, widgets
from app_qt.registry import discover_tools


def _resource(name):
    base = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base, name)


class ClickableCard(widgets.Card):
    """Thẻ công cụ ở Trang chủ — bấm để mở tool. Tự vẽ bóng + đổi màu khi hover
    (dùng lại widgets.Card thay QSS #ToolCard để bóng đồng nhất, không bị 'vệt
    xám' như QGraphicsDropShadowEffect trên Windows HiDPI)."""
    clicked = Signal()
    PAD = 12   # bóng nhỏ hơn thẻ trang; cũng là khoảng cách giữa các thẻ trong lưới

    def __init__(self, tool, parent=None):
        super().__init__(parent, pad=self.PAD, dy=5, alpha=6)
        self.setCursor(Qt.PointingHandCursor)
        self._hover = False
        self._hover_fill = QColor(theme.PALETTE["--card-bg-hover"])
        self._hover_border = QColor(theme.ACCENT)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(10)

        lay.addWidget(widgets.icon_badge(
            self, tool.icon, theme.category_color(tool.category), size=46))

        name = QLabel(tool.name, self)
        name.setObjectName("ToolCardName")
        name.setWordWrap(True)
        lay.addWidget(name)

        desc = QLabel(tool.description, self)
        desc.setObjectName("ToolCardDesc")
        desc.setWordWrap(True)
        lay.addWidget(desc)
        lay.addStretch(1)

    def _colors(self):
        if self._hover:
            return self._hover_fill, self._hover_border
        return super()._colors()

    def enterEvent(self, event):
        self._hover = True
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._hover = False
        self.update()
        super().leaveEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)


class _ResizeFilter(QObject):
    """Cho phép kéo giãn cửa sổ frameless ở 8 mép (dùng startSystemResize)."""
    M = 6   # bề rộng vùng bắt mép (px)

    def __init__(self, win):
        super().__init__(win)
        self.win = win

    def _edges_at(self, gpos):
        if self.win.isMaximized() or self.win.isFullScreen():
            return Qt.Edges()
        r = self.win.frameGeometry()
        x, y = gpos.x(), gpos.y()
        if not (r.left() - self.M <= x <= r.right() + self.M
                and r.top() - self.M <= y <= r.bottom() + self.M):
            return Qt.Edges()
        e = Qt.Edges()
        if abs(x - r.left()) <= self.M:
            e |= Qt.LeftEdge
        if abs(x - r.right()) <= self.M:
            e |= Qt.RightEdge
        if abs(y - r.top()) <= self.M:
            e |= Qt.TopEdge
        if abs(y - r.bottom()) <= self.M:
            e |= Qt.BottomEdge
        return e

    def eventFilter(self, obj, ev):
        t = ev.type()
        if t == QEvent.MouseMove and not (QApplication.mouseButtons() & Qt.LeftButton):
            e = self._edges_at(QCursor.pos())
            if e in (Qt.LeftEdge | Qt.TopEdge, Qt.RightEdge | Qt.BottomEdge):
                self.win.setCursor(Qt.SizeFDiagCursor)
            elif e in (Qt.RightEdge | Qt.TopEdge, Qt.LeftEdge | Qt.BottomEdge):
                self.win.setCursor(Qt.SizeBDiagCursor)
            elif e & (Qt.LeftEdge | Qt.RightEdge):
                self.win.setCursor(Qt.SizeHorCursor)
            elif e & (Qt.TopEdge | Qt.BottomEdge):
                self.win.setCursor(Qt.SizeVerCursor)
            else:
                self.win.unsetCursor()
        elif t == QEvent.MouseButtonPress and ev.button() == Qt.LeftButton:
            e = self._edges_at(QCursor.pos())
            if e:
                handle = self.win.windowHandle()
                if handle is not None:
                    handle.startSystemResize(e)
                    return True
        return False


class MainWindow(QMainWindow):
    MASTER_CATEGORY = "Master Data"

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Personal Toolbox")
        self.resize(1200, 760)
        self.setMinimumSize(1000, 640)
        self.setWindowFlag(Qt.FramelessWindowHint, True)   # bỏ title bar mặc định
        # Nền cửa sổ trong suốt → góc bo do QSS (border-radius) vẽ, ĐƯỢC khử răng
        # cưa (anti-alias). Trước đây dùng setMask: clip 1-bit, góc bị răng cưa và
        # nhìn đẹp/xấu tùy màu desktop. Cùng cách với AppDialog (dialogs.py).
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        try:
            self.setWindowIcon(QIcon(_resource("icon_app.ico")))
        except Exception:
            pass

        self.tools = discover_tools()
        self.nav_group = QButtonGroup(self)
        self.nav_group.setExclusive(True)
        self.nav_buttons = {}      # key -> QPushButton
        self.pages = {}            # key -> QWidget (cache)

        # Khung ngoài cùng bo góc + viền màu (nổi trên desktop). Sidebar cao full
        # bên trái; cột phải = dải kéo/nút cửa sổ (sáng, KHÔNG còn dải tối "đội mũ")
        # + vùng nội dung.
        shell = QFrame(self)
        shell.setObjectName("AppShell")
        shell.setProperty("maxed", False)   # phóng to → bỏ bo góc (QSS)
        self.shell = shell
        root = QHBoxLayout(shell)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        root.addWidget(self._build_sidebar())

        rightcol = QWidget(shell)
        rv = QVBoxLayout(rightcol)
        rv.setContentsMargins(0, 0, 0, 0)
        rv.setSpacing(0)
        rv.addWidget(self._build_titlebar())
        self.stack = QStackedWidget(rightcol)
        self.stack.setObjectName("Content")
        rv.addWidget(self.stack, 1)
        root.addWidget(rightcol, 1)
        self.setCentralWidget(shell)

        self._show("home")

        # Kéo giãn cửa sổ ở mép (title bar tự vẽ → không còn viền hệ điều hành).
        self._resizer = _ResizeFilter(self)
        QApplication.instance().installEventFilter(self._resizer)

        # Chạy tác vụ nền của tool bật auto_startup (sau khi cửa sổ đã hiện).
        QTimer.singleShot(600, self._run_startup_tasks)

    # ----------------------------------------------------------- title bar
    def _build_titlebar(self):
        bar = QFrame()
        bar.setObjectName("TitleBar")
        bar.setFixedHeight(38)
        h = QHBoxLayout(bar)
        h.setContentsMargins(12, 0, 0, 0)
        h.setSpacing(0)
        h.addStretch(1)
        self._btn_max = self._win_btn("win-max", self._toggle_max)
        h.addWidget(self._win_btn("win-min", self.showMinimized))
        h.addWidget(self._btn_max)
        h.addWidget(self._win_btn("x", self.close, close=True))
        bar.mousePressEvent = self._titlebar_press
        bar.mouseDoubleClickEvent = lambda _e: self._toggle_max()
        return bar

    def _win_btn(self, icon, cb, close=False):
        b = QPushButton()
        b.setObjectName("WinClose" if close else "WinBtn")
        b.setFixedSize(46, 38)
        b.setCursor(Qt.PointingHandCursor)
        b.setIcon(widgets.svg_icon(icon, theme.TEXT_MUTED, 15))
        b.setIconSize(QSize(15, 15))
        b.clicked.connect(lambda *_: cb())
        return b

    def _titlebar_press(self, event):
        if event.button() == Qt.LeftButton:
            handle = self.windowHandle()
            if handle is not None:
                handle.startSystemMove()

    def _toggle_max(self):
        # Đổi trạng thái; phần bo góc + icon nút do changeEvent cập nhật (bắt cả
        # khi phóng to/khôi phục bằng phím tắt hệ điều hành).
        self.showNormal() if self.isMaximized() else self.showMaximized()

    def _run_startup_tasks(self):
        for tool in self.tools:
            if not getattr(tool, "auto_startup", False):
                continue
            try:
                tool.startup(self)
            except Exception:
                pass   # một tool lỗi không được làm hỏng cả app

    def changeEvent(self, e):
        """Phóng to → vuông (bỏ bo góc + viền); khôi phục → bo góc lại."""
        super().changeEvent(e)
        if e.type() == QEvent.WindowStateChange:
            # Trong lúc xử lý event, isMaximized() còn trả trạng thái CŨ → đọc ở
            # vòng lặp sự kiện kế tiếp khi trạng thái đã ổn định.
            QTimer.singleShot(0, self._sync_window_style)

    def _sync_window_style(self):
        maxed = self.isMaximized() or self.isFullScreen()
        if self.shell.property("maxed") == maxed:
            return
        self.shell.setProperty("maxed", maxed)
        # Đổi property động → cần polish lại để QSS [maximized] có hiệu lực.
        # Nạp lại stylesheet của app = polish toàn bộ (đổi trạng thái hiếm → rẻ).
        app = QApplication.instance()
        app.setStyleSheet(app.styleSheet())
        self._btn_max.setIcon(widgets.svg_icon(
            "win-restore" if maxed else "win-max", theme.TEXT_MUTED, 15))

    # ----------------------------------------------------------- sidebar
    def _build_sidebar(self):
        sb = QFrame()
        sb.setObjectName("Sidebar")
        sb.setFixedWidth(264)
        outer = QVBoxLayout(sb)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # Thương hiệu = nút về Trang chủ (bấm để quay lại Home).
        brand = QFrame(sb)
        brand.setObjectName("Brand")
        brand.setCursor(Qt.PointingHandCursor)
        brand.mousePressEvent = lambda _e: self._show("home")
        b = QHBoxLayout(brand)
        b.setContentsMargins(12, 10, 12, 10)
        b.setSpacing(12)
        b.addWidget(widgets.icon_badge(brand, "🧰", theme.ACCENT, size=40))
        texts = QVBoxLayout()
        texts.setSpacing(1)
        n = QLabel("Personal Toolbox")
        n.setObjectName("BrandName")
        s = QLabel("Every task, one place")
        s.setObjectName("BrandSub")
        texts.addWidget(n)
        texts.addWidget(s)
        b.addLayout(texts)
        b.addStretch(1)
        outer.addWidget(brand)
        line = QFrame(); line.setObjectName("SidebarDivider"); line.setFrameShape(QFrame.HLine)
        outer.addWidget(line)

        # vùng cuộn chứa các mục nav
        nav_holder = QWidget()
        self._nav_layout = QVBoxLayout(nav_holder)
        self._nav_layout.setContentsMargins(0, 8, 0, 0)
        self._nav_layout.setSpacing(0)

        master = [t for t in self.tools if t.category == self.MASTER_CATEGORY]
        others = [t for t in self.tools if t.category != self.MASTER_CATEGORY]
        self._build_nav_group(others)
        if master:
            self._divider()
            self._build_nav_group(master)

        self._divider()
        self._add_nav("settings", "⚙", "Cài đặt", lambda: self._show("settings"), theme.ACCENT)
        self._nav_layout.addStretch(1)

        outer.addWidget(widgets.scroll_area(nav_holder), 1)

        ver = QLabel("v0.2.0 · PySide6"); ver.setObjectName("SidebarVersion")
        ver.setAlignment(Qt.AlignCenter)
        outer.addWidget(ver)
        return sb

    def _divider(self):
        line = QFrame(); line.setObjectName("SidebarDivider")
        line.setFrameShape(QFrame.HLine)
        self._nav_layout.addWidget(line)

    def _build_nav_group(self, tools):
        current = None
        for tool in tools:
            if tool.category != current:
                current = tool.category
                title = QLabel(current.upper())
                title.setObjectName("NavSectionTitle")
                self._nav_layout.addWidget(title)
            self._add_nav(tool.name, tool.icon, tool.name,
                          lambda t=tool: self._show_tool(t),
                          theme.category_color(tool.category))

    def _add_nav(self, key, emoji, text, on_click, color):
        btn = QPushButton("  " + text)
        btn.setObjectName("NavItem")
        btn.setCheckable(True)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setIconSize(QSize(20, 20))
        name = icons.name_for(emoji)

        def paint(checked):
            btn.setIcon(widgets.svg_icon(name, "#ffffff" if checked else color, 20))

        paint(False)
        btn.toggled.connect(paint)
        # clicked truyền tham số `checked`(bool) → nuốt bằng *_ để on_click chạy đúng.
        btn.clicked.connect(lambda *_: on_click())
        self.nav_group.addButton(btn)
        self.nav_buttons[key] = btn
        self._nav_layout.addWidget(btn)

    # ----------------------------------------------------------- điều hướng
    def _show(self, key, builder=None):
        if key not in self.pages:
            if key == "home":
                self.pages[key] = self._build_home()
            elif key == "settings":
                self.pages[key] = self._build_settings()
            elif builder is not None:
                self.pages[key] = builder()
            self.stack.addWidget(self.pages[key])
        self.stack.setCurrentWidget(self.pages[key])
        if key in self.nav_buttons:
            self.nav_buttons[key].setChecked(True)
        elif key == "home":
            # Home không có mục nav → bỏ chọn mọi mục (nhóm exclusive cần mở tạm).
            self.nav_group.setExclusive(False)
            for btn in self.nav_buttons.values():
                btn.setChecked(False)
            self.nav_group.setExclusive(True)

    def _show_tool(self, tool):
        self._show(tool.name, builder=lambda: self._build_tool_page(tool))

    # ----------------------------------------------------------- khung trang
    # Thẻ (widgets.Card) TỰ VẼ bóng và tự chừa CARD_PAD px mỗi phía cho bóng, nên
    # KHÔNG cần holder/lề trang chừa thêm. Mép NHÌN THẤY của thẻ = mép widget +
    # CARD_PAD, nên để thẻ nằm ở ~38px như cũ: lề trang = 38 − CARD_PAD; tiêu đề
    # thụt thêm CARD_PAD để thẳng hàng mép thẻ nhìn thấy. Nội dung KHÔNG-phải-thẻ
    # trong body (nút, tab, splitter…) cần tự thêm lề ngang CARD_PAD để cùng hàng.
    CONTENT_EDGE = 38    # mép trái nội dung nhìn thấy, tính từ mép vùng nội dung

    def _page_shell(self, icon, color, title, subtitle, body, fills_height=False):
        """Khung trang: header (chip + tiêu đề) + body (bọc scroll nếu cần)."""
        pad = widgets.CARD_PAD
        side = self.CONTENT_EDGE - pad
        page = QWidget()
        lay = QVBoxLayout(page)
        # lề ngang (side) = lề dưới → padding từ viền app tới thẻ đều nhau
        # (x == bottom). Cả hai + CARD_PAD = CONTENT_EDGE ở mép thẻ nhìn thấy.
        lay.setContentsMargins(side, 8, side, side)
        lay.setSpacing(0)

        header = QHBoxLayout()
        header.setContentsMargins(pad, 0, pad, 0)   # thẳng hàng mép thẻ nhìn thấy
        header.setSpacing(14)
        header.addWidget(widgets.icon_badge(page, icon, color, size=46))
        texts = QVBoxLayout(); texts.setSpacing(2)
        t = QLabel(title); t.setObjectName("PageTitle")
        st = QLabel(subtitle); st.setObjectName("PageSubtitle"); st.setWordWrap(False)
        texts.addWidget(t); texts.addWidget(st)
        header.addLayout(texts); header.addStretch(1)
        lay.addLayout(header)
        lay.addSpacing(22)   # khoảng cách tiêu đề → thẻ (trước do CARD_PAD chừa,
        #                      nay CARD_PAD=0 nên đặt thẳng ~22px ở đây)

        if fills_height:
            lay.addWidget(body, 1)
        else:
            holder = QWidget()
            hl = QVBoxLayout(holder)
            hl.setContentsMargins(0, 0, 0, 0)
            hl.addWidget(body)
            hl.addStretch(1)
            lay.addWidget(widgets.scroll_area(holder), 1)
        return page

    def _build_tool_page(self, tool):
        body = tool.build()
        return self._page_shell(
            tool.icon, theme.category_color(tool.category),
            tool.name, tool.description, body,
            fills_height=getattr(tool, "fills_height", False))

    def _build_settings(self):
        body = settings_page.build()
        return self._page_shell(
            "⚙", theme.ACCENT, "Cài đặt",
            "Thiết lập dùng chung cho toàn bộ ứng dụng.", body)

    # ----------------------------------------------------------- trang chủ
    def _build_home(self):
        pad = ClickableCard.PAD              # thẻ tự chừa pad cho bóng
        side = self.CONTENT_EDGE - pad
        page = QWidget()
        outer = QVBoxLayout(page)
        outer.setContentsMargins(side, 10, side, side)   # x == bottom (như trang tool)
        outer.setSpacing(0)

        # thụt pad để tiêu đề thẳng hàng mép trái lưới thẻ (thẻ cũng thụt pad)
        t = QLabel("Your everyday toolbox.")
        t.setObjectName("PageTitle")
        t.setContentsMargins(pad, 0, 0, 0)
        outer.addWidget(t)
        s = QLabel("Chọn một công cụ bên dưới để bắt đầu.")
        s.setObjectName("PageSubtitle")
        s.setContentsMargins(pad, 0, 0, 0)
        outer.addWidget(s)
        outer.addSpacing(2)

        grid_holder = QWidget()
        grid = QGridLayout(grid_holder)
        # thẻ tự chừa pad cho bóng → khoảng cách thấy được giữa thẻ ≈ 2*pad;
        # spacing lưới = 0 để không dôi thêm.
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setSpacing(0)
        cols = 3
        home_tools = [t for t in self.tools if getattr(t, "show_on_home", True)]
        for i, tool in enumerate(home_tools):
            card = ClickableCard(tool)
            card.clicked.connect(lambda tl=tool: self._show_tool(tl))
            r, c = divmod(i, cols)
            grid.addWidget(card, r, c)
        for c in range(cols):
            grid.setColumnStretch(c, 1)

        outer.addWidget(widgets.scroll_area(grid_holder), 1)
        return page
