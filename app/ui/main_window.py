"""Cửa sổ chính: thanh bên điều hướng + vùng nội dung."""
import os
import sys
import tkinter as tk

import ttkbootstrap as ttk

from app.core.registry import discover_tools
from app.ui import theme, widgets


def _resource(name):
    """Đường dẫn tới file tài nguyên, đúng cả khi chạy dev lẫn bản .exe.

    PyInstaller giải nén tài nguyên vào thư mục tạm sys._MEIPASS lúc chạy.
    """
    base = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base, name)


class NavItem(tk.Frame):
    """Một mục trong thanh bên (emoji màu + tên), có hover & active."""

    def __init__(self, parent, emoji, name, color, on_click):
        super().__init__(parent, bg=theme.SIDEBAR_BG, cursor="hand2")
        self._on_click = on_click
        self.active = False

        # pill thụt vào lề (bo cảm giác 'cục' của vùng active)
        self.pill = tk.Frame(self, bg=theme.SIDEBAR_BG)
        self.pill.pack(fill="x", padx=10, pady=1)

        self.bar = tk.Frame(self.pill, bg=theme.SIDEBAR_BG, width=3)
        self.bar.pack(side="left", fill="y")

        self.badge = widgets.icon_glyph(self.pill, emoji, theme.SIDEBAR_BG, size=20)
        self.badge.pack(side="left", padx=(11, 11), pady=8)

        self.label = tk.Label(
            self.pill, text=name, bg=theme.SIDEBAR_BG, fg=theme.SIDEBAR_FG,
            font=(theme.FONT_FAMILY, 10), anchor="w",
        )
        self.label.pack(side="left", fill="x", expand=True)

        for w in (self, self.pill, self.badge, self.label):
            w.bind("<Button-1>", lambda _e: self._on_click())
            w.bind("<Enter>", self._enter)
            w.bind("<Leave>", self._leave)

    def _enter(self, _=None):
        if not self.active:
            self._paint(theme.SIDEBAR_BG_HOVER, theme.SIDEBAR_FG_ACTIVE)

    def _leave(self, _=None):
        if not self.active:
            self._paint(theme.SIDEBAR_BG, theme.SIDEBAR_FG)

    def set_active(self, active):
        self.active = active
        if active:
            self._paint(theme.SIDEBAR_BG_ACTIVE, theme.SIDEBAR_FG_ACTIVE, bold=True)
        else:
            self._paint(theme.SIDEBAR_BG, theme.SIDEBAR_FG, bold=False)

    def _paint(self, bg, fg, bold=None):
        for w in (self.pill, self.label):
            w.config(bg=bg)
        self.badge.config(bg=bg)
        self.bar.config(bg=theme.ACCENT if self.active else bg)
        self.label.config(fg=fg)
        if bold is not None:
            weight = "bold" if bold else "normal"
            self.label.config(font=(theme.FONT_FAMILY, 10, weight))


class ToolCard(tk.Frame):
    """Thẻ công cụ ở màn hình chính, bấm để mở tool."""

    def __init__(self, parent, tool, on_click):
        super().__init__(
            parent, bg=theme.CARD_BG, cursor="hand2",
            highlightbackground=theme.BORDER, highlightthickness=1,
        )
        color = theme.category_color(tool.category)

        pad = tk.Frame(self, bg=theme.CARD_BG)
        pad.pack(fill="both", expand=True, padx=20, pady=18)

        widgets.icon_badge(
            pad, tool.icon, color, theme.CARD_BG,
            size=48, radius=14, font_size=21,
        ).pack(anchor="w")

        tk.Label(
            pad, text=tool.name, bg=theme.CARD_BG, fg=theme.TEXT,
            font=(theme.FONT_FAMILY, 12, "bold"), anchor="w", justify="left",
        ).pack(anchor="w", fill="x", pady=(12, 3))
        tk.Label(
            pad, text=tool.description, bg=theme.CARD_BG, fg=theme.MUTED,
            font=(theme.FONT_FAMILY, 9), wraplength=210,
            anchor="w", justify="left",
        ).pack(anchor="w", fill="x")

        for w in [self, pad] + list(pad.winfo_children()):
            w.bind("<Button-1>", lambda _e: on_click(tool))
            w.bind("<Enter>", self._enter)
            w.bind("<Leave>", self._leave)

    def _enter(self, _=None):
        self.config(highlightbackground=theme.ACCENT)

    def _leave(self, _=None):
        self.config(highlightbackground=theme.BORDER)


class MainWindow(ttk.Window):
    def __init__(self):
        super().__init__(themename=theme.BOOTSTYLE_THEME)
        self.title("Personal Toolbox")
        self.geometry("1140x710")
        self.minsize(960, 600)
        try:
            self.iconbitmap(_resource("icon_app.ico"))
        except Exception:
            pass

        self.tools = discover_tools()
        self.nav_items = {}

        self._build_sidebar()
        tk.Frame(self, bg=theme.BORDER, width=1).pack(side="left", fill="y")
        self.content = tk.Frame(self, bg=theme.CONTENT_BG)
        self.content.pack(side="right", fill="both", expand=True)

        self.show_home()

        # Chạy các tác vụ nền (vd: quét lịch phỏng vấn buổi sáng) sau khi
        # cửa sổ đã hiện — tránh chặn lúc khởi động.
        self.after(500, self._run_startup_tasks)

    def _run_startup_tasks(self):
        """Gọi startup() của những tool bật auto_startup."""
        for tool in self.tools:
            if not getattr(tool, "auto_startup", False):
                continue
            try:
                tool.startup(self)
            except Exception:
                pass   # một tool lỗi không được làm hỏng cả app

    # ----- Thanh bên -----
    def _build_sidebar(self):
        sb = tk.Frame(self, bg=theme.SIDEBAR_BG, width=252)
        sb.pack(side="left", fill="y")
        sb.pack_propagate(False)

        brand = tk.Frame(sb, bg=theme.SIDEBAR_BG)
        brand.pack(fill="x", pady=(20, 14), padx=18)
        widgets.icon_badge(
            brand, "🧰", theme.ACCENT, theme.SIDEBAR_BG,
            size=38, radius=11, font_size=17,
        ).pack(side="left", padx=(0, 11))
        texts = tk.Frame(brand, bg=theme.SIDEBAR_BG)
        texts.pack(side="left")
        tk.Label(
            texts, text="Personal Toolbox", bg=theme.SIDEBAR_BG,
            fg=theme.SIDEBAR_FG_ACTIVE, font=(theme.FONT_FAMILY, 12, "bold"),
        ).pack(anchor="w")
        tk.Label(
            texts, text="Hộp đồ nghề cá nhân", bg=theme.SIDEBAR_BG,
            fg=theme.SIDEBAR_MUTED, font=(theme.FONT_FAMILY, 9),
        ).pack(anchor="w")

        # đường kẻ ngăn cách mảnh
        tk.Frame(sb, bg=theme.SIDEBAR_BG_HOVER, height=1).pack(
            fill="x", padx=18, pady=(2, 6))

        nav = tk.Frame(sb, bg=theme.SIDEBAR_BG)
        nav.pack(fill="both", expand=True)

        home = NavItem(nav, "🏠", "Trang chủ", theme.ACCENT, self.show_home)
        home.pack(fill="x")
        self.nav_items["home"] = home

        current_cat = None
        for tool in self.tools:
            if tool.category != current_cat:
                current_cat = tool.category
                tk.Label(
                    nav, text=current_cat.upper(), bg=theme.SIDEBAR_BG,
                    fg=theme.SIDEBAR_MUTED, font=(theme.FONT_FAMILY, 8, "bold"),
                    anchor="w",
                ).pack(fill="x", padx=18, pady=(16, 4))
            item = NavItem(
                nav, tool.icon, tool.name, theme.category_color(tool.category),
                lambda t=tool: self.show_tool(t),
            )
            item.pack(fill="x")
            self.nav_items[tool.name] = item

        tk.Label(
            sb, text="v0.1.0", bg=theme.SIDEBAR_BG, fg=theme.SIDEBAR_MUTED,
            font=(theme.FONT_FAMILY, 8),
        ).pack(side="bottom", pady=12)

    def _set_active(self, key):
        for k, item in self.nav_items.items():
            item.set_active(k == key)

    def _clear_content(self):
        for child in self.content.winfo_children():
            child.destroy()

    # ----- Màn hình chính -----
    def show_home(self):
        self._set_active("home")
        self._clear_content()

        wrap = tk.Frame(self.content, bg=theme.CONTENT_BG)
        wrap.pack(fill="both", expand=True, padx=40, pady=34)

        tk.Label(
            wrap, text="Chào mừng trở lại 👋", bg=theme.CONTENT_BG,
            fg=theme.TEXT, font=(theme.FONT_FAMILY, 22, "bold"),
        ).pack(anchor="w")
        tk.Label(
            wrap,
            text="Chọn một công cụ ở thanh bên, hoặc bấm vào thẻ bên dưới để bắt đầu.",
            bg=theme.CONTENT_BG, fg=theme.MUTED, font=(theme.FONT_FAMILY, 11),
        ).pack(anchor="w", pady=(6, 26))

        grid = tk.Frame(wrap, bg=theme.CONTENT_BG)
        grid.pack(fill="both", expand=True)
        cols = 3
        for i, tool in enumerate(self.tools):
            card = ToolCard(grid, tool, self.show_tool)
            r, c = divmod(i, cols)
            card.grid(row=r, column=c, sticky="nsew", padx=9, pady=9)
        for c in range(cols):
            grid.columnconfigure(c, weight=1, uniform="cards")

    # ----- Màn hình một tool -----
    def show_tool(self, tool):
        self._set_active(tool.name)
        self._clear_content()

        header = tk.Frame(self.content, bg=theme.CONTENT_BG)
        header.pack(fill="x", padx=40, pady=(32, 0))

        title = tk.Frame(header, bg=theme.CONTENT_BG)
        title.pack(fill="x")
        widgets.icon_badge(
            title, tool.icon, theme.category_color(tool.category),
            theme.CONTENT_BG, size=44, radius=13, font_size=20,
        ).pack(side="left", padx=(0, 14))
        title_texts = tk.Frame(title, bg=theme.CONTENT_BG)
        title_texts.pack(side="left", fill="x", expand=True)
        tk.Label(
            title_texts, text=tool.name, bg=theme.CONTENT_BG, fg=theme.TEXT,
            font=(theme.FONT_FAMILY, 18, "bold"), anchor="w",
        ).pack(anchor="w")
        tk.Label(
            title_texts, text=tool.description, bg=theme.CONTENT_BG,
            fg=theme.MUTED, font=(theme.FONT_FAMILY, 10),
            anchor="w", justify="left",
        ).pack(anchor="w", pady=(2, 0))

        body = tool.build(self.content)
        body.pack(fill="both", expand=True, padx=40, pady=22)
