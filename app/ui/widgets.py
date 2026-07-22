"""Các widget dựng sẵn để tool tái sử dụng cho gọn.

Tất cả trả về một biến (StringVar/BooleanVar) để sau này gắn logic thật.
"""
import math
import tkinter as tk
from tkinter import filedialog

import ttkbootstrap as ttk

from app.ui import theme


def section_label(parent, text, bg=None):
    """Tiêu đề một nhóm trường nhập."""
    bg = bg or theme.CARD_BG
    tk.Label(
        parent, text=text, bg=bg, fg=theme.TEXT,
        font=(theme.FONT_FAMILY, 10, "bold"),
    ).pack(anchor="w", pady=(10, 8))


def hint(parent, text, bg=None):
    """Dòng ghi chú nhỏ, màu nhạt."""
    bg = bg or theme.CARD_BG
    tk.Label(
        parent, text=text, bg=bg, fg=theme.MUTED,
        font=(theme.FONT_FAMILY, 9), justify="left", wraplength=620,
    ).pack(anchor="w", pady=4)


def file_row(parent, label, mode="file", bg=None):
    """Ô chọn file/thư mục: nhãn + ô hiển thị đường dẫn + nút Chọn."""
    bg = bg or theme.CARD_BG
    block = tk.Frame(parent, bg=bg)
    block.pack(fill="x", pady=(6, 10))
    tk.Label(
        block, text=label, bg=bg, fg=theme.TEXT,
        font=(theme.FONT_FAMILY, 9),
    ).pack(anchor="w", pady=(0, 4))

    row = tk.Frame(block, bg=bg)
    row.pack(fill="x")
    var = tk.StringVar()
    ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True, ipady=4)

    def browse():
        if mode == "folder":
            path = filedialog.askdirectory()
        elif mode == "save":
            path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        else:
            path = filedialog.askopenfilename()
        if path:
            var.set(path)

    button(
        row, text="Chọn…", variant="neutral", icon="folder", command=browse,
    ).pack(side="left", padx=(8, 0))
    return var


def export_target_row(parent, label, bg=None):
    """Ô chọn ĐÍCH xuất: chọn Thư mục (tạo file mới) hoặc File Excel (nối tiếp).

    Trả về StringVar chứa đường dẫn. Bên gọi tự phân biệt thư mục / file.
    """
    bg = bg or theme.CARD_BG
    block = tk.Frame(parent, bg=bg)
    block.pack(fill="x", pady=(6, 10))
    tk.Label(
        block, text=label, bg=bg, fg=theme.TEXT,
        font=(theme.FONT_FAMILY, 9),
    ).pack(anchor="w", pady=(0, 4))

    row = tk.Frame(block, bg=bg)
    row.pack(fill="x")
    var = tk.StringVar()
    ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True, ipady=4)

    def pick_folder():
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    def pick_file():
        path = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx"), ("Tất cả", "*.*")])
        if path:
            var.set(path)

    button(
        row, text="Thư mục", variant="neutral", icon="folder", command=pick_folder,
    ).pack(side="left", padx=(8, 0))
    button(
        row, text="File Excel", variant="neutral", icon="file", command=pick_file,
    ).pack(side="left", padx=(6, 0))
    return var


def _only_digits(proposed):
    """Cho phép chuỗi rỗng hoặc toàn chữ số (dùng cho validatecommand)."""
    return proposed == "" or proposed.isdigit()


def digit_entry(parent, textvariable, **kw):
    """Ô nhập chỉ cho phép gõ chữ số."""
    vcmd = (parent.register(_only_digits), "%P")
    return ttk.Entry(
        parent, textvariable=textvariable,
        validate="key", validatecommand=vcmd, **kw,
    )


def text_row(parent, label, placeholder="", bg=None):
    """Ô nhập chữ một dòng."""
    bg = bg or theme.CARD_BG
    block = tk.Frame(parent, bg=bg)
    block.pack(fill="x", pady=(6, 10))
    tk.Label(
        block, text=label, bg=bg, fg=theme.TEXT,
        font=(theme.FONT_FAMILY, 9),
    ).pack(anchor="w", pady=(0, 4))
    var = tk.StringVar(value=placeholder)
    ttk.Entry(block, textvariable=var).pack(fill="x", ipady=4)
    return var


def text_area(parent, label, value="", height=8, bg=None):
    """Ô nhập chữ nhiều dòng. Trả về widget tk.Text (đọc bằng .get)."""
    bg = bg or theme.CARD_BG
    block = tk.Frame(parent, bg=bg)
    block.pack(fill="x", pady=(6, 10))
    tk.Label(
        block, text=label, bg=bg, fg=theme.TEXT,
        font=(theme.FONT_FAMILY, 9),
    ).pack(anchor="w", pady=(0, 4))
    holder = tk.Frame(
        block, bg="#ffffff", relief="solid", bd=1,
        highlightthickness=1, highlightbackground=theme.BORDER,
    )
    holder.pack(fill="x")
    box = tk.Text(
        holder, height=height, wrap="word", relief="flat", bd=0,
        font=(theme.FONT_FAMILY, 10), bg="#ffffff", fg=theme.TEXT,
        highlightthickness=0, padx=8, pady=6,
    )
    vsb = ttk.Scrollbar(holder, orient="vertical", command=box.yview)
    box.configure(yscrollcommand=vsb.set)
    box.pack(side="left", fill="both", expand=True)
    vsb.pack(side="right", fill="y")

    def _on_wheel(event):
        # Nếu nội dung còn phần bị ẩn -> cuộn trong ô này và chặn cuộn trang;
        # nếu đã hiện hết -> để sự kiện rơi xuống trang (ScrollableFrame).
        first, last = box.yview()
        if first <= 0.0 and last >= 1.0:
            return None
        box.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"

    box.bind("<MouseWheel>", _on_wheel)

    if value:
        box.insert("1.0", value)
    return box


def checkbox(parent, label, checked=True):
    """Công tắc bật/tắt một tùy chọn."""
    var = tk.BooleanVar(value=checked)
    ttk.Checkbutton(
        parent, text=label, variable=var, bootstyle="round-toggle",
    ).pack(anchor="w", pady=3)
    return var


class ScrollableFrame(tk.Frame):
    """Frame cuộn dọc. Thêm nội dung vào thuộc tính .inner."""

    def __init__(self, parent, bg=None, **kwargs):
        bg = bg or theme.CONTENT_BG
        super().__init__(parent, bg=bg, **kwargs)

        self._canvas = tk.Canvas(self, bg=bg, highlightthickness=0)
        self._vsb = ttk.Scrollbar(self, orient="vertical", command=self._canvas.yview)
        self.inner = tk.Frame(self._canvas, bg=bg)

        self._win_id = self._canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self._canvas.configure(yscrollcommand=self._vsb.set)

        self._vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._canvas.bind("<Configure>", self._on_canvas_resize)
        self.inner.bind("<Configure>", self._on_inner_resize)
        self._canvas.bind("<Enter>", lambda _: self._canvas.bind_all("<MouseWheel>", self._on_wheel))
        self._canvas.bind("<Leave>", lambda _: self._canvas.unbind_all("<MouseWheel>"))

    def _on_canvas_resize(self, event):
        self._canvas.itemconfig(self._win_id, width=event.width)

    def _on_inner_resize(self, event):
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _on_wheel(self, event):
        self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


def dropdown(parent, label, options, bg=None):
    """Danh sách lựa chọn (combobox)."""
    bg = bg or theme.CARD_BG
    block = tk.Frame(parent, bg=bg)
    block.pack(fill="x", pady=(6, 10))
    tk.Label(
        block, text=label, bg=bg, fg=theme.TEXT,
        font=(theme.FONT_FAMILY, 9),
    ).pack(anchor="w", pady=(0, 4))
    var = tk.StringVar(value=options[0])
    ttk.Combobox(
        block, textvariable=var, values=options, state="readonly",
    ).pack(fill="x", ipady=2)
    return var


# ----- Render icon bằng Pillow (mượt, có khử răng cưa, emoji màu thật) -----
try:
    from PIL import Image, ImageDraw, ImageFont, ImageTk
    _PIL_OK = True
except Exception:
    _PIL_OK = False

_EMOJI_FONT_PATH = r"C:\Windows\Fonts\seguiemj.ttf"
_img_cache = {}   # giữ tham chiếu PhotoImage để không bị thu hồi
_rounded_ready = set()   # tên style nút bo góc đã dựng (tránh dựng lại)


def _hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _tint(hex_color, frac):
    """Trộn màu với trắng để ra tông nhạt (frac = tỉ lệ màu gốc)."""
    r, g, b = _hex_to_rgb(hex_color)
    mix = lambda v: int(v * frac + 255 * (1 - frac))
    return (mix(r), mix(g), mix(b))


def _render_icon(emoji, size, chip_rgb, radius, ratio):
    """Tạo ảnh icon: (tùy chọn) nền bo góc + emoji màu, khử răng cưa bằng SSAA."""
    ss = 4
    big = Image.new("RGBA", (size * ss, size * ss), (0, 0, 0, 0))
    if chip_rgb is not None:
        ImageDraw.Draw(big).rounded_rectangle(
            [0, 0, size * ss - 1, size * ss - 1],
            radius=radius * ss, fill=chip_rgb + (255,),
        )
    img = big.resize((size, size), Image.LANCZOS)

    try:
        font = ImageFont.truetype(_EMOJI_FONT_PATH, max(1, int(size * ratio)))
        ImageDraw.Draw(img).text(
            (size / 2, size / 2 + 1), emoji, font=font,
            embedded_color=True, anchor="mm",
        )
    except Exception:
        pass
    return img


# ----- Icon vector tự vẽ cho nút (trắng, nét, đồng cỡ) -----
# Emoji của Windows là font bitmap, thu nhỏ 16px thì vỡ & mỗi cái một cỡ.
# Nên nút dùng icon tự vẽ: vẽ ở 64px rồi thu nhỏ (LANCZOS) → sắc & đồng đều.
_ICON_W = 64


def _ic_plus(d, im, c):
    d.rounded_rectangle((28, 13, 36, 51), 3, fill=c)
    d.rounded_rectangle((13, 28, 51, 36), 3, fill=c)


def _ic_pencil(d, im, c):
    lay = Image.new("RGBA", (_ICON_W, _ICON_W), (0, 0, 0, 0))
    dl = ImageDraw.Draw(lay)
    dl.rounded_rectangle((26, 14, 38, 42), 2, fill=c)      # thân
    dl.polygon([(26, 44), (38, 44), (32, 55)], fill=c)     # đầu nhọn
    dl.rounded_rectangle((26, 7, 38, 12), 2, fill=c)       # tẩy
    im.alpha_composite(lay.rotate(-40, resample=Image.BICUBIC, center=(32, 32)))


def _ic_trash(d, im, c):
    d.rounded_rectangle((13, 16, 51, 22), 2, fill=c)       # nắp
    d.rounded_rectangle((26, 9, 38, 17), 2, fill=c)        # quai
    d.polygon([(18, 22), (46, 22), (43, 54), (21, 54)], fill=c)  # thân
    for x in (27, 32, 37):                                 # đục rãnh (xuyên thấu)
        d.line((x, 27, x, 49), fill=(0, 0, 0, 0), width=2)


def _ic_folder(d, im, c):
    d.polygon([(12, 20), (26, 20), (30, 25), (52, 25), (52, 50), (12, 50)], fill=c)


def _ic_download(d, im, c):
    d.rounded_rectangle((28, 9, 36, 34), 2, fill=c)        # thân mũi tên
    d.polygon([(20, 30), (44, 30), (32, 48)], fill=c)      # đầu mũi tên
    d.rounded_rectangle((13, 50, 51, 56), 2, fill=c)       # khay


def _ic_refresh(d, im, c):
    d.arc((13, 13, 51, 51), start=-70, end=200, fill=c, width=7)
    a = math.radians(-70)
    ex, ey = 32 + 19 * math.cos(a), 32 + 19 * math.sin(a)
    d.polygon([(ex + 9, ey - 2), (ex - 3, ey - 9), (ex - 4, ey + 6)], fill=c)


def _ic_search(d, im, c):
    d.ellipse((14, 14, 40, 40), outline=c, width=7)
    d.line((39, 39, 53, 53), fill=c, width=8)


def _ic_x(d, im, c):
    d.line((17, 17, 47, 47), fill=c, width=8)
    d.line((47, 17, 17, 47), fill=c, width=8)


def _ic_save(d, im, c):
    d.rounded_rectangle((11, 11, 53, 53), 4, fill=c)
    d.rectangle((32, 11, 44, 24), fill=(0, 0, 0, 0))         # cửa trượt
    d.rounded_rectangle((19, 34, 45, 53), 1, fill=(0, 0, 0, 0))  # nhãn
    d.rounded_rectangle((23, 37, 41, 50), 1, fill=c)


def _ic_copy(d, im, c):
    d.rounded_rectangle((24, 12, 50, 42), 3, outline=c, width=5)
    d.rounded_rectangle((14, 22, 40, 52), 3, fill=c)
    d.rounded_rectangle((19, 27, 35, 47), 2, fill=(0, 0, 0, 0))


def _ic_check(d, im, c):
    d.line((15, 33, 27, 46), fill=c, width=9)
    d.line((27, 46, 50, 17), fill=c, width=9)
    d.ellipse((23, 42, 31, 50), fill=c)


def _ic_ban(d, im, c):
    d.ellipse((12, 12, 52, 52), outline=c, width=6)
    d.line((21, 21, 43, 43), fill=c, width=6)


def _ic_mail(d, im, c):
    d.rounded_rectangle((9, 16, 55, 48), 3, fill=c)
    d.line((11, 19, 32, 35), fill=(0, 0, 0, 0), width=4)
    d.line((53, 19, 32, 35), fill=(0, 0, 0, 0), width=4)


def _ic_calendar(d, im, c):
    d.rounded_rectangle((18, 8, 24, 18), 1, fill=c)
    d.rounded_rectangle((40, 8, 46, 18), 1, fill=c)
    d.rounded_rectangle((10, 13, 54, 54), 3, fill=c)
    d.rectangle((10, 26, 54, 29), fill=(0, 0, 0, 0))
    for gy in (35, 45):
        for gx in (18, 30, 42):
            d.rectangle((gx, gy, gx + 6, gy + 5), fill=(0, 0, 0, 0))


def _ic_chevron_left(d, im, c):
    d.line((40, 13, 23, 32), fill=c, width=8)
    d.line((23, 32, 40, 51), fill=c, width=8)


def _ic_chevron_right(d, im, c):
    d.line((24, 13, 41, 32), fill=c, width=8)
    d.line((41, 32, 24, 51), fill=c, width=8)


def _ic_play(d, im, c):
    d.polygon([(22, 14), (22, 50), (50, 32)], fill=c)


def _ic_sparkles(d, im, c):
    d.polygon([(27, 8), (31, 24), (47, 29), (31, 34),
               (27, 50), (23, 34), (7, 29), (23, 24)], fill=c)
    d.polygon([(48, 10), (50, 17), (57, 19), (50, 21),
               (48, 28), (46, 21), (39, 19), (46, 17)], fill=c)


def _ic_file(d, im, c):
    d.rounded_rectangle((16, 9, 48, 55), 3, fill=c)
    d.polygon([(40, 9), (48, 17), (40, 17)], fill=(0, 0, 0, 0))   # góc gập
    for ly in (28, 36, 44):
        d.line((23, ly, 41, ly), fill=(0, 0, 0, 0), width=3)


_ICON_DRAWERS = {
    "plus": _ic_plus, "pencil": _ic_pencil, "trash": _ic_trash,
    "folder": _ic_folder, "download": _ic_download, "refresh": _ic_refresh,
    "search": _ic_search, "x": _ic_x, "save": _ic_save, "copy": _ic_copy,
    "check": _ic_check, "ban": _ic_ban, "mail": _ic_mail,
    "calendar": _ic_calendar, "chevron_left": _ic_chevron_left,
    "chevron_right": _ic_chevron_right, "play": _ic_play,
    "sparkles": _ic_sparkles, "file": _ic_file,
}


def button_icon(name, size=18, gap=7, color=(255, 255, 255, 255)):
    """Trả về PhotoImage icon `name` (trắng mặc định) kèm khoảng trống bên phải
    `gap` — dùng cho ttk.Button(image=..., compound='left') để mọi nút cùng cỡ
    icon và cùng khoảng cách icon↔chữ. None nếu không có Pillow."""
    if not _PIL_OK or name not in _ICON_DRAWERS:
        return None
    key = ("btnicon", name, size, gap, color)
    if key not in _img_cache:
        im = Image.new("RGBA", (_ICON_W, _ICON_W), (0, 0, 0, 0))
        _ICON_DRAWERS[name](ImageDraw.Draw(im), im, color)
        small = im.resize((size, size), Image.LANCZOS)
        canvas = Image.new("RGBA", (size + gap, size), (0, 0, 0, 0))
        canvas.paste(small, (0, 0), small)
        _img_cache[key] = ImageTk.PhotoImage(canvas)
    return _img_cache[key]


def rounded_button_style(key, fill, fg, hover=None, active=None,
                         radius=8, padding=(14, 0), font_size=10, bg=None):
    """Tạo (một lần) style nút bo góc — nền vẽ bằng ảnh PIL, dùng cơ chế
    9-slice (`border`) để góc luôn tròn dù nút co giãn theo chữ.

    Trả về tên style để truyền vào ttk.Button(style=...). Nếu không có Pillow
    thì trả về None để bên gọi tự dùng bootstyle thường.

    key: định danh ASCII (không dùng emoji) để đặt tên style/element.
    bg:  màu nền vùng chứa nút (để 4 góc trong suốt hòa đúng nền, không lộ
         viền trắng). Mặc định nền thẻ.
    """
    import ttkbootstrap as _ttk
    bg = bg or theme.CARD_BG
    style_name = f"{key}.Rounded.TButton"
    if not _PIL_OK:
        return None
    if style_name in _rounded_ready:
        return style_name

    hover = hover or fill
    active = active or hover
    style = _ttk.Style.get_instance()

    def _mk(color):
        ss = 4
        side = radius * 2 + 4            # ảnh nhỏ, phần giữa để 9-slice kéo giãn
        big = Image.new("RGBA", (side * ss, side * ss), (0, 0, 0, 0))
        ImageDraw.Draw(big).rounded_rectangle(
            [0, 0, side * ss - 1, side * ss - 1],
            radius=radius * ss, fill=_hex_to_rgb(color) + (255,))
        return ImageTk.PhotoImage(big.resize((side, side), Image.LANCZOS))

    img_n, img_h, img_a = _mk(fill), _mk(hover), _mk(active)
    _img_cache[(style_name, "n")] = img_n
    _img_cache[(style_name, "h")] = img_h
    _img_cache[(style_name, "a")] = img_a

    elem = f"{key}.rbtn"
    try:
        style.element_create(elem, "image", img_n,
                             ("pressed", img_a), ("active", img_h),
                             border=radius, sticky="nsew")
    except tk.TclError:
        pass  # element đã tồn tại — bỏ qua
    style.layout(style_name, [
        (elem, {"sticky": "nsew", "children": [
            ("Button.padding", {"sticky": "nsew", "children": [
                ("Button.label", {"sticky": "nsew"}),
            ]}),
        ]}),
    ])
    style.configure(style_name, font=(theme.FONT_FAMILY, font_size),
                    foreground=fg, background=bg,
                    padding=padding, anchor="center",
                    relief="flat", borderwidth=0, focusthickness=0)
    style.map(style_name,
              foreground=[("disabled", "#cbd5e1")],
              background=[("active", bg), ("!active", bg)])
    _rounded_ready.add(style_name)
    return style_name


# ----- Nút chuẩn của app: bo góc + màu theo variant + icon vector -----
BUTTON_VARIANTS = {
    "primary": ("#6366f1", "#4f46e5", "#ffffff"),   # indigo — hành động chính
    "success": ("#22c55e", "#16a34a", "#ffffff"),   # xanh lá — tạo/lưu/xác nhận
    "danger":  ("#ef4444", "#dc2626", "#ffffff"),   # đỏ — xóa/từ chối
    "warning": ("#f59e0b", "#d97706", "#ffffff"),   # cam — mở/cảnh báo nhẹ
    "info":    ("#3b82f6", "#2563eb", "#ffffff"),   # xanh dương — sửa/thông tin
    "neutral": ("#e2e8f0", "#cbd5e1", "#334155"),   # xám nhạt — phụ (Hủy/Bỏ qua)
}


def button(parent, text="", variant="primary", icon=None, command=None,
           icon_only=False, **kw):
    """Nút chuẩn toàn app: nền bo góc màu `variant` + icon vector `icon` (trái)
    + chữ `text`. Trả về ttk.Button (bên gọi tự .pack()/.grid()).

    variant ∈ BUTTON_VARIANTS. icon ∈ tên trong _ICON_DRAWERS (hoặc None).
    icon_only=True → chỉ hiện icon (nút vuông nhỏ, vd ◀ ▶ ✖).
    Không có Pillow → tự lùi về ttk.Button bootstyle thường.
    """
    fill, hover, fg = BUTTON_VARIANTS.get(variant, BUTTON_VARIANTS["primary"])
    try:
        bg = parent.cget("background")
    except Exception:
        bg = theme.CARD_BG
    if not (isinstance(bg, str) and bg.startswith("#") and len(bg) == 7):
        bg = theme.CARD_BG

    style_name = rounded_button_style(
        f"btn_{variant}_{bg.lstrip('#')}", fill, fg, hover=hover, bg=bg)
    img = button_icon(icon, color=_hex_to_rgb(fg) + (255,)) if icon else None

    if style_name is None:                    # không có Pillow → fallback
        boot = {"neutral": "secondary"}.get(variant, variant)
        return ttk.Button(parent, text=text or "", bootstyle=boot,
                          command=command, **kw)
    if img is not None and not icon_only:
        b = ttk.Button(parent, image=img, text=text, compound="left",
                       style=style_name, command=command, **kw)
    elif img is not None:
        b = ttk.Button(parent, image=img, style=style_name,
                       command=command, **kw)
    else:
        b = ttk.Button(parent, text=text, style=style_name,
                       command=command, **kw)
    if img is not None:
        b.image = img
    return b


def polish_comboboxes(name="TCombobox", pad_y=6):
    """Làm combobox trông sạch, đồng bộ với ô nhập (nền trắng, bỏ highlight xanh
    khi readonly, cao & padding dễ nhìn). Truyền `name` để chỉ áp cho một style
    riêng (không ảnh hưởng combobox ở tool khác). Gọi lại nhiều lần vô hại."""
    import ttkbootstrap as _ttk
    style = _ttk.Style.get_instance()
    style.configure(name, padding=(10, pad_y), arrowsize=14,
                    borderwidth=1, relief="flat")
    # Đưa downarrow vào trong field và KHÔNG cho giãn dọc (sticky rỗng) → nó
    # được căn giữa theo chiều cao, vẽ đúng 1 lần. (Trước để "ns" khiến ttk
    # lát hình mũi tên thành nhiều cái khi field cao lên.)
    style.layout(name, [
        ("Combobox.field", {"sticky": "nswe", "children": [
            ("Combobox.downarrow", {"side": "right", "sticky": ""}),
            ("Combobox.padding", {"sticky": "nswe", "children": [
                ("Combobox.textarea", {"sticky": "nswe"}),
            ]}),
        ]}),
    ])
    style.map(
        name,
        fieldbackground=[("readonly", theme.CARD_BG)],
        background=[("readonly", theme.CARD_BG)],
        foreground=[("readonly", theme.TEXT)],
        selectbackground=[("readonly", theme.CARD_BG)],
        selectforeground=[("readonly", theme.TEXT)],
        bordercolor=[("focus", theme.ACCENT), ("!focus", theme.BORDER)],
        arrowcolor=[("readonly", theme.MUTED)],
    )


def icon_badge(parent, emoji, color, bg, size=46, radius=13, font_size=None):
    """Chip icon: nền tông nhạt của màu nhóm + emoji màu (cho thẻ/header)."""
    if _PIL_OK:
        key = ("chip", emoji, color, size, radius)
        if key not in _img_cache:
            _img_cache[key] = ImageTk.PhotoImage(
                _render_icon(emoji, size, _tint(color, 0.18), radius, 0.58))
        lbl = tk.Label(parent, image=_img_cache[key], bg=bg, bd=0)
        lbl.image = _img_cache[key]
        return lbl
    return tk.Label(parent, text=emoji, bg=bg,
                    font=(theme.FONT_EMOJI, int(size * 0.45)))


def icon_glyph(parent, emoji, bg, size=22):
    """Chỉ emoji màu (không nền) — dùng cho sidebar, sạch và không bị 'từng cục'."""
    if _PIL_OK:
        key = ("glyph", emoji, size)
        if key not in _img_cache:
            _img_cache[key] = ImageTk.PhotoImage(
                _render_icon(emoji, size, None, 0, 0.92))
        lbl = tk.Label(parent, image=_img_cache[key], bg=bg, bd=0)
        lbl.image = _img_cache[key]
        return lbl
    return tk.Label(parent, text=emoji, bg=bg,
                    font=(theme.FONT_EMOJI, int(size * 0.7)))


# ----- Hiệu ứng loading cho mọi nút bấm -----

def install_loading_buttons(loading_text="⏳ Đang xử lý…"):
    """Bọc lại ttk.Button để mọi nút hiện trạng thái 'đang xử lý' khi bấm.

    Gọi đúng một lần lúc khởi động (trước khi dựng UI). Nhờ vậy tất cả nút
    trong dự án tự có hiệu ứng loading mà không phải sửa từng nơi: khi bấm,
    nút bị vô hiệu hóa và đổi chữ thành trạng thái chờ, xử lý xong thì trả
    lại như cũ. Nút bị hủy giữa chừng (vd lệnh đóng hộp thoại) được bỏ qua.
    """
    orig = ttk.Button
    if getattr(orig, "_loading_wrapped", False):
        return  # đã cài rồi

    class _LoadingButton(orig):
        _loading_wrapped = True

        def __init__(self, master=None, **kw):
            cmd = kw.get("command")
            if callable(cmd):
                self._user_command = cmd
                kw["command"] = self._run_with_loading
            else:
                self._user_command = None
            super().__init__(master, **kw)

        def _run_with_loading(self):
            try:
                prev_text = self.cget("text")
            except tk.TclError:
                self._user_command()
                return
            try:
                self.configure(text=loading_text)
                self.state(["disabled"])
                self.update_idletasks()
            except tk.TclError:
                pass
            try:
                self._user_command()
            finally:
                try:
                    self.configure(text=prev_text)
                    self.state(["!disabled"])
                except tk.TclError:
                    pass  # nút đã bị hủy bởi chính lệnh (vd đóng hộp thoại)

    ttk.Button = _LoadingButton


# ----- Cho phép sao chép chữ ở mọi nơi -----

def _selected_text(widget):
    """Lấy phần chữ đang bôi đen (nếu có) của widget."""
    try:
        return widget.selection_get()
    except tk.TclError:
        return ""


def _full_text(widget):
    """Lấy toàn bộ nội dung chữ của widget (Label/Text/Entry/Button)."""
    if isinstance(widget, tk.Text):
        try:
            return widget.get("1.0", "end-1c")
        except tk.TclError:
            return ""
    if isinstance(widget, (tk.Entry, ttk.Entry, ttk.Combobox)):
        try:
            return widget.get()
        except tk.TclError:
            return ""
    # Label / Button…: đọc từ text, nếu rỗng thì đọc từ biến textvariable
    try:
        txt = widget.cget("text")
    except tk.TclError:
        return ""
    if not txt:
        try:
            varname = widget.cget("textvariable")
            if varname:
                txt = widget.getvar(varname)
        except tk.TclError:
            pass
    return txt or ""


def install_copy_support(root):
    """Gắn menu chuột phải "Sao chép" cho mọi widget trong ứng dụng.

    Gọi đúng một lần lúc khởi động. Text/Entry/Combobox đã bôi đen + Ctrl+C
    được sẵn, hàm này bổ sung menu chuột phải và — quan trọng nhất — cho phép
    sao chép chữ ở các nhãn (tk.Label) vốn không thể bôi đen trong Tkinter.
    """
    menu = tk.Menu(root, tearoff=0)
    state = {"widget": None}

    def _is_editable(w):
        return isinstance(w, (tk.Text, tk.Entry, ttk.Entry, ttk.Combobox))

    def _to_clipboard(text):
        if not text:
            return
        root.clipboard_clear()
        root.clipboard_append(text)

    def _copy():
        w = state["widget"]
        if w is None:
            return
        _to_clipboard(_selected_text(w) or _full_text(w))

    def _cut():
        w = state["widget"]
        if w is None:
            return
        sel = _selected_text(w)
        if not sel:
            return
        _to_clipboard(sel)
        try:
            w.delete("sel.first", "sel.last")
        except tk.TclError:
            pass

    def _paste():
        w = state["widget"]
        if w is None:
            return
        try:
            data = root.clipboard_get()
        except tk.TclError:
            return
        try:
            w.event_generate("<<Paste>>")
        except tk.TclError:
            try:
                w.insert("insert", data)
            except tk.TclError:
                pass

    def _select_all():
        w = state["widget"]
        if w is None:
            return
        try:
            if isinstance(w, tk.Text):
                w.tag_add("sel", "1.0", "end-1c")
            else:
                w.select_range(0, "end")
        except tk.TclError:
            pass

    def _popup(event):
        w = event.widget
        state["widget"] = w
        menu.delete(0, "end")
        menu.add_command(label="Sao chép", command=_copy)
        if _is_editable(w):
            menu.add_command(label="Cắt", command=_cut)
            menu.add_command(label="Dán", command=_paste)
            menu.add_separator()
            menu.add_command(label="Chọn tất cả", command=_select_all)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    # Bind cho mọi widget (kể cả tạo sau này).
    root.bind_all("<Button-3>", _popup, add="+")
