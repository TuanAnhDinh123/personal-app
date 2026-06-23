"""Các widget dựng sẵn để tool tái sử dụng cho gọn.

Tất cả trả về một biến (StringVar/BooleanVar) để sau này gắn logic thật.
"""
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

    ttk.Button(
        row, text="Chọn…", bootstyle="secondary-outline", command=browse,
    ).pack(side="left", padx=(8, 0))
    return var


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
    box = tk.Text(
        block, height=height, wrap="word", relief="solid", bd=1,
        font=(theme.FONT_FAMILY, 10), bg="#ffffff", fg=theme.TEXT,
        highlightthickness=1, highlightbackground=theme.BORDER,
        padx=8, pady=6,
    )
    box.pack(fill="x")
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
