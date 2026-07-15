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

    ttk.Button(
        row, text="📁 Thư mục", bootstyle="secondary-outline", command=pick_folder,
    ).pack(side="left", padx=(8, 0))
    ttk.Button(
        row, text="📄 File Excel", bootstyle="secondary-outline", command=pick_file,
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
