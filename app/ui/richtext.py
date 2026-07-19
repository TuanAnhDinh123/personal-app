"""Ô soạn thảo văn bản có định dạng (rich text) → xuất HTML.

Dùng cho các ô nội dung email cần in đậm / nghiêng / gạch chân / màu chữ /
gạch đầu dòng. Bên trong là một tk.Text gắn "tag" định dạng; khi cần gửi mail
thì serialize sang HTML đơn giản (b/i/u/span màu/br) để đặt vào HTMLBody của
Outlook. Có thể nạp lại HTML đó về đúng định dạng (set_html) để chỉnh tiếp.

Giới hạn: đây là editor mini — chỉ định dạng ký tự cơ bản (không chèn ảnh, bảng…).
Riêng trường hợp vừa in đậm VỪA in nghiêng trên cùng đoạn chữ thì màn hình có
thể chỉ hiện một kiểu, nhưng HTML xuất ra vẫn đúng (<b><i>…</i></b>).
"""
import html as _html
import re
from html.parser import HTMLParser

import tkinter as tk
from tkinter import colorchooser

import ttkbootstrap as ttk

from app.ui import theme

_INLINE = {
    "bold": ("<b>", "</b>"),
    "italic": ("<i>", "</i>"),
    "underline": ("<u>", "</u>"),
}


# --------------------------------------------------------------- Text -> HTML
def text_to_html(text):
    """Chuyển nội dung + tag định dạng của một tk.Text thành HTML."""
    active = set()          # bold / italic / underline đang bật
    color = [None]          # màu chữ đang bật (hoặc None)

    def wrap(chunk):
        pre, post = "", ""
        if color[0]:
            pre += f'<span style="color:{color[0]}">'
            post = "</span>" + post
        for name in ("bold", "italic", "underline"):
            if name in active:
                open_tag, close_tag = _INLINE[name]
                pre += open_tag
                post = close_tag + post
        return pre + _html.escape(chunk) + post

    out = []
    for key, value, _idx in text.dump("1.0", "end-1c", tag=True, text=True):
        if key == "tagon":
            if value in _INLINE:
                active.add(value)
            elif value.startswith("color="):
                color[0] = value[len("color="):]
        elif key == "tagoff":
            if value in _INLINE:
                active.discard(value)
            elif value.startswith("color="):
                color[0] = None
        elif key == "text":
            parts = value.split("\n")
            for i, part in enumerate(parts):
                if i:
                    out.append("<br>")
                if part:
                    out.append(wrap(part))
    return "".join(out)


# --------------------------------------------------------------- HTML -> Text
def _parse_color(style):
    m = re.search(r'color\s*:\s*([^;]+)', style or "", re.IGNORECASE)
    return m.group(1).strip() if m else None


class _Loader(HTMLParser):
    """Phân tích HTML đơn giản → chuỗi phẳng + danh sách vùng cần gắn tag."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.buf = []
        self.pos = 0
        self.stack = []     # (htmltag, ourname|None, start_pos)
        self.spans = []     # (ourname, start, end)

    def _emit(self, s):
        self.buf.append(s)
        self.pos += len(s)

    def _at_line_start(self):
        return self.pos == 0 or (self.buf and self.buf[-1].endswith("\n"))

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == "br":
            self._emit("\n")
            return
        if tag in ("p", "div") and not self._at_line_start():
            self._emit("\n")
        if tag == "li":
            if not self._at_line_start():
                self._emit("\n")
            self._emit("• ")

        ourname = None
        if tag in ("b", "strong"):
            ourname = "bold"
        elif tag in ("i", "em"):
            ourname = "italic"
        elif tag == "u":
            ourname = "underline"
        elif tag == "span":
            c = _parse_color(dict(attrs).get("style"))
            if c:
                ourname = f"color={c}"
        self.stack.append((tag, ourname, self.pos))

    def handle_startendtag(self, tag, attrs):
        if tag.lower() == "br":
            self._emit("\n")

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag == "br":
            return
        for i in range(len(self.stack) - 1, -1, -1):
            if self.stack[i][0] == tag:
                _, ourname, start = self.stack.pop(i)
                if ourname and self.pos > start:
                    self.spans.append((ourname, start, self.pos))
                break
        if tag in ("p", "div"):
            self._emit("\n")

    def handle_data(self, data):
        self._emit(data)


# ------------------------------------------------------------------- Widget
class RichText(tk.Frame):
    """Ô soạn thảo có thanh công cụ định dạng, đọc/ghi bằng HTML."""

    def __init__(self, parent, height=12, bg=None):
        bg = bg or theme.CARD_BG
        super().__init__(self._normalize(parent), bg=bg)
        self._base_size = 10
        self._color_tags = set()

        self._build_toolbar(bg)

        holder = tk.Frame(
            self, bg="#ffffff", relief="solid", bd=1,
            highlightthickness=1, highlightbackground=theme.BORDER,
        )
        holder.pack(fill="both", expand=True)
        self.text = tk.Text(
            holder, height=height, wrap="word", relief="flat", bd=0,
            font=(theme.FONT_FAMILY, self._base_size), bg="#ffffff",
            fg=theme.TEXT, highlightthickness=0, padx=8, pady=6, undo=True,
        )
        vsb = ttk.Scrollbar(holder, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=vsb.set)
        self.text.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        fam, sz = theme.FONT_FAMILY, self._base_size
        self.text.tag_configure("bold", font=(fam, sz, "bold"))
        self.text.tag_configure("italic", font=(fam, sz, "italic"))
        self.text.tag_configure("underline", underline=True)

        self.text.bind("<MouseWheel>", self._on_wheel)

    @staticmethod
    def _normalize(parent):
        return parent

    # -------------------------------------------------------------- toolbar
    def _build_toolbar(self, bg):
        bar = tk.Frame(self, bg=bg)
        bar.pack(fill="x", pady=(0, 4))

        def tb(txt, cmd, **font_kw):
            b = tk.Button(
                bar, text=txt, command=cmd, relief="groove", bd=1,
                bg="#f2f4f9", fg=theme.TEXT, activebackground="#e2e6f0",
                cursor="hand2", padx=8, pady=1,
                font=font_kw.get("font", (theme.FONT_FAMILY, 9)),
            )
            b.pack(side="left", padx=(0, 4))
            return b

        tb("B", lambda: self._toggle("bold"), font=(theme.FONT_FAMILY, 10, "bold"))
        tb("I", lambda: self._toggle("italic"), font=(theme.FONT_FAMILY, 10, "italic"))
        tb("U", lambda: self._toggle("underline"), font=(theme.FONT_FAMILY, 10, "underline"))
        tb("A🎨", self._pick_color)
        tb("• List", self._bullet)
        tb("✕ Định dạng", self._clear_format)

    # ------------------------------------------------------------ thao tác
    def _selection(self):
        try:
            return self.text.index("sel.first"), self.text.index("sel.last")
        except tk.TclError:
            return None

    def _toggle(self, tag):
        sel = self._selection()
        if not sel:
            return
        start, end = sel
        if tag in self.text.tag_names("sel.first"):
            self.text.tag_remove(tag, start, end)
        else:
            self.text.tag_add(tag, start, end)
        self.text.focus_set()

    def _ensure_color_tag(self, color):
        name = f"color={color}"
        if name not in self._color_tags:
            try:
                self.text.tag_configure(name, foreground=color)
            except tk.TclError:
                return None
            self._color_tags.add(name)
        return name

    def _pick_color(self):
        sel = self._selection()
        if not sel:
            return
        start, end = sel
        _rgb, hexv = colorchooser.askcolor(parent=self, title="Chọn màu chữ")
        if not hexv:
            return
        for t in list(self._color_tags):
            self.text.tag_remove(t, start, end)
        name = self._ensure_color_tag(hexv)
        if name:
            self.text.tag_add(name, start, end)
        self.text.focus_set()

    def _bullet(self):
        sel = self._selection()
        if sel:
            l1 = int(sel[0].split(".")[0])
            l2 = int(sel[1].split(".")[0])
        else:
            l1 = l2 = int(self.text.index("insert").split(".")[0])
        for ln in range(l1, l2 + 1):
            self.text.insert(f"{ln}.0", "• ")
        self.text.focus_set()

    def _clear_format(self):
        sel = self._selection()
        if not sel:
            return
        start, end = sel
        for t in ("bold", "italic", "underline"):
            self.text.tag_remove(t, start, end)
        for t in list(self._color_tags):
            self.text.tag_remove(t, start, end)

    def _on_wheel(self, event):
        first, last = self.text.yview()
        if first <= 0.0 and last >= 1.0:
            return None            # đã hiện hết -> để trang ngoài cuộn
        self.text.yview_scroll(int(-event.delta / 120), "units")
        return "break"

    # ------------------------------------------------------------ đọc / ghi
    def get_html(self):
        return text_to_html(self.text)

    def get_text(self):
        return self.text.get("1.0", "end-1c")

    def set_html(self, html):
        self.text.delete("1.0", "end")
        loader = _Loader()
        loader.feed(html or "")
        loader.close()
        self.text.insert("1.0", "".join(loader.buf))
        for name, start, end in loader.spans:
            if name.startswith("color="):
                if not self._ensure_color_tag(name[len("color="):]):
                    continue
            self.text.tag_add(
                name, f"1.0 + {start} chars", f"1.0 + {end} chars")