"""PDF → Text: trích văn bản từ PDF (Word xuất ra hoặc bản scan) + đối chiếu.

Cách hoạt động:
  - Trang nào có sẵn lớp text (PDF từ Word) → trích thẳng, nhanh & chính xác.
  - Trang nào là ảnh scan (không có lớp text) → OCR bằng Tesseract (offline).
Giao diện chia hai cột: bên trái là ảnh trang gốc, bên phải là văn bản
trích được (sửa trực tiếp) để người dùng đối chiếu rồi sao chép / lưu ra .txt.

Phụ thuộc:
  pip install pymupdf pytesseract
  + cài Tesseract-OCR (kèm gói tiếng Việt) — chỉ cần cho PDF scan.
"""
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import ttkbootstrap as ttk

from app.core.base_tool import BaseTool
from app.ui import theme, widgets

try:
    import fitz  # PyMuPDF
    _FITZ_OK = True
except ImportError:
    _FITZ_OK = False

try:
    import pytesseract
    from PIL import Image
    _PYTESSERACT_OK = True
except ImportError:
    _PYTESSERACT_OK = False


# Ngưỡng ký tự để coi một trang là "có lớp text". Dưới ngưỡng → coi là scan.
_MIN_TEXT_CHARS = 12
_OCR_DPI_OPTIONS = {"Nhanh (200 DPI)": 200, "Cân bằng (300 DPI)": 300, "Nét (400 DPI)": 400}


# ---------------------------------------------------------------------------
# Core logic (độc lập với UI, dễ test)
# ---------------------------------------------------------------------------

def tesseract_available():
    """True nếu pytesseract import được VÀ tìm thấy file thực thi Tesseract."""
    if not _PYTESSERACT_OK:
        return False
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def _ocr_page(page, lang, dpi):
    """Render một trang PDF ra ảnh rồi OCR, trả về chuỗi text."""
    zoom = dpi / 72.0
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    return pytesseract.image_to_string(img, lang=lang).strip()


def extract_pages(pdf_path, ocr_lang="vie+eng", ocr_dpi=300,
                  force_ocr=False, progress=None):
    """Trích text từng trang. Trả về list dict: {text, source}.

    source: "text"  = lấy từ lớp text có sẵn
            "ocr"   = nhận dạng bằng Tesseract
            "scan"  = là trang scan nhưng chưa OCR được (thiếu Tesseract)
    `progress(done, total, source)` được gọi sau mỗi trang (nếu truyền vào).
    """
    can_ocr = tesseract_available()
    pages = []
    with fitz.open(pdf_path) as doc:
        total = doc.page_count
        for i in range(total):
            page = doc[i]
            text = page.get_text("text").strip()
            source = "text"

            if force_ocr or len(text) < _MIN_TEXT_CHARS:
                if can_ocr:
                    text = _ocr_page(page, ocr_lang, ocr_dpi)
                    source = "ocr"
                elif source == "text" and len(text) < _MIN_TEXT_CHARS:
                    source = "scan"  # cần OCR nhưng không có Tesseract

            pages.append({"text": text, "source": source})
            if progress:
                progress(i + 1, total, source)
    return pages


# ---------------------------------------------------------------------------
# Tool (UI)
# ---------------------------------------------------------------------------

_SOURCE_BADGE = {
    "text": ("Lớp text", "info"),
    "ocr": ("OCR", "success"),
    "scan": ("Scan – chưa OCR", "warning"),
}


class PdfToTextTool(BaseTool):
    name = "PDF → Text"
    description = "Trích văn bản từ PDF (Word xuất ra hoặc bản scan) và đối chiếu với bản gốc."
    icon = "📝"
    category = "Tệp & Tài liệu"
    order = 15
    action_label = "Trích văn bản"

    # ----- Giao diện: ghi đè build() để dùng toàn bộ vùng nội dung -----
    def build(self, parent):
        self._pages = []          # kết quả trích: [{text, source}, ...]
        self._cur = 0             # trang đang xem (0-based)
        self._doc = None          # fitz.Document đang mở để render ảnh
        self._page_img = None     # giữ tham chiếu PhotoImage hiện tại
        self._busy = False

        outer = tk.Frame(parent, bg=theme.CONTENT_BG)

        if not _FITZ_OK:
            self._build_missing_lib(outer)
            return outer

        self._build_controls(outer)
        self._build_viewer(outer)
        self._build_footer(outer)
        self._refresh_view()
        return outer

    def _build_missing_lib(self, parent):
        card = self._card(parent)
        widgets.section_label(card, "Thiếu thư viện")
        widgets.hint(
            card,
            "Tính năng này cần PyMuPDF. Hãy cài rồi mở lại app:\n"
            "    pip install pymupdf pytesseract",
        )

    def _card(self, parent, **pack):
        card = tk.Frame(
            parent, bg=theme.CARD_BG,
            highlightbackground=theme.BORDER, highlightthickness=1,
        )
        card.pack(fill="x", **pack)
        inner = tk.Frame(card, bg=theme.CARD_BG)
        inner.pack(fill="both", expand=True, padx=20, pady=16)
        return inner

    # ----- Thanh điều khiển: chọn file + tùy chọn + nút trích -----
    def _build_controls(self, parent):
        card = self._card(parent, pady=(0, 12))

        self._file_var = tk.StringVar()
        row = tk.Frame(card, bg=theme.CARD_BG)
        row.pack(fill="x")
        tk.Label(
            row, text="File PDF", bg=theme.CARD_BG, fg=theme.TEXT,
            font=(theme.FONT_FAMILY, 9),
        ).pack(anchor="w", pady=(0, 4))
        pick = tk.Frame(card, bg=theme.CARD_BG)
        pick.pack(fill="x")
        ttk.Entry(pick, textvariable=self._file_var).pack(
            side="left", fill="x", expand=True, ipady=4)

        def browse():
            path = filedialog.askopenfilename(
                filetypes=[("PDF", "*.pdf"), ("Tất cả", "*.*")])
            if path:
                self._file_var.set(path)

        ttk.Button(pick, text="Chọn…", bootstyle="secondary-outline",
                   command=browse).pack(side="left", padx=(8, 0))

        # Hàng tùy chọn OCR
        opts = tk.Frame(card, bg=theme.CARD_BG)
        opts.pack(fill="x", pady=(12, 0))

        tk.Label(opts, text="Ngôn ngữ OCR", bg=theme.CARD_BG, fg=theme.TEXT,
                 font=(theme.FONT_FAMILY, 9)).pack(side="left")
        self._lang_var = tk.StringVar(value="vie+eng")
        ttk.Entry(opts, textvariable=self._lang_var, width=12).pack(
            side="left", padx=(8, 18), ipady=2)

        tk.Label(opts, text="Độ nét scan", bg=theme.CARD_BG, fg=theme.TEXT,
                 font=(theme.FONT_FAMILY, 9)).pack(side="left")
        self._dpi_var = tk.StringVar(value="Cân bằng (300 DPI)")
        ttk.Combobox(opts, textvariable=self._dpi_var, state="readonly",
                     width=20, values=list(_OCR_DPI_OPTIONS)).pack(
            side="left", padx=(8, 18), ipady=1)

        self._force_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opts, text="Luôn OCR (kể cả khi có lớp text)",
                        variable=self._force_var,
                        bootstyle="round-toggle").pack(side="left")

        # Nút trích + trạng thái
        bar = tk.Frame(card, bg=theme.CARD_BG)
        bar.pack(fill="x", pady=(14, 0))
        self._run_btn = ttk.Button(
            bar, text=self.action_label, bootstyle="primary",
            command=self._start_extract)
        self._run_btn.pack(side="left", ipadx=10, ipady=3)
        self._status = tk.Label(
            bar, text="", bg=theme.CARD_BG, fg=theme.MUTED,
            font=(theme.FONT_FAMILY, 9))
        self._status.pack(side="left", padx=14)

        if not tesseract_available():
            widgets.hint(
                card,
                "💡 PDF từ Word sẽ trích được ngay. Với PDF scan cần cài Tesseract-OCR: "
                "tải bản Windows tại github.com/UB-Mannheim/tesseract/wiki, khi cài nhớ "
                "tick gói ngôn ngữ \"Vietnamese\". Cài xong mở lại app là dùng được.",
            )

    # ----- Vùng đối chiếu: ảnh gốc (trái) | văn bản trích (phải) -----
    def _build_viewer(self, parent):
        paned = ttk.Panedwindow(parent, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # --- Cột trái: trang gốc ---
        left = tk.Frame(paned, bg=theme.CARD_BG,
                        highlightbackground=theme.BORDER, highlightthickness=1)
        nav = tk.Frame(left, bg=theme.CARD_BG)
        nav.pack(fill="x", padx=10, pady=8)
        tk.Label(nav, text="Bản gốc", bg=theme.CARD_BG, fg=theme.TEXT,
                 font=(theme.FONT_FAMILY, 10, "bold")).pack(side="left")
        self._prev_btn = ttk.Button(nav, text="◀", width=3,
                                    bootstyle="secondary-outline",
                                    command=lambda: self._go(-1))
        self._next_btn = ttk.Button(nav, text="▶", width=3,
                                    bootstyle="secondary-outline",
                                    command=lambda: self._go(1))
        self._page_lbl = tk.Label(nav, text="– / –", bg=theme.CARD_BG,
                                  fg=theme.MUTED, font=(theme.FONT_FAMILY, 9))
        self._src_badge = ttk.Label(nav, text="", bootstyle="secondary")
        self._next_btn.pack(side="right")
        self._page_lbl.pack(side="right", padx=8)
        self._prev_btn.pack(side="right")
        self._src_badge.pack(side="right", padx=12)

        self._canvas = tk.Canvas(left, bg="#e9edf3", highlightthickness=0,
                                 height=520)
        self._canvas.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self._canvas.bind("<Configure>", lambda _e: self._render_page())

        # --- Cột phải: văn bản trích được ---
        right = tk.Frame(paned, bg=theme.CARD_BG,
                         highlightbackground=theme.BORDER, highlightthickness=1)
        head = tk.Frame(right, bg=theme.CARD_BG)
        head.pack(fill="x", padx=10, pady=8)
        tk.Label(head, text="Văn bản trích được (có thể sửa)", bg=theme.CARD_BG,
                 fg=theme.TEXT, font=(theme.FONT_FAMILY, 10, "bold")).pack(side="left")

        txt_wrap = tk.Frame(right, bg=theme.CARD_BG)
        txt_wrap.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self._text = tk.Text(
            txt_wrap, wrap="word", relief="solid", bd=1, undo=True,
            font=(theme.FONT_FAMILY, 10), bg="#ffffff", fg=theme.TEXT,
            highlightthickness=1, highlightbackground=theme.BORDER,
            padx=10, pady=8, state="disabled",
        )
        vsb = ttk.Scrollbar(txt_wrap, orient="vertical", command=self._text.yview)
        self._text.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._text.pack(side="left", fill="both", expand=True)

        paned.add(left, weight=1)
        paned.add(right, weight=1)

    # ----- Chân: sao chép / lưu -----
    def _build_footer(self, parent):
        bar = tk.Frame(parent, bg=theme.CONTENT_BG)
        bar.pack(fill="x", pady=(12, 0))
        self._copy_btn = ttk.Button(bar, text="📋 Sao chép trang này",
                                   bootstyle="secondary",
                                   command=self._copy_current)
        self._copy_btn.pack(side="left", ipady=2)
        self._save_btn = ttk.Button(bar, text="💾 Lưu toàn bộ ra .txt",
                                   bootstyle="success",
                                   command=self._save_txt)
        self._save_btn.pack(side="left", padx=10, ipady=2)

    # ----- Trích xuất (chạy nền) -----
    def _start_extract(self):
        if self._busy:
            return
        path = self._file_var.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showerror("Lỗi", "Vui lòng chọn một file PDF hợp lệ.")
            return
        if self._force_var.get() and not tesseract_available():
            messagebox.showwarning(
                "Chưa có Tesseract",
                "Bạn bật \"Luôn OCR\" nhưng chưa cài Tesseract-OCR.\n"
                "Hãy cài Tesseract hoặc tắt tùy chọn này.")
            return

        self._busy = True
        self._run_btn.config(state="disabled")
        self._set_status("Đang mở PDF…")

        lang = self._lang_var.get().strip() or "vie+eng"
        dpi = _OCR_DPI_OPTIONS.get(self._dpi_var.get(), 300)
        force = self._force_var.get()

        def worker():
            try:
                def prog(done, total, _src):
                    self._ui(lambda: self._set_status(
                        f"Đang xử lý trang {done}/{total}…"))
                pages = extract_pages(path, ocr_lang=lang, ocr_dpi=dpi,
                                      force_ocr=force, progress=prog)
            except Exception as exc:
                self._ui(lambda e=exc: self._on_error(e))
                return
            self._ui(lambda: self._on_done(path, pages))

        threading.Thread(target=worker, daemon=True).start()

    def _on_error(self, exc):
        self._busy = False
        self._run_btn.config(state="normal")
        self._set_status("")
        messagebox.showerror("Lỗi", f"Không trích được văn bản:\n{exc}")

    def _on_done(self, path, pages):
        self._busy = False
        self._run_btn.config(state="normal")

        # Đóng doc cũ, mở doc mới để render ảnh
        if self._doc is not None:
            try:
                self._doc.close()
            except Exception:
                pass
        self._doc = fitz.open(path)
        self._pages = pages
        self._cur = 0

        n_ocr = sum(1 for p in pages if p["source"] == "ocr")
        n_scan = sum(1 for p in pages if p["source"] == "scan")
        msg = f"✅ Xong {len(pages)} trang."
        if n_ocr:
            msg += f" OCR: {n_ocr} trang."
        if n_scan:
            msg += f" ⚠ {n_scan} trang scan chưa OCR (thiếu Tesseract)."
        self._set_status(msg)
        self._refresh_view()

    # ----- Hiển thị -----
    def _go(self, delta):
        if not self._pages:
            return
        self._commit_text()  # lưu chỉnh sửa của trang hiện tại
        self._cur = max(0, min(len(self._pages) - 1, self._cur + delta))
        self._refresh_view()

    def _refresh_view(self):
        has = bool(self._pages)
        state = "normal" if has else "disabled"
        for b in (self._prev_btn, self._next_btn, self._copy_btn, self._save_btn):
            b.config(state=state)

        if not has:
            self._page_lbl.config(text="– / –")
            self._src_badge.config(text="")
            self._set_text("", editable=False)
            self._canvas.delete("all")
            return

        page = self._pages[self._cur]
        self._page_lbl.config(text=f"{self._cur + 1} / {len(self._pages)}")
        label, style = _SOURCE_BADGE.get(page["source"], ("", "secondary"))
        self._src_badge.config(text=label, bootstyle=style)
        self._prev_btn.config(state="normal" if self._cur > 0 else "disabled")
        self._next_btn.config(
            state="normal" if self._cur < len(self._pages) - 1 else "disabled")

        self._set_text(page["text"], editable=True)
        self._render_page()

    def _render_page(self):
        """Render trang PDF hiện tại vừa khít chiều rộng canvas."""
        if not self._pages or self._doc is None:
            return
        cw = self._canvas.winfo_width()
        ch = self._canvas.winfo_height()
        if cw <= 1 or ch <= 1:
            return
        page = self._doc[self._cur]
        rect = page.rect
        zoom = max(0.1, min(cw / rect.width, ch / rect.height))
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        from PIL import ImageTk
        self._page_img = ImageTk.PhotoImage(img)
        self._canvas.delete("all")
        self._canvas.create_image(cw // 2, ch // 2, image=self._page_img)

    def _set_text(self, value, editable):
        self._text.config(state="normal")
        self._text.delete("1.0", "end")
        if value:
            self._text.insert("1.0", value)
        self._text.config(state="normal" if editable else "disabled")

    def _commit_text(self):
        """Lưu nội dung đang sửa vào trang hiện tại."""
        if self._pages and 0 <= self._cur < len(self._pages):
            self._pages[self._cur]["text"] = self._text.get("1.0", "end-1c")

    # ----- Sao chép / lưu -----
    def _copy_current(self):
        if not self._pages:
            return
        self._commit_text()
        text = self._pages[self._cur]["text"]
        self._text.clipboard_clear()
        self._text.clipboard_append(text)
        self._set_status(f"📋 Đã sao chép trang {self._cur + 1}.")

    def _save_txt(self):
        if not self._pages:
            return
        self._commit_text()
        default = "ket_qua.txt"
        src = self._file_var.get().strip()
        if src:
            default = os.path.splitext(os.path.basename(src))[0] + ".txt"
        path = filedialog.asksaveasfilename(
            defaultextension=".txt", initialfile=default,
            filetypes=[("Văn bản", "*.txt")])
        if not path:
            return
        parts = []
        for i, p in enumerate(self._pages):
            parts.append(f"--- Trang {i + 1} ---\n{p['text']}")
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n\n".join(parts))
        except Exception as exc:
            messagebox.showerror("Lỗi", f"Không lưu được file:\n{exc}")
            return
        self._set_status(f"💾 Đã lưu: {os.path.basename(path)}")

    # ----- Tiện ích luồng -----
    def _ui(self, fn):
        """Lên lịch chạy fn trên luồng giao diện."""
        self._text.after(0, fn)

    def _set_status(self, text):
        self._status.config(text=text)

    def build_body(self, parent):
        # Không dùng — đã ghi đè build(). Bắt buộc cài đặt vì là abstract.
        pass
