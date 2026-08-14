"""PDF → Text: đọc PDF thành văn bản giữ bố cục, để đọc và copy sang chỗ khác.

Bố cục màn hình: thẻ điều khiển ở trên, còn lại là KHUNG ĐỌC chiếm trọn bề
ngang. Ảnh trang gốc chỉ hiện khi bấm "Show original" — trước đây nó chiếm cố
định nửa màn hình nên cả hai bên đều chật.

Khung đọc CHỈ ĐỌC và luôn hiển thị bản đã dựng định dạng (tiêu đề, danh sách,
bảng). Không có chế độ sửa tay: mục đích của tool là lấy text ra, mà đã cho sửa
thì phải có chỗ lưu bản sửa — nửa vời sẽ mất nội dung khi chuyển trang.

Trích xuất chạy trong ProgressDialog (có thanh tiến trình + nút Hủy). Luồng giao
diện KHÔNG gọi thẳng vào MuPDF; ảnh trang lấy qua pdf_text.render_png() trả về
PNG bytes — MuPDF không thread-safe, dùng chéo luồng làm app tắt ngang.
"""
import os

from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QApplication, QFileDialog, QFrame, QHBoxLayout, QLabel, QScrollArea,
    QSplitter, QVBoxLayout, QWidget,
)

from app.core import pdf_text, settings
from app.core.pdf_text import (
    MODE_AUTO, MODE_OCR, MODE_TEXT, OCR_DPI_OPTIONS,
    SRC_EMPTY, SRC_OCR, SRC_SCAN, SRC_TEXT,
)
from app_qt import theme, widgets
from app_qt.base_tool import BaseTool
from app_qt.components.progress_dialog import ProgressDialog

try:
    import pymupdf  # noqa: F401  (chỉ để biết thư viện đã cài chưa)
    _PYMUPDF_OK = True
except ImportError:
    try:
        import fitz  # noqa: F401
        _PYMUPDF_OK = True
    except ImportError:
        _PYMUPDF_OK = False

# Nhãn hiển thị của từng chế độ xử lý (thứ tự = thứ tự trong ô chọn).
_MODES = [
    ("Auto — OCR only scanned pages", MODE_AUTO),
    ("Text layer only (no AI, offline)", MODE_TEXT),
    ("Force OCR on every page", MODE_OCR),
]

# nhãn + màu badge theo nguồn text của trang
_SOURCE_BADGE = {
    SRC_TEXT:  ("Text layer", "--info"),
    SRC_OCR:   ("OCR", "--success"),
    SRC_SCAN:  ("Scanned – not read", "--warning"),
    SRC_EMPTY: ("Empty page", "--cat-default"),
}


def _segmented(parent, layout, options, on_select):
    """Dải nút chọn-một: nút đang bật tô đậm, các nút còn lại để mờ.

    Dùng thay cho kiểu một nút bấm-đổi-nhãn: nhìn nhãn không biết đó là thứ
    ĐANG xem hay thứ sẽ chuyển sang, bấm xong vẫn không rõ mình đang ở đâu.

    `options` là list (nhãn, icon, tooltip). Trả về (hàm chọn, danh sách nút).
    """
    buttons = []
    for index, (label, icon, tip) in enumerate(options):
        btn = widgets.button(parent, label, variant="neutral", icon=icon,
                             command=lambda i=index: on_select(i))
        btn.setToolTip(tip)
        layout.addWidget(btn)
        buttons.append((btn, icon))

    def select(active):
        for index, (btn, icon) in enumerate(buttons):
            on = index == active
            btn.setProperty("variant", "primary" if on else "neutral")
            # Icon tô theo màu chữ của nút nên phải vẽ lại khi đổi variant.
            btn.setIcon(widgets.svg_icon(icon, "#ffffff" if on else theme.TEXT, 16))
            btn.style().unpolish(btn)
            btn.style().polish(btn)

    return select, [btn for btn, _ in buttons]


class _ImageView(QScrollArea):
    """Vùng xem ảnh trang gốc — vẽ lại vừa bề rộng mỗi khi đổi kích thước."""

    def __init__(self, on_resize):
        super().__init__()
        self._on_resize = on_resize
        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.NoFrame)
        self.setStyleSheet("QScrollArea{border:none;background:#eef1f6;}")
        self._label = QLabel()
        self._label.setAlignment(Qt.AlignCenter)
        self.setWidget(self._label)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._on_resize()

    def show_pixmap(self, pm):
        if pm is None or pm.isNull():
            self._label.clear()
            return
        width = max(50, self.viewport().width() - 16)
        self._label.setPixmap(pm.scaledToWidth(min(width, pm.width()),
                                               Qt.SmoothTransformation))

    def clear(self):
        self._label.clear()


class PdfToTextTool(BaseTool):
    name = "PDF → Text"
    description = "Turn a PDF into text that keeps its layout, ready to copy."
    icon = "📝"
    category = "Files & Documents"
    order = 15
    fills_height = True

    def build(self, parent=None):
        self._pages = []
        self._cur = 0
        self._path = ""
        self._pixmaps = {}      # index trang → QPixmap đã render (nhớ để khỏi vẽ lại)
        self._whole_doc = True
        # Giữ cờ riêng thay vì hỏi _left.isVisible(): widget trả về False chừng
        # nào trang tool chưa được hiện, nên trạng thái đọc ra sẽ sai.
        self._show_original = False

        page = QWidget(parent)
        self._page = page
        outer = QVBoxLayout(page)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(12)

        if not _PYMUPDF_OK:
            card = widgets.Card()
            lay = QVBoxLayout(card)
            lay.setContentsMargins(20, 16, 20, 16)
            widgets.section_label(card, "Missing library")
            widgets.hint(card, "This feature needs PyMuPDF. Install it and reopen "
                               "the app:\n    pip install pymupdf")
            outer.addWidget(card)
            outer.addStretch(1)
            return page

        self._build_controls(outer)
        self._build_reader(outer)
        self._build_footer(outer)
        self._refresh()
        return page

    def build_body(self, parent):
        pass

    # ------------------------------------------------------------ điều khiển
    def _build_controls(self, outer):
        card = widgets.Card()
        lay = QVBoxLayout(card)
        lay.setContentsMargins(20, 14, 20, 16)
        lay.setSpacing(6)

        self.var_file = widgets.file_row(card, "PDF file", mode="file")

        opts = QHBoxLayout()
        opts.setSpacing(8)
        opts.addWidget(QLabel("Mode"))
        self.cb_mode = widgets.ComboBox()
        self.cb_mode.addItems([label for label, _ in _MODES])
        self.cb_mode.setMinimumWidth(240)
        opts.addWidget(self.cb_mode)
        opts.addSpacing(12)
        opts.addWidget(QLabel("Scan quality"))
        self.cb_dpi = widgets.ComboBox()
        self.cb_dpi.addItems(list(OCR_DPI_OPTIONS))
        self.cb_dpi.setCurrentText("Balanced (200 DPI)")
        opts.addWidget(self.cb_dpi)
        opts.addSpacing(12)
        self._run_btn = widgets.button(card, "Extract text", variant="primary",
                                       icon="play", command=self._start)
        opts.addWidget(self._run_btn)
        self._status = QLabel("")
        self._status.setObjectName("Hint")
        opts.addWidget(self._status)
        opts.addStretch(1)
        lay.addLayout(opts)

        widgets.hint(card, "Input: a PDF file, scanned ones included. Output: text "
                           "that keeps headings, lists and tables — copy it out or "
                           "save to a file.")
        outer.addWidget(card)

    # ------------------------------------------------------------ khung đọc
    def _build_reader(self, outer):
        self._split = QSplitter(Qt.Horizontal)

        # --- trái: ảnh trang gốc (ẩn mặc định) ---
        self._left = QFrame()
        self._left.setObjectName("Card")
        ll = QVBoxLayout(self._left)
        ll.setContentsMargins(12, 10, 12, 12)
        ll.setSpacing(8)
        head = QLabel("Original page")
        head.setObjectName("SectionLabel")
        ll.addWidget(head)
        self._imgview = _ImageView(self._draw_preview)
        ll.addWidget(self._imgview, 1)
        self._split.addWidget(self._left)
        self._left.hide()

        # --- phải: văn bản ---
        right = QFrame()
        right.setObjectName("Card")
        rl = QVBoxLayout(right)
        rl.setContentsMargins(12, 10, 12, 12)
        rl.setSpacing(8)

        bar = QHBoxLayout()
        bar.setSpacing(6)
        self._set_scope, self._scope_btns = _segmented(right, bar, [
            ("Whole document", "file-text", "Show every page as one document."),
            ("Page by page", "files", "Show one page at a time, so you can compare "
                                      "it with the original."),
        ], self._choose_scope)
        bar.addSpacing(10)
        self._prev_btn = widgets.button(right, "", variant="neutral",
                                        icon="chevron-left", command=lambda: self._go(-1))
        self._page_lbl = QLabel("– / –")
        self._page_lbl.setObjectName("Hint")
        self._next_btn = widgets.button(right, "", variant="neutral",
                                        icon="chevron-right", command=lambda: self._go(1))
        bar.addWidget(self._prev_btn)
        bar.addWidget(self._page_lbl)
        bar.addWidget(self._next_btn)
        bar.addSpacing(10)
        self._badge = QLabel("")
        bar.addWidget(self._badge)
        bar.addStretch(1)
        self._orig_btn = widgets.button(right, "Show original", variant="neutral",
                                        icon="files", command=self._toggle_original)
        bar.addWidget(self._orig_btn)
        rl.addLayout(bar)

        # Chỉ đọc: tool này để LẤY text ra, không phải để soạn thảo. Cho sửa mà
        # không có chỗ lưu lại thì bấm qua trang khác là mất, gây hiểu nhầm.
        self._text = widgets.TextEdit()
        self._text.setReadOnly(True)
        rl.addWidget(self._text, 1)
        self._split.addWidget(right)

        self._split.setSizes([420, 700])
        self._split.setStretchFactor(1, 1)
        # Thẻ tự chừa CARD_PAD cho mép; splitter không phải thẻ nên tự thêm lề
        # ngang để thẳng hàng với mép thẻ điều khiển ở trên.
        holder = QWidget()
        hl = QVBoxLayout(holder)
        hl.setContentsMargins(widgets.CARD_PAD, 0, widgets.CARD_PAD, 0)
        hl.addWidget(self._split)
        outer.addWidget(holder, 1)

    def _build_footer(self, outer):
        bar = QHBoxLayout()
        bar.setContentsMargins(widgets.CARD_PAD, 0, widgets.CARD_PAD, 0)
        bar.setSpacing(8)
        self._copy_btn = widgets.button(None, "Copy", variant="neutral",
                                        icon="copy", command=self._copy)
        self._md_btn = widgets.button(None, "Save as .md", variant="neutral",
                                      icon="save", command=lambda: self._save("md"))
        self._txt_btn = widgets.button(None, "Save as .txt", variant="success",
                                       icon="save", command=lambda: self._save("txt"))
        for b in (self._copy_btn, self._md_btn, self._txt_btn):
            bar.addWidget(b)
        bar.addStretch(1)
        outer.addLayout(bar)

    # ------------------------------------------------------------ trích xuất
    def _start(self):
        path = self.var_file.get().strip()
        if not path or not os.path.isfile(path):
            self.error("Error", "Please choose a valid PDF file.")
            return

        mode = _MODES[self.cb_mode.currentIndex()][1]
        dpi = OCR_DPI_OPTIONS.get(self.cb_dpi.currentText(), 200)
        cfg = settings.load()
        api_key = (cfg.get("api_key") or "").strip()
        model = (cfg.get("ai_model") or "").strip()

        if mode == MODE_OCR and not api_key:
            self.error("Gemini API key missing",
                       "Reading scanned pages needs a Gemini API key.\n"
                       "Add one in Settings, or switch the mode to "
                       '"Text layer only".')
            return

        try:
            total = pdf_text.page_count(path)
        except Exception as exc:
            self.error("Error", f"Couldn't open this PDF:\n{exc}")
            return

        def job(ctx):
            def progress(done, _total, source):
                ctx.step(1)
                ctx.status(f"Page {done}/{total} — "
                           f"{'read by AI' if source == SRC_OCR else 'text layer'}")

            try:
                return pdf_text.extract(
                    path, mode=mode, ocr_dpi=dpi, api_key=api_key, model=model,
                    progress=progress, should_cancel=lambda: ctx.cancelled)
            except pdf_text.Cancelled:
                return None     # Hủy là lựa chọn của người dùng, không phải lỗi

        def on_finish(dlg, pages):
            if pages is None:
                dlg.set_final_status("Cancelled.")
                return
            self._load(path, pages)
            dlg.set_final_status(self._summary(pages))

        subtitle = f"{os.path.basename(path)} — {total} pages"
        dlg = ProgressDialog(self._page, "Extracting text…", total=total,
                             subtitle=subtitle)
        dlg.start(job, on_finish)

    def _summary(self, pages):
        ocr = sum(1 for p in pages if p["source"] == SRC_OCR)
        scan = sum(1 for p in pages if p["source"] == SRC_SCAN)
        msg = f"Done — {len(pages)} pages."
        if ocr:
            msg += f" {ocr} read by AI."
        if scan:
            msg += (f" {scan} scanned pages skipped: add a Gemini API key in "
                    "Settings to read them.")
        return msg

    def _load(self, path, pages):
        self._path = path
        self._pages = pages
        self._cur = 0
        self._pixmaps.clear()
        self._status.setText(self._summary(pages))
        self._refresh()

    # ------------------------------------------------------------ hiển thị
    def _choose_scope(self, index):
        self._whole_doc = index == 0
        if self._whole_doc and self._show_original:
            self._set_original(False)
        self._refresh()

    def _toggle_original(self):
        self._set_original(not self._show_original)

    def _set_original(self, show):
        self._show_original = show
        self._left.setVisible(show)
        self._orig_btn.setText("Hide original" if show else "Show original")
        if show:
            # Ảnh trang chỉ có nghĩa khi đang xem từng trang.
            if self._whole_doc:
                self._choose_scope(1)
            self._draw_preview()

    def _go(self, delta):
        if not self._pages:
            return
        self._cur = max(0, min(len(self._pages) - 1, self._cur + delta))
        self._refresh()

    def _refresh(self):
        has = bool(self._pages)
        for btn in ([self._copy_btn, self._md_btn, self._txt_btn, self._orig_btn]
                    + self._scope_btns):
            btn.setEnabled(has)
        self._set_scope(0 if self._whole_doc else 1)
        per_page = has and not self._whole_doc
        for btn in (self._prev_btn, self._next_btn):
            btn.setEnabled(per_page)

        if not has:
            self._page_lbl.setText("– / –")
            self._badge.clear()
            self._text.clear()
            self._text.setReadOnly(True)
            self._imgview.clear()
            return

        if per_page:
            self._page_lbl.setText(f"{self._cur + 1} / {len(self._pages)}")
            self._prev_btn.setEnabled(self._cur > 0)
            self._next_btn.setEnabled(self._cur < len(self._pages) - 1)
            self._set_badge(self._pages[self._cur]["source"])
        else:
            self._page_lbl.setText(f"{len(self._pages)} pages")
            self._badge.clear()

        self._text.setMarkdown(self._current_markdown())
        if per_page and self._show_original:
            self._draw_preview()

    def _set_badge(self, source):
        label, token = _SOURCE_BADGE.get(source, ("", "--cat-default"))
        color = theme.PALETTE.get(token, theme.PALETTE["--cat-default"])
        r, g, b = widgets._hex_to_rgb(color)
        self._badge.setText(label)
        self._badge.setStyleSheet(
            f"background: rgba({r},{g},{b},0.15); color: {color};"
            "border-radius: 8px; padding: 3px 9px; font-size: 11px; font-weight: 600;")

    def _current_markdown(self):
        if not self._pages:
            return ""
        if self._whole_doc:
            return pdf_text.join_pages(self._pages, separators=False)
        return self._pages[self._cur]["markdown"]

    def _draw_preview(self):
        if not self._pages or self._whole_doc or not self._show_original:
            return
        pm = self._pixmaps.get(self._cur)
        if pm is None:
            try:
                png = pdf_text.render_png(self._path, self._cur)
            except Exception:
                png = None
            if not png:
                self._imgview.clear()
                return
            pm = QPixmap()
            pm.loadFromData(png, "PNG")
            self._pixmaps[self._cur] = pm
        self._imgview.show_pixmap(pm)

    # ------------------------------------------------------------ copy / lưu
    def _copy(self):
        if not self._pages:
            return
        # Copy THẲNG TỪ Ô ĐANG HIỂN THỊ: Qt tự đặt lên clipboard cả bản HTML lẫn
        # bản chữ trần, nên dán vào Word/Outlook giữ được tiêu đề - danh sách -
        # bảng. Không tự dựng QMimeData rồi setMimeData(): PySide6 và clipboard
        # cùng nhận quyền sở hữu đối tượng đó, lúc thoát app double-free làm
        # tiến trình chết ngang (kiểm chứng được bằng ~12 dòng Qt thuần).
        cursor = self._text.textCursor()
        self._text.selectAll()
        self._text.copy()
        self._text.setTextCursor(cursor)     # bỏ vệt bôi đen vừa tạo
        self._status.setText("Copied." if self._whole_doc
                             else f"Copied page {self._cur + 1}.")

    def _save(self, kind):
        if not self._pages:
            return
        base = os.path.splitext(os.path.basename(self._path))[0] or "result"
        filters = {"md": "Markdown (*.md)", "txt": "Text (*.txt)"}
        path, _ = QFileDialog.getSaveFileName(
            self._page, "Save text", f"{base}.{kind}", filters[kind])
        if not path:
            return
        markdown = pdf_text.join_pages(self._pages, separators=False)
        content = markdown if kind == "md" else pdf_text.to_plain_text(markdown)
        try:
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(content)
        except OSError as exc:
            self.error("Error", f"Couldn't save the file:\n{exc}")
            return
        self._status.setText(f"Saved: {os.path.basename(path)}")
