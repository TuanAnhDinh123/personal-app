"""PDF → Text: trích văn bản từ PDF (Word/scan) + đối chiếu — bản PySide6.

Logic trích (extract_pages, OCR Tesseract) dùng lại nguyên từ module Tk cũ.
Giao diện Qt: thanh điều khiển + splitter (ảnh trang gốc | văn bản sửa được),
trích chạy luồng nền (Task) cập nhật trạng thái tại chỗ.
"""
import os

from PySide6.QtCore import Qt
from PySide6.QtGui import QImage, QPixmap
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QFileDialog, QFrame, QHBoxLayout,
    QLabel, QLineEdit, QScrollArea, QSplitter, QTextEdit, QVBoxLayout, QWidget,
)

from app.core.pdf_text import (
    _OCR_DPI_OPTIONS, extract_pages, tesseract_available,
)
from app_qt import dialogs, theme, widgets
from app_qt.base_tool import BaseTool
from app_qt.components.task import Task

try:
    import fitz  # PyMuPDF
    _FITZ_OK = True
except ImportError:
    _FITZ_OK = False

# nhãn + màu badge theo nguồn text của trang
_SOURCE_BADGE = {
    "text": ("Text layer", theme.PALETTE["--info"]),
    "ocr":  ("OCR", theme.PALETTE["--success"]),
    "scan": ("Scan – no OCR", theme.PALETTE["--warning"]),
}


class _ImageView(QScrollArea):
    """Vùng xem ảnh trang PDF — gọi lại on_resize để render vừa bề rộng."""

    def __init__(self, on_resize):
        super().__init__()
        self._on_resize = on_resize
        self.setWidgetResizable(True)
        self.setStyleSheet("QScrollArea{border:none;background:#eef1f6;}")
        self._label = QLabel()
        self._label.setAlignment(Qt.AlignCenter)
        self.setWidget(self._label)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._on_resize()

    def set_pixmap(self, pm):
        self._label.setPixmap(pm)

    def clear(self):
        self._label.clear()


class PdfToTextTool(BaseTool):
    name = "PDF → Text"
    description = "Extract text from PDFs (Word-exported or scanned) and compare with the original."
    icon = "📝"
    category = "Files & Documents"
    order = 15
    fills_height = True

    def build(self, parent=None):
        self._pages = []
        self._cur = 0
        self._doc = None
        self._task = None
        self._busy = False

        page = QWidget(parent)
        self._page = page
        outer = QVBoxLayout(page)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(12)

        if not _FITZ_OK:
            card, lay = self._card()
            widgets.section_label(card, "Missing library")
            widgets.hint(card, "This feature needs PyMuPDF. Install it and reopen the app:\n"
                               "    pip install pymupdf pytesseract")
            outer.addWidget(card)
            outer.addStretch(1)
            return page

        self._build_controls(outer)
        self._build_viewer(outer)
        self._build_footer(outer)
        self._refresh_view()
        return page

    def build_body(self, parent):
        pass

    def _card(self):
        card = widgets.Card()
        lay = QVBoxLayout(card)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(6)
        return card, lay

    # ---------------------------------------------------------- điều khiển
    def _build_controls(self, outer):
        card, lay = self._card()
        self.var_file = widgets.file_row(card, "PDF file", mode="file")

        opts = QHBoxLayout()
        opts.setSpacing(8)
        opts.addWidget(QLabel("OCR language"))
        self.ent_lang = QLineEdit("vie+eng")
        self.ent_lang.setFixedWidth(110)
        opts.addWidget(self.ent_lang)
        opts.addSpacing(12)
        opts.addWidget(QLabel("Scan quality"))
        self.cb_dpi = widgets.ComboBox()
        self.cb_dpi.addItems(list(_OCR_DPI_OPTIONS))
        self.cb_dpi.setCurrentText("Balanced (300 DPI)")
        opts.addWidget(self.cb_dpi)
        opts.addSpacing(12)
        self.chk_force = QCheckBox("Always OCR (even if a text layer exists)")
        opts.addWidget(self.chk_force)
        opts.addStretch(1)
        lay.addLayout(opts)

        bar = QHBoxLayout()
        self._run_btn = widgets.button(card, "Extract text", variant="primary",
                                       icon="play", command=self._start_extract)
        bar.addWidget(self._run_btn)
        self._status = QLabel("")
        self._status.setObjectName("Hint")
        bar.addWidget(self._status)
        bar.addStretch(1)
        lay.addLayout(bar)

        if not tesseract_available():
            widgets.hint(card, "💡 Word-exported PDFs extract instantly. Scanned PDFs need "
                               "Tesseract-OCR (github.com/UB-Mannheim/tesseract/wiki — "
                               "include the Vietnamese pack).")
        outer.addWidget(card)

    # ---------------------------------------------------------- viewer
    def _build_viewer(self, outer):
        split = QSplitter(Qt.Horizontal)

        # trái: ảnh gốc + nav
        left = QFrame()
        left.setObjectName("Card")
        ll = QVBoxLayout(left)
        ll.setContentsMargins(12, 10, 12, 12)
        ll.setSpacing(8)
        nav = QHBoxLayout()
        title = QLabel("Original")
        title.setObjectName("SectionLabel")
        nav.addWidget(title)
        nav.addStretch(1)
        self._badge = QLabel("")
        nav.addWidget(self._badge)
        self._prev_btn = widgets.button(left, "", variant="neutral", command=lambda: self._go(-1))
        self._prev_btn.setText("‹")
        self._page_lbl = QLabel("– / –")
        self._page_lbl.setObjectName("Hint")
        self._next_btn = widgets.button(left, "", variant="neutral", command=lambda: self._go(1))
        self._next_btn.setText("›")
        nav.addWidget(self._prev_btn)
        nav.addWidget(self._page_lbl)
        nav.addWidget(self._next_btn)
        ll.addLayout(nav)
        self._imgview = _ImageView(self._render_page)
        ll.addWidget(self._imgview, 1)
        split.addWidget(left)

        # phải: văn bản trích được
        right = QFrame()
        right.setObjectName("Card")
        rl = QVBoxLayout(right)
        rl.setContentsMargins(12, 10, 12, 12)
        rl.setSpacing(8)
        rtitle = QLabel("Extracted text (editable)")
        rtitle.setObjectName("SectionLabel")
        rl.addWidget(rtitle)
        self._text = widgets.TextEdit()
        self._text.setAcceptRichText(False)
        rl.addWidget(self._text, 1)
        split.addWidget(right)

        split.setSizes([500, 500])
        # Thẻ điều khiển tự chừa CARD_PAD cho bóng → splitter (không phải thẻ) bọc
        # thêm lề ngang CARD_PAD để thẳng hàng mép thẻ nhìn thấy.
        split_holder = QWidget()
        shl = QVBoxLayout(split_holder)
        shl.setContentsMargins(widgets.CARD_PAD, 0, widgets.CARD_PAD, 0)
        shl.addWidget(split)
        outer.addWidget(split_holder, 1)

    def _build_footer(self, outer):
        bar = QHBoxLayout()
        bar.setContentsMargins(widgets.CARD_PAD, 0, widgets.CARD_PAD, 0)
        self._copy_btn = widgets.button(None, "Copy this page", variant="neutral",
                                        icon="copy", command=self._copy_current)
        self._save_btn = widgets.button(None, "Save all as .txt", variant="success",
                                        icon="save", command=self._save_txt)
        bar.addWidget(self._copy_btn)
        bar.addWidget(self._save_btn)
        bar.addStretch(1)
        outer.addLayout(bar)

    # ---------------------------------------------------------- trích (nền)
    def _start_extract(self):
        if self._busy:
            return
        path = self.var_file.get().strip()
        if not path or not os.path.isfile(path):
            self.error("Error", "Please choose a valid PDF file.")
            return
        if self.chk_force.isChecked() and not tesseract_available():
            self.error("Tesseract missing",
                       'You enabled "Always OCR" but Tesseract-OCR isn\'t installed.\n'
                       "Install it or turn off this option.")
            return

        self._busy = True
        self._run_btn.setEnabled(False)
        self._status.setText("Opening PDF…")
        lang = self.ent_lang.text().strip() or "vie+eng"
        dpi = _OCR_DPI_OPTIONS.get(self.cb_dpi.currentText(), 300)
        force = self.chk_force.isChecked()

        def work(emit):
            return extract_pages(path, ocr_lang=lang, ocr_dpi=dpi, force_ocr=force,
                                 progress=lambda d, t, s: emit((d, t)))

        self._task = Task(work, self._page)
        self._task.signals.progress.connect(
            lambda dt: self._status.setText(f"Processing page {dt[0]}/{dt[1]}…"))
        self._task.signals.finished.connect(lambda pages: self._on_done(path, pages))
        self._task.signals.failed.connect(self._on_error)
        self._task.start()

    def _on_error(self, msg):
        self._busy = False
        self._run_btn.setEnabled(True)
        self._status.setText("")
        self.error("Error", f"Couldn't extract text:\n{msg}")

    def _on_done(self, path, pages):
        self._busy = False
        self._run_btn.setEnabled(True)
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
        msg = f"✅ Done, {len(pages)} pages."
        if n_ocr:
            msg += f" OCR: {n_ocr} pages."
        if n_scan:
            msg += f" ⚠ {n_scan} scanned pages not OCR'd (Tesseract missing)."
        self._status.setText(msg)
        self._refresh_view()

    # ---------------------------------------------------------- hiển thị
    def _go(self, delta):
        if not self._pages:
            return
        self._commit_text()
        self._cur = max(0, min(len(self._pages) - 1, self._cur + delta))
        self._refresh_view()

    def _refresh_view(self):
        has = bool(self._pages)
        for b in (self._prev_btn, self._next_btn, self._copy_btn, self._save_btn):
            b.setEnabled(has)
        if not has:
            self._page_lbl.setText("– / –")
            self._badge.setText("")
            self._text.clear()
            self._text.setReadOnly(True)
            self._imgview.clear()
            return
        page = self._pages[self._cur]
        self._page_lbl.setText(f"{self._cur + 1} / {len(self._pages)}")
        label, color = _SOURCE_BADGE.get(page["source"], ("", theme.PALETTE["--cat-default"]))
        r, g, b = widgets._hex_to_rgb(color)
        self._badge.setText(label)
        self._badge.setStyleSheet(
            f"background: rgba({r},{g},{b},0.15); color: {color};"
            "border-radius: 8px; padding: 3px 9px; font-size: 11px; font-weight: 600;")
        self._prev_btn.setEnabled(self._cur > 0)
        self._next_btn.setEnabled(self._cur < len(self._pages) - 1)
        self._text.setReadOnly(False)
        self._text.setPlainText(page["text"])
        self._render_page()

    def _render_page(self):
        if not self._pages or self._doc is None:
            return
        vw = self._imgview.viewport().width() - 16
        if vw <= 1:
            return
        page = self._doc[self._cur]
        rect = page.rect
        zoom = max(0.1, min(3.0, vw / rect.width))
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        fmt = QImage.Format_RGBA8888 if pix.alpha else QImage.Format_RGB888
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt).copy()
        self._imgview.set_pixmap(QPixmap.fromImage(img))

    def _commit_text(self):
        if self._pages and 0 <= self._cur < len(self._pages):
            self._pages[self._cur]["text"] = self._text.toPlainText()

    # ---------------------------------------------------------- sao chép / lưu
    def _copy_current(self):
        if not self._pages:
            return
        self._commit_text()
        QApplication.clipboard().setText(self._pages[self._cur]["text"])
        self._status.setText(f"📋 Copied page {self._cur + 1}.")

    def _save_txt(self):
        if not self._pages:
            return
        self._commit_text()
        src = self.var_file.get().strip()
        default = (os.path.splitext(os.path.basename(src))[0] + ".txt") if src else "result.txt"
        path, _ = QFileDialog.getSaveFileName(self._page, "Save text", default,
                                              "Text (*.txt)")
        if not path:
            return
        parts = [f"--- Page {i + 1} ---\n{p['text']}" for i, p in enumerate(self._pages)]
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n\n".join(parts))
        except Exception as exc:
            self.error("Error", f"Couldn't save the file:\n{exc}")
            return
        self._status.setText(f"💾 Saved: {os.path.basename(path)}")
