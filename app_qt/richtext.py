"""Ô soạn thảo có định dạng (rich text) → HTML — bản PySide6.

Gọn hơn hẳn bản Tk: QTextEdit hỗ trợ sẵn in đậm/nghiêng/gạch chân/màu/gạch đầu
dòng và xuất/nhập HTML (toHtml/setHtml). Giữ nguyên API get_html/get_text/
set_html để các tool (reminder…) dùng như cũ.
"""
import re

from PySide6.QtGui import QColor, QTextCharFormat, QTextListFormat
from PySide6.QtWidgets import (
    QColorDialog, QHBoxLayout, QPushButton, QTextEdit, QVBoxLayout, QWidget,
)


def _looks_like_html(s):
    return bool(s) and re.search(r"<[a-zA-Z/!]", s) is not None


class RichText(QWidget):
    """QTextEdit + thanh công cụ định dạng; đọc/ghi bằng HTML."""

    def __init__(self, parent=None, height=12):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)

        bar = QHBoxLayout()
        bar.setSpacing(4)
        bar.addWidget(self._tb("B", self._bold, bold=True))
        bar.addWidget(self._tb("I", self._italic, italic=True))
        bar.addWidget(self._tb("U", self._underline, underline=True))
        bar.addWidget(self._tb("A🎨", self._color))
        bar.addWidget(self._tb("• List", self._bullet))
        bar.addWidget(self._tb("✕ Định dạng", self._clear))
        bar.addStretch(1)
        lay.addLayout(bar)

        self.text = QTextEdit(self)
        self.text.setMinimumHeight(max(80, height * 20))
        lay.addWidget(self.text)

    def _tb(self, label, cmd, bold=False, italic=False, underline=False):
        b = QPushButton(label)
        b.setProperty("variant", "neutral")
        b.clicked.connect(lambda: cmd())
        f = b.font()
        f.setBold(bold)
        f.setItalic(italic)
        f.setUnderline(underline)
        b.setFont(f)
        b.setFixedHeight(30)
        return b

    # ---------------------------------------------------------- định dạng
    def _bold(self):
        w = self.text.fontWeight()
        from PySide6.QtGui import QFont
        self.text.setFontWeight(QFont.Normal if w > QFont.Normal else QFont.Bold)
        self.text.setFocus()

    def _italic(self):
        self.text.setFontItalic(not self.text.fontItalic())
        self.text.setFocus()

    def _underline(self):
        self.text.setFontUnderline(not self.text.fontUnderline())
        self.text.setFocus()

    def _color(self):
        col = QColorDialog.getColor(QColor("#1f2735"), self, "Chọn màu chữ")
        if col.isValid():
            self.text.setTextColor(col)
        self.text.setFocus()

    def _bullet(self):
        cursor = self.text.textCursor()
        cursor.createList(QTextListFormat.ListDisc)
        self.text.setFocus()

    def _clear(self):
        cursor = self.text.textCursor()
        cursor.setCharFormat(QTextCharFormat())
        self.text.setFocus()

    # ---------------------------------------------------------- đọc / ghi
    def get_html(self):
        return self.text.toHtml()

    def get_text(self):
        return self.text.toPlainText()

    def set_html(self, html):
        if _looks_like_html(html):
            self.text.setHtml(html)
        else:
            self.text.setPlainText(html or "")
