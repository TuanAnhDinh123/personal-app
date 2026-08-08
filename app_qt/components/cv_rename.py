"""Chuẩn hóa tên file CV — modal cấu hình mã CV + bảng xem trước đổi tên.

Dùng ở tool "AI CV Scan" (nút *Normalize file names*): mở `RenameConfigDialog`
để nhập thư mục + mã CV + từ nhiễu, bấm xem trước thì hiện bảng sửa được rồi đổi
tên hàng loạt. Toàn bộ logic tách tên / ghép tên file nằm ở `app.core.cv_scan`,
file này chỉ là giao diện.
"""
import os
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView, QHBoxLayout, QHeaderView, QLabel, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QWidget,
)

from app.core import config
from app.core.cv_scan import (
    DEFAULTS, SECTION, _CV_EXTENSIONS, _build_filename, _extract_name,
    _parse_noise, _seq_code,
)
from app_qt import dialogs, widgets
from app_qt.components.modal import ModalDialog


def cv_files(folder: str):
    """Danh sách file CV (PDF/DOC/DOCX) trong thư mục, sắp theo tên."""
    return sorted(p for p in Path(folder).iterdir()
                  if p.is_file() and p.suffix.lower() in _CV_EXTENSIONS)


class RenameConfigDialog(ModalDialog):
    """Cấu hình đổi tên CV; cấu hình lưu ở section `scan_cv` dùng chung."""

    def __init__(self, parent, default_folder=""):
        super().__init__(parent, "sm")
        cfg = config.load(SECTION, DEFAULTS)
        card, lay = self.build_shell("Normalize CV file names")

        form = QWidget()
        col = QVBoxLayout(form)
        col.setContentsMargins(0, 0, 8, 0)
        col.setSpacing(4)

        self.var_folder = widgets.file_row(form, "CV folder", mode="folder")
        # Ưu tiên thư mục đang chọn ở màn hình gọi tới; chưa có thì lấy cấu hình cũ.
        self.var_folder.set(default_folder or cfg["folder"])

        widgets.section_label(form, "CV code")
        row = QWidget(form)
        h = QHBoxLayout(row)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(12)
        left, right = QWidget(row), QWidget(row)
        h.addWidget(left, 1)
        h.addWidget(right, 1)
        self.var_prefix = widgets.digit_entry(left, "Prefix code (4 digits)", "2506")
        self.var_prefix.set(cfg["prefix"])
        self.var_start = widgets.digit_entry(right, "Start code (2 digits, incrementing)", "01")
        self.var_start.set(cfg["start"])
        col.addWidget(row)
        widgets.hint(form, "Example: prefix=2506, start=01 → 250601_Nguyen Van A.pdf, "
                           "250602_Tran Thi B.pdf, …")

        widgets.section_label(form, "Words to strip from candidate names")
        self.noise_box = widgets.text_area(
            form, "One word/phrase per line (or comma-separated):",
            value=cfg["noise_keywords"], height=10)
        col.addStretch(1)

        sa = widgets.scroll_area(form)
        lay.addWidget(sa, 1)
        self.set_grow_region(sa)

        foot = QHBoxLayout()
        foot.setContentsMargins(0, 6, 0, 0)
        foot.addWidget(widgets.button(card, "Preview & rename", variant="success",
                                      icon="pencil", command=self._preview))
        foot.addWidget(widgets.button(card, "Save config", variant="neutral",
                                      icon="save", command=self._save_config))
        foot.addWidget(widgets.button(card, "Cancel", variant="neutral", icon="x",
                                      command=self.reject))
        foot.addStretch(1)
        lay.addLayout(foot)

    # ------------------------------------------------------------------ config
    def _collect(self):
        return {
            "folder": self.var_folder.get().strip(),
            "prefix": self.var_prefix.get().strip(),
            "start": self.var_start.get().strip(),
            "noise_keywords": self.noise_box.get(),
        }

    def _save_config(self):
        config.save(SECTION, self._collect())
        dialogs.success(self, "Saved", "Config saved ✅")

    # ----------------------------------------------------------- xem trước
    def _preview(self):
        folder = self.var_folder.get().strip()
        if not folder or not os.path.isdir(folder):
            dialogs.error(self, "Missing folder", "Please choose the CV folder.")
            return
        files = cv_files(folder)
        if not files:
            dialogs.info(self, "No files", "No PDF/DOC/DOCX files found in the folder.")
            return
        config.save(SECTION, self._collect())
        noise = _parse_noise(self.noise_box.get())
        dlg = RenamePreviewDialog(self, files, self.var_prefix.get().strip(),
                                  self.var_start.get().strip() or "01", noise)
        if dlg.exec():
            self.accept()   # đổi tên xong thì đóng luôn hộp cấu hình


class RenamePreviewDialog(ModalDialog):
    """Bảng xem trước đổi tên — sửa được cột 'Candidate name', tên mới tự cập nhật."""

    def __init__(self, parent, files, prefix, start_str, noise):
        super().__init__(parent, "lg")
        self._files = files
        self._prefix = prefix
        self._suffixes = []
        self._codes = []

        card, lay = self.build_shell(
            f"Found {len(files)} files — double-click the middle column to edit the name",
            spacing=10)

        sub = QLabel("Candidate names are extracted from the original filenames; "
                     "the 'New filename' column updates live.")
        sub.setObjectName("DialogMsg")
        lay.addWidget(sub)

        self.table = QTableWidget(len(files), 3)
        self.table.setHorizontalHeaderLabels(
            ["Original filename", "Candidate name (editable)", "New filename (preview)"])
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(32)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.Interactive)
        hdr.setSectionResizeMode(1, QHeaderView.Interactive)
        hdr.setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.setColumnWidth(0, 260)
        self.table.setColumnWidth(1, 220)

        for idx, p in enumerate(files):
            code = _seq_code(start_str, idx)
            name = _extract_name(p.stem, noise)
            self._codes.append(code)
            self._suffixes.append(p.suffix)
            c0 = QTableWidgetItem(p.name)
            c0.setFlags(c0.flags() & ~Qt.ItemIsEditable)
            c1 = QTableWidgetItem(name)
            c2 = QTableWidgetItem(_build_filename(name, prefix, code, p.suffix))
            c2.setFlags(c2.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(idx, 0, c0)
            self.table.setItem(idx, 1, c1)
            self.table.setItem(idx, 2, c2)

        self.table.itemChanged.connect(self._on_edit)
        lay.addWidget(self.table, 1)
        self.set_grow_region(self.table)   # cao theo cỡ lg, tự co khi màn hình thấp

        foot = QHBoxLayout()
        foot.addWidget(widgets.button(card, "Rename all", variant="success",
                                      icon="check", command=self._do_rename))
        foot.addWidget(widgets.button(card, "Cancel", variant="neutral", icon="x",
                                      command=self.reject))
        foot.addStretch(1)
        lay.addLayout(foot)

    def _on_edit(self, item):
        if item.column() != 1:
            return
        r = item.row()
        name = item.text().strip()
        new_file = _build_filename(name, self._prefix, self._codes[r], self._suffixes[r])
        self.table.blockSignals(True)
        self.table.item(r, 2).setText(new_file)
        self.table.blockSignals(False)

    def _do_rename(self):
        renamed = skipped = 0
        errors = []
        for r, p in enumerate(self._files):
            cname = self.table.item(r, 1).text().strip()
            if not cname:
                skipped += 1
                continue
            new_name = _build_filename(cname, self._prefix, self._codes[r], p.suffix)
            new_path = p.parent / new_name
            if new_path == p:
                skipped += 1
                continue
            try:
                p.rename(new_path)
                renamed += 1
            except Exception as exc:
                errors.append(f"{p.name}: {exc}")
        msg = f"Renamed {renamed} files."
        if skipped:
            msg += f"\n{skipped} files skipped (unchanged name)."
        if errors:
            msg += "\n\nErrors:\n" + "\n".join(errors[:8])
        self.accept()
        if errors:
            dialogs.warning(self.parent(), "Done (with errors)", msg)
        else:
            dialogs.success(self.parent(), "Done", msg)
