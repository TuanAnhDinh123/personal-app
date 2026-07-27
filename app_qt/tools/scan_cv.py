"""Quét CV: đổi tên hàng loạt + trích xuất Email/SĐT ra Excel — bản PySide6.

Toàn bộ logic (đọc PDF/DOCX, regex email/SĐT, tách tên/ID, ghi template Excel)
tách riêng ở app.core.cv_scan. Chỉ dựng lại giao diện Qt:
2 tab (Đổi tên / Trích xuất), bảng xem trước sửa được, trích xuất chạy luồng nền.
"""
import datetime
import os
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView, QHBoxLayout, QHeaderView, QLabel,
    QTableWidget, QTableWidgetItem, QTabWidget, QVBoxLayout, QWidget,
)

from app.core import config
from app.core.cv_scan import (
    DEFAULTS, SECTION, _CV_EXTENSIONS, _batch_from_folder, _build_filename,
    _extract_cv_text, _extract_name, _find_email, _find_phone,
    _open_existing_workbook, _open_template_workbook, _parse_noise,
    _safe_filename, _seq_code, _split_id_name, _write_candidates,
)
from app_qt import dialogs, widgets
from app_qt.base_tool import BaseTool
from app_qt.components.modal import ModalDialog
from app_qt.components.progress_dialog import ProgressDialog

try:
    import openpyxl  # noqa: F401
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False


def _card():
    card = widgets.Card()
    lay = QVBoxLayout(card)
    lay.setContentsMargins(24, 20, 24, 20)
    lay.setSpacing(6)
    return card, lay


class ScanCvTool(BaseTool):
    name = "CV Scan"
    description = "Batch-rename CV files and extract emails & phone numbers to Excel."
    icon = "📇"
    category = "Files & Documents"
    order = 10

    def build(self, parent=None):
        cfg = config.load(SECTION, DEFAULTS)
        page = QWidget(parent)
        outer = QVBoxLayout(page)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(14)
        self._page = page

        # ---- Card dùng chung: thư mục + từ nhiễu + lưu cấu hình ----
        shared, sl = _card()
        widgets.section_label(shared, "CV folder")
        self.var_folder = widgets.file_row(shared, "Folder", mode="folder")
        self.var_folder.set(cfg["folder"])
        widgets.section_label(shared, "Words to strip from candidate names")
        self.noise_box = widgets.text_area(
            shared, "One word/phrase per line (or comma-separated):",
            value=cfg["noise_keywords"], height=5)
        save_row = QHBoxLayout()
        save_row.addWidget(widgets.button(shared, "Save config", variant="neutral",
                                          icon="save", command=self._save_config))
        save_row.addStretch(1)
        sl.addLayout(save_row)
        outer.addWidget(shared)

        # ---- Tabs ----
        # Thẻ "shared" tự chừa CARD_PAD cho bóng → tab (không phải thẻ) bọc trong
        # container thêm lề ngang CARD_PAD để thẳng hàng mép thẻ nhìn thấy.
        tabs = QTabWidget()
        tabs.addTab(self._build_rename_tab(cfg), "Rename files")
        tabs.addTab(self._build_extract_tab(), "Export to Excel")
        tabs_holder = QWidget()
        thl = QVBoxLayout(tabs_holder)
        thl.setContentsMargins(widgets.CARD_PAD, 0, widgets.CARD_PAD, 0)
        thl.addWidget(tabs)
        outer.addWidget(tabs_holder)
        outer.addStretch(1)
        return page

    def build_body(self, parent):
        pass

    # ------------------------------------------------------------- tab đổi tên
    def _build_rename_tab(self, cfg):
        tab = QWidget()
        lay = QVBoxLayout(tab)
        lay.setContentsMargins(20, 18, 20, 18)
        lay.setSpacing(6)
        widgets.section_label(tab, "CV code")

        row = QWidget(tab)
        h = QHBoxLayout(row)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(12)
        left, right = QWidget(), QWidget()
        h.addWidget(left, 1)
        h.addWidget(right, 1)
        self.var_prefix = widgets.digit_entry(left, "Prefix code (4 digits)", "2506")
        self.var_prefix.set(cfg["prefix"])
        self.var_start = widgets.digit_entry(right, "Start code (2 digits, incrementing)", "01")
        self.var_start.set(cfg["start"])
        lay.addWidget(row)

        widgets.hint(tab, "Example: prefix=2506, start=01 → 250601_Nguyen Van A.pdf, "
                          "250602_Tran Thi B.pdf, …")
        act = QHBoxLayout()
        act.setContentsMargins(0, 14, 0, 0)
        act.addWidget(widgets.button(tab, "Rename CV files", variant="success",
                                     icon="pencil", command=self.run))
        act.addStretch(1)
        lay.addLayout(act)
        lay.addStretch(1)
        return tab

    # ------------------------------------------------------------ tab trích xuất
    def _build_extract_tab(self):
        tab = QWidget()
        lay = QVBoxLayout(tab)
        lay.setContentsMargins(20, 18, 20, 18)
        lay.setSpacing(6)
        widgets.section_label(tab, "Export data to Excel")
        self.var_dept = widgets.text_row(tab, "Department (fills the APPLYING FOR column)")
        self.var_output = widgets.export_target_row(tab, "Export summary to")
        widgets.hint(tab, "• Pick 📁 a folder → create a new Excel from the template.\n"
                          "• Pick 📄 an Excel file → append to the 'Candidates' sheet.")
        widgets.section_label(tab, "Fields to extract")
        self.chk_name = widgets.checkbox(tab, "Candidate name (without ID)")
        self.chk_id = widgets.checkbox(tab, "ID (prefix + start code from filename)")
        self.chk_email = widgets.checkbox(tab, "Email")
        self.chk_phone = widgets.checkbox(tab, "Phone")
        widgets.hint(tab, "Name & ID come from the filename (rename them in the previous tab "
                          "first). Email & phone are read from the CV content (PDF/DOCX). "
                          "Legacy .doc isn't supported.")
        act = QHBoxLayout()
        act.setContentsMargins(0, 14, 0, 0)
        act.addWidget(widgets.button(tab, "Export to Excel", variant="primary",
                                     icon="file", command=self._run_extract))
        act.addStretch(1)
        lay.addLayout(act)
        lay.addStretch(1)
        return tab

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
        self.info("Saved", "Config saved ✅")

    def _files(self):
        folder = self.var_folder.get().strip()
        if not folder or not os.path.isdir(folder):
            self.error("Missing folder", "Please choose the CV folder.")
            return None, None
        files = sorted(p for p in Path(folder).iterdir()
                       if p.is_file() and p.suffix.lower() in _CV_EXTENSIONS)
        if not files:
            self.info("No files", "No PDF/DOC/DOCX files found in the folder.")
            return folder, None
        return folder, files

    # -------------------------------------------------------------- đổi tên
    def run(self):
        folder, files = self._files()
        if not files:
            return
        prefix = self.var_prefix.get().strip()
        start_str = self.var_start.get().strip() or "01"
        noise = _parse_noise(self.noise_box.get())
        config.save(SECTION, self._collect())
        _RenamePreview(self._page, files, prefix, start_str, noise).exec()

    # -------------------------------------------------------------- trích xuất
    def _run_extract(self):
        if not _OPENPYXL_OK:
            self.error("Missing library", "openpyxl is required to export Excel:\n  pip install openpyxl")
            return
        folder, files = self._files()
        if not files:
            return
        target = self.var_output.get().strip()
        if not target:
            self.error("Missing target",
                       "Pick a folder (new file) or an Excel file (append).")
            return
        dept = self.var_dept.get().strip()
        want = dict(name=self.chk_name.get(), id=self.chk_id.get(),
                    email=self.chk_email.get(), phone=self.chk_phone.get())
        if not any(want.values()):
            self.error("No fields selected", "Choose at least one field to extract.")
            return
        config.save(SECTION, self._collect())

        noise = _parse_noise(self.noise_box.get())
        batch = _batch_from_folder(folder)
        need_text = want["email"] or want["phone"]
        total = len(files)

        def job(ctx):
            rows, errors = [], []
            for i, p in enumerate(files, start=1):
                if ctx.cancelled:
                    break
                ctx.status(f"({i}/{total}) {p.name}")
                text = ""
                if need_text:
                    try:
                        text = _extract_cv_text(p)
                    except Exception as exc:
                        errors.append(f"{p.name}: {exc}")
                        ctx.log(f"⚠ {p.name}: {exc}")
                cv_id, cv_name = _split_id_name(p.stem, noise)
                row = {}
                if batch is not None:
                    row["batch"] = batch
                if want["name"]:
                    row["name"] = cv_name
                if want["id"]:
                    row["id"] = cv_id
                if dept:
                    row["apply"] = dept
                if want["email"]:
                    row["email"] = _find_email(text)
                if want["phone"]:
                    row["phone"] = _find_phone(text)
                rows.append(row)
                ctx.log(f"✓ {p.name}")
                ctx.step()
            return rows, errors

        def on_finish(dlg, result):
            rows, errors = result
            try:
                if os.path.isdir(target):
                    wb, ws = _open_template_workbook()
                    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    base = _safe_filename(f"Candidates_{dept}" if dept else "Candidates")
                    out = os.path.join(target, f"{base}_{stamp}.xlsx")
                    mode = "new"
                elif os.path.isfile(target):
                    wb, ws = _open_existing_workbook(target)
                    out, mode = target, "appended"
                elif target.lower().endswith(".xlsx") and os.path.isdir(os.path.dirname(target)):
                    wb, ws = _open_template_workbook()
                    out, mode = target, "new"
                else:
                    dlg.set_final_status("Invalid target.")
                    dlg.log("⚠ Pick a valid FOLDER or .xlsx FILE.")
                    return
                _write_candidates(ws, rows)
                wb.save(out)
            except PermissionError:
                dlg.set_final_status("Can't write file.")
                dlg.log(f"⚠ Open in Excel? Close it and retry:\n{target}")
                return
            except Exception as exc:
                dlg.set_final_status("Excel write error.")
                dlg.log(f"⚠ {exc}")
                return
            dlg.set_final_status(f"Done — {len(rows)} CVs ({mode}).")
            dlg.log(f"\n✅ Saved {len(rows)} CVs to:\n{out}")
            if errors:
                dlg.log(f"⚠ {len(errors)} files couldn't be read (see above).")

        dlg = ProgressDialog(self._page, "Extracting CVs…", total=total,
                             subtitle=f"Reading {total} CV files")
        dlg.start(job, on_finish)


class _RenamePreview(ModalDialog):
    """Bảng xem trước đổi tên — sửa được cột 'Tên ứng viên', tên mới tự cập nhật."""

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
