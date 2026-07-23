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
    QAbstractItemView, QDialog, QFrame, QHBoxLayout, QHeaderView, QLabel,
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
from app_qt.components.progress_dialog import ProgressDialog

try:
    import openpyxl  # noqa: F401
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False


def _card():
    card = QFrame()
    card.setObjectName("Card")
    widgets.add_shadow(card)
    lay = QVBoxLayout(card)
    lay.setContentsMargins(24, 20, 24, 20)
    lay.setSpacing(6)
    return card, lay


class ScanCvTool(BaseTool):
    name = "Quét CV"
    description = "Đổi tên hàng loạt file CV và trích xuất Email, Số điện thoại ra Excel."
    icon = "📇"
    category = "Tệp & Tài liệu"
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
        widgets.section_label(shared, "Thư mục chứa CV")
        self.var_folder = widgets.file_row(shared, "Thư mục", mode="folder")
        self.var_folder.set(cfg["folder"])
        widgets.section_label(shared, "Từ cần xóa khi trích tên ứng viên")
        self.noise_box = widgets.text_area(
            shared, "Mỗi từ / cụm từ một dòng (hoặc cách nhau bởi dấu phẩy):",
            value=cfg["noise_keywords"], height=5)
        save_row = QHBoxLayout()
        save_row.addWidget(widgets.button(shared, "Lưu cấu hình", variant="neutral",
                                          icon="save", command=self._save_config))
        save_row.addStretch(1)
        sl.addLayout(save_row)
        outer.addWidget(shared)

        # ---- Tabs ----
        tabs = QTabWidget()
        tabs.addTab(self._build_rename_tab(cfg), "Đổi tên file")
        tabs.addTab(self._build_extract_tab(), "Trích xuất Excel")
        outer.addWidget(tabs)
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
        widgets.section_label(tab, "Mã CV")

        row = QWidget(tab)
        h = QHBoxLayout(row)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(12)
        left, right = QWidget(), QWidget()
        h.addWidget(left, 1)
        h.addWidget(right, 1)
        self.var_prefix = widgets.digit_entry(left, "Prefix code (4 số)", "2506")
        self.var_prefix.set(cfg["prefix"])
        self.var_start = widgets.digit_entry(right, "Start code (2 số, tăng dần)", "01")
        self.var_start.set(cfg["start"])
        lay.addWidget(row)

        widgets.hint(tab, "Ví dụ: prefix=2506, start=01 → 250601_Nguyen Van A.pdf, "
                          "250602_Tran Thi B.pdf, …")
        act = QHBoxLayout()
        act.setContentsMargins(0, 14, 0, 0)
        act.addWidget(widgets.button(tab, "Đổi tên file CV", variant="success",
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
        widgets.section_label(tab, "Trích xuất dữ liệu ra Excel")
        self.var_dept = widgets.text_row(tab, "Phòng ban (điền vào cột APPLYING FOR)")
        self.var_output = widgets.export_target_row(tab, "Xuất bảng tổng hợp ra")
        widgets.hint(tab, "• Chọn 📁 Thư mục → tạo file Excel mới theo template.\n"
                          "• Chọn 📄 File Excel → nối tiếp vào sheet 'Candidates'.")
        widgets.section_label(tab, "Trường cần lấy")
        self.chk_name = widgets.checkbox(tab, "Tên ứng viên (không kèm ID)")
        self.chk_id = widgets.checkbox(tab, "ID (prefix + start code từ tên file)")
        self.chk_email = widgets.checkbox(tab, "Email")
        self.chk_phone = widgets.checkbox(tab, "Số điện thoại")
        widgets.hint(tab, "Tên & ID tách từ tên file (nên đổi tên ở tab trước). "
                          "Email & SĐT đọc từ nội dung CV (PDF/DOCX). File .doc cũ chưa hỗ trợ.")
        act = QHBoxLayout()
        act.setContentsMargins(0, 14, 0, 0)
        act.addWidget(widgets.button(tab, "Trích xuất ra Excel", variant="primary",
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
        self.info("Đã lưu", "Đã lưu cấu hình ✅")

    def _files(self):
        folder = self.var_folder.get().strip()
        if not folder or not os.path.isdir(folder):
            self.error("Thiếu thư mục", "Vui lòng chọn thư mục chứa CV.")
            return None, None
        files = sorted(p for p in Path(folder).iterdir()
                       if p.is_file() and p.suffix.lower() in _CV_EXTENSIONS)
        if not files:
            self.info("Không có file", "Không tìm thấy file PDF/DOC/DOCX trong thư mục.")
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
            self.error("Thiếu thư viện", "Cần openpyxl để xuất Excel:\n  pip install openpyxl")
            return
        folder, files = self._files()
        if not files:
            return
        target = self.var_output.get().strip()
        if not target:
            self.error("Thiếu đích xuất",
                       "Chọn thư mục (tạo file mới) hoặc file Excel (nối tiếp).")
            return
        dept = self.var_dept.get().strip()
        want = dict(name=self.chk_name.get(), id=self.chk_id.get(),
                    email=self.chk_email.get(), phone=self.chk_phone.get())
        if not any(want.values()):
            self.error("Chưa chọn trường", "Hãy chọn ít nhất một trường để trích xuất.")
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
                    mode = "mới"
                elif os.path.isfile(target):
                    wb, ws = _open_existing_workbook(target)
                    out, mode = target, "nối tiếp"
                elif target.lower().endswith(".xlsx") and os.path.isdir(os.path.dirname(target)):
                    wb, ws = _open_template_workbook()
                    out, mode = target, "mới"
                else:
                    dlg.set_final_status("Đích không hợp lệ.")
                    dlg.log("⚠ Hãy chọn THƯ MỤC hoặc FILE .xlsx hợp lệ.")
                    return
                _write_candidates(ws, rows)
                wb.save(out)
            except PermissionError:
                dlg.set_final_status("Không ghi được file.")
                dlg.log(f"⚠ File đang mở trong Excel? Hãy đóng rồi thử lại:\n{target}")
                return
            except Exception as exc:
                dlg.set_final_status("Lỗi ghi Excel.")
                dlg.log(f"⚠ {exc}")
                return
            dlg.set_final_status(f"Hoàn thành — {len(rows)} CV ({mode}).")
            dlg.log(f"\n✅ Đã lưu {len(rows)} CV vào:\n{out}")
            if errors:
                dlg.log(f"⚠ {len(errors)} file không đọc được nội dung (xem trên).")

        dlg = ProgressDialog(self._page, "Đang trích xuất CV…", total=total,
                             subtitle=f"Đọc {total} file CV")
        dlg.start(job, on_finish)


class _RenamePreview(QDialog):
    """Bảng xem trước đổi tên — sửa được cột 'Tên ứng viên', tên mới tự cập nhật."""

    def __init__(self, parent, files, prefix, start_str, noise):
        super().__init__(parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setModal(True)
        self._files = files
        self._prefix = prefix
        self._suffixes = []
        self._codes = []
        self._drag = None

        shell = QVBoxLayout(self)
        shell.setContentsMargins(24, 24, 24, 24)
        card = QFrame(self)
        card.setObjectName("Dialog")
        card.setMinimumWidth(900)
        widgets.add_shadow(card, blur=48, dy=12, alpha=70)
        shell.addWidget(card)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(24, 20, 24, 18)
        lay.setSpacing(10)

        t = QLabel(f"Tìm thấy {len(files)} file — double-click cột giữa để chỉnh tên")
        t.setObjectName("DialogTitle")
        lay.addWidget(t)
        sub = QLabel("Tên ứng viên tự trích từ tên file gốc; cột 'Tên file mới' cập nhật ngay.")
        sub.setObjectName("DialogMsg")
        lay.addWidget(sub)

        self.table = QTableWidget(len(files), 3)
        self.table.setHorizontalHeaderLabels(
            ["Tên file gốc", "Tên ứng viên (sửa được)", "Tên file mới (xem trước)"])
        self.table.verticalHeader().setVisible(False)
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

        foot = QHBoxLayout()
        foot.addWidget(widgets.button(card, "Đổi tên tất cả", variant="success",
                                      icon="check", command=self._do_rename))
        foot.addWidget(widgets.button(card, "Hủy", variant="neutral", icon="x",
                                      command=self.reject))
        foot.addStretch(1)
        lay.addLayout(foot)
        self.resize(960, 620)

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
        msg = f"Đã đổi tên {renamed} file."
        if skipped:
            msg += f"\n{skipped} file bỏ qua (tên không đổi)."
        if errors:
            msg += "\n\nLỗi:\n" + "\n".join(errors[:8])
        self.accept()
        if errors:
            dialogs.warning(self.parent(), "Hoàn thành (có lỗi)", msg)
        else:
            dialogs.success(self.parent(), "Hoàn thành", msg)

    # kéo di chuyển
    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag = e.globalPosition().toPoint() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, e):
        if self._drag is not None and e.buttons() & Qt.LeftButton:
            self.move(e.globalPosition().toPoint() - self._drag)

    def mouseReleaseEvent(self, e):
        self._drag = None
