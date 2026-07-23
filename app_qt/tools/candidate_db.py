"""Quản lý CV ứng viên & danh mục tuyển dụng (SQLite) — bản PySide6.

Port của app/tools/candidate_db.py. Tầng dữ liệu (app.core.cv_repository,
app.core.cv_schema) dùng lại 100% — chỉ dựng lại giao diện bằng Qt:
    • Tool chính "Quản lý CV ứng viên": tìm kiếm + bảng + CRUD + nhập Excel.
    • 3 trang Master Data: Bộ phận · Vị trí · JD (dùng CrudTablePanel).
"""
import os

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox, QDialog, QFileDialog, QFrame, QHBoxLayout, QLabel, QLineEdit,
    QVBoxLayout, QWidget,
)

from app.core import cv_repository as repo
from app.core import cv_schema
from app_qt import dialogs, widgets
from app_qt.base_tool import BaseTool
from app_qt.components.crud_panel import CrudTablePanel
from app_qt.components.form_dialog import FormDialog
from app_qt.components.table import DataTable

try:
    import openpyxl
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

_NONE = "— Chưa chọn —"
_ALL_POS = "— Mọi vị trí —"

# Cột bảng ỨNG VIÊN: (khóa, tiêu đề, rộng, canh lề).
_CAND_COLUMNS = [
    ("candidate_id",    "ID",         50,  "center"),
    ("full_name",       "Họ tên",     180, "w"),
    ("date_of_birth",   "Ngày sinh",  95,  "center"),
    ("email",           "Email",      210, "w"),
    ("phone",           "SĐT",        115, "w"),
    ("position_title",  "Vị trí",     160, "w"),
    ("department_name", "Bộ phận",    140, "w"),
    ("fit_score",       "Điểm",       60,  "center"),
    ("status",          "Trạng thái", 100, "center"),
    ("applied_at",      "Ngày nộp",   105, "center"),
]

_EXCEL_HEADER_MAP = {
    "họ tên":           "full_name",
    "ngày sinh":        "date_of_birth",
    "email":            "email",
    "số điện thoại":    "phone",
    "điểm phù hợp":     "fit_score",
    "đánh giá phù hợp": "fit_summary",
    "ưu điểm":          "strengths",
    "nhược điểm":       "weaknesses",
    "tên file":         "cv_file_path",
}


def _num(text, kind):
    s = str(text).strip()
    if not s:
        return None
    try:
        return int(float(s)) if kind == "int" else float(s)
    except ValueError:
        return None


def _dept_options():
    return {d["department_name"] or f"#{d['department_id']}": d["department_id"]
            for d in repo.list_departments()}


def _position_options():
    return {p["position_title"] or f"#{p['position_id']}": p["position_id"]
            for p in repo.list_positions()}


# ═══════════════════════════ MASTER DATA specs ══════════════════════════
def _master_specs():
    return {
        "department": {
            "title": "bộ phận", "pk": "department_id",
            "list_fn": repo.list_departments,
            "get": repo.get_department, "insert": repo.insert_department,
            "update": repo.update_department, "delete": repo.delete_department,
            "columns": [
                ("department_id", "ID", 50),
                ("department_name", "Tên bộ phận", 200),
                ("manager_name", "Quản lý", 150),
                ("description", "Mô tả", 220),
            ],
            "form": [
                {"key": "department_name", "label": "Tên bộ phận (*)",
                 "kind": "text", "required": True},
                {"key": "manager_name", "label": "Người quản lý", "kind": "text"},
                {"key": "description", "label": "Mô tả", "kind": "textarea", "height": 3},
            ],
        },
        "position": {
            "title": "vị trí", "pk": "position_id",
            "list_fn": repo.list_positions,
            "get": repo.get_position, "insert": repo.insert_position,
            "update": repo.update_position, "delete": repo.delete_position,
            "columns": [
                ("position_id", "ID", 50),
                ("position_code", "Mã", 90),
                ("position_title", "Vị trí", 200),
                ("department_name", "Bộ phận", 150),
                ("level", "Cấp", 90),
                ("headcount", "SL", 55),
                ("status", "Trạng thái", 110),
            ],
            "form": [
                {"key": "department_id", "label": "Bộ phận", "kind": "dropdown",
                 "options": _dept_options},
                {"key": "position_code", "label": "Mã vị trí", "kind": "text"},
                {"key": "position_title", "label": "Tên vị trí (*)",
                 "kind": "text", "required": True},
                {"key": "level", "label": "Cấp bậc", "kind": "text"},
                {"key": "headcount", "label": "Số lượng cần tuyển", "kind": "int"},
                {"key": "status", "label": "Trạng thái", "kind": "choice",
                 "choices": cv_schema.POSITION_STATUS_CHOICES, "allow_empty": True},
            ],
        },
        "jd": {
            "title": "JD", "pk": "jd_id",
            "list_fn": repo.list_job_descriptions,
            "get": repo.get_job_description, "insert": repo.insert_job_description,
            "update": repo.update_job_description, "delete": repo.delete_job_description,
            "columns": [
                ("jd_id", "ID", 50),
                ("jd_title", "Tiêu đề JD", 230),
                ("position_title", "Vị trí", 180),
                ("created_at", "Ngày tạo", 140),
            ],
            "form": [
                {"key": "position_id", "label": "Vị trí", "kind": "dropdown",
                 "options": _position_options},
                {"key": "jd_title", "label": "Tiêu đề JD (*)",
                 "kind": "text", "required": True},
                {"key": "jd_file_path", "label": "File JD (đường dẫn trên máy)",
                 "kind": "file"},
            ],
        },
    }


def _card(parent):
    """Thẻ trắng chiếm hết chỗ, có shadow — khung chung cho trang full-height."""
    card = QFrame(parent)
    card.setObjectName("Card")
    widgets.add_shadow(card)
    lay = QVBoxLayout(card)
    lay.setContentsMargins(22, 20, 22, 18)
    lay.setSpacing(10)
    return card, lay


# ═══════════════════════════ TRANG MASTER DATA ══════════════════════════
class _MasterPageTool(BaseTool):
    category = "Master Data"
    show_on_home = False
    fills_height = True
    spec_key = ""

    def build(self, parent=None):
        repo.init_db()
        card, lay = _card(parent)
        spec = _master_specs()[self.spec_key]
        lay.addWidget(CrudTablePanel(spec), 1)
        return card

    def build_body(self, parent):
        pass


class DepartmentTool(_MasterPageTool):
    name = "Bộ phận"
    description = "Danh mục phòng ban / bộ phận."
    icon = "🏢"
    order = 10
    spec_key = "department"


class PositionTool(_MasterPageTool):
    name = "Vị trí tuyển dụng"
    description = "Danh mục vị trí cần tuyển."
    icon = "💼"
    order = 20
    spec_key = "position"


class JobDescriptionTool(_MasterPageTool):
    name = "Mô tả công việc (JD)"
    description = "Danh mục JD gắn với từng vị trí."
    icon = "📋"
    order = 30
    spec_key = "jd"


# ═══════════════════════════════ TOOL CHÍNH ═════════════════════════════
class CandidateDbTool(BaseTool):
    name = "Quản lý CV ứng viên"
    description = "Tìm kiếm ứng viên, quản lý bộ phận/vị trí/JD, nhập hàng loạt (SQLite)."
    icon = "🙋"
    category = "Tuyển dụng"
    order = 10
    fills_height = True

    def build(self, parent=None):
        repo.init_db()
        card, lay = _card(parent)
        self._root = card

        widgets.section_label(card, "Tìm kiếm ứng viên")
        self._build_search_bar(lay)
        hint = QLabel("Gõ tên / email / SĐT rồi Enter. Lọc thêm theo vị trí và trạng thái.")
        hint.setObjectName("Hint")
        lay.addWidget(hint)

        self._build_toolbar(lay)

        self.table = DataTable(_CAND_COLUMNS, pk="candidate_id",
                               stretch_key="email", on_double=lambda _id: self._edit())
        lay.addWidget(self.table, 1)

        self.count_lbl = QLabel("")
        self.count_lbl.setObjectName("Hint")
        lay.addWidget(self.count_lbl)

        self._reload()
        return card

    def build_body(self, parent):
        pass

    # -------------------------------------------------------------- tìm kiếm
    def _build_search_bar(self, lay):
        bar = QHBoxLayout()
        bar.setSpacing(8)
        self.ent_kw = QLineEdit()
        self.ent_kw.setPlaceholderText("Tìm theo tên, email, số điện thoại…")
        self.ent_kw.returnPressed.connect(self._reload)
        bar.addWidget(self.ent_kw, 1)

        self.cb_pos = QComboBox()
        self.cb_pos.setMinimumWidth(180)
        self.cb_pos.addItem(_ALL_POS)
        bar.addWidget(self.cb_pos)

        self.cb_status = QComboBox()
        self.cb_status.setMinimumWidth(130)
        self.cb_status.addItems(["Tất cả"] + cv_schema.STATUS_CHOICES)
        bar.addWidget(self.cb_status)

        bar.addWidget(widgets.button(None, "Tìm", variant="primary", icon="search",
                                     command=self._reload))
        bar.addWidget(widgets.button(None, "", variant="neutral", icon="x",
                                     command=self._clear_filters))
        lay.addLayout(bar)

    def _build_toolbar(self, lay):
        bar = QHBoxLayout()
        bar.setSpacing(6)
        B = widgets.button
        bar.addWidget(B(None, "Thêm mới", variant="success", icon="plus", command=self._add))
        bar.addWidget(B(None, "Sửa", variant="info", icon="pencil", command=self._edit))
        bar.addWidget(B(None, "Xóa", variant="danger", icon="trash", command=self._delete))
        bar.addWidget(B(None, "Mở CV", variant="warning", icon="folder", command=self._open_cv))
        bar.addWidget(B(None, "Nhập từ Excel", variant="primary", icon="download",
                        command=self._batch_import))
        bar.addStretch(1)
        bar.addWidget(B(None, "Tải lại", variant="neutral", icon="refresh", command=self._reload))
        lay.addLayout(bar)

    # -------------------------------------------------------------- dữ liệu
    def _reload(self):
        pos_opts = _position_options()
        names = [_ALL_POS] + list(pos_opts)
        cur = self.cb_pos.currentText()
        self.cb_pos.blockSignals(True)
        self.cb_pos.clear()
        self.cb_pos.addItems(names)
        self.cb_pos.setCurrentText(cur if cur in names else _ALL_POS)
        self.cb_pos.blockSignals(False)

        pos_id = pos_opts.get(self.cb_pos.currentText())
        status = self.cb_status.currentText()
        status = "" if status == "Tất cả" else status

        rows = repo.search_candidates(self.ent_kw.text(), pos_id, status)
        self.table.set_rows(rows)
        self.count_lbl.setText(
            f"Hiển thị {len(rows)} ứng viên · Tổng trong DB: {repo.count_candidates()}")

    def _clear_filters(self):
        self.ent_kw.clear()
        self.cb_pos.setCurrentText(_ALL_POS)
        self.cb_status.setCurrentText("Tất cả")
        self._reload()

    def _selected_id(self):
        cid = self.table.selected_id()
        if cid is None:
            dialogs.info(self._root, "Chưa chọn", "Vui lòng chọn một ứng viên trong bảng.")
        return cid

    # ------------------------------------------------------------- form specs
    def _candidate_form_specs(self):
        return [
            {"kind": "section", "label": "Thông tin cá nhân"},
            {"key": "full_name", "label": "Họ và tên (*)", "kind": "text", "required": True},
            {"key": "email", "label": "Email", "kind": "text"},
            {"key": "phone", "label": "Số điện thoại", "kind": "text"},
            {"key": "date_of_birth", "label": "Ngày sinh (dd/mm/yyyy)", "kind": "text"},
            {"key": "gender", "label": "Giới tính", "kind": "choice",
             "choices": cv_schema.GENDER_CHOICES, "allow_empty": True},
            {"key": "address", "label": "Địa chỉ", "kind": "text"},
            {"kind": "section", "label": "Ứng tuyển"},
            {"key": "position_id", "label": "Vị trí ứng tuyển", "kind": "dropdown",
             "options": _position_options},
            {"key": "years_experience", "label": "Số năm kinh nghiệm", "kind": "int"},
            {"key": "education", "label": "Học vấn", "kind": "text"},
            {"key": "applied_at", "label": "Ngày nộp (yyyy-mm-dd)", "kind": "text"},
            {"key": "status", "label": "Trạng thái", "kind": "choice",
             "choices": cv_schema.STATUS_CHOICES},
            {"key": "source", "label": "Nguồn CV", "kind": "text"},
            {"key": "cv_file_path", "label": "File CV (đường dẫn trên máy)", "kind": "file",
             "filetypes": [("PDF/Word", "*.pdf *.doc *.docx"), ("Tất cả", "*.*")]},
            {"kind": "section", "label": "Đánh giá (từ quét CV)"},
            {"key": "fit_score", "label": "Điểm phù hợp (0-100)", "kind": "decimal"},
            {"key": "fit_summary", "label": "Nhận xét phù hợp", "kind": "textarea", "height": 3},
            {"key": "strengths", "label": "Ưu điểm", "kind": "textarea", "height": 3},
            {"key": "weaknesses", "label": "Nhược điểm", "kind": "textarea", "height": 3},
            {"key": "note", "label": "Ghi chú", "kind": "textarea", "height": 3},
        ]

    def _add(self):
        def _save(data):
            dups = repo.find_duplicates(data.get("email"), data.get("phone"))
            if dups and not self._confirm_duplicate(dups):
                return False
            repo.insert_candidate(data)
            self._reload()

        FormDialog(self._root, "Thêm ứng viên mới",
                   self._candidate_form_specs(), None, on_save=_save).run()

    def _edit(self):
        cid = self._selected_id()
        if cid is None:
            return
        current = repo.get_candidate(cid)

        def _save(data):
            dups = repo.find_duplicates(data.get("email"), data.get("phone"), exclude_id=cid)
            if dups and not self._confirm_duplicate(dups):
                return False
            repo.update_candidate(cid, data)
            self._reload()

        FormDialog(self._root, "Sửa ứng viên",
                   self._candidate_form_specs(), current, on_save=_save).run()

    def _confirm_duplicate(self, dups) -> bool:
        lines = "\n".join(
            f"  • #{d['candidate_id']} {d['full_name'] or ''}"
            f"  ({d['email'] or '—'} / {d['phone'] or '—'})"
            for d in dups[:8])
        more = "" if len(dups) <= 8 else f"\n  … và {len(dups) - 8} người khác"
        return dialogs.confirm(
            self._root, "Có thể trùng ứng viên",
            f"Đã có {len(dups)} ứng viên trùng email hoặc số điện thoại:\n\n"
            f"{lines}{more}\n\nVẫn lưu ứng viên này?",
            ok_label="Vẫn lưu", cancel_label="Hủy")

    def _delete(self):
        cid = self._selected_id()
        if cid is None:
            return
        row = repo.get_candidate(cid)
        name = row["full_name"] if row else f"#{cid}"
        if dialogs.confirm(self._root, "Xác nhận xóa",
                           f'Xóa ứng viên "{name}" khỏi DB?', ok_label="Xóa"):
            repo.delete_candidate(cid)
            self._reload()

    # ------------------------------------------------------------- mở file CV
    def _open_cv(self):
        cid = self._selected_id()
        if cid is None:
            return
        row = repo.get_candidate(cid)
        path = (row["cv_file_path"] or "").strip() if row else ""
        if path and os.path.isfile(path):
            self._launch(path)
            return
        if path:
            msg = (f"Không tìm thấy file CV ở đường dẫn đã lưu:\n{path}\n\n"
                   "Có thể file đã bị di chuyển hoặc đổi tên. Chọn lại vị trí file bây giờ?")
        else:
            msg = "Ứng viên này chưa gắn file CV. Chọn file bây giờ?"
        if not dialogs.confirm(self._root, "Không tìm thấy file", msg,
                               ok_label="Chọn file"):
            return
        new_path, _ = QFileDialog.getOpenFileName(
            self._root, "Chọn lại file CV", "",
            "PDF/Word (*.pdf *.doc *.docx);;Tất cả (*.*)")
        if not new_path:
            return
        repo.set_cv_file_path(cid, new_path)
        self._reload()
        self._launch(new_path)

    def _launch(self, path):
        try:
            os.startfile(path)
        except AttributeError:
            import subprocess
            subprocess.Popen(["xdg-open", path])
        except Exception as exc:
            dialogs.error(self._root, "Lỗi mở file", f"Không mở được file:\n{exc}")

    # ----------------------------------------------------- nhập hàng loạt Excel
    def _batch_import(self):
        if not _OPENPYXL_OK:
            dialogs.error(self._root, "Thiếu thư viện",
                          "Cần openpyxl để đọc Excel:\n  pip install openpyxl")
            return
        path, _ = QFileDialog.getOpenFileName(
            self._root, "Chọn file Excel kết quả quét CV", "",
            "Excel (*.xlsx);;Tất cả (*.*)")
        if not path:
            return
        try:
            rows = self._read_excel(path)
        except Exception as exc:
            dialogs.error(self._root, "Lỗi đọc file", f"Không đọc được Excel:\n{exc}")
            return
        if not rows:
            dialogs.info(self._root, "Trống", "Không tìm thấy dòng dữ liệu hợp lệ.")
            return
        if not dialogs.confirm(self._root, "Xác nhận nhập",
                               f"Đọc được {len(rows)} ứng viên từ file.\n\nNhập vào DB?",
                               ok_label="Nhập"):
            return

        folder = ""
        if any((r.get("cv_file_path") or "").strip() for r in rows):
            folder = QFileDialog.getExistingDirectory(
                self._root, "Thư mục chứa các file CV (bỏ qua nếu không có)") or ""

        added = 0
        dups = []
        seen = set()
        for rec in rows:
            self._apply_cv_folder(rec, folder)
            email = (rec.get("email") or "").strip().lower()
            phone = (rec.get("phone") or "").strip()
            keys = set()
            if email:
                keys.add(("e", email))
            if phone:
                keys.add(("p", phone))
            is_dup = bool(keys & seen) or bool(
                repo.find_duplicates(rec.get("email"), rec.get("phone")))
            if is_dup:
                dups.append(rec)
            else:
                repo.insert_candidate(rec)
                added += 1
                seen |= keys

        added_dup = 0
        if dups and self._confirm_import_dups(dups):
            for rec in dups:
                repo.insert_candidate(rec)
                added_dup += 1

        self._reload()
        msg = f"Đã nhập {added} ứng viên (không trùng)."
        if dups:
            msg += (f"\nĐã nhập thêm {added_dup} bản trùng."
                    if added_dup else f"\nBỏ qua {len(dups)} bản trùng.")
        dialogs.success(self._root, "Hoàn tất", msg)

    @staticmethod
    def _apply_cv_folder(rec, folder):
        fname = (rec.get("cv_file_path") or "").strip()
        if fname and folder:
            full = os.path.join(folder, fname)
            rec["cv_file_path"] = full if os.path.isfile(full) else fname

    def _confirm_import_dups(self, dups) -> bool:
        dlg = QDialog(self._root)
        dlg.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        dlg.setAttribute(Qt.WA_TranslucentBackground)
        dlg.setModal(True)
        shell = QVBoxLayout(dlg)
        shell.setContentsMargins(24, 24, 24, 24)
        card = QFrame(dlg)
        card.setObjectName("Dialog")
        card.setMinimumWidth(620)
        widgets.add_shadow(card, blur=48, dy=12, alpha=70)
        shell.addWidget(card)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(24, 20, 24, 18)
        lay.setSpacing(12)

        title = QLabel("Ứng viên trùng")
        title.setObjectName("DialogTitle")
        lay.addWidget(title)
        desc = QLabel(f"Có {len(dups)} ứng viên trùng email hoặc SĐT (với người đã có "
                      "trong DB hoặc trùng nhau trong file):")
        desc.setObjectName("DialogMsg")
        desc.setWordWrap(True)
        lay.addWidget(desc)

        table = DataTable([("full_name", "Họ tên", 220), ("email", "Email", 240),
                           ("phone", "SĐT", 120)])
        table.set_rows(dups)
        lay.addWidget(table, 1)

        result = {"ok": False}
        foot = QHBoxLayout()
        foot.addWidget(widgets.button(
            card, "Vẫn nhập các bản trùng", variant="success", icon="check",
            command=lambda: (result.update(ok=True), dlg.accept())))
        foot.addWidget(widgets.button(card, "Bỏ qua", variant="neutral", icon="ban",
                                      command=dlg.reject))
        foot.addStretch(1)
        lay.addLayout(foot)

        dlg.resize(680, 460)
        dlg.exec()
        return result["ok"]

    @staticmethod
    def _read_excel(path):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if not header:
            wb.close()
            return []
        col_key = {}
        for idx, title in enumerate(header):
            if title is None:
                continue
            key = _EXCEL_HEADER_MAP.get(str(title).strip().lower())
            if key:
                col_key[idx] = key
        rows = []
        for values in ws.iter_rows(min_row=2, values_only=True):
            rec = {}
            for idx, key in col_key.items():
                if idx < len(values):
                    v = values[idx]
                    rec[key] = v if v is not None else ""
            rec["fit_score"] = _num(rec.get("fit_score", ""), "decimal")
            if (rec.get("full_name") or "").strip():
                rows.append(rec)
        wb.close()
        return rows
