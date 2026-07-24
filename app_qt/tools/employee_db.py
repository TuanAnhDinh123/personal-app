"""Quản lý nhân viên (SQLite) — bản PySide6.

Giao diện dựng theo mẫu tool "Quản lý CV ứng viên": tìm kiếm (từ khóa + lọc bộ
phận / giới tính / level) + bảng liệt kê đầy đủ cột (có checkbox) + CRUD.
Tầng dữ liệu dùng lại app.core.cv_repository (bảng `employees`).
"""
import datetime

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFileDialog, QHBoxLayout, QLabel, QLineEdit, QVBoxLayout,
)

from app.core import cv_repository as repo
from app.core import cv_schema
from app_qt import dialogs, theme, widgets
from app_qt.base_tool import BaseTool
from app_qt.components.form_dialog import FormDialog
from app_qt.components.modal import ModalDialog
from app_qt.components.table import DataTable

try:
    import openpyxl
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

# Sentinel: cột "bộ phận" trong Excel là TEXT → cần tra ra department_id.
_DEPT_TEXT = "__department_name__"

# Map tiêu đề cột trong file Excel → field trong DB (theo ảnh mapping người dùng
# gửi). Khóa đã CHUẨN HÓA (viết thường + bỏ khoảng trắng thừa). Có kèm vài
# biến thể/alias cho chắc, khớp cả khi tiêu đề hơi khác.
_EXCEL_HEADER_MAP = {
    "ec":                             "code",
    "emp code":                       "code",
    "employee code":                  "code",
    "code":                           "code",
    "globalemp code":                 "global_code",
    "global emp code":                "global_code",
    "global code":                    "global_code",
    "global_code":                    "global_code",
    "full name":                      "full_name",
    "fullname":                       "full_name",
    "surname":                        "surname",
    "name":                           "name",
    "middle name (only for vietnam)": "middle_name",
    "middle name":                    "middle_name",
    "date of birth":                  "date_of_birth",
    "dob":                            "date_of_birth",
    "gender":                         "gender",
    "education level":                "education",
    "education":                      "education",
    "phone number":                   "phone",
    "phone":                          "phone",
    "personal email address":         "email",
    "email":                          "email",
    "job level":                      "level",
    "level":                          "level",
    "business unit (department)":     _DEPT_TEXT,
    "business unit":                  _DEPT_TEXT,
    "department":                     _DEPT_TEXT,
    "street (address)":               "address",
    "address":                        "address",
}


def _norm(text):
    """Chuẩn hóa tiêu đề cột / tên bộ phận: về chữ thường, gộp khoảng trắng."""
    return " ".join(str(text).strip().lower().split())


def _cell_str(value):
    """Đổi 1 ô Excel → chuỗi. Ngày/giờ → 'dd/mm/yyyy'; còn lại strip chuỗi."""
    if value is None:
        return ""
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime("%d/%m/%Y")
    return str(value).strip()

# ─────────────────────────────────────────────────────────────────────────
#  BỀ RỘNG (px) CÁC CỘT BẢNG NHÂN VIÊN — chỉnh tùy ý ở đây.
# ─────────────────────────────────────────────────────────────────────────
EMP_COL_WIDTHS = {
    "employee_id":     40,
    "code":            90,
    "global_code":     100,
    "full_name":       170,
    "surname":         90,
    "middle_name":     100,
    "name":            90,
    "date_of_birth":   95,
    "gender":          70,
    "education":       130,
    "phone":           120,
    "email":           210,
    "level":           90,
    "department_name": 140,
    "address":         200,
}

_W = EMP_COL_WIDTHS

# Cột bảng NHÂN VIÊN: (khóa, tiêu đề, rộng, canh lề). Liệt kê đầy đủ mọi cột.
_EMP_COLUMNS = [
    ("employee_id",     "ID",          _W["employee_id"],     "center"),
    ("code",            "Mã NV",       _W["code"],            "w"),
    ("global_code",     "Global Code", _W["global_code"],     "w"),
    ("full_name",       "Họ tên",      _W["full_name"],       "w"),
    ("surname",         "Họ",          _W["surname"],         "w"),
    ("middle_name",     "Tên đệm",     _W["middle_name"],     "w"),
    ("name",            "Tên",         _W["name"],            "w"),
    ("date_of_birth",   "Ngày sinh",   _W["date_of_birth"],   "center"),
    ("gender",          "Giới tính",   _W["gender"],          "center"),
    ("education",       "Học vấn",     _W["education"],       "w"),
    ("phone",           "SĐT",         _W["phone"],           "w"),
    ("email",           "Email",       _W["email"],           "w"),
    ("level",           "Level",       _W["level"],           "center"),
    ("department_name", "Bộ phận",     _W["department_name"], "w"),
    ("address",         "Địa chỉ",     _W["address"],         "w"),
]


def _dept_options():
    return {d["department_name"] or f"#{d['department_id']}": d["department_id"]
            for d in repo.list_departments()}


class EmployeeDbTool(BaseTool):
    name = "Quản lý nhân viên"
    description = "Tìm kiếm, thêm/sửa/xóa hồ sơ nhân viên (SQLite)."
    icon = "👥"
    category = "Nhân sự"
    order = 10
    fills_height = True

    # Dựng thẳng thẻ full-height (giống CandidateDbTool) thay cho khung mặc định.
    def build(self, parent=None):
        repo.init_db()
        card = widgets.Card(parent)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(22, 20, 22, 18)
        lay.setSpacing(10)
        self._root = card

        widgets.section_label(card, "Tìm kiếm nhân viên")
        self._build_search_bar(lay)
        self._build_toolbar(lay)

        self.table = DataTable(_EMP_COLUMNS, pk="employee_id",
                               stretch_key="email", on_double=self._edit,
                               checkable=True)
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
        # Hàng 1: tìm theo MÃ NV — dán nguyên cột mã từ Excel (nhiều mã cách nhau
        # bởi dấu cách), khớp CHÍNH XÁC từng mã.
        self.ent_codes = QLineEdit()
        self.ent_codes.setPlaceholderText(
            "Tìm theo mã NV — dán nhiều mã, cách nhau bởi dấu cách…")
        self.ent_codes.setClearButtonEnabled(True)
        self.ent_codes.addAction(widgets.svg_icon("idcard", theme.TEXT_MUTED, 16),
                                 QLineEdit.LeadingPosition)
        self.ent_codes.editingFinished.connect(self._reload)
        lay.addWidget(self.ent_codes)

        # Hàng 2: ô tìm free-text (rộng hơn) + các ô lọc select + nút đặt lại.
        filters = QHBoxLayout()
        filters.setSpacing(10)
        self.ent_kw = QLineEdit()
        self.ent_kw.setPlaceholderText("Tìm kiếm…")
        self.ent_kw.setClearButtonEnabled(True)
        self.ent_kw.addAction(widgets.svg_icon("search", theme.TEXT_MUTED, 16),
                              QLineEdit.LeadingPosition)
        self.ent_kw.editingFinished.connect(self._reload)
        # Ô free-text: cùng CHIỀU CAO với ô select (QLineEdit & QComboBox chung
        # QSS → cao bằng nhau), chỉ chiếm phần rộng còn lại nên NHÌN dài hơn các
        # ô select một chút. Canh giữa theo chiều dọc để thẳng hàng với ô select
        # (ô select cao 54px, ô text ~36px → tự canh giữa cho khớp).
        filters.addWidget(self.ent_kw, 1, Qt.AlignVCenter)

        # Ba ô select để bề rộng CỐ ĐỊNH & bằng nhau (nếu không, combo tự giãn
        # theo nội dung dài nhất — vd tên bộ phận — nuốt hết chỗ của ô text).
        self.sel_dept = widgets.FilterSelect("Bộ phận")
        self.sel_gender = widgets.FilterSelect("Giới tính")
        self.sel_level = widgets.FilterSelect("Level")
        self.sel_gender.set_options(cv_schema.GENDER_CHOICES)
        self.sel_level.set_options(cv_schema.EMPLOYEE_LEVEL_CHOICES)
        for w in (self.sel_dept, self.sel_gender, self.sel_level):
            w.setFixedWidth(180)
            w.changed.connect(self._reload)
            filters.addWidget(w, 0)
        filters.addWidget(widgets.button(None, "Đặt lại", variant="neutral",
                                         icon="eraser", command=self._clear_filters), 0)
        lay.addLayout(filters)

    def _build_toolbar(self, lay):
        bar = QHBoxLayout()
        bar.setSpacing(6)
        B = widgets.button
        bar.addWidget(B(None, "Thêm mới", variant="success", icon="plus", command=self._add))
        bar.addWidget(B(None, "Enroll", variant="info", icon="award",
                        command=self._enroll_to_course))
        bar.addWidget(B(None, "Nhập từ Excel", variant="primary", icon="download",
                        command=self._batch_import))
        bar.addStretch(1)
        bar.addWidget(B(None, "Tải lại", variant="neutral", icon="refresh", command=self._reload))
        lay.addLayout(bar)

    # -------------------------------------------------------------- dữ liệu
    def _reload(self):
        dept_opts = _dept_options()
        self.sel_dept.set_options(dept_opts.keys())
        dept_id = dept_opts.get(self.sel_dept.value())

        rows = repo.search_employees(
            self.ent_kw.text(), department_id=dept_id,
            gender=self.sel_gender.value(), level=self.sel_level.value(),
            codes=self.ent_codes.text().split())
        self.table.set_rows(rows)
        self.count_lbl.setText(
            f"Hiển thị {len(rows)} nhân viên · Tổng trong DB: {repo.count_employees()}")

    def _clear_filters(self):
        self.ent_kw.clear()
        self.ent_codes.clear()
        for w in (self.sel_dept, self.sel_gender, self.sel_level):
            w.clear()
        self._reload()

    def _selected_id(self):
        eid = self.table.selected_id()
        if eid is None:
            dialogs.info(self._root, "Chưa chọn", "Vui lòng chọn một nhân viên trong bảng.")
        return eid

    # -------------------------------------------- ghi danh vào khóa học (Enroll)
    # Tick chọn nhiều nhân viên → chọn 1 khóa học trong modal → OK: thêm TẤT CẢ
    # người đã tick vào bảng course_employees (bỏ qua người đã ghi danh trước đó).
    def _enroll_to_course(self):
        rows = self.table.checked_rows()
        if not rows:
            dialogs.info(self._root, "Chưa chọn",
                         "Hãy tick chọn ít nhất một nhân viên để thêm vào khóa học.")
            return
        courses = repo.list_courses()
        if not courses:
            dialogs.info(self._root, "Chưa có khóa học",
                         "Chưa có khóa học nào trong hệ thống để thêm nhân viên vào.")
            return
        course_id = self._pick_course(courses, len(rows))
        if course_id is None:
            return

        added = skipped = 0
        for row in rows:
            eid = row["employee_id"]
            if eid is None:
                continue
            rid = repo.enroll_employee(
                course_id, eid, {"status": cv_schema.COURSE_STATUS_CHOICES[0]})
            if rid:
                added += 1
            else:
                skipped += 1   # đã ghi danh trước đó (unique course_id+employee_id)

        title = next((c["title"] for c in courses
                      if c["course_id"] == course_id), "") or f"#{course_id}"
        msg = f'Đã thêm {added} nhân viên vào khóa "{title}".'
        if skipped:
            msg += f"\nBỏ qua {skipped} người đã ghi danh trước đó."
        dialogs.success(self._root, "Hoàn tất", msg)

    def _pick_course(self, courses, n_selected):
        """Modal chọn 1 khóa học. Trả về course_id đã chọn, hoặc None nếu hủy."""
        dlg = ModalDialog(self._root, "sm")
        card, lay = dlg.build_shell("Enroll to course")

        info = QLabel(f"Thêm {n_selected} nhân viên đã chọn vào khóa học:")
        info.setObjectName("DialogMsg")
        info.setWordWrap(True)
        lay.addWidget(info)

        lbl = QLabel("Khóa học")
        lbl.setObjectName("FieldLabel")
        lay.addWidget(lbl)
        combo = widgets.ComboBox(card)
        combo.addItems([self._course_label(c) for c in courses])
        lay.addWidget(combo)
        lay.addStretch(1)

        result = {"id": None}

        def _ok():
            i = combo.currentIndex()
            if 0 <= i < len(courses):
                result["id"] = courses[i]["course_id"]
            dlg.accept()

        foot = QHBoxLayout()
        foot.addWidget(widgets.button(card, "OK", variant="success", icon="check",
                                      command=_ok))
        foot.addWidget(widgets.button(card, "Hủy", variant="neutral", icon="x",
                                      command=dlg.reject))
        foot.addStretch(1)
        lay.addLayout(foot)
        dlg.exec()
        return result["id"]

    @staticmethod
    def _course_label(c):
        title = (c["title"] or "").strip() or f"#{c['course_id']}"
        date = (str(c["date"]).strip() if c["date"] else "")
        return f"{title} · {date}" if date else title

    # ----------------------------------------------------- nhập hàng loạt Excel
    def _batch_import(self):
        if not _OPENPYXL_OK:
            dialogs.error(self._root, "Thiếu thư viện",
                          "Cần openpyxl để đọc Excel:\n  pip install openpyxl")
            return
        path, _ = QFileDialog.getOpenFileName(
            self._root, "Chọn file Excel danh sách nhân viên", "",
            "Excel (*.xlsx *.xlsm);;Tất cả (*.*)")
        if not path:
            return
        try:
            rows, unknown = self._read_excel(path)
        except Exception as exc:
            dialogs.error(self._root, "Lỗi đọc file", f"Không đọc được Excel:\n{exc}")
            return
        if not rows:
            dialogs.info(self._root, "Trống", "Không tìm thấy dòng dữ liệu hợp lệ.")
            return

        note = ""
        if unknown:
            note = ("\n\nCác cột không nhận diện được (bỏ qua): "
                    + ", ".join(unknown[:10])
                    + (" …" if len(unknown) > 10 else ""))
        if not dialogs.confirm(
                self._root, "Xác nhận nhập",
                f"Đọc được {len(rows)} nhân viên từ file.\n\nNhập vào DB?{note}",
                ok_label="Nhập"):
            return

        # Tra department_id theo TÊN bộ phận (khớp không phân biệt hoa/thường).
        dept_by_name = {_norm(d["department_name"]): d["department_id"]
                        for d in repo.list_departments()
                        if d["department_name"]}

        added = 0
        missing_depts = set()   # tên bộ phận trong file nhưng không có trong DB
        for rec in rows:
            dept_text = rec.pop(_DEPT_TEXT, "")
            if dept_text:
                dept_id = dept_by_name.get(_norm(dept_text))
                if dept_id is not None:
                    rec["department_id"] = dept_id
                else:
                    missing_depts.add(dept_text)
            repo.insert_employee(rec)
            added += 1

        self._reload()
        msg = f"Đã nhập {added} nhân viên."
        if missing_depts:
            names = ", ".join(sorted(missing_depts)[:10])
            more = " …" if len(missing_depts) > 10 else ""
            msg += ("\n\nKhông tìm thấy bộ phận (để trống liên kết): "
                    f"{names}{more}\nTạo các bộ phận này ở trang 'Bộ phận' rồi "
                    "nhập lại nếu cần gắn đúng.")
        dialogs.success(self._root, "Hoàn tất", msg)

    @staticmethod
    def _read_excel(path):
        """Đọc file Excel → (list rec, list tiêu đề cột không nhận diện được).

        Mỗi rec là dict {field DB → giá trị}, riêng cột bộ phận giữ TEXT dưới
        khóa sentinel `_DEPT_TEXT` (sẽ tra ra id ở bước sau). Nếu thiếu full_name
        thì ghép từ surname + middle_name + name (thứ tự tên tiếng Việt).
        """
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if not header:
            wb.close()
            return [], []

        col_key = {}       # chỉ số cột → field DB
        unknown = []       # tiêu đề không map được (để báo lại)
        for idx, title in enumerate(header):
            if title is None or not str(title).strip():
                continue
            key = _EXCEL_HEADER_MAP.get(_norm(title))
            if key:
                col_key[idx] = key
            else:
                unknown.append(str(title).strip())

        rows = []
        for values in ws.iter_rows(min_row=2, values_only=True):
            rec = {}
            for idx, key in col_key.items():
                if idx < len(values):
                    v = _cell_str(values[idx])
                    if v == "":
                        continue
                    rec[key] = v
            if not rec.get("full_name"):
                parts = [rec.get("surname"), rec.get("middle_name"), rec.get("name")]
                composed = " ".join(p for p in parts if p)
                if composed:
                    rec["full_name"] = composed
            # Bỏ dòng rỗng hoàn toàn (không có định danh nào).
            if any(rec.get(k) for k in ("full_name", "code", "global_code", "email")):
                rows.append(rec)
        wb.close()
        return rows, unknown

    # ------------------------------------------------------------- form specs
    def _employee_form_specs(self):
        return [
            {"kind": "section", "label": "Định danh"},
            {"key": "code", "label": "Mã nhân viên", "kind": "text"},
            {"key": "global_code", "label": "Global Code", "kind": "text"},
            {"kind": "section", "label": "Thông tin cá nhân"},
            {"key": "full_name", "label": "Họ và tên (*)", "kind": "text", "required": True},
            {"key": "surname", "label": "Họ", "kind": "text"},
            {"key": "middle_name", "label": "Tên đệm", "kind": "text"},
            {"key": "name", "label": "Tên", "kind": "text"},
            {"key": "date_of_birth", "label": "Ngày sinh (dd/mm/yyyy)", "kind": "text"},
            {"key": "gender", "label": "Giới tính", "kind": "choice",
             "choices": cv_schema.GENDER_CHOICES, "allow_empty": True},
            {"key": "education", "label": "Trình độ học vấn", "kind": "text"},
            {"key": "phone", "label": "Số điện thoại", "kind": "text"},
            {"key": "email", "label": "Email", "kind": "text"},
            {"key": "address", "label": "Địa chỉ", "kind": "text"},
            {"kind": "section", "label": "Công việc"},
            {"key": "level", "label": "Cấp bậc (level)", "kind": "choice",
             "choices": cv_schema.EMPLOYEE_LEVEL_CHOICES, "allow_empty": True},
            {"key": "department_id", "label": "Bộ phận", "kind": "dropdown",
             "options": _dept_options},
        ]

    def _add(self):
        def _save(data):
            repo.insert_employee(data)
            self._reload()

        FormDialog(self._root, "Thêm nhân viên mới",
                   self._employee_form_specs(), None, on_save=_save).run()

    def _edit(self, eid=None):
        if eid is None:
            eid = self._selected_id()
        if eid is None:
            return
        current = repo.get_employee(eid)

        def _save(data):
            repo.update_employee(eid, data)
            self._reload()

        FormDialog(self._root, "Sửa nhân viên",
                   self._employee_form_specs(), current,
                   on_save=_save, on_delete=lambda: self._delete(eid)).run()

    def _delete(self, eid):
        """Xóa nhân viên; trả về False nếu người dùng hủy xác nhận (giữ form mở)."""
        row = repo.get_employee(eid)
        name = (row["full_name"] if row and row["full_name"] else f"#{eid}")
        if not dialogs.confirm(self._root, "Xác nhận xóa",
                               f'Xóa nhân viên "{name}" khỏi DB?', ok_label="Xóa"):
            return False
        repo.delete_employee(eid)
        self._reload()
        return True
