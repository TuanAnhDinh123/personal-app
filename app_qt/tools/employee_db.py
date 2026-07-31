"""Quản lý nhân viên (SQLite) — bản PySide6.

Giao diện dựng theo mẫu tool "Quản lý CV ứng viên": tìm kiếm (từ khóa + lọc bộ
phận / giới tính / level) + bảng liệt kê đầy đủ cột (có checkbox) + CRUD.
Tầng dữ liệu dùng lại app.core.cv_repository (bảng `employees`).
"""
import datetime
import re

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFileDialog, QHBoxLayout, QLabel, QLineEdit, QVBoxLayout,
)

from app.core import config
from app.core import cv_repository as repo
from app.core import cv_schema
from app_qt import dialogs, theme, widgets
from app_qt.base_tool import BaseTool
from app_qt.components.column_picker import ColumnPicker
from app_qt.components.form_dialog import FormDialog
from app_qt.components.modal import ModalDialog
from app_qt.components.table import DataTable

try:
    import openpyxl
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

# Sentinel: cột trong Excel là TEXT, cần tra ra id ở bảng master trước khi ghi.
_DEPT_TEXT = "__department_short_name__"   # tra theo departments.short_name
_LEVEL_TEXT = "__level_name__"             # tra theo levels.level_name

# Map tiêu đề cột trong file Excel → field trong DB (theo ảnh mapping người dùng
# gửi). Khóa đã CHUẨN HÓA (viết thường + bỏ khoảng trắng thừa). Có kèm vài
# biến thể/alias cho chắc, khớp cả khi tiêu đề hơi khác.
#
# "Function (Common)" tra ra department_id qua departments.short_name (KHÔNG
# dùng "Business Unit (Department)" nữa — cột đó chỉ còn text hiển thị, khi
# xem sẽ mapping qua bảng master departments).
_EXCEL_HEADER_MAP = {
    "ec":                             "code",
    "emp code":                       "code",
    "employee code":                  "code",
    "code":                           "code",
    "globalempcode":                  "global_code",
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
    "job level":                      _LEVEL_TEXT,
    "function (common)":              _DEPT_TEXT,
    "function":                       _DEPT_TEXT,
    "street (address)":               "address",
    "address":                        "address",
    "position status":                "status",
    "status":                         "status",
}

# Tiêu đề cột CỐ TÌNH bỏ qua khi import (không báo "unrecognized column" vì đã
# biết rõ lý do bỏ): thông tin trùng lặp, không dùng, hoặc chỉ hiển thị qua
# mapping bảng master thay vì lưu thẳng text.
_EXCEL_IGNORED_HEADERS = {
    "legal entity (company)",
    "stt",
    "job title with level (no use)",
    "business unit (department)",
    "business unit",
    "department",
    "job title (description)",
    "birthday",
    "(old) phone number",
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


_PHONE_SPLIT_RE = re.compile(r"[,;/\n]+")


def _normalize_phones(text):
    """Chuẩn hóa 1 ô số điện thoại → chuỗi nhiều số ngăn cách bởi "; ".

    Cột `phone` có thể chứa nhiều số (1 nhân viên nhiều SĐT) — file Excel
    thường gộp chung 1 ô, ngăn cách bởi dấu phẩy/chấm phẩy/gạch chéo/xuống
    dòng. Ghép lại đồng nhất về "; " để hiển thị & tìm kiếm (LIKE) nhất quán.
    """
    parts = [p.strip() for p in _PHONE_SPLIT_RE.split(text) if p.strip()]
    return "; ".join(parts)


# Dò dòng header THẬT trong N dòng đầu file — file gốc thường có vài dòng
# tiêu đề/logo/ghi chú phía trên (đôi khi bị ẨN) trước khi tới dòng tên cột.
_HEADER_SCAN_ROWS = 20
_HEADER_MIN_MATCHES = 3    # số cột khớp tối thiểu để coi 1 dòng là header thật
_KNOWN_HEADERS = set(_EXCEL_HEADER_MAP) | _EXCEL_IGNORED_HEADERS


def _find_header_row(ws):
    """Trả về SỐ THỨ TỰ dòng chứa tên cột thật (1-based).

    Quét `_HEADER_SCAN_ROWS` dòng đầu, chọn dòng có nhiều ô khớp tên cột đã
    biết (`_EXCEL_HEADER_MAP`/`_EXCEL_IGNORED_HEADERS`) nhất. Không dòng nào đạt
    tối thiểu `_HEADER_MIN_MATCHES` → coi dòng 1 là header (hành vi cũ, để
    không vỡ với file đơn giản không có phần tiêu đề thừa phía trên).
    """
    best_row, best_score = 1, -1
    for row_idx, row in enumerate(
            ws.iter_rows(min_row=1, max_row=_HEADER_SCAN_ROWS, values_only=True), start=1):
        score = sum(1 for v in row if v and _norm(v) in _KNOWN_HEADERS)
        if score > best_score:
            best_row, best_score = row_idx, score
    return best_row if best_score >= _HEADER_MIN_MATCHES else 1


def _normalize_status(text):
    """Chuẩn hóa cột 'Position status' trong Excel → 1 trong
    cv_schema.EMPLOYEE_STATUS_CHOICES ("Working"/"Resigned").

    Giá trị không nhận diện được (không chứa từ khóa quen) → trả "" (bỏ trống)
    thay vì gán liều, tránh ghi sai trạng thái đang làm/đã nghỉ của nhân viên.
    """
    t = _norm(text)
    if "active" in t:
        return "Working"
    if any(kw in t for kw in ("terminat", "resign", "leave", "inactive")):
        return "Resigned"
    return ""

# ─────────────────────────────────────────────────────────────────────────
#  BỀ RỘNG (px) CÁC CỘT BẢNG NHÂN VIÊN — chỉnh tùy ý ở đây.
# ─────────────────────────────────────────────────────────────────────────
EMP_COL_WIDTHS = {
    "employee_id":     56,   # vừa đủ 4 ký tự (kể cả padding 8px 2 bên)
    "code":            90,
    "global_code":     100,
    "full_name":       170,
    "surname":         90,
    "middle_name":     100,
    "name":            90,
    "date_of_birth":   95,
    "gender":          70,
    "education":       130,
    "phone":           160,
    "email":           210,
    "level_name":      90,
    "department_name": 140,
    "address":         200,
    "status":          90,
}

_W = EMP_COL_WIDTHS

# Cột bảng NHÂN VIÊN: (khóa, tiêu đề, rộng, canh lề). Liệt kê đầy đủ mọi cột.
_EMP_COLUMNS = [
    ("employee_id",     "ID",          _W["employee_id"],     "center"),
    ("code",            "Emp code",    _W["code"],            "w"),
    ("global_code",     "Global code", _W["global_code"],     "w"),
    ("full_name",       "Full name",   _W["full_name"],       "w"),
    ("surname",         "Surname",     _W["surname"],         "w"),
    ("middle_name",     "Middle name", _W["middle_name"],     "w"),
    ("name",            "Name",        _W["name"],            "w"),
    ("date_of_birth",   "Date of birth", _W["date_of_birth"], "center"),
    ("gender",          "Gender",      _W["gender"],          "center"),
    ("education",       "Education",   _W["education"],       "w"),
    ("phone",           "Phone",       _W["phone"],           "w"),
    ("email",           "Email",       _W["email"],           "w"),
    ("level_name",      "Level",       _W["level_name"],      "center"),
    ("department_name", "Department",  _W["department_name"], "w"),
    ("address",         "Address",     _W["address"],         "w"),
    ("status",          "Status",      _W["status"],          "center"),
]

# Bảng còn được bổ sung nhiều cột nữa → UI KHÔNG hiện hết. Đây là các cột hiện
# MẶC ĐỊNH; người dùng bật/tắt thêm ở dropdown "Columns" (lựa chọn được lưu lại).
_EMP_DEFAULT_COLUMNS = [
    "employee_id", "code", "global_code", "full_name",
    "date_of_birth", "phone", "email", "department_name", "status",
]

# Section cấu hình để nhớ tập cột người dùng đã chọn (%APPDATA%/…/config.json).
_CFG_SECTION = "employee_db"
_CFG_COLUMNS = "visible_columns"


def _dept_options():
    return {d["department_name"] or f"#{d['department_id']}": d["department_id"]
            for d in repo.list_departments()}


def _level_options():
    """Danh sách cho ô lọc/form Level: lấy theo bảng danh mục `levels` (Master
    Data → Levels), giữ đúng thứ tự sort_order. {tên hiển thị: level_id}."""
    return {r["level_name"] or f"#{r['level_id']}": r["level_id"]
            for r in repo.list_levels()}


class _DuplicateCodesDialog(ModalDialog):
    """Modal cảnh báo mã NV bị trùng khi import Excel (đã có sẵn trong DB).

    `dups` = list sqlite3.Row bảng `employees` (employee_id, code, full_name)
    bị trùng. Hiện trong DataTable để có thể copy (chọn ô → Ctrl+C, hoặc chuột
    phải → Copy — DataTable đã hỗ trợ sẵn). Trả về True (tiếp tục import, bỏ
    qua các dòng trùng) / False (hủy import) qua .run().
    """

    def __init__(self, parent, dups):
        super().__init__(parent, "sm")
        self._result = False
        card, lay = self.build_shell(f"Duplicate employee codes · {len(dups)}")

        desc = QLabel("These employee codes already exist in the database. "
                     "They will be skipped if you continue:")
        desc.setObjectName("DialogMsg")
        desc.setWordWrap(True)
        lay.addWidget(desc)

        rows = [{"code": d["code"] or "", "full_name": d["full_name"] or ""}
                for d in dups]
        table = DataTable([
            ("code", "Emp code", 140),
            ("full_name", "Existing employee", 240),
        ])
        table.set_rows(rows)
        table.setMinimumHeight(min(260, self.modal_h))
        lay.addWidget(table, 1)
        self.set_grow_region(table)

        hint = QLabel("Select a cell and press Ctrl+C (or right-click → Copy) "
                     "to copy the code.")
        hint.setObjectName("Hint")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        foot = QHBoxLayout()
        foot.addWidget(widgets.button(card, "Import the rest, skip duplicates",
                                      variant="success", icon="check",
                                      command=lambda: self._choose(True)))
        foot.addWidget(widgets.button(card, "Cancel import", variant="neutral",
                                      icon="x", command=lambda: self._choose(False)))
        foot.addStretch(1)
        lay.addLayout(foot)

    def _choose(self, value):
        self._result = value
        self.accept()

    def run(self):
        self.exec()
        return self._result


class EmployeeDbTool(BaseTool):
    name = "Employees"
    description = "Search, manage work status, export reports."
    icon = "👥"
    category = "Human Resources"
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

        widgets.section_label(card, "Search employees")
        self._build_search_bar(lay)

        # Dựng bảng TRƯỚC thanh nút (nút "Columns" cần tham chiếu tới bảng),
        # nhưng thêm vào layout SAU để thứ tự hiển thị vẫn là: nút → bảng.
        self.table = DataTable(_EMP_COLUMNS, pk="employee_id",
                               stretch_key="email", on_double=self._edit,
                               checkable=True)
        self._build_toolbar(lay)
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
            "Search by employee code — paste multiple codes separated by spaces…")
        self.ent_codes.setClearButtonEnabled(True)
        self.ent_codes.addAction(widgets.svg_icon("idcard", theme.TEXT_MUTED, 16),
                                 QLineEdit.LeadingPosition)
        self.ent_codes.editingFinished.connect(self._reload)
        lay.addWidget(self.ent_codes)

        # Hàng 2: ô tìm free-text (rộng hơn) + các ô lọc select + nút đặt lại.
        filters = QHBoxLayout()
        filters.setSpacing(10)
        self.ent_kw = QLineEdit()
        self.ent_kw.setPlaceholderText("Search…")
        self.ent_kw.setClearButtonEnabled(True)
        self.ent_kw.addAction(widgets.svg_icon("search", theme.TEXT_MUTED, 16),
                              QLineEdit.LeadingPosition)
        self.ent_kw.editingFinished.connect(self._reload)
        # Ô free-text: cùng CHIỀU CAO với ô select (QLineEdit & QComboBox chung
        # QSS → cao bằng nhau), chỉ chiếm phần rộng còn lại nên NHÌN dài hơn các
        # ô select một chút. Canh giữa theo chiều dọc để thẳng hàng với ô select
        # (ô select cao 54px, ô text ~36px → tự canh giữa cho khớp).
        filters.addWidget(self.ent_kw, 1, Qt.AlignVCenter)

        # Các ô select để bề rộng CỐ ĐỊNH & bằng nhau (nếu không, combo tự giãn
        # theo nội dung dài nhất — vd tên bộ phận — nuốt hết chỗ của ô text).
        self.sel_dept = widgets.FilterSelect("Department")
        self.sel_gender = widgets.FilterSelect("Gender")
        self.sel_level = widgets.FilterSelect("Level")
        self.sel_status = widgets.FilterSelect("Status")
        self.sel_gender.set_options(cv_schema.GENDER_CHOICES)
        self.sel_status.set_options(cv_schema.EMPLOYEE_STATUS_CHOICES)
        # self.sel_dept / self.sel_level nạp options ở _reload() (đọc từ DB).
        for w in (self.sel_dept, self.sel_gender, self.sel_level, self.sel_status):
            w.setFixedWidth(180)
            w.changed.connect(self._reload)
            filters.addWidget(w, 0)
        filters.addWidget(widgets.button(None, "Reset", variant="neutral",
                                         icon="eraser", command=self._clear_filters), 0)
        lay.addLayout(filters)

    def _build_toolbar(self, lay):
        bar = QHBoxLayout()
        bar.setSpacing(6)
        B = widgets.button
        bar.addWidget(B(None, "Add", variant="success", icon="plus", command=self._add))
        bar.addWidget(B(None, "Enroll", variant="info", icon="award",
                        command=self._enroll_to_course))
        bar.addWidget(B(None, "Import from Excel", variant="primary", icon="download",
                        command=self._batch_import))
        bar.addStretch(1)
        bar.addWidget(self._build_column_picker())
        bar.addWidget(B(None, "Reload", variant="neutral", icon="refresh", command=self._reload))
        lay.addLayout(bar)

    def _build_column_picker(self):
        """Dropdown tích chọn cột hiển thị; nhớ lựa chọn qua config.json."""
        saved = config.load(_CFG_SECTION).get(_CFG_COLUMNS)

        def _save(keys):
            cfg = config.load(_CFG_SECTION)
            cfg[_CFG_COLUMNS] = list(keys)
            config.save(_CFG_SECTION, cfg)

        self.col_picker = ColumnPicker(self.table, _EMP_DEFAULT_COLUMNS,
                                       on_change=_save)
        if saved:
            self.col_picker.set_keys(saved, notify=False)
        return self.col_picker

    # -------------------------------------------------------------- dữ liệu
    def _reload(self):
        dept_opts = _dept_options()
        self.sel_dept.set_options(dept_opts.keys())
        dept_id = dept_opts.get(self.sel_dept.value())
        # Nạp lại danh mục Level mỗi lần tìm → thêm/sửa cấp bậc ở trang Master
        # Data → Levels là thấy ngay, không cần mở lại tool.
        level_opts = _level_options()
        self.sel_level.set_options(level_opts.keys())
        level_id = level_opts.get(self.sel_level.value())

        rows = repo.search_employees(
            self.ent_kw.text(), department_id=dept_id,
            gender=self.sel_gender.value(), level_id=level_id,
            status=self.sel_status.value(),
            codes=self.ent_codes.text().split())
        self.table.set_rows(rows)
        self.count_lbl.setText(
            f"Showing {len(rows)} employees · Total in DB: {repo.count_employees()}")

    def _clear_filters(self):
        self.ent_kw.clear()
        self.ent_codes.clear()
        for w in (self.sel_dept, self.sel_gender, self.sel_level, self.sel_status):
            w.clear()
        self._reload()

    def _selected_id(self):
        eid = self.table.selected_id()
        if eid is None:
            dialogs.info(self._root, "Nothing selected", "Please select an employee in the table.")
        return eid

    # -------------------------------------------- ghi danh vào khóa học (Enroll)
    # Tick chọn nhiều nhân viên → chọn 1 khóa học trong modal → OK: thêm TẤT CẢ
    # người đã tick vào bảng course_employees (bỏ qua người đã ghi danh trước đó).
    def _enroll_to_course(self):
        rows = self.table.checked_rows()
        if not rows:
            dialogs.info(self._root, "Nothing selected",
                         "Tick at least one employee to enroll in a course.")
            return
        courses = repo.list_courses()
        if not courses:
            dialogs.info(self._root, "No courses",
                         "There are no courses to enroll employees into yet.")
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
        msg = f'Enrolled {added} employees in "{title}".'
        if skipped:
            msg += f"\nSkipped {skipped} already enrolled."
        dialogs.success(self._root, "Done", msg)

    def _pick_course(self, courses, n_selected):
        """Modal chọn 1 khóa học. Trả về course_id đã chọn, hoặc None nếu hủy."""
        dlg = ModalDialog(self._root, "sm")
        card, lay = dlg.build_shell("Enroll to course")

        info = QLabel(f"Enroll {n_selected} selected employees in a course:")
        info.setObjectName("DialogMsg")
        info.setWordWrap(True)
        lay.addWidget(info)

        lbl = QLabel("Course")
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
        foot.addWidget(widgets.button(card, "Cancel", variant="neutral", icon="x",
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
            dialogs.error(self._root, "Missing library",
                          "openpyxl is required to read Excel:\n  pip install openpyxl")
            return
        path, _ = QFileDialog.getOpenFileName(
            self._root, "Choose the employee list Excel file", "",
            "Excel (*.xlsx *.xlsm);;All files (*.*)")
        if not path:
            return
        try:
            rows, unknown = self._read_excel(path)
        except Exception as exc:
            dialogs.error(self._root, "Read error", f"Couldn't read Excel:\n{exc}")
            return
        if not rows:
            dialogs.info(self._root, "Empty", "No valid data rows found.")
            return

        note = ""
        if unknown:
            note = ("\n\nUnrecognized columns (skipped): "
                    + ", ".join(unknown[:10])
                    + (" …" if len(unknown) > 10 else ""))
        if not dialogs.confirm(
                self._root, "Confirm import",
                f"Found {len(rows)} employees in the file.\n\nImport into the DB?{note}",
                ok_label="Import"):
            return

        # Mã NV (`code`) đã có sẵn trong DB → hỏi lại trước khi ghi trùng.
        dup_rows = repo.find_employees_by_codes([r.get("code") for r in rows])
        skip_codes = set()
        if dup_rows:
            if not _DuplicateCodesDialog(self._root, dup_rows).run():
                return
            skip_codes = {_norm(d["code"]) for d in dup_rows if d["code"]}

        # Tra department_id theo MÃ VIẾT TẮT (short_name) — cột "Function
        # (Common)" trong file, khớp không phân biệt hoa/thường.
        dept_by_short = {_norm(d["short_name"]): d["department_id"]
                         for d in repo.list_departments()
                         if d["short_name"]}
        # Tra level_id theo tên cấp bậc trong danh mục `levels`.
        level_by_name = {_norm(l["level_name"]): l["level_id"]
                         for l in repo.list_levels()
                         if l["level_name"]}

        added = 0
        skipped = 0
        missing_depts = set()    # mã bộ phận trong file nhưng không có trong DB
        missing_levels = set()   # tên cấp bậc trong file nhưng không có trong danh mục
        for rec in rows:
            if skip_codes and _norm(rec.get("code", "")) in skip_codes:
                skipped += 1
                continue
            dept_text = rec.pop(_DEPT_TEXT, "")
            if dept_text:
                dept_id = dept_by_short.get(_norm(dept_text))
                if dept_id is not None:
                    rec["department_id"] = dept_id
                else:
                    missing_depts.add(dept_text)
            level_text = rec.pop(_LEVEL_TEXT, "")
            if level_text:
                level_id = level_by_name.get(_norm(level_text))
                if level_id is not None:
                    rec["level_id"] = level_id
                else:
                    missing_levels.add(level_text)
            repo.insert_employee(rec)
            added += 1

        self._reload()
        msg = f"Imported {added} employees."
        if skipped:
            msg += f"\n\nSkipped {skipped} duplicate employee code(s)."
        if missing_depts:
            names = ", ".join(sorted(missing_depts)[:10])
            more = " …" if len(missing_depts) > 10 else ""
            msg += ("\n\nDepartments not found by short name (link left empty): "
                    f"{names}{more}\nCreate/fix them on the 'Departments' page "
                    "and re-import if you need the link.")
        if missing_levels:
            names = ", ".join(sorted(missing_levels)[:10])
            more = " …" if len(missing_levels) > 10 else ""
            msg += ("\n\nJob levels not found in the 'Levels' master data "
                    f"(link left empty): {names}{more}\nAdd them on the "
                    "'Levels' page and re-import if you need the link.")
        dialogs.success(self._root, "Done", msg)

    @staticmethod
    def _read_excel(path):
        """Đọc file Excel → (list rec, list tiêu đề cột không nhận diện được).

        Mỗi rec là dict {field DB → giá trị}, riêng cột bộ phận / cấp bậc giữ
        TEXT dưới khóa sentinel `_DEPT_TEXT` / `_LEVEL_TEXT` (sẽ tra ra id ở
        bước sau, xem `_batch_import`). Nếu thiếu full_name thì ghép từ
        surname + middle_name + name (thứ tự tên tiếng Việt).

        Dòng header KHÔNG chắc luôn ở dòng 1 — file gốc thường có vài dòng tiêu
        đề/logo/ghi chú phía trên (có thể bị ẨN) trước khi tới dòng tên cột thật
        → dò tìm dòng đó bằng `_find_header_row` thay vì đinh cứng dòng 1.
        """
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        header_row = _find_header_row(ws)
        header = next(
            ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True), None)
        if not header:
            wb.close()
            return [], []

        col_key = {}       # chỉ số cột → field DB
        unknown = []       # tiêu đề không map được (để báo lại)
        for idx, title in enumerate(header):
            if title is None or not str(title).strip():
                continue
            norm_title = _norm(title)
            key = _EXCEL_HEADER_MAP.get(norm_title)
            if key:
                col_key[idx] = key
            elif norm_title not in _EXCEL_IGNORED_HEADERS:
                unknown.append(str(title).strip())

        rows = []
        for values in ws.iter_rows(min_row=header_row + 1, values_only=True):
            rec = {}
            for idx, key in col_key.items():
                if idx < len(values):
                    v = _cell_str(values[idx])
                    if v == "":
                        continue
                    rec[key] = v
            if rec.get("phone"):
                rec["phone"] = _normalize_phones(rec["phone"])
            if rec.get("status"):
                rec["status"] = _normalize_status(rec["status"])
                if not rec["status"]:
                    del rec["status"]
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
            {"kind": "section", "label": "Identity"},
            {"key": "code", "label": "Employee code", "kind": "text"},
            {"key": "global_code", "label": "Global code", "kind": "text"},
            {"kind": "section", "label": "Personal info"},
            {"key": "full_name", "label": "Full name (*)", "kind": "text", "required": True},
            {"key": "surname", "label": "Surname", "kind": "text"},
            {"key": "middle_name", "label": "Middle name", "kind": "text"},
            {"key": "name", "label": "Name", "kind": "text"},
            {"key": "date_of_birth", "label": "Date of birth (dd/mm/yyyy)", "kind": "text"},
            {"key": "gender", "label": "Gender", "kind": "choice",
             "choices": cv_schema.GENDER_CHOICES, "allow_empty": True},
            {"key": "education", "label": "Education", "kind": "text"},
            {"key": "phone", "label": 'Phone (separate multiple with "; ")', "kind": "text"},
            {"key": "email", "label": "Email", "kind": "text"},
            {"key": "address", "label": "Address", "kind": "text"},
            {"kind": "section", "label": "Job"},
            {"key": "level_id", "label": "Level", "kind": "dropdown",
             "options": _level_options},
            {"key": "department_id", "label": "Department", "kind": "dropdown",
             "options": _dept_options},
            {"key": "status", "label": "Status", "kind": "choice",
             "choices": cv_schema.EMPLOYEE_STATUS_CHOICES, "allow_empty": True},
        ]

    def _add(self):
        def _save(data):
            repo.insert_employee(data)
            self._reload()

        FormDialog(self._root, "Add employee",
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

        FormDialog(self._root, "Edit employee",
                   self._employee_form_specs(), current,
                   on_save=_save, on_delete=lambda: self._delete(eid)).run()

    def _delete(self, eid):
        """Xóa nhân viên; trả về False nếu người dùng hủy xác nhận (giữ form mở)."""
        row = repo.get_employee(eid)
        name = (row["full_name"] if row and row["full_name"] else f"#{eid}")
        if not dialogs.confirm(self._root, "Confirm delete",
                               f'Delete employee "{name}" from the DB?', ok_label="Delete"):
            return False
        repo.delete_employee(eid)
        self._reload()
        return True
