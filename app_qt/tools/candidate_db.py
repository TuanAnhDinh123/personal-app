"""Quản lý CV ứng viên & danh mục tuyển dụng (SQLite) — bản PySide6.

Port của app/tools/candidate_db.py. Tầng dữ liệu (app.core.cv_repository,
app.core.cv_schema) dùng lại 100% — chỉ dựng lại giao diện bằng Qt:
    • Tool chính "Quản lý CV ứng viên": tìm kiếm + bảng + CRUD + nhập Excel.
    • 6 trang Master Data (dùng CrudTablePanel): Bộ phận · Loại nhân viên ·
      Cấp bậc · Cost center · Vị trí · Khóa học.

Mỗi vị trí chỉ có ĐÚNG 1 mô tả công việc (JD) nên JD không còn trang riêng —
tiêu đề + file JD nhập ngay trong form của trang "Vị trí tuyển dụng".
"""
import os
import unicodedata
from pathlib import Path

from PySide6.QtCore import QDate, QDateTime, QTime, Qt
from PySide6.QtGui import QCursor
from PySide6.QtWidgets import (
    QApplication, QCalendarWidget, QFileDialog, QFrame, QHBoxLayout,
    QLabel, QLineEdit, QTimeEdit, QToolTip, QVBoxLayout, QWidget,
)

from app.core import cv_repository as repo
from app.core import cv_schema
from app.core import outlook
from app.core.cv_scan import (
    _open_existing_workbook, _open_template_workbook, _split_id_name,
    _write_candidates,
)
from app_qt import dialogs, theme, widgets
from app_qt.base_tool import BaseTool
from app_qt.components.crud_panel import CrudTablePanel
from app_qt.components.dialog_base import build_dialog_shell
from app_qt.components.form_dialog import FormDialog
from app_qt.components.modal import ModalDialog
from app_qt.components.table import DataTable
from app_qt.richtext import RichText

try:
    import openpyxl
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

# ─────────────────────────────────────────────────────────────────────────
#  BỀ RỘNG (px) CÁC CỘT BẢNG ỨNG VIÊN — chỉnh tùy ý ở đây.
#  (Cột checkbox 'chọn' nằm ở app_qt/components/table.py → CHECK_COL_WIDTH.)
# ─────────────────────────────────────────────────────────────────────────
CAND_COL_WIDTHS = {
    "candidate_id":    56,   # vừa đủ 4 ký tự (kể cả padding 8px 2 bên)
    "full_name":       180,
    "email":           210,
    "phone":           120,
    "position_title":  100,
    "fit_score":       60,
    "cv_file_path":    160,
    "department_name": 140,
    "batch":           90,
    "status":          100,
    "date_of_birth":   95,
    "applied_at":      105,
    "note":            200,
}

_W = CAND_COL_WIDTHS

# Cột bảng ỨNG VIÊN: (khóa, tiêu đề, rộng, canh lề[, formatter]).
_CAND_COLUMNS = [
    ("candidate_id",    "ID",            _W["candidate_id"],    "center"),
    ("full_name",       "Full name",     _W["full_name"],       "w"),
    ("email",           "Email",         _W["email"],           "w"),
    ("phone",           "Phone",         _W["phone"],           "w"),
    ("position_title",  "Position",      _W["position_title"],  "w"),
    ("fit_score",       "Score",         _W["fit_score"],       "center"),
    ("cv_file_path",    "CV",            _W["cv_file_path"],    "w",
     lambda v: os.path.basename(str(v)) if v else ""),
    ("department_name", "Department",    _W["department_name"], "w"),
    ("batch",           "Batch",         _W["batch"],           "center"),
    ("status",          "Status",        _W["status"],          "center"),
    ("date_of_birth",   "Date of birth", _W["date_of_birth"],   "center"),
    ("applied_at",      "Applied",       _W["applied_at"],      "center"),
    ("note",            "Note",          _W["note"],            "w",
     lambda v: (str(v).replace("\n", " ")[:60] + "…")
     if v and len(str(v)) > 60 else (str(v).replace("\n", " ") if v else "")),
]

# Map tiêu đề cột Excel (đã hạ chữ thường) → khóa dữ liệu. Giữ CẢ nhãn tiếng Anh
# (bản mới) lẫn nhãn tiếng Việt (file cũ) để vẫn nhập được file xuất trước đây.
_EXCEL_HEADER_MAP = {
    "batch":            "batch",
    "full name":        "full_name",
    "date of birth":    "date_of_birth",
    "email":            "email",
    "phone":            "phone",
    "fit score":        "fit_score",
    "fit summary":      "fit_summary",
    "strengths":        "strengths",
    "weaknesses":       "weaknesses",
    "file name":        "cv_file_path",
    "cv path":          "cv_file_path",
    # --- nhãn tiếng Việt cũ (tương thích ngược) ---
    "họ tên":           "full_name",
    "ngày sinh":        "date_of_birth",
    "số điện thoại":    "phone",
    "điểm phù hợp":     "fit_score",
    "đánh giá phù hợp": "fit_summary",
    "ưu điểm":          "strengths",
    "nhược điểm":       "weaknesses",
    "tên file":         "cv_file_path",
    "đường dẫn cv":     "cv_file_path",
}


def _num(text, kind):
    s = str(text).strip()
    if not s:
        return None
    try:
        return int(float(s)) if kind == "int" else float(s)
    except ValueError:
        return None


def _strip_accents(s):
    """Bỏ dấu tiếng Việt: 'Tuấn Anh' → 'Tuan Anh' (xử lý riêng đ/Đ)."""
    s = s.replace("đ", "d").replace("Đ", "D")
    return "".join(c for c in unicodedata.normalize("NFD", s)
                   if unicodedata.category(c) != "Mn")


def _given_name(full_name):
    """Tên gọi (không dấu) từ họ tên đầy đủ — lấy TỪ CUỐI: 'Đinh Tuấn Anh' → 'Anh'."""
    parts = (full_name or "").strip().split()
    return _strip_accents(parts[-1]) if parts else ""


def _fill_template(text, mapping):
    """Thay các placeholder {name}{possion}{position}{date}{time}{time_start}
    {time_end} trong `text` bằng giá trị thật."""
    if not text:
        return text or ""
    for key, val in mapping.items():
        text = text.replace("{" + key + "}", val)
    return text


# Tên thứ/tháng tiếng Anh cố định — KHÔNG dùng locale máy (đang là tiếng Việt)
# để định dạng ngày luôn ra dạng 'Fri, 24th Jul 2026'.
_WEEKDAYS_EN = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
_MONTHS_EN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _ordinal(n):
    """Số thứ tự tiếng Anh: 1→1st, 2→2nd, 3→3rd, 4→4th, 11→11th, 21→21st…"""
    if 11 <= n % 100 <= 13:
        return f"{n}th"
    return f"{n}{ {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th') }"


def _fmt_date_en(dt):
    """datetime → 'Fri, 24th Jul 2026'."""
    return (f"{_WEEKDAYS_EN[dt.weekday()]}, {_ordinal(dt.day)} "
            f"{_MONTHS_EN[dt.month - 1]} {dt.year}")


def _fmt_time_en(dt):
    """datetime → '08:30 AM' (12 giờ, có AM/PM, không phụ thuộc locale)."""
    suffix = "AM" if dt.hour < 12 else "PM"
    hour12 = dt.hour % 12 or 12
    return f"{hour12:02d}:{dt.minute:02d} {suffix}"


def _picker_qss():
    """QSS riêng cho lịch + ô giờ trong hộp chọn ngày giờ.

    QSS toàn cục cho QTableView/::item bị QCalendarWidget kế thừa (khiến ô ngày
    bị bo góc, đệm sai, header ngày/tuần co lại thành '…'). Ghi đè tại đây bằng
    selector cụ thể hơn để lịch hiển thị gọn gàng, đúng tông màu app.
    """
    P = theme.PALETTE
    return f"""
    QCalendarWidget QWidget {{ alternate-background-color: {P['--input-bg']};
        color: {P['--text']}; }}
    QCalendarWidget QAbstractItemView {{
        background: {P['--input-bg']}; color: {P['--text']};
        selection-background-color: {P['--accent']}; selection-color: #ffffff;
        outline: none; border: none; border-radius: 0;
        gridline-color: transparent; padding: 2px;
    }}
    QCalendarWidget QAbstractItemView::item {{
        border: none; border-radius: 8px; padding: 2px; }}
    QCalendarWidget QAbstractItemView:disabled {{ color: {P['--text-faint']}; }}
    QCalendarWidget QHeaderView::section {{
        background: transparent; color: {P['--text-muted']};
        border: none; padding: 4px 0; font-weight: 600; }}
    QCalendarWidget #qt_calendar_navigationbar {{
        background: {P['--card-bg']};
        border-top-left-radius: 10px; border-top-right-radius: 10px; }}
    QCalendarWidget QToolButton {{
        color: {P['--text']}; background: transparent; font-size: 13px;
        font-weight: 600; padding: 5px 12px; border-radius: 8px; margin: 3px; }}
    QCalendarWidget QToolButton:hover {{ background: {P['--accent-soft']}; }}
    QCalendarWidget QToolButton::menu-indicator {{ image: none; }}
    QCalendarWidget QMenu {{ background: {P['--card-bg']}; color: {P['--text']};
        border: 1px solid {P['--border-strong']}; border-radius: 8px; }}
    QCalendarWidget QSpinBox {{ background: {P['--input-bg']}; color: {P['--text']};
        border: 1px solid {P['--border-strong']}; border-radius: 8px;
        padding: 2px 6px; }}
    QTimeEdit {{ background: {P['--input-bg']}; color: {P['--text']};
        border: 1px solid {P['--border-strong']}; border-radius: 10px;
        padding: 8px 10px; }}
    QTimeEdit:focus {{ border: 1px solid {P['--accent']}; }}
    """


def _dept_options():
    return {d["department_name"] or f"#{d['department_id']}": d["department_id"]
            for d in repo.list_departments()}


def _position_options():
    return {p["position_title"] or f"#{p['position_id']}": p["position_id"]
            for p in repo.list_positions()}


def _course_type_options():
    """Tên loại khóa học → mã số lưu trong DB (inhouse=0, external=1, funded=2)."""
    return {name: i for i, name in enumerate(cv_schema.COURSE_TYPE_CHOICES)}


def _course_type_label(v):
    """Mã số course_type → nhãn hiển thị; giá trị lạ/rỗng → '—'."""
    try:
        return cv_schema.COURSE_TYPE_CHOICES[int(v)]
    except (ValueError, TypeError, IndexError):
        return "—"


def _txt(row, key):
    """Đọc row[key] an toàn (sqlite3.Row/dict) → chuỗi đã strip, None → ''."""
    try:
        v = row[key]
    except (KeyError, IndexError):
        return ""
    return "" if v is None else str(v).strip()


def _score_color(score):
    """Màu chip điểm phù hợp: cao→xanh, trung bình→cam, thấp→đỏ."""
    try:
        s = float(score)
    except (TypeError, ValueError):
        return theme.PALETTE["--text-muted"]
    if s >= 80:
        return theme.PALETTE["--success"]
    if s >= 50:
        return theme.PALETTE["--warning"]
    return theme.PALETTE["--danger"]


def _chip(parent, text, color):
    """Nhãn pill nhỏ: nền tông nhạt của `color`, chữ `color`."""
    lbl = QLabel(text, parent)
    r, g, b = widgets._hex_to_rgb(color)
    lbl.setStyleSheet(
        f"background: rgba({r},{g},{b},0.15); color:{color}; border-radius:10px;"
        " padding:3px 10px; font-size:12px; font-weight:600;")
    return lbl


def _copy_chip(parent, value):
    """Chip hiển thị `value` + icon copy; bấm để sao chép vào clipboard."""
    chip = QFrame(parent)
    chip.setObjectName("CopyChip")
    chip.setCursor(Qt.PointingHandCursor)
    h = QHBoxLayout(chip)
    h.setContentsMargins(9, 3, 9, 3)
    h.setSpacing(5)
    ico = QLabel(chip)
    ico.setPixmap(widgets.svg_pixmap("copy", theme.PALETTE["--text-muted"], 13))
    h.addWidget(ico)
    txt = QLabel(value, chip)
    txt.setObjectName("CopyChipText")
    h.addWidget(txt)

    def _do_copy(_e):
        QApplication.clipboard().setText(value)
        QToolTip.showText(QCursor.pos(), "Copied", chip)

    chip.mousePressEvent = _do_copy
    return chip


# ═══════════════════════════ MASTER DATA specs ══════════════════════════
def _master_specs():
    return {
        "department": {
            "title": "department", "pk": "department_id",
            "list_fn": repo.list_departments,
            "get": repo.get_department, "insert": repo.insert_department,
            "update": repo.update_department, "delete": repo.delete_department,
            "columns": [
                ("department_id", "ID", 50),
                ("department_name", "Department name", 200),
                ("short_name", "Short code", 100),
                ("manager_name", "Manager", 150),
                ("description", "Description", 220),
            ],
            "form": [
                {"key": "department_name", "label": "Department name (*)",
                 "kind": "text", "required": True},
                {"key": "short_name", "label": "Short code (e.g. FIN, IT, R&D)",
                 "kind": "text"},
                {"key": "manager_name", "label": "Manager", "kind": "text"},
                {"key": "description", "label": "Description", "kind": "textarea", "height": 3},
            ],
        },
        "position": {
            "title": "position", "pk": "position_id",
            "modal_size": "md",   # form vị trí có mẫu mail dài → dùng cỡ md
            "list_fn": repo.list_positions,
            "get": repo.get_position, "insert": repo.insert_position,
            "update": repo.update_position, "delete": repo.delete_position,
            "columns": [
                ("position_id", "ID", 50),
                ("position_code", "Code", 90),
                ("position_title", "Position", 190),
                ("department_name", "Department", 140),
                ("level", "Level", 80),
                ("headcount", "Qty", 50),
                ("status", "Status", 105),
                ("jd_file_path", "JD file", 160, "w",
                 lambda v: os.path.basename(str(v)) if v else "—"),
                ("mail_subject", "Mail template", 180, "w",
                 lambda v: (str(v).replace("\n", " ")[:50] + "…")
                 if v and len(str(v)) > 50 else (str(v) if v else "—")),
            ],
            "form": [
                {"key": "department_id", "label": "Department", "kind": "dropdown",
                 "options": _dept_options},
                {"key": "position_code", "label": "Position code", "kind": "text"},
                {"key": "position_title", "label": "Position title (*)",
                 "kind": "text", "required": True},
                {"key": "level", "label": "Level", "kind": "text"},
                {"key": "headcount", "label": "Headcount", "kind": "int"},
                {"key": "status", "label": "Status", "kind": "choice",
                 "choices": cv_schema.POSITION_STATUS_CHOICES, "allow_empty": True},
                # Mỗi vị trí chỉ có 1 JD → nhập ngay tại form vị trí (không còn
                # trang master "Mô tả công việc (JD)" riêng).
                {"kind": "section", "label": "Job description (JD)"},
                {"key": "jd_file_path", "label": "JD file (local path)",
                 "kind": "file",
                 "filetypes": [("PDF/Word/Text", "*.pdf *.doc *.docx *.txt"),
                               ("All files", "*.*")]},
                {"kind": "section", "label": "Interview invite email template"},
                {"key": "mail_cc", "label": "CC (separate emails with ;)",
                 "kind": "text"},
                {"key": "mail_subject", "label": "Email subject", "kind": "text"},
                {"key": "mail_body", "label": "Email body (use {name} "
                 "{possion} {date} {time_start} {time_end})", "kind": "richtext",
                 "height": 20, "grow": True},
            ],
        },
        # ── Danh mục nhân sự (nạp sẵn từ Code.xlsx — xem cv_schema.SEED_DATA) ──
        "employee_type": {
            "title": "employee type", "pk": "employee_type_id",
            "list_fn": repo.list_employee_types,
            "get": repo.get_employee_type, "insert": repo.insert_employee_type,
            "update": repo.update_employee_type, "delete": repo.delete_employee_type,
            "columns": [
                ("employee_type_id", "ID", 50),
                ("code", "Code", 100),
                ("collar", "Collar", 140),
                ("description", "Description", 300),
            ],
            "form": [
                {"key": "code", "label": "Code (*) — e.g. WC, WCA, IBC, DBCA",
                 "kind": "text", "required": True},
                {"key": "collar", "label": "Collar", "kind": "choice",
                 "choices": cv_schema.COLLAR_CHOICES, "allow_empty": True},
                {"key": "description", "label": "Description",
                 "kind": "textarea", "height": 3},
            ],
        },
        "cost_center": {
            "title": "cost center", "pk": "cost_center_id",
            "list_fn": repo.list_cost_centers,
            "get": repo.get_cost_center, "insert": repo.insert_cost_center,
            "update": repo.update_cost_center, "delete": repo.delete_cost_center,
            "columns": [
                ("cost_center_id", "ID", 50),
                ("code", "Cost center", 110),
                ("group_function", "Group function", 130),
                ("name", "Name", 190),
                ("description", "Description", 260),
            ],
            "form": [
                {"key": "code", "label": "Cost center code (*) — e.g. VN1001",
                 "kind": "text", "required": True},
                {"key": "group_function", "label": "Group function",
                 "kind": "choice", "choices": cv_schema.GROUP_FUNCTION_CHOICES,
                 "allow_empty": True},
                {"key": "name", "label": "Name (optional)", "kind": "text"},
                {"key": "description", "label": "Description",
                 "kind": "textarea", "height": 3},
            ],
        },
        "level": {
            "title": "level", "pk": "level_id",
            "list_fn": repo.list_levels,
            "get": repo.get_level, "insert": repo.insert_level,
            "update": repo.update_level, "delete": repo.delete_level,
            "columns": [
                ("level_id", "ID", 50),
                ("level_name", "Level", 190),
                ("sort_order", "Order", 70, "center"),
                ("description", "Description", 300),
            ],
            "form": [
                {"key": "level_name", "label": "Level name (*) — e.g. Manager",
                 "kind": "text", "required": True},
                {"key": "sort_order", "label": "Display order (smaller = higher)",
                 "kind": "int"},
                {"key": "description", "label": "Description",
                 "kind": "textarea", "height": 3},
            ],
        },
        "course": {
            "title": "course", "pk": "course_id",
            "modal_size": "md",   # form có ô nội dung dài → dùng cỡ md
            "list_fn": repo.list_courses,
            "get": repo.get_course, "insert": repo.insert_course,
            "update": repo.update_course, "delete": repo.delete_course,
            "columns": [
                ("course_id", "ID", 50),
                ("title", "Course title", 230),
                ("course_type", "Type", 100, "w", _course_type_label),
                ("date", "Date", 110),
                ("location", "Location", 160),
                ("content", "Content", 240, "w",
                 lambda v: (str(v).replace("\n", " ")[:60] + "…")
                 if v and len(str(v)) > 60 else (str(v) if v else "—")),
            ],
            "form": [
                {"key": "title", "label": "Course title (*)",
                 "kind": "text", "required": True},
                {"key": "course_type", "label": "Course type", "kind": "dropdown",
                 "options": _course_type_options},
                {"key": "date", "label": "Date held (yyyy-mm-dd)", "kind": "text"},
                {"key": "location", "label": "Location", "kind": "text"},
                {"key": "content", "label": "Content", "kind": "textarea",
                 "height": 6, "grow": True},
            ],
        },
    }


def _card(parent):
    """Thẻ trắng chiếm hết chỗ, có shadow — khung chung cho trang full-height."""
    card = widgets.Card(parent)
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
    name = "Departments"
    description = "Department directory."
    icon = "🏢"
    order = 10
    spec_key = "department"


class EmployeeTypeTool(_MasterPageTool):
    name = "Employee types"
    description = "Employee type codes (WC/WCA/IBC/IBCA/DBC/DBCA) and their collar group."
    icon = "🏷"
    order = 12
    spec_key = "employee_type"


class LevelTool(_MasterPageTool):
    name = "Levels"
    description = "Job level directory (Director, Manager, Officer…)."
    icon = "🎖"
    order = 14
    spec_key = "level"


class CostCenterTool(_MasterPageTool):
    name = "Cost centers"
    description = "Cost centers & group functions — group employees to compute team running cost."
    icon = "🏦"
    order = 16
    spec_key = "cost_center"


class PositionTool(_MasterPageTool):
    name = "Positions"
    description = "Open positions (with each position's JD & email template)."
    icon = "💼"
    order = 20
    spec_key = "position"


class CourseTool(_MasterPageTool):
    name = "Courses"
    description = "Training / course directory."
    icon = "🎓"
    order = 40
    spec_key = "course"


# ═════════════════════ MODAL XEM CHI TIẾT ỨNG VIÊN ══════════════════════
class _CandidateDetailDialog(ModalDialog):
    """Modal lớn xem chi tiết toàn bộ ứng viên trong danh sách hiện tại.

    Mỗi ứng viên là một thẻ; phần nổi bật nhất là NHẬN XÉT CỦA AI
    (điểm phù hợp, nhận xét, ưu điểm, nhược điểm) từ bước quét CV.
    """

    def __init__(self, parent, rows):
        super().__init__(parent, "lg")
        card, lay = self.build_shell(f"Candidate details · {len(rows)}")

        body = QWidget()
        col = QVBoxLayout(body)
        col.setContentsMargins(0, 0, 8, 0)
        col.setSpacing(12)
        if not rows:
            empty = QLabel("The current list is empty.")
            empty.setObjectName("DialogMsg")
            col.addWidget(empty)
        for row in rows:
            col.addWidget(self._candidate_card(body, row))
        col.addStretch(1)
        sa = widgets.scroll_area(body)
        lay.addWidget(sa, 1)
        self.set_grow_region(sa)   # cao theo cỡ lg, tự co khi màn hình thấp

        foot = QHBoxLayout()
        foot.addStretch(1)
        foot.addWidget(widgets.button(card, "Close", variant="neutral", icon="x",
                                      command=self.reject))
        lay.addLayout(foot)

    # ------------------------------------------------------------- thẻ 1 ứng viên
    def _candidate_card(self, parent, row):
        box = QFrame(parent)
        box.setObjectName("DetailCard")
        v = QVBoxLayout(box)
        v.setContentsMargins(16, 14, 16, 14)
        v.setSpacing(8)

        # Hàng tiêu đề: tên + chip điểm/trạng thái
        head = QHBoxLayout()
        head.setSpacing(8)
        cid = _txt(row, "candidate_id")
        name = QLabel(f"#{cid}  {_txt(row, 'full_name') or '(no name)'}", box)
        name.setObjectName("DetailName")
        head.addWidget(name, 1)
        score = _txt(row, "fit_score")
        if score:
            head.addWidget(_chip(box, f"Score {score}", _score_color(score)))
        status = _txt(row, "status")
        if status:
            head.addWidget(_chip(box, status, theme.PALETTE["--info"]))
        v.addLayout(head)

        # Hàng thông tin phụ (bôi-chọn được để copy tay nếu cần)
        meta = " · ".join(p for p in (
            _txt(row, "position_title"), _txt(row, "department_name"),
            (f"DOB: {_txt(row, 'date_of_birth')}" if _txt(row, "date_of_birth") else ""),
            (f"Applied: {_txt(row, 'applied_at')}" if _txt(row, "applied_at") else ""),
        ) if p)
        if meta:
            lbl = QLabel(meta, box)
            lbl.setObjectName("DetailMeta")
            lbl.setWordWrap(True)
            lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
            v.addWidget(lbl)

        # Email / SĐT: chip bấm-để-sao-chép
        email, phone = _txt(row, "email"), _txt(row, "phone")
        if email or phone:
            chips = QHBoxLayout()
            chips.setSpacing(8)
            if email:
                chips.addWidget(_copy_chip(box, email))
            if phone:
                chips.addWidget(_copy_chip(box, phone))
            chips.addStretch(1)
            v.addLayout(chips)

        v.addWidget(self._ai_box(box, row))

        note = _txt(row, "note")
        if note:
            v.addLayout(self._para(box, "Note", note))
        return box

    # ----------------------------------------------------- hộp nhận xét của AI
    def _ai_box(self, parent, row):
        box = QFrame(parent)
        box.setObjectName("AIBox")
        v = QVBoxLayout(box)
        v.setContentsMargins(14, 12, 14, 12)
        v.setSpacing(6)

        header = QHBoxLayout()
        header.setSpacing(6)
        ico = QLabel(box)
        ico.setPixmap(widgets.svg_pixmap("sparkles", theme.PALETTE["--accent"], 16))
        header.addWidget(ico, 0, Qt.AlignVCenter)
        h = QLabel("AI assessment", box)
        h.setObjectName("AIHeader")
        header.addWidget(h, 1)
        v.addLayout(header)

        summary = _txt(row, "fit_summary")
        strengths = _txt(row, "strengths")
        weaknesses = _txt(row, "weaknesses")

        if not any((summary, strengths, weaknesses)):
            empty = QLabel("No AI assessment for this candidate yet.", box)
            empty.setObjectName("AIEmpty")
            v.addWidget(empty)
            return box

        if summary:
            v.addLayout(self._para(box, "Fit summary", summary))
        if strengths or weaknesses:
            two = QHBoxLayout()
            two.setSpacing(12)
            two.addLayout(self._para(box, "Strengths", strengths or "—"), 1)
            two.addLayout(self._para(box, "Weaknesses", weaknesses or "—"), 1)
            v.addLayout(two)
        return box

    @staticmethod
    def _para(parent, label, value):
        """Khối nhãn nhỏ + đoạn văn bản (wrap). Trả về QVBoxLayout."""
        col = QVBoxLayout()
        col.setSpacing(2)
        lbl = QLabel(label, parent)
        lbl.setObjectName("AILabel")
        col.addWidget(lbl)
        txt = QLabel(value, parent)
        txt.setObjectName("AIText")
        txt.setWordWrap(True)
        txt.setTextInteractionFlags(Qt.TextSelectableByMouse)
        col.addWidget(txt)
        return col


# ═══════════════════════════════ TOOL CHÍNH ═════════════════════════════
class CandidateDbTool(BaseTool):
    name = "Candidate Manager"
    description = "Search candidates, manage departments/positions, bulk import (SQLite)."
    icon = "🙋"
    category = "Recruitment"
    order = 10
    fills_height = True

    def build(self, parent=None):
        repo.init_db()
        card, lay = _card(parent)
        self._root = card

        widgets.section_label(card, "Search candidates")
        self._build_search_bar(lay)

        self._build_toolbar(lay)

        self.table = DataTable(_CAND_COLUMNS, pk="candidate_id",
                               stretch_key="email", on_double=self._edit,
                               link_keys={"cv_file_path"}, on_link=self._on_file_link,
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
        # Ô tìm kiếm toàn văn: quét MỌI field text. Tìm ngay khi rời ô (tab /
        # click ra ngoài) hoặc nhấn Enter — editingFinished bao cả hai.
        self.ent_kw = QLineEdit()
        self.ent_kw.setPlaceholderText("Search…")
        self.ent_kw.setClearButtonEnabled(True)
        self.ent_kw.addAction(widgets.svg_icon("search", theme.TEXT_MUTED, 16),
                              QLineEdit.LeadingPosition)
        self.ent_kw.editingFinished.connect(self._reload)
        lay.addWidget(self.ent_kw)

        # Hàng ô lọc dạng select 'nhãn nổi' — chọn 1 option là tìm luôn.
        filters = QHBoxLayout()
        filters.setSpacing(10)
        self.sel_pos = widgets.FilterSelect("Position")
        self.sel_dept = widgets.FilterSelect("Department")
        self.sel_status = widgets.FilterSelect("Status")
        self.sel_batch = widgets.FilterSelect("Batch")
        self.sel_status.set_options(cv_schema.STATUS_CHOICES)
        for w in (self.sel_pos, self.sel_dept, self.sel_status, self.sel_batch):
            w.changed.connect(self._reload)
            filters.addWidget(w, 1)
        filters.addWidget(widgets.button(None, "Reset", variant="neutral",
                                         icon="eraser", command=self._clear_filters), 0)
        lay.addLayout(filters)

    def _build_toolbar(self, lay):
        bar = QHBoxLayout()
        bar.setSpacing(6)
        B = widgets.button
        bar.addWidget(B(None, "Add", variant="success", icon="plus", command=self._add))
        bar.addWidget(B(None, "View details", variant="info", icon="sparkles",
                        command=self._show_details))
        bar.addWidget(B(None, "Send email", variant="info", icon="mail",
                        command=self._send_mail))
        bar.addWidget(B(None, "Import from Excel", variant="primary", icon="download",
                        command=self._batch_import))
        self._btn_export = B(None, "Export to Excel", variant="warning", icon="save",
                             command=self._export_excel)
        bar.addWidget(self._btn_export)
        bar.addStretch(1)
        bar.addWidget(B(None, "Reload", variant="neutral", icon="refresh", command=self._reload))
        lay.addLayout(bar)

    # -------------------------------------------------------------- dữ liệu
    def _reload(self):
        # Nạp lại danh sách vị trí / bộ phận / batch (giữ lựa chọn cũ nếu còn).
        pos_opts = _position_options()
        dept_opts = _dept_options()
        self.sel_pos.set_options(pos_opts.keys())
        self.sel_dept.set_options(dept_opts.keys())
        self.sel_batch.set_options(repo.list_batches())

        pos_id = pos_opts.get(self.sel_pos.value())
        dept_id = dept_opts.get(self.sel_dept.value())

        rows = repo.search_candidates(
            self.ent_kw.text(), pos_id, self.sel_status.value(),
            department_id=dept_id, batch=self.sel_batch.value())
        self._rows = rows
        self.table.set_rows(rows)
        self.count_lbl.setText(
            f"Showing {len(rows)} candidates · Total in DB: {repo.count_candidates()}")

    def _clear_filters(self):
        self.ent_kw.clear()
        for w in (self.sel_pos, self.sel_dept, self.sel_status, self.sel_batch):
            w.clear()
        self._reload()

    def _selected_id(self):
        cid = self.table.selected_id()
        if cid is None:
            dialogs.info(self._root, "Nothing selected", "Please select a candidate in the table.")
        return cid

    def _show_details(self):
        rows = self.table.checked_rows()
        if not rows:
            dialogs.info(self._root, "Nothing selected",
                         "Tick at least one candidate in the table to view details.")
            return
        _CandidateDetailDialog(self._root, rows).exec()

    # ------------------------------------------------- gửi mail mời phỏng vấn
    # Tick ĐÚNG 1 ứng viên → chọn ngày giờ → soạn mail (điền sẵn từ mẫu của VỊ
    # TRÍ ứng tuyển, đã thay {name}{possion}{date}{time}) → gửi qua Outlook.
    def _send_mail(self):
        if not outlook.available():
            dialogs.warning(self._root, "Outlook required",
                            "Sending email needs Outlook on Windows (pywin32).")
            return
        rows = self.table.checked_rows()
        if len(rows) != 1:
            dialogs.warning(
                self._root, "Select exactly one candidate",
                "Please tick EXACTLY ONE candidate in the table to send an email.")
            return
        row = rows[0]
        email = _txt(row, "email")
        if not email:
            dialogs.warning(self._root, "Missing email",
                            "This candidate has no email — can't send.")
            return

        pos_id = row["position_id"] if "position_id" in row.keys() else None
        pos = repo.get_position(pos_id) if pos_id else None
        if pos is None:
            dialogs.warning(
                self._root, "No position",
                "This candidate has no position, so there's no email template. "
                "Assign a position (and write its template under Positions).")
            return

        picked = self._pick_datetime()
        if picked is None:
            return
        start, end = picked
        title = _txt(pos, "position_title")
        sd, ed = start.toPython(), end.toPython()
        start_hm, end_hm = _fmt_time_en(sd), _fmt_time_en(ed)
        mapping = {
            "name":       _given_name(_txt(row, "full_name")),
            "possion":    title,
            "position":   title,
            "date":       _fmt_date_en(sd),    # 'Fri, 24th Jul 2026'
            "time":       start_hm,            # tương thích cũ: {time} = giờ bắt đầu
            "time_start": start_hm,            # '08:30 AM'
            "time_end":   end_hm,
        }
        subject = _fill_template(_txt(pos, "mail_subject"), mapping)
        body_html = _fill_template(
            pos["mail_body"] if "mail_body" in pos.keys() and pos["mail_body"] else "",
            mapping)
        ctx = {"full_name": _txt(row, "full_name"), "position_title": title,
               "start": start, "end": end}
        self._compose_mail(email, _txt(pos, "mail_cc"), subject, body_html, ctx)

    def _pick_datetime(self):
        """Hộp thoại chọn NGÀY + GIỜ bắt đầu/kết thúc phỏng vấn.

        Dùng lịch INLINE (không phải popup) có style riêng — tránh lỗi hiển thị
        do QSS bảng toàn cục & popup trong modal frameless. Trả về (start, end)
        dạng QDateTime, hoặc None nếu hủy.
        """
        dlg, card, lay = build_dialog_shell(self._root, "Pick interview date & time",
                                            size="sm")
        lbl = QLabel("Interview date:")
        lbl.setObjectName("FieldLabel")
        lay.addWidget(lbl)

        cal = QCalendarWidget(card)
        cal.setStyleSheet(_picker_qss())
        cal.setGridVisible(False)
        cal.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)   # bỏ cột số tuần
        cal.setHorizontalHeaderFormat(QCalendarWidget.ShortDayNames)
        cal.setNavigationBarVisible(True)
        cal.setFirstDayOfWeek(Qt.Monday)
        cal.setMinimumDate(QDate.currentDate())
        cal.setSelectedDate(QDate.currentDate().addDays(1))   # mặc định: ngày mai
        cal.setMinimumHeight(260)
        lay.addWidget(cal)

        # Hàng chọn giờ bắt đầu / kết thúc.
        times = QHBoxLayout()
        times.setSpacing(12)

        def _time_col(label, default):
            col = QVBoxLayout(); col.setSpacing(4)
            cap = QLabel(label); cap.setObjectName("FieldLabel")
            col.addWidget(cap)
            te = QTimeEdit(card)
            te.setDisplayFormat("HH:mm")
            te.setStyleSheet(_picker_qss())
            te.setTime(default)
            col.addWidget(te)
            times.addLayout(col, 1)
            return te

        start_te = _time_col("Start time", QTime(9, 0))
        end_te = _time_col("End time", QTime(9, 30))
        lay.addLayout(times)

        result = {"val": None}

        def _confirm():
            d = cal.selectedDate()
            start = QDateTime(d, start_te.time())
            end = QDateTime(d, end_te.time())
            if end <= start:
                dialogs.warning(dlg, "Invalid time",
                                "End time must be after start time.")
                return
            result["val"] = (start, end)
            dlg.accept()

        foot = QHBoxLayout()
        foot.addWidget(widgets.button(card, "Continue", variant="primary",
                                      icon="check", command=_confirm))
        foot.addWidget(widgets.button(card, "Cancel", variant="neutral", icon="x",
                                      command=dlg.reject))
        foot.addStretch(1)
        lay.addLayout(foot)
        dlg.exec()
        return result["val"]

    def _compose_mail(self, to, cc, subject, body_html, ctx):
        """Hộp thoại soạn/duyệt mail → gửi qua Outlook (HTML) → tạo lịch phỏng vấn.

        `ctx` = {full_name, position_title, start, end} để tạo cuộc hẹn (appointment)
        trên lịch cá nhân ngay sau khi gửi mail thành công.
        """
        dlg, card, lay = build_dialog_shell(self._root, "Compose & send email", size="md")

        def field(label):
            lb = QLabel(label); lb.setObjectName("FieldLabel")
            lay.addWidget(lb)

        field("To")
        to_w = QLineEdit(to); lay.addWidget(to_w)
        field("CC")
        cc_w = QLineEdit(cc); lay.addWidget(cc_w)
        field("Subject")
        subj_w = QLineEdit(subject); lay.addWidget(subj_w)
        field("Body")
        body_w = RichText(card, height=12)
        body_w.set_html(body_html)
        lay.addWidget(body_w)

        foot = QHBoxLayout()

        def do_send():
            to_value = to_w.text().strip()
            if not to_value:
                dialogs.warning(dlg, "Missing recipient",
                                "Please enter a recipient email.")
                return
            try:
                outlook.send_mail(to_value, subj_w.text().strip(),
                                  body_w.get_text(), cc=cc_w.text().strip(),
                                  html=body_w.get_html())
            except Exception as exc:
                dialogs.error(dlg, "Send failed", f"Couldn't send:\n{exc}")
                return
            # Gửi xong → đặt lịch phỏng vấn (appointment cá nhân). Lỗi tạo lịch
            # KHÔNG hủy việc đã gửi mail — chỉ báo để người dùng tự thêm tay.
            appt_err = self._create_appointment(ctx, body_w.get_text())
            dlg.accept()
            if appt_err is None:
                dialogs.success(self._root, "Done",
                                "Email sent and interview scheduled ✅")
            else:
                dialogs.warning(
                    self._root, "Email sent",
                    f"Email sent, but the calendar event failed:\n{appt_err}")

        foot.addWidget(widgets.button(card, "Send", variant="primary", icon="mail",
                                      command=do_send))
        foot.addWidget(widgets.button(card, "Cancel", variant="neutral", icon="x",
                                      command=dlg.reject))
        foot.addStretch(1)
        lay.addLayout(foot)
        body_w.setMinimumHeight(round(dlg.modal_h * 0.5))   # vùng nội dung cao theo cỡ md
        dlg.exec()

    @staticmethod
    def _create_appointment(ctx, body_text):
        """Tạo cuộc hẹn phỏng vấn trên lịch Outlook. Trả về None nếu OK, hoặc
        chuỗi mô tả lỗi nếu thất bại (để bên gọi báo mà không chặn việc gửi mail)."""
        start, end = ctx["start"], ctx["end"]
        duration = max(15, start.secsTo(end) // 60)   # phút; tối thiểu 15
        who = " ".join(p for p in (ctx.get("full_name"),
                                   ctx.get("position_title")) if p)
        subject = f"Interview {who}".strip()
        try:
            outlook.create_appointment(
                subject, start.toPython(), duration_minutes=duration,
                body=body_text or "")
        except Exception as exc:   # noqa: BLE001 — báo lại cho UI, không chặn luồng
            return str(exc)
        return None

    # --------------------------------------------------- xuất Excel (các dòng đã tick)
    # Dùng lại đúng logic "Quét CV → Trích xuất Excel": ghi vào template có sẵn
    # (sheet "Candidates"). Chọn file MỚI → tạo theo template; chọn file CÓ SẴN
    # đúng mẫu → ghi nối tiếp. (Tương lai gỡ tool "Quét CV", tính năng nằm ở đây.)
    def _export_excel(self):
        if not _OPENPYXL_OK:
            dialogs.error(self._root, "Missing library",
                          "openpyxl is required to export Excel:\n  pip install openpyxl")
            return
        rows = self.table.checked_rows()
        if not rows:
            dialogs.info(self._root, "Nothing selected",
                         "Tick at least one candidate in the table to export.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self._root, "Export selected candidates to Excel", "Candidates.xlsx",
            "Excel (*.xlsx)", "", QFileDialog.Option.DontConfirmOverwrite)
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"

        export_rows = [self._to_export_row(r) for r in rows]
        self._set_export_loading(True)
        try:
            if os.path.isfile(path):
                wb, ws = _open_existing_workbook(path)
                mode = "appended"
            else:
                wb, ws = _open_template_workbook()
                mode = "new"
            _write_candidates(ws, export_rows)
            wb.save(path)
        except PermissionError:
            dialogs.error(self._root, "Can't write file",
                          f"Is the Excel file open? Close it and retry:\n{path}")
            return
        except Exception as exc:
            dialogs.error(self._root, "Excel export error", str(exc))
            return
        finally:
            self._set_export_loading(False)

        if dialogs.confirm(
                self._root, "Done",
                f"Exported {len(export_rows)} candidates ({mode}) to:\n{path}\n\nOpen now?",
                ok_label="Open", cancel_label="Close"):
            self._launch(path)

    def _set_export_loading(self, loading):
        """Bật/tắt trạng thái 'đang xuất' cho nút Xuất Excel (khóa nút + đổi chữ).

        Xuất chạy đồng bộ trên luồng chính; ép vẽ lại ngay để nút hiện trạng thái
        loading TRƯỚC khi bắt đầu ghi file (thao tác nặng làm UI đứng một chút).
        """
        btn = self._btn_export
        if loading:
            self._export_label = btn.text()
            btn.setEnabled(False)
            btn.setText("Exporting…")
        else:
            btn.setText(getattr(self, "_export_label", "Export to Excel"))
            btn.setEnabled(True)
        QApplication.processEvents()

    @staticmethod
    def _to_export_row(row):
        """Map 1 dòng ứng viên (DB) → dict theo template Candidates của cv_scan.

        Cột template: batch · id · name · apply (bộ phận) · email · phone.
        `id` (mã CV) bóc từ tên file CV nếu có dạng '<số>_<tên>'.
        """
        rec = {}
        batch = _txt(row, "batch")
        if batch:
            rec["batch"] = int(batch) if batch.isdigit() else batch
        name = _txt(row, "full_name")
        if name:
            rec["name"] = name
        cv_path = _txt(row, "cv_file_path")
        if cv_path:
            cv_id, _ = _split_id_name(Path(cv_path).stem, [])
            if cv_id:
                rec["id"] = cv_id
        dept = _txt(row, "department_name")
        if dept:
            rec["apply"] = dept
        email = _txt(row, "email")
        if email:
            rec["email"] = email
        phone = _txt(row, "phone")
        if phone:
            rec["phone"] = phone
        return rec

    # ------------------------------------------------------------- form specs
    def _candidate_form_specs(self):
        return [
            {"kind": "section", "label": "Personal info"},
            {"key": "full_name", "label": "Full name (*)", "kind": "text", "required": True},
            {"key": "email", "label": "Email", "kind": "text"},
            {"key": "phone", "label": "Phone", "kind": "text"},
            {"key": "date_of_birth", "label": "Date of birth (dd/mm/yyyy)", "kind": "text"},
            {"key": "address", "label": "Address", "kind": "text"},
            {"kind": "section", "label": "Application"},
            {"key": "position_id", "label": "Position applied for", "kind": "dropdown",
             "options": _position_options},
            {"key": "years_experience", "label": "Years of experience", "kind": "int"},
            {"key": "education", "label": "Education", "kind": "text"},
            {"key": "applied_at", "label": "Applied date (yyyy-mm-dd)", "kind": "text"},
            {"key": "status", "label": "Status", "kind": "choice",
             "choices": cv_schema.STATUS_CHOICES},
            {"key": "source", "label": "CV source", "kind": "text"},
            {"key": "batch", "label": "Batch (scan round — number)", "kind": "int"},
            {"key": "cv_file_path", "label": "CV file (local path)", "kind": "file",
             "filetypes": [("PDF/Word", "*.pdf *.doc *.docx"), ("All files", "*.*")]},
            {"kind": "section", "label": "Assessment (from CV scan)"},
            {"key": "fit_score", "label": "Fit score (0-100)", "kind": "decimal"},
            {"key": "fit_summary", "label": "Fit summary", "kind": "textarea", "height": 3},
            {"key": "strengths", "label": "Strengths", "kind": "textarea", "height": 3},
            {"key": "weaknesses", "label": "Weaknesses", "kind": "textarea", "height": 3},
            {"key": "note", "label": "Note", "kind": "textarea", "height": 3},
        ]

    def _add(self):
        def _save(data):
            dups = repo.find_duplicates(data.get("email"), data.get("phone"))
            if dups and not self._confirm_duplicate(dups):
                return False
            repo.insert_candidate(data)
            self._reload()

        FormDialog(self._root, "Add candidate",
                   self._candidate_form_specs(), None, on_save=_save).run()

    def _edit(self, cid=None):
        if cid is None:
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

        FormDialog(self._root, "Edit candidate",
                   self._candidate_form_specs(), current,
                   on_save=_save, on_delete=lambda: self._delete(cid)).run()

    def _confirm_duplicate(self, dups) -> bool:
        lines = "\n".join(
            f"  • #{d['candidate_id']} {d['full_name'] or ''}"
            f"  ({d['email'] or '—'} / {d['phone'] or '—'})"
            for d in dups[:8])
        more = "" if len(dups) <= 8 else f"\n  … and {len(dups) - 8} more"
        return dialogs.confirm(
            self._root, "Possible duplicate",
            f"{len(dups)} candidates already share this email or phone:\n\n"
            f"{lines}{more}\n\nSave this candidate anyway?",
            ok_label="Save anyway", cancel_label="Cancel")

    def _delete(self, cid):
        """Xóa ứng viên; trả về False nếu người dùng hủy xác nhận (giữ form mở)."""
        row = repo.get_candidate(cid)
        name = row["full_name"] if row else f"#{cid}"
        if not dialogs.confirm(self._root, "Confirm delete",
                               f'Delete candidate "{name}" from the DB?', ok_label="Delete"):
            return False
        repo.delete_candidate(cid)
        self._reload()
        return True

    # ------------------------------------------------------------- mở file CV
    def _on_file_link(self, row, _key):
        """Click vào tên file trong bảng → mở file CV của ứng viên đó."""
        cid = row["candidate_id"]
        if cid is not None:
            self._open_cv(cid)

    def _open_cv(self, cid):
        row = repo.get_candidate(cid)
        path = (row["cv_file_path"] or "").strip() if row else ""
        if path and os.path.isfile(path):
            self._launch(path)
            return
        if path:
            msg = (f"The CV file wasn't found at the saved path:\n{path}\n\n"
                   "It may have been moved or renamed. Locate the file now?")
        else:
            msg = "This candidate has no CV file attached. Choose one now?"
        if not dialogs.confirm(self._root, "File not found", msg,
                               ok_label="Choose file"):
            return
        new_path, _ = QFileDialog.getOpenFileName(
            self._root, "Locate the CV file", "",
            "PDF/Word (*.pdf *.doc *.docx);;All files (*.*)")
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
            dialogs.error(self._root, "Open error", f"Couldn't open the file:\n{exc}")

    # ----------------------------------------------------- nhập hàng loạt Excel
    def _batch_import(self):
        if not _OPENPYXL_OK:
            dialogs.error(self._root, "Missing library",
                          "openpyxl is required to read Excel:\n  pip install openpyxl")
            return
        path, _ = QFileDialog.getOpenFileName(
            self._root, "Choose the CV-scan result Excel file", "",
            "Excel (*.xlsx);;All files (*.*)")
        if not path:
            return
        try:
            rows = self._read_excel(path)
        except Exception as exc:
            dialogs.error(self._root, "Read error", f"Couldn't read Excel:\n{exc}")
            return
        if not rows:
            dialogs.info(self._root, "Empty", "No valid data rows found.")
            return
        if not dialogs.confirm(self._root, "Confirm import",
                               f"Found {len(rows)} candidates in the file.\n\nImport into the DB?",
                               ok_label="Import"):
            return

        # Chỉ hỏi thư mục CV khi còn đường dẫn TƯƠNG ĐỐI cần ghép (file Excel cũ
        # chỉ có tên file). File mới từ tool quét AI đã ghi sẵn đường dẫn tuyệt
        # đối nên bỏ qua bước này.
        folder = ""
        if any(self._needs_cv_folder(r) for r in rows):
            folder = QFileDialog.getExistingDirectory(
                self._root, "Folder with the CV files (skip if none)") or ""

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
        msg = f"Imported {added} candidates (no duplicates)."
        if dups:
            msg += (f"\nAlso imported {added_dup} duplicates."
                    if added_dup else f"\nSkipped {len(dups)} duplicates.")
        dialogs.success(self._root, "Done", msg)

    @staticmethod
    def _needs_cv_folder(rec):
        """True nếu cột CV là đường dẫn TƯƠNG ĐỐI cần ghép với thư mục gốc.

        Đường dẫn tuyệt đối (file mới từ tool quét AI đã ghi sẵn) thì không cần
        hỏi lại thư mục.
        """
        path = (rec.get("cv_file_path") or "").strip()
        return bool(path) and not os.path.isabs(path)

    @staticmethod
    def _apply_cv_folder(rec, folder):
        fname = (rec.get("cv_file_path") or "").strip()
        if fname and folder and not os.path.isabs(fname):
            full = os.path.join(folder, fname)
            rec["cv_file_path"] = full if os.path.isfile(full) else fname

    def _confirm_import_dups(self, dups) -> bool:
        dlg = ModalDialog(self._root, "md")
        card, lay = dlg.build_shell("Duplicate candidates")

        desc = QLabel(f"{len(dups)} candidates share an email or phone (with existing "
                      "records or within the file):")
        desc.setObjectName("DialogMsg")
        desc.setWordWrap(True)
        lay.addWidget(desc)

        table = DataTable([("full_name", "Full name", 220), ("email", "Email", 240),
                           ("phone", "Phone", 120)])
        table.set_rows(dups)
        table.setMinimumHeight(min(360, dlg.modal_h))
        lay.addWidget(table, 1)

        result = {"ok": False}
        foot = QHBoxLayout()
        foot.addWidget(widgets.button(
            card, "Import duplicates anyway", variant="success", icon="check",
            command=lambda: (result.update(ok=True), dlg.accept())))
        foot.addWidget(widgets.button(card, "Skip", variant="neutral", icon="ban",
                                      command=dlg.reject))
        foot.addStretch(1)
        lay.addLayout(foot)

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
                    v = "" if v is None else v
                    # 'Tên file' và 'Đường dẫn CV' cùng map vào cv_file_path;
                    # đừng để một cột rỗng đè lên giá trị đã đọc được từ cột kia.
                    if v == "" and str(rec.get(key, "")).strip():
                        continue
                    rec[key] = v
            rec["fit_score"] = _num(rec.get("fit_score", ""), "decimal")
            rec["batch"] = _num(rec.get("batch", ""), "int")
            if (rec.get("full_name") or "").strip():
                rows.append(rec)
        wb.close()
        return rows
