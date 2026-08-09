"""Quản lý CV ứng viên & danh mục tuyển dụng (SQLite) — bản PySide6.

Port của app/tools/candidate_db.py. Tầng dữ liệu (app.core.cv_repository,
app.core.cv_schema) dùng lại 100% — chỉ dựng lại giao diện bằng Qt:
    • Tool chính "Quản lý CV ứng viên": tìm kiếm + bảng + CRUD + nhập Excel.
    • 7 trang Master Data (dùng CrudTablePanel): Bộ phận · Loại nhân viên ·
      Cấp bậc · Cost center · Vị trí · Mẫu mail · Khóa học.

Mỗi vị trí chỉ có ĐÚNG 1 mô tả công việc (JD) nên JD không còn trang riêng —
tiêu đề + file JD nhập ngay trong form của trang "Vị trí tuyển dụng".

Mẫu mail nằm ở bảng dùng chung `mail_templates` (trang "Mail templates"); mỗi vị
trí chỉ TRỎ tới 3 mẫu, tương ứng 3 vòng phỏng vấn (cv_schema.INTERVIEW_ROUNDS).
"""
import os
import re
import unicodedata
from pathlib import Path

from PySide6.QtCore import QDate, QDateTime, QTime, Qt
from PySide6.QtGui import QCursor, QTextDocument
from PySide6.QtWidgets import (
    QApplication, QCalendarWidget, QFileDialog, QFrame, QHBoxLayout,
    QLabel, QLineEdit, QToolTip, QVBoxLayout, QWidget,
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

try:
    import openpyxl  # noqa: F401
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

# Nhãn hiển thị cho hồ sơ chưa có trạng thái (cột status rỗng/NULL).
_NO_STATUS = "(no status yet)"

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
    "status":          130,   # đủ chỗ cho nhãn dài nhất ("Fail Probation Period")
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


def _strip_accents(s):
    """Bỏ dấu tiếng Việt: 'Tuấn Anh' → 'Tuan Anh' (xử lý riêng đ/Đ)."""
    s = s.replace("đ", "d").replace("Đ", "D")
    return "".join(c for c in unicodedata.normalize("NFD", s)
                   if unicodedata.category(c) != "Mn")


def _given_name(full_name):
    """Tên gọi (không dấu) từ họ tên đầy đủ — lấy TỪ CUỐI: 'Đinh Tuấn Anh' → 'Anh'."""
    parts = (full_name or "").strip().split()
    return _strip_accents(parts[-1]) if parts else ""


# Một placeholder có thể bị trình soạn thảo rich text cắt vụn bằng thẻ định dạng
# ('{<span style="font-weight:600;">possion</span>}' khi chỉ bôi đậm chữ bên
# trong), nên phần giữa hai dấu ngoặc chấp nhận cả thẻ HTML lẫn &nbsp;.
_PLACEHOLDER_RE = re.compile(r"\{((?:<[^>]+>|&nbsp;|[^<>{}])*)\}")
_TAG_RE = re.compile(r"<[^>]+>")


def _fill_template(text, mapping, escape=False):
    """Thay các placeholder {name}{possion}{position}{date}{time}{time_start}
    {time_end} trong `text` bằng giá trị thật.

    Placeholder bôi đậm/đổi màu một phần vẫn nhận đúng tên khóa, và định dạng đó
    được giữ lại cho giá trị thay vào. `escape=True` khi `text` là HTML — giá trị
    thay vào được escape để tên kiểu "R&D Engineer" không phá cấu trúc HTML.
    """
    if not text:
        return text or ""

    def _value(key):
        val = mapping[key]
        if escape:
            val = val.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return val

    def _replace(match):
        inner = match.group(1)
        key = _TAG_RE.sub("", inner).replace("&nbsp;", " ").strip()
        if key not in mapping:
            return match.group(0)
        # Giữ nguyên các thẻ định dạng nằm trong placeholder, đặt giá trị vào chỗ
        # đoạn chữ đầu tiên: '{<b>possion</b>}' → '<b>Sales Executive</b>'.
        parts, filled = [], False
        for token in re.split(r"(<[^>]+>)", inner):
            if _TAG_RE.fullmatch(token):
                parts.append(token)
            elif token and not filled:
                parts.append(_value(key))
                filled = True
        if not filled:
            parts.append(_value(key))
        return "".join(parts)

    return _PLACEHOLDER_RE.sub(_replace, text)


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
    """


def _dept_options():
    return {d["department_name"] or f"#{d['department_id']}": d["department_id"]
            for d in repo.list_departments()}


def _position_options():
    return {p["position_title"] or f"#{p['position_id']}": p["position_id"]
            for p in repo.list_positions()}


def _mail_template_options():
    """Mẫu mail → id, nhãn kèm loại cho dễ chọn: 'Interview Round 1 · Sales R1'."""
    opts = {}
    for t in repo.list_mail_templates():
        name = (t["name"] or f"#{t['mail_template_id']}").strip()
        kind = (t["type"] or "").strip()
        opts[f"{kind} · {name}" if kind else name] = t["mail_template_id"]
    return opts


def _course_type_options():
    """Tên loại khóa học → mã số lưu trong DB (inhouse=0, external=1, funded=2)."""
    return {name: i for i, name in enumerate(cv_schema.COURSE_TYPE_CHOICES)}


def _course_type_label(v):
    """Mã số course_type → nhãn hiển thị; giá trị lạ/rỗng → '—'."""
    try:
        return cv_schema.COURSE_TYPE_CHOICES[int(v)]
    except (ValueError, TypeError, IndexError):
        return "—"


def _text_preview(value, limit=60):
    """Rút gọn giá trị về MỘT dòng cho ô bảng (rỗng → '—').

    Nội dung mẫu mail lưu dạng HTML nên bóc lấy chữ thuần trước khi cắt.
    """
    s = str(value or "").strip()
    if not s:
        return "—"
    if s.startswith("<"):
        doc = QTextDocument()
        doc.setHtml(s)
        s = doc.toPlainText()
    s = " ".join(s.split())
    return s[:limit] + "…" if len(s) > limit else s


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


def _launch_file(parent, path):
    """Mở file bằng ứng dụng mặc định của hệ điều hành."""
    try:
        os.startfile(path)
    except AttributeError:
        import subprocess
        subprocess.Popen(["xdg-open", path])
    except Exception as exc:
        dialogs.error(parent, "Open error", f"Couldn't open the file:\n{exc}")


def _open_jd_link(panel, row, _key):
    """Click cột 'JD file' ở trang Positions → mở file JD của vị trí đó.

    File không còn ở đường dẫn đã lưu thì cho chọn lại, lưu lại vào DB rồi mở.
    """
    pid = row["position_id"]
    path = _txt(row, "jd_file_path")
    if path and os.path.isfile(path):
        _launch_file(panel, path)
        return
    if not dialogs.confirm(
            panel, "File not found",
            f"The JD file wasn't found at the saved path:\n{path}\n\n"
            "It may have been moved or renamed. Locate the file now?",
            ok_label="Choose file"):
        return
    new_path, _ = QFileDialog.getOpenFileName(
        panel, "Locate the JD file", "",
        "PDF/Word/Text (*.pdf *.doc *.docx *.txt);;All files (*.*)")
    if not new_path:
        return
    repo.update_position(pid, {"jd_file_path": new_path})
    panel.reload()
    _launch_file(panel, new_path)


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
            # Tên file JD hiển thị dạng link — bấm để mở file (như cột CV ở
            # màn hình ứng viên).
            "link_keys": {"jd_file_path"},
            "on_link": _open_jd_link,
            "columns": [
                ("position_id", "ID", 50),
                ("position_code", "Code", 90),
                ("position_title", "Position", 190),
                ("department_name", "Department", 140),
                ("level", "Level", 80),
                ("headcount", "Qty", 50),
                ("status", "Status", 105),
                # Ô rỗng (không phải "—") khi chưa gắn JD: cột này là link nên
                # chỉ ô có chữ mới được tô màu/gạch chân & bấm được.
                ("jd_file_path", "JD file", 160, "w",
                 lambda v: os.path.basename(str(v)) if v else ""),
                ("mail_template_r1_name", "Mail · round 1", 150, "w",
                 lambda v: _text_preview(v, 30)),
                ("mail_template_r2_name", "Mail · round 2", 150, "w",
                 lambda v: _text_preview(v, 30)),
                ("mail_template_r3_name", "Mail · round 3", 150, "w",
                 lambda v: _text_preview(v, 30)),
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
                # 3 vòng phỏng vấn → 3 mẫu mail lấy từ màn hình "Mail templates".
                {"kind": "section", "label": "Interview invite email templates"},
                *[{"key": col, "label": label, "kind": "dropdown",
                   "options": _mail_template_options}
                  for col, label, _ in cv_schema.INTERVIEW_ROUNDS],
            ],
        },
        "mail_template": {
            "title": "mail template", "pk": "mail_template_id",
            "modal_size": "md",   # form có ô soạn nội dung mail dài → cỡ md
            "list_fn": repo.list_mail_templates,
            "get": repo.get_mail_template, "insert": repo.insert_mail_template,
            "update": repo.update_mail_template, "delete": repo.delete_mail_template,
            "duplicate": repo.duplicate_mail_template,
            "columns": [
                ("mail_template_id", "ID", 50),
                ("name", "Template name", 200),
                ("type", "Type", 150),
                ("mail_subject", "Subject", 240, "w", _text_preview),
                ("mail_cc", "CC", 180, "w", _text_preview),
                ("mail_body", "Body", 260, "w", _text_preview),
            ],
            "form": [
                {"key": "name", "label": "Template name (*)",
                 "kind": "text", "required": True},
                {"key": "type", "label": "Type", "kind": "choice",
                 "choices": cv_schema.MAIL_TEMPLATE_TYPE_CHOICES, "allow_empty": True},
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
    description = "Open positions (with each position's JD & 3 interview-round email templates)."
    icon = "💼"
    order = 20
    spec_key = "position"


class MailTemplateTool(_MasterPageTool):
    name = "Mail templates"
    description = "Email templates (interview rounds, thank-you, notification…) — duplicate to reuse."
    icon = "✉"
    order = 22
    spec_key = "mail_template"


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
    description = "Search candidates, manage offer status, send offer emails."
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
        self.sel_status.set_options(cv_schema.CANDIDATE_STATUS_CHOICES)
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
        # Toolbar chia hai vùng: BÊN TRÁI là thao tác trên các hồ sơ đang tick,
        # BÊN PHẢI (sau stretch) là thao tác cấp trang. "Add" nằm bên phải, tông
        # neutral vì hồ sơ chủ yếu vào DB qua tool Quét CV bằng AI — nhập tay chỉ
        # còn là trường hợp lẻ (ứng viên walk-in, giới thiệu nội bộ).
        bar.addWidget(B(None, "View details", variant="info", icon="sparkles",
                        command=self._show_details))
        bar.addWidget(B(None, "Update status", variant="primary", icon="check",
                        command=self._bulk_status))
        bar.addWidget(B(None, "Send email", variant="info", icon="mail",
                        command=self._send_mail))
        self._btn_export = B(None, "Export to Excel", variant="warning", icon="save",
                             command=self._export_excel)
        bar.addWidget(self._btn_export)
        bar.addStretch(1)
        bar.addWidget(B(None, "Add", variant="neutral", icon="plus", command=self._add))
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

    # ------------------------------------------ đổi trạng thái hàng loạt
    # Tick NHIỀU ứng viên → tất cả phải đang CÙNG một trạng thái thì mới đổi
    # chung được; khác nhau thì báo lỗi kèm danh sách để người dùng tick lại.
    def _common_status(self, rows, what):
        """Trạng thái CHUNG của các dòng đã tick — None nếu chúng lệch nhau.

        Lệch thì báo lỗi kèm danh sách từng nhóm để người dùng tick lại; `what`
        là tên thao tác đang làm, ghép vào câu báo lỗi.
        """
        groups = {}
        for row in rows:
            label = _txt(row, "status").strip() or _NO_STATUS
            groups.setdefault(label, []).append(
                _txt(row, "full_name") or f"#{row['candidate_id']}")
        if len(groups) > 1:
            detail = "\n\n".join(f"{status}:\n• " + "\n• ".join(names)
                                 for status, names in groups.items())
            dialogs.error(
                self._root, "Statuses don't match",
                f"{what} only works when every ticked candidate is at the "
                "same status. Right now they are at different ones:\n\n" + detail)
            return None
        return next(iter(groups))

    def _bulk_status(self):
        rows = self.table.checked_rows()
        if not rows:
            dialogs.info(self._root, "Nothing selected",
                         "Tick at least one candidate in the table to update status.")
            return

        current = self._common_status(rows, "A bulk update")
        if current is None:
            return
        self._ask_bulk_status(rows, current)

    def _ask_bulk_status(self, rows, current):
        """Modal xác nhận: trạng thái hiện tại (cố định) → trạng thái mới (sửa được).

        Ô trạng thái mới điền sẵn bước kế tiếp trong luồng; hồ sơ đang ở điểm
        dừng thì để nguyên trạng thái hiện tại. Chỉ bấm OK mới ghi xuống DB.
        """
        dlg, card, lay = build_dialog_shell(
            self._root, f"Update status — {len(rows)} candidate(s)")

        line = QHBoxLayout()
        line.setSpacing(10)
        lbl = QLabel("Current status:")
        lbl.setObjectName("FieldLabel")
        line.addWidget(lbl)
        line.addWidget(_chip(card, current, theme.PALETTE["--info"]))
        line.addStretch(1)
        lay.addLayout(line)

        lbl_new = QLabel("Move to:")
        lbl_new.setObjectName("FieldLabel")
        lay.addWidget(lbl_new)
        combo = widgets.ComboBox(card)
        combo.addItems(cv_schema.CANDIDATE_STATUS_CHOICES)
        nxt = cv_schema.candidate_next_status(
            "" if current == _NO_STATUS else current)
        combo.setCurrentIndex(max(0, combo.findText(nxt or current)))
        lay.addWidget(combo)

        names = QLabel("• " + "\n• ".join(
            _txt(r, "full_name") or f"#{r['candidate_id']}" for r in rows))
        names.setObjectName("Hint")
        names.setWordWrap(True)
        sa = widgets.scroll_area(names)
        sa.setMaximumHeight(dlg.modal_h)   # danh sách dài thì cuộn, ngắn thì vừa khít
        lay.addWidget(sa, 1)

        def do_update():
            status = combo.currentText()
            failed = []
            for row in rows:
                cid = row["candidate_id"]
                try:
                    repo.update_candidate(cid, {"status": status})
                except Exception as exc:   # noqa: BLE001 — gom lỗi báo một lần
                    failed.append(f"#{cid} — {exc}")
            dlg.accept()
            self._reload()
            if failed:
                dialogs.error(self._root, "Update failed",
                              "Couldn't update:\n• " + "\n• ".join(failed))
            else:
                dialogs.success(
                    self._root, "Status updated",
                    f"Moved {len(rows)} candidate(s) to \"{status}\".")

        foot = QHBoxLayout()
        foot.addWidget(widgets.button(card, "OK", variant="primary", icon="check",
                                      command=do_update))
        foot.addWidget(widgets.button(card, "Cancel", variant="neutral", icon="x",
                                      command=dlg.reject))
        foot.addStretch(1)
        lay.addLayout(foot)
        dlg.exec()

    # ------------------------------------------------- gửi mail cho ứng viên
    # Tick MỘT HOẶC NHIỀU ứng viên → chọn loại mail muốn gửi:
    #   • VÒNG phỏng vấn (1/2/3) → chọn ngày giờ cho từng người → mở bấy nhiêu cửa
    #     sổ MEETING của Outlook đã điền sẵn (nội dung lấy từ MẪU MAIL mà vị trí
    #     ứng tuyển gán cho vòng đó). Người dùng duyệt rồi bấm Send — Outlook vừa
    #     gửi mail mời, vừa tạo lịch, vừa đặt phòng.
    #   • Thư CẢM ƠN ĐÃ ỨNG TUYỂN → mẫu chọn thẳng trong modal (không gắn theo vị
    #     trí), mở cửa sổ MAIL THƯỜNG: không hỏi giờ, không tạo lịch.
    # Cả hai đều thay placeholder {name}{possion}{date}{time} trước khi mở cửa sổ.
    def _send_mail(self):
        if not outlook.available():
            dialogs.warning(self._root, "Outlook required",
                            "Sending email needs Outlook on Windows (pywin32).")
            return
        rows = self.table.checked_rows()
        if not rows:
            dialogs.warning(
                self._root, "Nothing selected",
                "Tick at least one candidate in the table to send an email.")
            return

        # Cùng một lượt gửi thì các hồ sơ phải đang ở CÙNG trạng thái — trạng thái
        # đó cho biết đang mời vòng nào, hiện luôn trong hộp chọn loại mail.
        current = self._common_status(rows, "Sending emails")
        if current is None:
            return

        choice = self._pick_mail_kind(len(rows), current)
        if choice is None:
            return
        kind, value = choice
        if kind == "thank_you":
            self._send_thank_you(rows, repo.get_mail_template(value))
            return
        tpl_col, round_label, after_status = cv_schema.INTERVIEW_ROUNDS[value]

        # Loại trước các ứng viên không đủ dữ liệu để soạn thư mời.
        jobs, skipped = [], []
        for row in rows:
            who = _txt(row, "full_name") or f"#{row['candidate_id']}"
            email = _txt(row, "email")
            if not email:
                skipped.append(f"{who} — no email address")
                continue
            pos_id = row["position_id"] if "position_id" in row.keys() else None
            pos = repo.get_position(pos_id) if pos_id else None
            if pos is None:
                skipped.append(f"{who} — no position, so no email template")
                continue
            tpl_id = pos[tpl_col] if tpl_col in pos.keys() else None
            tpl = repo.get_mail_template(tpl_id) if tpl_id else None
            if tpl is None:
                skipped.append(
                    f"{who} — position \"{_txt(pos, 'position_title')}\" has no "
                    f"{round_label} email template")
                continue
            jobs.append((row, pos, tpl, email, who))
        if not jobs:
            dialogs.warning(self._root, "Nothing to send",
                            "None of the selected candidates can be invited:\n• "
                            + "\n• ".join(skipped))
            return

        # Chọn giờ cho từng ứng viên trước, mở cửa sổ sau — người dùng chọn xong
        # một lượt rồi mới phải làm việc với Outlook.
        picks, previous = [], None
        for idx, (row, pos, tpl, email, who) in enumerate(jobs, 1):
            label = who if len(jobs) == 1 else f"{who}  ({idx}/{len(jobs)})"
            picked = self._pick_datetime(label, previous, batch=len(jobs) > 1)
            if picked is None:
                skipped.append(f"{who} — skipped")
                continue
            picks.append((row, pos, tpl, email, who, picked))
            previous = picked
        if not picks:
            return

        opened, failed, no_cv = [], [], []
        for row, pos, tpl, email, who, (start, end) in picks:
            subject, body_html = self._meeting_content(row, pos, tpl, start, end)
            cv_path = _txt(row, "cv_file_path")
            if cv_path and os.path.isfile(cv_path):
                attachments = [cv_path]
            else:
                attachments = []
                no_cv.append(f"{who} — " + ("CV file not found: " + cv_path
                                            if cv_path else "no CV file on record"))
            err = self._open_meeting(email, _txt(tpl, "mail_cc"), subject, body_html,
                                     start.toPython(), end.toPython(), attachments)
            if err is None:
                opened.append((row, who))
            else:
                failed.append(f"{who} — {err}")
        self._report_meetings(opened, failed, skipped, no_cv, after_status)

    def _pick_mail_kind(self, count, current=""):
        """Hỏi loại mail muốn gửi: 3 vòng phỏng vấn, hoặc thư cảm ơn đã ứng tuyển.

        Hiện luôn trạng thái hiện tại của các hồ sơ đã tick (đều giống nhau — xem
        _common_status) và chọn sẵn vòng suy ra từ trạng thái đó (Short List →
        vòng 1, First Interview → vòng 2…, xem cv_schema.interview_round_for_status);
        vẫn đổi được.

        Mẫu của 3 vòng lấy theo VỊ TRÍ ứng tuyển, còn thư cảm ơn không gắn với vị
        trí nên chọn thẳng mẫu ở ô bên dưới (chỉ hiện khi chọn loại này).

        Trả về ("round", chỉ số trong INTERVIEW_ROUNDS) / ("thank_you", id mẫu
        mail), hoặc None nếu hủy.
        """
        dlg, card, lay = build_dialog_shell(
            self._root, f"Send email — {count} candidate(s)", size="sm")

        if current:
            line = QHBoxLayout()
            line.setSpacing(10)
            lbl = QLabel("Current status:")
            lbl.setObjectName("FieldLabel")
            line.addWidget(lbl)
            line.addWidget(_chip(card, current, theme.PALETTE["--info"]))
            line.addStretch(1)
            lay.addLayout(line)

        note = QLabel("Pick the email to send:")
        note.setObjectName("DialogMsg")
        note.setWordWrap(True)
        lay.addWidget(note)

        thank_you = cv_schema.MAIL_TEMPLATE_TYPE_THANK_YOU
        combo = widgets.ComboBox(card)
        combo.addItems([label for _, label, _ in cv_schema.INTERVIEW_ROUNDS]
                       + [thank_you])
        combo.setCurrentIndex(cv_schema.interview_round_for_status(current))
        lay.addWidget(combo)

        # Ô chọn mẫu thư cảm ơn — chỉ hiện khi chọn loại đó ở ô trên.
        tpl_opts = {(t["name"] or f"#{t['mail_template_id']}").strip():
                    t["mail_template_id"]
                    for t in repo.list_mail_templates(thank_you)}
        tpl_lbl = QLabel("Template:")
        tpl_lbl.setObjectName("FieldLabel")
        tpl_combo = widgets.ComboBox(card)
        tpl_combo.addItems(list(tpl_opts) or ["(no template of this type yet)"])
        tpl_combo.setEnabled(bool(tpl_opts))
        lay.addWidget(tpl_lbl)
        lay.addWidget(tpl_combo)

        def _sync(idx):
            picked_thank_you = idx == combo.count() - 1
            tpl_lbl.setVisible(picked_thank_you)
            tpl_combo.setVisible(picked_thank_you)

        combo.currentIndexChanged.connect(_sync)
        _sync(combo.currentIndex())

        result = {"val": None}

        def _confirm():
            idx = combo.currentIndex()
            if idx < len(cv_schema.INTERVIEW_ROUNDS):
                result["val"] = ("round", idx)
            elif not tpl_opts:
                dialogs.warning(
                    dlg, "No template",
                    f'There is no "{thank_you}" mail template yet. '
                    "Create one on the Mail templates page first.")
                return
            else:
                result["val"] = ("thank_you", tpl_opts[tpl_combo.currentText()])
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

    # ------------------------------------------------- thư cảm ơn đã ứng tuyển
    def _send_thank_you(self, rows, tpl):
        """Mở cửa sổ MAIL THƯỜNG của Outlook cho từng ứng viên đã tick.

        Không hỏi giờ, không tạo lịch, không đính kèm CV — chỉ điền sẵn người
        nhận / CC / tiêu đề / nội dung từ mẫu rồi để người dùng bấm Send. Trạng
        thái ứng viên GIỮ NGUYÊN (thư cảm ơn không đổi giai đoạn tuyển dụng).
        """
        if tpl is None:
            dialogs.error(self._root, "Template missing",
                          "That mail template no longer exists.")
            return
        opened, failed, skipped = [], [], []
        for row in rows:
            who = _txt(row, "full_name") or f"#{row['candidate_id']}"
            email = _txt(row, "email")
            if not email:
                skipped.append(f"{who} — no email address")
                continue
            subject, body_html = self._mail_content(row, tpl)
            try:
                outlook.create_mail(email, subject, cc=_txt(tpl, "mail_cc"),
                                    html=body_html)
            except Exception as exc:   # noqa: BLE001 — gom lỗi báo một lần
                failed.append(f"{who} — {exc}")
                continue
            opened.append(who)

        lines = []
        if failed:
            lines.append("Couldn't open a draft for:\n• " + "\n• ".join(failed))
        if skipped:
            lines.append("Not emailed:\n• " + "\n• ".join(skipped))
        if lines:
            report = dialogs.error if failed else dialogs.warning
            report(self._root, "Some emails were left out", "\n\n".join(lines))
        if opened:
            dialogs.success(
                self._root, "Drafts opened",
                f"{len(opened)} email draft(s) opened in Outlook — review each "
                "window then press Send there.\n\n• " + "\n• ".join(opened))

    @staticmethod
    def _mail_content(row, tpl):
        """Tiêu đề + nội dung HTML của mail thường (không có ngày/giờ phỏng vấn).

        Chỉ điền các placeholder liên quan tới ứng viên; {date}/{time…} nếu lỡ có
        trong mẫu thì GIỮ NGUYÊN để người dùng nhìn thấy mà sửa trước khi gửi.
        """
        pos_id = row["position_id"] if "position_id" in row.keys() else None
        pos = repo.get_position(pos_id) if pos_id else None
        title = _txt(pos, "position_title") if pos is not None else ""
        mapping = {
            "name":     _given_name(_txt(row, "full_name")),
            "possion":  title,
            "position": title,
        }
        body = tpl["mail_body"] if "mail_body" in tpl.keys() and tpl["mail_body"] else ""
        return (_fill_template(_txt(tpl, "mail_subject"), mapping),
                _fill_template(body, mapping, escape=True))

    @staticmethod
    def _meeting_content(row, pos, tpl, start, end):
        """Tiêu đề + nội dung HTML của thư mời, điền từ mẫu mail của vòng đang mời."""
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
        body = tpl["mail_body"] if "mail_body" in tpl.keys() and tpl["mail_body"] else ""
        return (_fill_template(_txt(tpl, "mail_subject"), mapping),
                _fill_template(body, mapping, escape=True))

    def _report_meetings(self, opened, failed, skipped, no_cv=(), after_status=""):
        """Kết thúc một lượt gửi: báo trước các trường hợp lỗi/bỏ qua/thiếu CV
        (nếu có), rồi mời cập nhật trạng thái cho những ứng viên đã mở thư mời."""
        lines = []
        if failed:
            lines.append("Couldn't open a meeting for:\n• " + "\n• ".join(failed))
        if skipped:
            lines.append("Not invited:\n• " + "\n• ".join(skipped))
        if no_cv:
            lines.append("Invite opened without a CV attached:\n• " + "\n• ".join(no_cv))
        if lines:
            if failed:
                dialogs.error(self._root, "Some invites failed", "\n\n".join(lines))
            else:
                title = "Some candidates were left out" if skipped else "Missing CV file"
                dialogs.warning(self._root, title, "\n\n".join(lines))
        if opened:
            self._ask_status_update(opened, after_status)

    def _ask_status_update(self, entries, after_status=""):
        """Modal cập nhật trạng thái cho các ứng viên vừa mở thư mời.

        Điền sẵn trạng thái ứng với VÒNG vừa mời (`after_status` — xem
        cv_schema.INTERVIEW_ROUNDS) nhưng đổi được từng người. Chỉ bấm Update mới
        ghi xuống DB; đóng modal thì trạng thái giữ nguyên như cũ. `entries` là
        list (row, tên hiển thị).
        """
        dlg, card, lay = build_dialog_shell(self._root, "Update candidate status")

        note = QLabel(
            f"{len(entries)} meeting invite(s) opened in Outlook — review each window, "
            "add a meeting room if you need one, then press Send there.\n\n"
            "New status for these candidates:")
        note.setObjectName("DialogMsg")
        note.setWordWrap(True)
        lay.addWidget(note)

        body = QWidget()
        col = QVBoxLayout(body)
        col.setContentsMargins(0, 0, 8, 0)
        col.setSpacing(8)
        selects = []
        for row, who in entries:
            line = QHBoxLayout()
            line.setSpacing(10)
            name = QLabel(who)
            name.setObjectName("FieldLabel")
            name.setWordWrap(True)
            line.addWidget(name, 1)
            combo = widgets.ComboBox(body)
            combo.addItems(cv_schema.CANDIDATE_STATUS_CHOICES)
            combo.setCurrentIndex(max(0, combo.findText(after_status)))
            line.addWidget(combo, 1)
            col.addLayout(line)
            selects.append((row["candidate_id"], combo))
        col.addStretch(1)
        sa = widgets.scroll_area(body)
        sa.setMaximumHeight(dlg.modal_h)   # danh sách dài thì cuộn, ngắn thì vừa khít
        lay.addWidget(sa, 1)

        def do_update():
            failed = []
            for cid, combo in selects:
                try:
                    repo.update_candidate(cid, {"status": combo.currentText()})
                except Exception as exc:   # noqa: BLE001 — gom lỗi báo một lần
                    failed.append(f"#{cid} — {exc}")
            dlg.accept()
            self._reload()
            if failed:
                dialogs.error(self._root, "Update failed",
                              "Couldn't update:\n• " + "\n• ".join(failed))
            else:
                dialogs.success(self._root, "Status updated",
                                f"Updated {len(selects)} candidate(s).")

        foot = QHBoxLayout()
        foot.addWidget(widgets.button(card, "Update", variant="primary", icon="check",
                                      command=do_update))
        foot.addWidget(widgets.button(card, "Keep current status", variant="neutral",
                                      icon="x", command=dlg.reject))
        foot.addStretch(1)
        lay.addLayout(foot)
        dlg.exec()

    def _pick_datetime(self, who="", previous=None, batch=False):
        """Hộp thoại chọn NGÀY + GIỜ bắt đầu/kết thúc phỏng vấn của MỘT ứng viên.

        Dùng lịch INLINE (không phải popup) có style riêng — tránh lỗi hiển thị
        do QSS bảng toàn cục & popup trong modal frameless. Trả về (start, end)
        dạng QDateTime, hoặc None nếu bỏ qua.

        `who` hiện trên tiêu đề để biết đang xếp lịch cho ai. `previous` là lựa
        chọn của ứng viên liền trước — mặc định của ứng viên này nối tiếp ngay
        sau đó, cùng ngày và cùng độ dài, vì phỏng vấn thường xếp liền nhau.
        `batch` đổi nút hủy thành "Skip" cho rõ là chỉ bỏ qua một người.
        """
        title = f"Interview time — {who}" if who else "Pick interview date & time"
        dlg, card, lay = build_dialog_shell(self._root, title, size="sm")
        lbl = QLabel("Interview date:")
        lbl.setObjectName("FieldLabel")
        lay.addWidget(lbl)

        day = QDate.currentDate().addDays(1)          # mặc định: ngày mai
        first, second = QTime(9, 0), QTime(9, 30)
        if previous is not None:
            prev_start, prev_end = previous
            day = prev_end.date()
            first = prev_end.time()
            second = first.addSecs(max(1800, prev_start.secsTo(prev_end)))

        cal = QCalendarWidget(card)
        cal.setStyleSheet(_picker_qss())
        cal.setGridVisible(False)
        cal.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)   # bỏ cột số tuần
        cal.setHorizontalHeaderFormat(QCalendarWidget.ShortDayNames)
        cal.setNavigationBarVisible(True)
        cal.setFirstDayOfWeek(Qt.Monday)
        cal.setMinimumDate(QDate.currentDate())
        cal.setSelectedDate(day)
        cal.setMinimumHeight(260)
        lay.addWidget(cal)

        # Hàng chọn giờ bắt đầu / kết thúc.
        times = QHBoxLayout()
        times.setSpacing(12)

        def _time_slots():
            slots = []
            t = QTime(8, 0)
            end = QTime(20, 0)
            while t <= end:
                slots.append(QTime(t))
                t = t.addSecs(30 * 60)
            return slots

        def _time_col(label, default):
            col = QVBoxLayout(); col.setSpacing(4)
            cap = QLabel(label); cap.setObjectName("FieldLabel")
            col.addWidget(cap)
            combo = widgets.ComboBox(card)
            for t in _time_slots():
                combo.addItem(t.toString("HH:mm"), t)
            idx = combo.findText(default.toString("HH:mm"))
            if idx < 0:      # giờ gợi ý rơi ngoài khung 08:00–20:00
                idx = combo.count() - 1 if default > QTime(20, 0) else 0
            combo.setCurrentIndex(idx)
            col.addWidget(combo)
            times.addLayout(col, 1)
            return combo

        start_te = _time_col("Start time", first)
        end_te = _time_col("End time", second)
        lay.addLayout(times)

        result = {"val": None}

        def _confirm():
            d = cal.selectedDate()
            start = QDateTime(d, start_te.currentData())
            end = QDateTime(d, end_te.currentData())
            if end <= start:
                dialogs.warning(dlg, "Invalid time",
                                "End time must be after start time.")
                return
            result["val"] = (start, end)
            dlg.accept()

        foot = QHBoxLayout()
        foot.addWidget(widgets.button(card, "Continue", variant="primary",
                                      icon="check", command=_confirm))
        foot.addWidget(widgets.button(card, "Skip" if batch else "Cancel",
                                      variant="neutral", icon="x", command=dlg.reject))
        foot.addStretch(1)
        lay.addLayout(foot)
        dlg.exec()
        return result["val"]

    @staticmethod
    def _open_meeting(to, cc, subject, body_html, start, end, attachments=()):
        """Mở cửa sổ meeting của Outlook đã điền sẵn ứng viên, giờ, nội dung và
        file đính kèm (CV của ứng viên).

        Nội dung mẫu là rich text nên truyền dạng HTML để giữ định dạng; bản
        thuần chỉ dùng khi Outlook không chèn được HTML. CC của mẫu mail thành
        người tham dự tùy chọn (meeting không có CC). Trả về None nếu mở được,
        hoặc chuỗi lỗi để bên gọi gộp vào bảng tổng kết cuối lượt.
        """
        doc = QTextDocument()
        doc.setHtml(body_html)
        try:
            outlook.create_meeting(subject, start, end, to, optional=cc,
                                   html=body_html, body=doc.toPlainText(),
                                   attachments=list(attachments))
        except Exception as exc:   # noqa: BLE001 — gom lỗi lại để báo một lần
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
             "choices": cv_schema.CANDIDATE_STATUS_CHOICES},
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
        _launch_file(self._root, path)
