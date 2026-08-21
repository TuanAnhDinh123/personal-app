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

from PySide6.QtCore import QDate, QDateTime, QTime, Qt
from PySide6.QtGui import QCursor, QTextDocument
from PySide6.QtWidgets import (
    QApplication, QFileDialog, QFrame, QHBoxLayout,
    QLabel, QLineEdit, QToolTip, QVBoxLayout, QWidget,
)

from app.core import candidate_export
from app.core import cv_repository as repo
from app.core import cv_schema
from app.core import outlook
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
    "current_title":   170,
    "years_experience": 60,
    "email":           210,
    "phone":           120,
    "position_title":  120,
    "ai_score":        60,
    "cv_file_path":    150,
    "source":          110,
    "department_name": 130,
    "batch":           70,
    "status":          130,   # đủ chỗ cho nhãn dài nhất ("Fail Probation Period")
    "final_status":    120,
    "pool_status":     110,
    "applied_at":      105,
    "note":            180,
}

_W = CAND_COL_WIDTHS


def _short(value, limit=60):
    """Rút gọn về một dòng cho ô bảng (rỗng → chuỗi rỗng)."""
    s = str(value or "").replace("\n", " ").strip()
    return (s[:limit] + "…") if len(s) > limit else s


def _source_cell(value):
    """Cột Source = SÀN cung cấp CV. Hồ sơ do tool quét AI nạp vào mang dấu
    `CANDIDATE_SOURCE_AUTO` — đó là đường vào app chứ chưa phải sàn, nên hiện
    dấu gạch để thấy ngay dòng nào còn phải điền tay (chuột phải → Update source).
    """
    text = str(value or "").strip()
    return "" if not text or text == cv_schema.CANDIDATE_SOURCE_AUTO else text


def _years_cell(value):
    """Số năm kinh nghiệm TẠI THỜI ĐIỂM CV — con số hôm nay xem ở màn hình chi tiết."""
    try:
        return f"{float(value):.1f}"
    except (TypeError, ValueError):
        return ""


# Cột bảng ỨNG VIÊN: (khóa, tiêu đề, rộng, canh lề[, formatter]).
# `status` / `position_title` / `ai_score` lấy từ ĐƠN ỨNG TUYỂN MỚI NHẤT và
# lượt AI chấm mới nhất của đơn đó (xem cv_repository._CANDIDATE_SELECT).
_CAND_COLUMNS = [
    ("candidate_id",     "ID",            _W["candidate_id"],     "center"),
    ("full_name",        "Full name",     _W["full_name"],        "w"),
    ("current_title",    "Current title", _W["current_title"],    "w"),
    ("years_experience", "Yrs",           _W["years_experience"], "center", _years_cell),
    ("email",            "Email",         _W["email"],            "w"),
    ("phone",            "Phone",         _W["phone"],            "w"),
    ("position_title",   "Applied for",   _W["position_title"],   "w"),
    ("status",           "Status",        _W["status"],           "center"),
    ("final_status",     "Result",        _W["final_status"],     "center"),
    ("ai_score",         "Score",         _W["ai_score"],         "center",
     lambda v: "" if v in (None, "") else str(int(float(v)))),
    ("cv_file_path",     "CV",            _W["cv_file_path"],     "w",
     lambda v: os.path.basename(str(v)) if v else ""),
    ("source",           "Source",        _W["source"],           "w", _source_cell),
    ("department_name",  "Department",    _W["department_name"],  "w"),
    ("pool_status",      "Pool",          _W["pool_status"],      "center"),
    ("batch",            "Batch",         _W["batch"],            "center"),
    ("applied_at",       "Applied",       _W["applied_at"],       "center",
     lambda v: str(v)[:10] if v else ""),
    ("note",             "Note",          _W["note"],             "w", _short),
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
    """Đọc row[key] an toàn → chuỗi đã strip. Cả `row` lẫn giá trị None đều ra ''.

    Nhận cả `row=None` để form nhập dùng chung một đường cho bản ghi đã có và
    bản ghi chưa tồn tại (vòng phỏng vấn chưa diễn ra → mọi ô rỗng).
    """
    if row is None:
        return ""
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


def _int_str(value):
    """Điểm số về dạng số nguyên để hiển thị ('82.0' → '82')."""
    try:
        return str(int(float(value)))
    except (TypeError, ValueError):
        return str(value or "")


def _result_color(value):
    """Màu cho kết luận Pass / Fail / Consideration."""
    return {"pass": theme.PALETTE["--success"],
            "fail": theme.PALETTE["--danger"],
            "consideration": theme.PALETTE["--warning"]}.get(
        str(value or "").strip().lower(), theme.PALETTE["--text-muted"])


def _experience_label(years):
    """Kinh nghiệm hiển thị HAI con số (xem cv_repository.experience_years).

    Con số thứ nhất chắc chắn đúng vì CV nói vậy; con số thứ hai chỉ là ước
    tính đến hôm nay nên có dấu ≈. Bằng nhau (CV vừa nhận) thì chỉ hiện một.
    """
    at_cv, today = years["at_cv"], years["today"]
    as_of = years["as_of"][:7]
    text = f"{at_cv:.1f} yrs"
    if as_of:
        text += f" (CV {as_of})"
    if abs(today - at_cv) >= 0.1:
        text += f"  ·  ≈ {today:.1f} yrs today"
    return text


def _section_box(parent, title, icon):
    """Khung con trong thẻ chi tiết: tiêu đề + icon. Trả về (khung, layout)."""
    box = QFrame(parent)
    box.setObjectName("AIBox")
    v = QVBoxLayout(box)
    v.setContentsMargins(14, 12, 14, 12)
    v.setSpacing(6)
    head = QHBoxLayout()
    head.setSpacing(6)
    ico = QLabel(box)
    ico.setPixmap(widgets.svg_pixmap(icon, theme.PALETTE["--accent"], 16))
    head.addWidget(ico, 0, Qt.AlignVCenter)
    lbl = QLabel(title, box)
    lbl.setObjectName("AIHeader")
    head.addWidget(lbl, 1)
    v.addLayout(head)
    return box, v


def _muted(parent, text):
    """Dòng chữ nhỏ, màu nhạt (dùng cho lịch sử & chú thích)."""
    lbl = QLabel(text, parent)
    lbl.setObjectName("AIEmpty")
    lbl.setWordWrap(True)
    lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
    return lbl


def _divider(parent):
    """Đường kẻ ngang mảnh ngăn các mục trong cùng một khung."""
    line = QFrame(parent)
    line.setFrameShape(QFrame.HLine)
    line.setStyleSheet(f"color: {theme.PALETTE['--border']}; max-height: 1px;")
    return line


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
                ("jrf_code", "JRF", 80),
                ("position_title", "Position", 190),
                ("department_name", "Department", 140),
                ("level", "Level", 80),
                ("headcount", "Qty", 50),
                ("status", "Status", 105),
                ("starting_date", "Start by", 100),
                ("salary_level", "Salary level", 120),
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
                {"key": "jrf_code", "label": "JRF code (job requisition)", "kind": "text"},
                {"key": "position_title", "label": "Position title (*)",
                 "kind": "text", "required": True},
                {"key": "level", "label": "Level", "kind": "text"},
                {"key": "headcount", "label": "Headcount", "kind": "int"},
                {"key": "status", "label": "Status", "kind": "choice",
                 "choices": cv_schema.POSITION_STATUS_CHOICES, "allow_empty": True},
                {"key": "starting_date", "label": "Needed by (yyyy-mm-dd)", "kind": "text"},
                {"key": "salary_level", "label": "Approved salary level", "kind": "text"},
                {"key": "required_experience", "label": "Required experience",
                 "kind": "text"},
                {"key": "description", "label": "Description", "kind": "textarea",
                 "height": 3},
                {"key": "note", "label": "Note", "kind": "textarea", "height": 3},
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
        # Danh mục kỹ năng chuẩn hóa: `aliases` là chỗ gom các cách viết khác
        # nhau để CV ghi "JS" vẫn khớp với JD đòi "JavaScript".
        "skill": {
            "title": "skill", "pk": "skill_id",
            "list_fn": repo.list_skills,
            "get": repo.get_skill, "insert": repo.insert_skill,
            "update": repo.update_skill, "delete": repo.delete_skill,
            "columns": [
                ("skill_id", "ID", 50),
                ("name", "Skill", 180),
                ("category", "Category", 120),
                ("aliases", "Also written as", 260),
                ("description", "Description", 240),
            ],
            "form": [
                {"key": "name", "label": "Skill name (*) — e.g. JavaScript",
                 "kind": "text", "required": True},
                {"key": "category", "label": "Category", "kind": "choice",
                 "choices": cv_schema.SKILL_CATEGORY_CHOICES, "allow_empty": True},
                {"key": "aliases", "label": "Also written as (separate with ;) — "
                 "e.g. JS; ECMAScript", "kind": "text"},
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


class SkillTool(_MasterPageTool):
    name = "Skills"
    description = ("Standardised skill list — aliases let a CV that says \"JS\" "
                   "match a JD that asks for \"JavaScript\".")
    icon = "🧩"
    order = 24
    spec_key = "skill"


class CourseTool(_MasterPageTool):
    name = "Courses"
    description = "Training / course directory."
    icon = "🎓"
    order = 40
    spec_key = "course"


# ═════════════════════ MODAL XEM CHI TIẾT ỨNG VIÊN ══════════════════════
class _CandidateDetailDialog(ModalDialog):
    """Modal lớn xem chi tiết các ứng viên đang tick.

    Mỗi ứng viên là một thẻ gồm 5 khối, đọc từ trên xuống là đủ hiểu hồ sơ:
        hồ sơ nghề nghiệp → LỊCH SỬ AI CHẤM (mọi lượt, không chỉ lượt cuối)
        → các vòng phỏng vấn kèm nhận xét → lịch sử liên hệ → các đơn ứng tuyển.
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
        cid = row["candidate_id"]
        box = QFrame(parent)
        box.setObjectName("DetailCard")
        v = QVBoxLayout(box)
        v.setContentsMargins(16, 14, 16, 14)
        v.setSpacing(8)

        # Hàng tiêu đề: tên + chip điểm / trạng thái / tình trạng trong pool
        head = QHBoxLayout()
        head.setSpacing(8)
        name = QLabel(f"#{cid}  {_txt(row, 'full_name') or '(no name)'}", box)
        name.setObjectName("DetailName")
        head.addWidget(name, 1)
        score = _txt(row, "ai_score")
        if score:
            head.addWidget(_chip(box, f"Score {_int_str(score)}", _score_color(score)))
        for key, color in (("status", "--info"), ("final_status", "--accent"),
                           ("pool_status", "--text-muted")):
            value = _txt(row, key)
            if value:
                head.addWidget(_chip(box, value, theme.PALETTE[color]))
        v.addLayout(head)

        # Hàng thông tin phụ (bôi-chọn được để copy tay nếu cần)
        meta = " · ".join(p for p in (
            _txt(row, "current_title"), _txt(row, "industry"),
            _txt(row, "position_title"), _txt(row, "department_name"),
            (f"DOB: {_txt(row, 'date_of_birth')}" if _txt(row, "date_of_birth") else ""),
            (f"Applied: {_txt(row, 'applied_at')[:10]}" if _txt(row, "applied_at") else ""),
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

        v.addWidget(self._profile_box(box, row))
        v.addWidget(self._ai_box(box, cid))
        v.addWidget(self._interview_box(box, cid))
        v.addWidget(self._history_box(box, cid))

        note = _txt(row, "note")
        if note:
            v.addLayout(self._para(box, "Note", note))
        return box

    # ------------------------------------------------------- hồ sơ nghề nghiệp
    def _profile_box(self, parent, row):
        """Kinh nghiệm (hai con số), kỹ năng, tóm tắt — phần KHÔNG dính JD nào."""
        cid = row["candidate_id"]
        box, v = _section_box(parent, "Profile", "idcard")

        years = repo.experience_years(cid, row)
        line = QHBoxLayout()
        line.setSpacing(8)
        line.addWidget(_chip(box, _experience_label(years), theme.PALETTE["--accent"]))
        if years["stale"]:
            line.addWidget(_chip(box, "Stale profile", theme.PALETTE["--warning"]))
        line.addStretch(1)
        v.addLayout(line)

        facts = " · ".join(p for p in (
            _txt(row, "education"), _txt(row, "major"), _txt(row, "languages"),
            _txt(row, "city"),
            (f"Expects {_txt(row, 'salary_note')}" if _txt(row, "salary_note") else ""),
            (f"Available {_txt(row, 'available_from')}"
             if _txt(row, "available_from") else ""),
        ) if p)
        if facts:
            v.addLayout(self._para(box, "Details", facts))
        if _txt(row, "profile_summary"):
            v.addLayout(self._para(box, "Summary", _txt(row, "profile_summary")))
        if _txt(row, "skills_text"):
            v.addLayout(self._para(box, "Skills", _txt(row, "skills_text")))

        jobs = repo.list_candidate_experiences(cid)
        if jobs:
            lines = []
            for j in jobs:
                span = f"{_txt(j, 'start_date')} → {_txt(j, 'end_date') or 'now'}"
                lines.append(f"• {_txt(j, 'job_title') or '(role)'} — "
                             f"{_txt(j, 'company') or '(company)'}  ({span})")
            v.addLayout(self._para(box, "Work history", "\n".join(lines)))
        return box

    # ----------------------------------------------------- lịch sử AI chấm điểm
    def _ai_box(self, parent, candidate_id):
        """MỌI lượt AI đã chấm, mới nhất trước — không ghi đè nên xem được cả quá trình."""
        box, v = _section_box(parent, "AI assessments", "sparkles")
        rows = repo.list_evaluations(candidate_id)
        if not rows:
            v.addWidget(_muted(box, "No AI assessment for this candidate yet."))
            return box

        for i, ev in enumerate(rows):
            if i:
                v.addWidget(_divider(box))
            head = QHBoxLayout()
            head.setSpacing(8)
            title = " · ".join(p for p in (
                _txt(ev, "position_title") or "(no position)",
                _txt(ev, "source"),
                _txt(ev, "evaluated_at")[:16],
            ) if p)
            lbl = QLabel(title, box)
            lbl.setObjectName("AILabel")
            head.addWidget(lbl, 1)
            score = _txt(ev, "ai_score")
            if score:
                head.addWidget(_chip(box, _int_str(score), _score_color(score)))
            v.addLayout(head)

            sub = " · ".join(p for p in (
                (f"model {_txt(ev, 'model')}" if _txt(ev, "model") else ""),
                (f"CV of {_txt(ev, 'cv_received_at')}"
                 if _txt(ev, "cv_received_at") else ""),
            ) if p)
            if sub:
                v.addWidget(_muted(box, sub))
            if _txt(ev, "summary"):
                v.addLayout(self._para(box, "Fit summary", _txt(ev, "summary")))
            if _txt(ev, "matched_skills") or _txt(ev, "missing_skills"):
                two = QHBoxLayout()
                two.setSpacing(12)
                two.addLayout(self._para(box, "Matched skills",
                                         _txt(ev, "matched_skills") or "—"), 1)
                two.addLayout(self._para(box, "Missing skills",
                                         _txt(ev, "missing_skills") or "—"), 1)
                v.addLayout(two)
            if _txt(ev, "strengths") or _txt(ev, "weaknesses"):
                two = QHBoxLayout()
                two.setSpacing(12)
                two.addLayout(self._para(box, "Strengths", _txt(ev, "strengths") or "—"), 1)
                two.addLayout(self._para(box, "Weaknesses", _txt(ev, "weaknesses") or "—"), 1)
                v.addLayout(two)
        return box

    # ------------------------------------------------------- các vòng phỏng vấn
    def _interview_box(self, parent, candidate_id):
        """Từng vòng: kết luận chung + nhận xét của từng người phỏng vấn."""
        box, v = _section_box(parent, "Interviews", "users")
        rows = repo.list_interviews(candidate_id=candidate_id)
        if not rows:
            v.addWidget(_muted(box, "No interview recorded yet."))
            return box

        for i, iv in enumerate(rows):
            if i:
                v.addWidget(_divider(box))
            head = QHBoxLayout()
            head.setSpacing(8)
            title = f"Round {_txt(iv, 'round') or '?'}"
            when = _txt(iv, "interview_date")[:16]
            if when:
                title += f" · {when}"
            for extra in (_txt(iv, "mode"), _txt(iv, "position_title")):
                if extra:
                    title += f" · {extra}"
            lbl = QLabel(title, box)
            lbl.setObjectName("AILabel")
            head.addWidget(lbl, 1)
            for value, color in ((_txt(iv, "overall_score"), _result_color(_txt(iv, "overall_score"))),
                                 (_txt(iv, "status"), theme.PALETTE["--text-muted"])):
                if value:
                    head.addWidget(_chip(box, value, color))
            v.addLayout(head)

            for fb in repo.list_interview_feedbacks(iv["interview_id"]):
                who = _txt(fb, "display_name") or "(interviewer)"
                bits = [b for b in (_txt(fb, "role"), _txt(fb, "job_title"),
                                    _txt(fb, "score")) if b]
                label = who + (f"  ({' · '.join(bits)})" if bits else "")
                v.addLayout(self._para(box, label, _txt(fb, "feedback") or "—"))
        return box

    # ---------------------------------------------------------- lịch sử liên hệ
    def _history_box(self, parent, candidate_id):
        """Mail đã gửi, cuộc gọi, đổi trạng thái, ghi chú — mỗi thứ một dòng."""
        box, v = _section_box(parent, "Contact history", "calendar")
        rows = repo.list_activities(candidate_id=candidate_id, limit=30)
        if not rows:
            v.addWidget(_muted(box, "This candidate has never been contacted."))
            return box

        for ac in rows:
            when = _txt(ac, "occurred_at")[:16]
            kind = _txt(ac, "type")
            if kind == "Status change":
                what = (f"{_txt(ac, 'from_status') or '(none)'} → "
                        f"{_txt(ac, 'to_status') or '(none)'}")
            else:
                what = _txt(ac, "subject") or _short(_txt(ac, "content"), 80) or "—"
            extra = _txt(ac, "position_title")
            line = f"{when}  ·  {kind}  ·  {what}" + (f"  ·  {extra}" if extra else "")
            v.addWidget(_muted(box, line))
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


# ═════════════ NHẬP NHANH (đúng các cột của file Excel xuất ra) ═════════
# Form sửa hồ sơ đầy đủ có ~30 ô (thông tin cá nhân · hồ sơ nghề nghiệp · nguyện
# vọng · đơn ứng tuyển) — quá dài cho việc thường xuyên nhất sau mỗi buổi phỏng
# vấn: gõ lại nhận xét mà người phỏng vấn vừa gửi về. Hộp này chỉ giữ đúng những
# ô có mặt trong file Excel (xem app/core/candidate_export.py) và bỏ các cột do
# máy sinh ra: Batch · ID (bóc từ tên file CV) · Score & AI Evaluation (AI chấm,
# sửa tay là hỏng lịch sử đánh giá).
_ROUND_TITLES = ["1st interview", "2nd interview", "3rd interview"]


class _FeedbackRow(QFrame):
    """Một người phỏng vấn + nhận xét của họ, trong một vòng.

    Ứng với MỘT dòng `interview_feedbacks`. Cột "INTERVIEW EVALUATION" của file
    Excel là các dòng này nối lại, nên nhập tách từng người ở đây thì lúc xuất
    mới ghép đúng tên với nhận xét tương ứng.

    Chỉ chọn NGƯỜI, không nhập vai trò: người phỏng vấn lấy từ `employees` nên
    chức danh và phòng ban đã có sẵn ở đó, hỏi lại là chép dữ liệu thừa.
    """

    def __init__(self, parent, interviewers, data=None):
        super().__init__(parent)
        self.setObjectName("AIBox")
        self.feedback_id = _get_val(data, "feedback_id")
        self._people = interviewers          # tên hiển thị → employee_id

        col = QVBoxLayout(self)
        col.setContentsMargins(12, 10, 12, 10)
        col.setSpacing(6)

        head = QHBoxLayout()
        head.setSpacing(8)
        self.name = widgets.ComboBox(self)
        self.name.addItems([""] + list(interviewers))
        # Người đã lưu mà nay không còn trong danh sách (đã nghỉ việc, hoặc là
        # khách mời) vẫn phải hiện đúng tên → thêm vào cuối danh sách.
        saved = _txt(data, "display_name") or _txt(data, "interviewer_name")
        if saved and self.name.findText(saved) < 0:
            self.name.addItem(saved)
        self.name.setCurrentText(saved)
        head.addWidget(self.name, 3)

        self.score = widgets.ComboBox(self)
        self.score.addItems([""] + cv_schema.INTERVIEW_SCORE_CHOICES)
        self.score.setCurrentText(_txt(data, "score"))
        head.addWidget(self.score, 2)

        head.addWidget(widgets.button(self, "", variant="neutral", icon="x",
                                      command=self._remove), 0)
        col.addLayout(head)

        self.feedback = widgets.TextEdit(self)
        self.feedback.setAcceptRichText(False)
        self.feedback.setFixedHeight(66)
        self.feedback.setPlaceholderText("What did they say about the candidate?")
        self.feedback.setPlainText(_txt(data, "feedback"))
        col.addWidget(self.feedback)

    def _remove(self):
        self.setParent(None)
        self.deleteLater()

    def value(self):
        """dict để ghi xuống `interview_feedbacks`; None nếu dòng bỏ trống."""
        name = self.name.currentText().strip()
        text = self.feedback.toPlainText().strip()
        if not name and not text:
            return None
        return {
            "feedback_id":      self.feedback_id,
            # Nối vào employees khi chọn người đang làm việc; tên lạ (khách mời,
            # người đã nghỉ) chỉ lưu tên.
            "employee_id":      self._people.get(name),
            "interviewer_name": name or None,
            "score":            self.score.currentText().strip() or None,
            "feedback":         text or None,
        }


class _QuickEditDialog(ModalDialog):
    """Nhập nhanh cho MỘT ứng viên: đơn ứng tuyển + 3 vòng phỏng vấn."""

    def __init__(self, parent, row):
        super().__init__(parent, "lg")
        self._row = row
        self._cid = row["candidate_id"]
        self._app_id = _get_val(row, "application_id")
        # tên hiển thị → employee_id, để nhận xét nối được vào `employees`.
        self._interviewers = {e["full_name"]: e["employee_id"]
                              for e in repo.list_interviewers() if e["full_name"]}
        # Buổi phỏng vấn đã có của từng vòng — mỗi vòng nhiều nhất một buổi
        # (chỉ mục duy nhất application_id + round).
        self._rounds = {}
        for iv in repo.list_interviews(candidate_id=self._cid):
            self._rounds.setdefault(iv["round"] or 1, iv)
        self._fb_rows = {}      # vòng → list _FeedbackRow đang hiện
        self._fb_boxes = {}     # vòng → layout để chèn dòng mới vào

        name = _txt(row, "full_name") or f"#{self._cid}"
        card, lay = self.build_shell(f"Update feedback · {name}")

        body = QWidget()
        col = QVBoxLayout(body)
        col.setContentsMargins(0, 0, 8, 0)
        col.setSpacing(12)
        col.addWidget(self._application_box(body))
        for n in (1, 2, 3):
            col.addWidget(self._round_box(body, n))
        col.addStretch(1)
        sa = widgets.scroll_area(body)
        lay.addWidget(sa, 1)
        self.set_grow_region(sa)

        foot = QHBoxLayout()
        foot.addWidget(widgets.button(card, "Save", variant="primary", icon="check",
                                      command=self._save))
        foot.addWidget(widgets.button(card, "Cancel", variant="neutral", icon="x",
                                      command=self.reject))
        foot.addStretch(1)
        lay.addLayout(foot)

    # ------------------------------------------------------------ đơn ứng tuyển
    def _application_box(self, parent):
        box, v = _section_box(parent, "Application", "idcard")
        if not self._app_id:
            v.addWidget(_muted(box, "This candidate isn't linked to any position "
                                    "yet. Double-click the row to open the full "
                                    "form and pick a position first — interview "
                                    "rounds hang off the application."))

        # Vị trí ứng tuyển CHỈ ĐỂ XEM: đổi vị trí là đổi cả luồng tuyển dụng
        # (JD, mẫu mail, các vòng đã có) nên chỉ làm ở form đầy đủ.
        v.addLayout(_labeled(box, "Applying for",
                             _readonly(box, _txt(self._row, "position_title"))))

        two = QHBoxLayout()
        two.setSpacing(12)
        self.f_status = widgets.ComboBox(box)
        self.f_status.addItems([""] + cv_schema.CANDIDATE_STATUS_CHOICES)
        self.f_status.setCurrentText(_txt(self._row, "status"))
        two.addLayout(_labeled(box, "Status", self.f_status), 1)

        self.f_result = widgets.ComboBox(box)
        self.f_result.addItems([""] + cv_schema.FINAL_STATUS_CHOICES)
        self.f_result.setCurrentText(_txt(self._row, "final_status"))
        two.addLayout(_labeled(box, "Result", self.f_result), 1)
        v.addLayout(two)

        self.f_ps_date = widgets.DateEdit(box)
        self.f_ps_date.set(_txt(self._row, "phone_screen_date"))
        v.addLayout(_labeled(box, "Phone screen date", self.f_ps_date))

        self.f_ps_note = widgets.TextEdit(box)
        self.f_ps_note.setAcceptRichText(False)
        self.f_ps_note.setFixedHeight(60)
        self.f_ps_note.setPlainText(_txt(self._row, "application_note"))
        v.addLayout(_labeled(box, "Phone screen note", self.f_ps_note))
        return box

    # -------------------------------------------------------- một vòng phỏng vấn
    def _round_box(self, parent, n):
        interview = self._rounds.get(n)
        box, v = _section_box(parent, _ROUND_TITLES[n - 1], "users")

        two = QHBoxLayout()
        two.setSpacing(12)
        date = widgets.DateEdit(box)
        date.set(_txt(interview, "interview_date"))
        two.addLayout(_labeled(box, "Date", date), 1)

        result = widgets.ComboBox(box)
        result.addItems([""] + cv_schema.INTERVIEW_SCORE_CHOICES)
        result.setCurrentText(_txt(interview, "overall_score"))
        two.addLayout(_labeled(box, "Final result", result), 1)
        v.addLayout(two)

        setattr(self, f"f_r{n}_date", date)
        setattr(self, f"f_r{n}_result", result)

        head = QHBoxLayout()
        lbl = QLabel("Interviewer feedback", box)
        lbl.setObjectName("AILabel")
        head.addWidget(lbl, 1)
        head.addWidget(widgets.button(box, "Add interviewer", variant="neutral",
                                      icon="plus",
                                      command=lambda: self._add_feedback(n)))
        v.addLayout(head)

        holder = QVBoxLayout()
        holder.setSpacing(8)
        v.addLayout(holder)
        self._fb_boxes[n] = holder
        self._fb_rows[n] = []

        existing = (repo.list_interview_feedbacks(interview["interview_id"])
                    if interview else [])
        for fb in existing:
            self._add_feedback(n, fb)
        if not existing:
            self._add_feedback(n)      # sẵn một dòng trống để gõ ngay
        return box

    def _add_feedback(self, n, data=None):
        row = _FeedbackRow(self, self._interviewers, data)
        self._fb_boxes[n].addWidget(row)
        self._fb_rows[n].append(row)

    # ------------------------------------------------------------------- lưu
    def _save(self):
        """Ghi đơn ứng tuyển + 3 vòng phỏng vấn. Vòng trống trơn thì bỏ qua."""
        try:
            app_id = self._save_application()
            for n in (1, 2, 3):
                self._save_round(n, app_id)
        except Exception as exc:   # noqa: BLE001 — báo một lần, giữ hộp thoại lại
            dialogs.error(self, "Save failed", str(exc))
            return
        self.accept()

    def _save_application(self):
        """Cập nhật đơn ứng tuyển đang hiển thị; trả về application_id.

        Không tạo đơn mới: vị trí ứng tuyển chỉ để xem ở hộp này nên chưa có đơn
        thì cũng chưa biết gắn vào vị trí nào.
        """
        if not self._app_id:
            return None
        repo.update_application(self._app_id, {
            "status":            self.f_status.currentText().strip() or None,
            "final_status":      self.f_result.currentText().strip() or None,
            "phone_screen_date": self.f_ps_date.get() or None,
            "note":              self.f_ps_note.toPlainText().strip() or None,
        })
        return self._app_id

    def _save_round(self, n, app_id):
        """Ghi một vòng. Chưa có đơn thì không ghi được (buổi PV treo vào đơn)."""
        entries = [v for v in (r.value() for r in self._fb_rows[n]) if v]
        data = {
            "interview_date": getattr(self, f"f_r{n}_date").get() or None,
            "overall_score":  getattr(self, f"f_r{n}_result").currentText().strip() or None,
        }
        interview = self._rounds.get(n)
        if not any(data.values()) and not entries and interview is None:
            return                     # vòng chưa diễn ra → không tạo dòng rỗng
        if not app_id:
            raise ValueError(
                f"{_ROUND_TITLES[n - 1]} needs an application — open the full form "
                "(double-click the row) and pick a position first.")
        data["candidate_id"] = self._cid
        interview_id = repo.save_interview(app_id, n, data)
        repo.save_interview_feedbacks(interview_id, entries)


def _labeled(parent, text, widget):
    """Nhãn nhỏ + ô nhập bên dưới. Trả về QVBoxLayout để nhét vào form."""
    col = QVBoxLayout()
    col.setSpacing(3)
    lbl = QLabel(text, parent)
    lbl.setObjectName("AILabel")
    col.addWidget(lbl)
    col.addWidget(widget)
    return col


def _readonly(parent, text):
    """Ô CHỈ ĐỂ XEM: trông như ô nhập nhưng không sửa và không nhận focus."""
    edit = QLineEdit(text or "—", parent)
    edit.setReadOnly(True)
    edit.setFocusPolicy(Qt.NoFocus)
    edit.setDisabled(True)      # dùng luôn tông chữ mờ của QSS cho ô khoá
    return edit


def _get_val(row, key):
    """row[key] an toàn cho sqlite3.Row / dict / None (giữ nguyên kiểu, không ép chuỗi)."""
    if row is None:
        return None
    try:
        return row[key]
    except (KeyError, IndexError):
        return None


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
                               checkable=True,
                               menu_actions=[
                                   ("Update feedback", self._quick_edit,
                                    {"single": True}),
                                   ("Update source", self._bulk_source)])
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
        self.sel_pool = widgets.FilterSelect("Pool")
        self.sel_batch = widgets.FilterSelect("Batch")
        self.sel_status.set_options(cv_schema.CANDIDATE_STATUS_CHOICES)
        self.sel_pool.set_options(cv_schema.POOL_STATUS_CHOICES)
        for w in (self.sel_pos, self.sel_dept, self.sel_status, self.sel_pool,
                  self.sel_batch):
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
            department_id=dept_id, batch=self.sel_batch.value(),
            pool_status=self.sel_pool.value())
        self._rows = rows
        self.table.set_rows(rows)
        self.count_lbl.setText(
            f"Showing {len(rows)} candidates · Total in DB: {repo.count_candidates()}")

    def _clear_filters(self):
        self.ent_kw.clear()
        for w in (self.sel_pos, self.sel_dept, self.sel_status, self.sel_pool,
                  self.sel_batch):
            w.clear()
        self._reload()

    # Trạng thái tuyển dụng nằm ở ĐƠN ỨNG TUYỂN, không nằm ở ứng viên — nên mọi
    # thao tác đổi trạng thái / gửi thư mời đều cần application_id. Ứng viên mới
    # nhập tay mà chưa gắn vị trí thì chưa có đơn nào.
    @staticmethod
    def _app_id(row):
        try:
            return row["application_id"]
        except (KeyError, IndexError):
            return None

    def _rows_with_application(self, rows, what):
        """Lọc ra các dòng đã có đơn ứng tuyển; báo tên những dòng chưa có."""
        ok = [r for r in rows if self._app_id(r)]
        missing = [_txt(r, "full_name") or f"#{r['candidate_id']}"
                   for r in rows if not self._app_id(r)]
        if missing:
            dialogs.warning(
                self._root, "No application yet",
                f"{what} needs an application (candidate + position). These "
                "candidates aren't linked to any position yet — edit them and "
                "pick a position first:\n\n• " + "\n• ".join(missing))
        return ok

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
        rows = self._rows_with_application(rows, "Updating status")
        if not rows:
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
                    # Ghi vào ĐƠN, đồng thời tự lưu một dòng lịch sử
                    # "Status change" cho ứng viên (xem repo.set_application_status).
                    repo.set_application_status(self._app_id(row), status)
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

    # ------------------------------------------ đổi NGUỒN CV hàng loạt
    # Nguồn = SÀN cung cấp CV (Itviec · VietnamWorks · LinkedIn…). Tool quét CV
    # bằng AI không suy ra được thông tin này — nó đọc cả thư mục nên không biết
    # từng file lấy ở đâu — nên đây là chỗ điền tay duy nhất. Vào bằng CHUỘT PHẢI
    # trên bảng (xem menu_actions của DataTable), không chiếm chỗ trên toolbar.
    # ------------------------------------------- nhập nhanh (chuột phải 1 dòng)
    def _quick_edit(self, rows):
        """Mở hộp nhập nhanh cho dòng chuột phải trỏ tới (khai báo `single` nên
        bảng chỉ gọi khi phạm vi đúng 1 dòng)."""
        if _QuickEditDialog(self._root, rows[0]).exec():
            self._reload()

    def _bulk_source(self, rows):
        """Đổi nguồn cho các dòng bảng đã giải sẵn — dòng chuột phải nếu nó chưa
        tick, cả nhóm tick nếu nó nằm trong nhóm."""
        self._ask_bulk_source(rows)

    def _ask_bulk_source(self, rows):
        """Popup nhỏ: chọn sàn cung cấp CV rồi ghi cho MỌI hồ sơ đang tick.

        Ô chọn nhập tay được — sàn mới chưa có trong `cv_schema` vẫn điền thẳng.
        Các hồ sơ đang cùng một nguồn thật thì điền sẵn nguồn đó.
        """
        dlg, card, lay = build_dialog_shell(
            self._root, f"Update source — {len(rows)} candidate(s)")

        hint = QLabel("Where did these CVs come from? Pick a job board / "
                      "headhunter, or type a new one.")
        hint.setObjectName("Hint")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        combo = widgets.ComboBox(card)
        combo.setEditable(True)
        combo.addItems(cv_schema.CANDIDATE_SOURCE_CHOICES)
        # Dấu "AI CV Scan" là đường hồ sơ vào app, không phải sàn → coi như chưa
        # có nguồn, để ô trống thay vì điền sẵn một giá trị sai.
        current = {s for s in (_txt(r, "source") for r in rows)
                   if s and s != cv_schema.CANDIDATE_SOURCE_AUTO}
        combo.setCurrentText(current.pop() if len(current) == 1 else "")
        lay.addWidget(combo)

        names = QLabel("• " + "\n• ".join(
            _txt(r, "full_name") or f"#{r['candidate_id']}" for r in rows))
        names.setObjectName("Hint")
        names.setWordWrap(True)
        sa = widgets.scroll_area(names)
        sa.setMaximumHeight(dlg.modal_h)   # danh sách dài thì cuộn, ngắn thì vừa khít
        lay.addWidget(sa, 1)

        def do_update():
            source = combo.currentText().strip()
            if not source:
                dialogs.warning(dlg, "No source",
                                "Pick a source from the list or type one in.")
                return
            failed = []
            for row in rows:
                cid = row["candidate_id"]
                try:
                    # Ghi cả 3 cột `source` (ứng viên · đơn đang hiện trên bảng ·
                    # bản CV mới nhất) để mọi màn hình đọc ra cùng một giá trị.
                    repo.set_candidate_source(cid, source, self._app_id(row))
                except Exception as exc:   # noqa: BLE001 — gom lỗi báo một lần
                    failed.append(f"#{cid} — {exc}")
            dlg.accept()
            self._reload()
            if failed:
                dialogs.error(self._root, "Update failed",
                              "Couldn't update:\n• " + "\n• ".join(failed))
            else:
                dialogs.success(
                    self._root, "Source updated",
                    f"Set {len(rows)} candidate(s) to \"{source}\".")

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

        # "Do Not Contact" là hàng rào CỨNG — chặn ngay tại đây, không chỉ ẩn
        # khỏi danh sách tìm kiếm.
        blocked = [_txt(r, "full_name") or f"#{r['candidate_id']}"
                   for r in rows if _txt(r, "pool_status") == "Do Not Contact"]
        if blocked:
            dialogs.error(
                self._root, "Marked as do not contact",
                "These candidates are marked \"Do Not Contact\" and must not be "
                "emailed. Untick them first:\n\n• " + "\n• ".join(blocked))
            return
        rows = self._rows_with_application(rows, "Sending emails")
        if not rows:
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
                self._record_invite(row, tpl, value + 1, subject, body_html,
                                    email, start, end)
            else:
                failed.append(f"{who} — {err}")
        self._report_meetings(opened, failed, skipped, no_cv, after_status)

    def _record_invite(self, row, tpl, round_no, subject, body_html, email,
                       start, end):
        """Ghi lại việc đã mời phỏng vấn: 1 dòng lịch sử liên hệ + 1 buổi PV.

        Buổi phỏng vấn được tạo ở trạng thái `Scheduled` với đúng giờ đã hẹn —
        sau khi phỏng vấn xong, HR mở màn hình chi tiết để nhập kết quả và nhận
        xét cho vòng đó.
        """
        application_id = self._app_id(row)
        try:
            activity_id = repo.log_activity({
                "candidate_id":     row["candidate_id"],
                "application_id":   application_id,
                "type":             "Email",
                "round":            round_no,
                "scheduled_at":     start.toPython().strftime("%Y-%m-%d %H:%M:%S"),
                "subject":          subject,
                "content":          body_html,
                "mail_template_id": tpl["mail_template_id"],
                "mail_to":          email,
                "mail_cc":          _txt(tpl, "mail_cc"),
                "result":           "Pending",
            })
            minutes = max(0, start.secsTo(end) // 60)
            repo.save_interview(application_id, round_no, {
                "interview_date":   start.toPython().strftime("%Y-%m-%d %H:%M:%S"),
                "duration_minutes": minutes,
                "status":           "Scheduled",
                "mail_activity_id": activity_id,
            })
        except Exception as exc:   # noqa: BLE001 — mail đã mở rồi, đừng chặn luồng
            dialogs.warning(self._root, "History not saved",
                            f"The invite was opened but couldn't be logged:\n{exc}")

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
            repo.log_activity({
                "candidate_id":     row["candidate_id"],
                "application_id":   self._app_id(row),
                "type":             "Email",
                "subject":          subject,
                "content":          body_html,
                "mail_template_id": tpl["mail_template_id"],
                "mail_to":          email,
                "mail_cc":          _txt(tpl, "mail_cc"),
            })
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
            selects.append((row, combo))
        col.addStretch(1)
        sa = widgets.scroll_area(body)
        sa.setMaximumHeight(dlg.modal_h)   # danh sách dài thì cuộn, ngắn thì vừa khít
        lay.addWidget(sa, 1)

        def do_update():
            failed = []
            for row, combo in selects:
                cid = row["candidate_id"]
                try:
                    repo.set_application_status(self._app_id(row), combo.currentText())
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

        cal = widgets.Calendar(card)     # đã style sẵn, bỏ cột số tuần, khoanh hôm nay
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
    # Sheet "Candidates" được dựng thẳng bằng code (app.core.candidate_export),
    # không đọc file .xlsx mẫu nào. Tên file MỚI → tạo mới; tên file ĐÃ CÓ → hỏi
    # ghi nối tiếp hay ghi đè (xem _ask_overwrite).
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
        # DontConfirmOverwrite: hộp thoại của Windows chỉ hỏi được có/không, mà ở
        # đây có tới ba lối đi (nối tiếp · ghi đè · hủy) nên tự hỏi lấy.
        path, _ = QFileDialog.getSaveFileName(
            self._root, "Export selected candidates to Excel", "Candidates.xlsx",
            "Excel (*.xlsx)", "", QFileDialog.Option.DontConfirmOverwrite)
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"

        append = False
        if os.path.isfile(path):
            append = self._ask_overwrite(path)
            if append is None:
                return

        self._set_export_loading(True)
        try:
            export_rows = candidate_export.rows_from_candidates(rows)
            candidate_export.export(
                path, export_rows,
                dropdowns=candidate_export.dropdown_sources(), append=append)
        except PermissionError:
            dialogs.error(self._root, "Can't write file",
                          f"Is the Excel file open? Close it and retry:\n{path}")
            return
        except Exception as exc:
            dialogs.error(self._root, "Excel export error", str(exc))
            return
        finally:
            self._set_export_loading(False)

        mode = "appended" if append else "new file"
        if dialogs.confirm(
                self._root, "Done",
                f"Exported {len(export_rows)} candidates ({mode}) to:\n{path}\n\nOpen now?",
                ok_label="Open", cancel_label="Close"):
            self._launch(path)

    def _ask_overwrite(self, path):
        """File trùng tên → hỏi ghi nối tiếp hay ghi đè.

        Trả về True (nối tiếp) / False (ghi đè) / None (hủy). Nối tiếp là lựa
        chọn chính vì đó là cách dùng thường ngày — gom nhiều đợt lọc vào cùng
        một bảng; ghi đè để tông cảnh báo vì nó XÓA dữ liệu cũ trong file.
        """
        name = os.path.basename(path)
        choice = dialogs.choose(
            self._root, "File already exists",
            f'"{name}" already exists. What should we do?\n\n'
            "• Append — add these candidates below the rows already in the file\n"
            "• Overwrite — replace the file with only these candidates "
            "(everything currently in it is lost)",
            [("Append", "primary", "append"),
             ("Overwrite", "warning", "overwrite")])
        if not choice:
            return None
        return choice == "append"

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

    # ------------------------------------------------------------- form specs
    # Form nhập tay gộp HAI bảng: phần hồ sơ con người (`candidates`) và phần
    # đơn ứng tuyển (`applications`). `_split_form_data` tách lại lúc lưu.
    _APPLICATION_KEYS = ("position_id", "status", "final_status", "applied_at",
                         "phone_screen_date")

    def _candidate_form_specs(self):
        return [
            {"kind": "section", "label": "Personal info"},
            {"key": "full_name", "label": "Full name (*)", "kind": "text", "required": True},
            {"key": "email", "label": "Email", "kind": "text"},
            {"key": "phone", "label": "Phone", "kind": "text"},
            {"key": "date_of_birth", "label": "Date of birth (dd/mm/yyyy)", "kind": "text"},
            {"key": "gender", "label": "Gender", "kind": "choice",
             "choices": cv_schema.GENDER_CHOICES, "allow_empty": True},
            {"key": "address", "label": "Address", "kind": "text"},
            {"key": "city", "label": "City / province", "kind": "text"},

            {"kind": "section", "label": "Professional profile"},
            {"key": "current_title", "label": "Current job title", "kind": "text"},
            {"key": "industry", "label": "Industry", "kind": "text"},
            {"key": "years_experience", "label": "Years of experience (at CV date)",
             "kind": "decimal"},
            {"key": "experience_as_of", "label": "…as of date (yyyy-mm-dd)", "kind": "text"},
            {"key": "education", "label": "Education", "kind": "text"},
            {"key": "major", "label": "Major", "kind": "text"},
            {"key": "languages", "label": "Languages (separate with ;)", "kind": "text"},
            {"key": "skills_text", "label": "Skills (separate with ;)",
             "kind": "textarea", "height": 3},
            {"key": "profile_summary", "label": "Profile summary",
             "kind": "textarea", "height": 3},

            {"kind": "section", "label": "Expectations"},
            {"key": "expected_salary", "label": "Expected salary (number)", "kind": "decimal"},
            {"key": "salary_note", "label": "Salary note (gross/net, range…)", "kind": "text"},
            {"key": "available_from", "label": "Available from (yyyy-mm-dd)", "kind": "text"},
            {"key": "willing_to_relocate", "label": "Willing to relocate",
             "kind": "choice", "choices": cv_schema.RELOCATE_CHOICES, "allow_empty": True},
            {"key": "preferred_location", "label": "Preferred location", "kind": "text"},

            {"kind": "section", "label": "Application"},
            {"key": "position_id", "label": "Position applied for", "kind": "dropdown",
             "options": _position_options},
            {"key": "status", "label": "Status", "kind": "choice",
             "choices": cv_schema.CANDIDATE_STATUS_CHOICES, "allow_empty": True},
            {"key": "final_status", "label": "Result", "kind": "choice",
             "choices": cv_schema.FINAL_STATUS_CHOICES, "allow_empty": True},
            {"key": "applied_at", "label": "Applied date (yyyy-mm-dd)", "kind": "text"},
            {"key": "phone_screen_date", "label": "Phone screen date (yyyy-mm-dd)",
             "kind": "text"},

            {"kind": "section", "label": "Talent pool"},
            {"key": "pool_status", "label": "Pool status", "kind": "choice",
             "choices": cv_schema.POOL_STATUS_CHOICES, "allow_empty": True},
            {"key": "source", "label": "Source", "kind": "text"},
            {"key": "note", "label": "Note", "kind": "textarea", "height": 3},
        ]

    @classmethod
    def _split_form_data(cls, data):
        """Dict từ form → (phần ứng viên, phần đơn ứng tuyển)."""
        app = {k: data[k] for k in cls._APPLICATION_KEYS if k in data}
        candidate = {k: v for k, v in data.items() if k not in cls._APPLICATION_KEYS}
        return candidate, app

    def _add(self):
        def _save(data):
            dups = repo.find_duplicates(data.get("email"), data.get("phone"))
            if dups and not self._confirm_duplicate(dups):
                return False
            candidate, app = self._split_form_data(data)
            cid = repo.insert_candidate(candidate)
            # Có chọn vị trí thì mở luôn một đơn ứng tuyển — trạng thái tuyển
            # dụng chỉ tồn tại trên đơn, không nằm ở hồ sơ ứng viên.
            if app.get("position_id"):
                app["candidate_id"] = cid
                repo.insert_application(app)
            self._reload()

        FormDialog(self._root, "Add candidate",
                   self._candidate_form_specs(), None, on_save=_save).run()

    def _edit(self, cid=None):
        if cid is None:
            cid = self._selected_id()
        if cid is None:
            return
        # Bản ghi đầy đủ = hồ sơ ứng viên + đơn mới nhất, để form điền sẵn cả hai.
        current = repo.get_candidate_full(cid)

        def _save(data):
            dups = repo.find_duplicates(data.get("email"), data.get("phone"), exclude_id=cid)
            if dups and not self._confirm_duplicate(dups):
                return False
            candidate, app = self._split_form_data(data)
            repo.update_candidate(cid, candidate)
            existing = repo.latest_application(cid)
            if existing is not None:
                # Đổi trạng thái qua set_application_status để lịch sử ghi lại
                # được bước chuyển; các ô còn lại cập nhật thẳng.
                status = app.pop("status", None)
                repo.update_application(existing["application_id"], app)
                if status and status != (existing["status"] or ""):
                    repo.set_application_status(existing["application_id"], status)
            elif app.get("position_id"):
                app["candidate_id"] = cid
                repo.insert_application(app)
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
        path = repo.candidate_cv_path(cid)
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
