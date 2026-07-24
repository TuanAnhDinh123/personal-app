"""Quản lý CV ứng viên & danh mục tuyển dụng (SQLite) — bản PySide6.

Port của app/tools/candidate_db.py. Tầng dữ liệu (app.core.cv_repository,
app.core.cv_schema) dùng lại 100% — chỉ dựng lại giao diện bằng Qt:
    • Tool chính "Quản lý CV ứng viên": tìm kiếm + bảng + CRUD + nhập Excel.
    • 3 trang Master Data: Bộ phận · Vị trí · JD (dùng CrudTablePanel).
"""
import os
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QCursor
from PySide6.QtWidgets import (
    QApplication, QDialog, QFileDialog, QFrame, QHBoxLayout, QLabel,
    QLineEdit, QToolTip, QVBoxLayout, QWidget,
)

from app.core import cv_repository as repo
from app.core import cv_schema
from app.core.cv_scan import (
    _open_existing_workbook, _open_template_workbook, _split_id_name,
    _write_candidates,
)
from app_qt import dialogs, theme, widgets
from app_qt.base_tool import BaseTool
from app_qt.components.crud_panel import CrudTablePanel
from app_qt.components.form_dialog import FormDialog
from app_qt.components.table import DataTable

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
    "candidate_id":    40,
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
    ("candidate_id",    "ID",         _W["candidate_id"],    "center"),
    ("full_name",       "Họ tên",     _W["full_name"],       "w"),
    ("email",           "Email",      _W["email"],           "w"),
    ("phone",           "SĐT",        _W["phone"],           "w"),
    ("position_title",  "Vị trí",     _W["position_title"],  "w"),
    ("fit_score",       "Điểm",       _W["fit_score"],       "center"),
    ("cv_file_path",    "CV",         _W["cv_file_path"],    "w",
     lambda v: os.path.basename(str(v)) if v else ""),
    ("department_name", "Bộ phận",    _W["department_name"], "w"),
    ("batch",           "Batch",      _W["batch"],           "center"),
    ("status",          "Trạng thái", _W["status"],          "center"),
    ("date_of_birth",   "Ngày sinh",  _W["date_of_birth"],   "center"),
    ("applied_at",      "Ngày nộp",   _W["applied_at"],      "center"),
    ("note",            "Ghi chú",    _W["note"],            "w",
     lambda v: (str(v).replace("\n", " ")[:60] + "…")
     if v and len(str(v)) > 60 else (str(v).replace("\n", " ") if v else "")),
]

_EXCEL_HEADER_MAP = {
    "batch":            "batch",
    "họ tên":           "full_name",
    "ngày sinh":        "date_of_birth",
    "email":            "email",
    "số điện thoại":    "phone",
    "điểm phù hợp":     "fit_score",
    "đánh giá phù hợp": "fit_summary",
    "ưu điểm":          "strengths",
    "nhược điểm":       "weaknesses",
    "tên file":         "cv_file_path",   # file cũ: chỉ có tên → cần hỏi thư mục CV
    "đường dẫn cv":     "cv_file_path",   # file mới: đã có sẵn đường dẫn đầy đủ
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
        QToolTip.showText(QCursor.pos(), "Đã sao chép", chip)

    chip.mousePressEvent = _do_copy
    return chip


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


# ═════════════════════ MODAL XEM CHI TIẾT ỨNG VIÊN ══════════════════════
class _CandidateDetailDialog(QDialog):
    """Modal lớn xem chi tiết toàn bộ ứng viên trong danh sách hiện tại.

    Mỗi ứng viên là một thẻ; phần nổi bật nhất là NHẬN XÉT CỦA AI
    (điểm phù hợp, nhận xét, ưu điểm, nhược điểm) từ bước quét CV.
    """

    def __init__(self, parent, rows):
        super().__init__(parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setModal(True)
        self._drag = None

        shell = QVBoxLayout(self)
        shell.setContentsMargins(24, 24, 24, 24)
        card = QFrame(self)
        card.setObjectName("Dialog")
        card.setMinimumSize(920, 620)
        shell.addWidget(card)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(22, 20, 22, 18)
        lay.setSpacing(12)

        head = QHBoxLayout()
        title = QLabel(f"Chi tiết ứng viên · {len(rows)} người")
        title.setObjectName("DialogTitle")
        head.addWidget(title, 1)
        close = QLabel("✕")
        close.setObjectName("DialogClose")
        close.setFixedSize(24, 24)
        close.setAlignment(Qt.AlignCenter)
        close.setCursor(Qt.PointingHandCursor)
        close.mousePressEvent = lambda _e: self.reject()
        head.addWidget(close, 0, Qt.AlignTop)
        lay.addLayout(head)

        body = QWidget()
        col = QVBoxLayout(body)
        col.setContentsMargins(0, 0, 8, 0)
        col.setSpacing(12)
        if not rows:
            empty = QLabel("Danh sách hiện tại đang trống.")
            empty.setObjectName("DialogMsg")
            col.addWidget(empty)
        for row in rows:
            col.addWidget(self._card(body, row))
        col.addStretch(1)
        lay.addWidget(widgets.scroll_area(body), 1)

        foot = QHBoxLayout()
        foot.addStretch(1)
        foot.addWidget(widgets.button(card, "Đóng", variant="neutral", icon="x",
                                      command=self.reject))
        lay.addLayout(foot)

        self.resize(1000, 760)

    # ------------------------------------------------------------- thẻ 1 ứng viên
    def _card(self, parent, row):
        box = QFrame(parent)
        box.setObjectName("DetailCard")
        v = QVBoxLayout(box)
        v.setContentsMargins(16, 14, 16, 14)
        v.setSpacing(8)

        # Hàng tiêu đề: tên + chip điểm/trạng thái
        head = QHBoxLayout()
        head.setSpacing(8)
        cid = _txt(row, "candidate_id")
        name = QLabel(f"#{cid}  {_txt(row, 'full_name') or '(chưa có tên)'}", box)
        name.setObjectName("DetailName")
        head.addWidget(name, 1)
        score = _txt(row, "fit_score")
        if score:
            head.addWidget(_chip(box, f"Điểm {score}", _score_color(score)))
        status = _txt(row, "status")
        if status:
            head.addWidget(_chip(box, status, theme.PALETTE["--info"]))
        v.addLayout(head)

        # Hàng thông tin phụ (bôi-chọn được để copy tay nếu cần)
        meta = " · ".join(p for p in (
            _txt(row, "position_title"), _txt(row, "department_name"),
            (f"NS: {_txt(row, 'date_of_birth')}" if _txt(row, "date_of_birth") else ""),
            (f"Nộp: {_txt(row, 'applied_at')}" if _txt(row, "applied_at") else ""),
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
            v.addLayout(self._para(box, "Ghi chú", note))
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
        h = QLabel("Nhận xét của AI", box)
        h.setObjectName("AIHeader")
        header.addWidget(h, 1)
        v.addLayout(header)

        summary = _txt(row, "fit_summary")
        strengths = _txt(row, "strengths")
        weaknesses = _txt(row, "weaknesses")

        if not any((summary, strengths, weaknesses)):
            empty = QLabel("Chưa có nhận xét từ AI cho ứng viên này.", box)
            empty.setObjectName("AIEmpty")
            v.addWidget(empty)
            return box

        if summary:
            v.addLayout(self._para(box, "Nhận xét phù hợp", summary))
        if strengths or weaknesses:
            two = QHBoxLayout()
            two.setSpacing(12)
            two.addLayout(self._para(box, "Ưu điểm", strengths or "—"), 1)
            two.addLayout(self._para(box, "Nhược điểm", weaknesses or "—"), 1)
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

    # kéo di chuyển (frameless)
    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag = e.globalPosition().toPoint() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, e):
        if self._drag is not None and e.buttons() & Qt.LeftButton:
            self.move(e.globalPosition().toPoint() - self._drag)

    def mouseReleaseEvent(self, e):
        self._drag = None


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
        self.ent_kw.setPlaceholderText("Tìm kiếm…")
        self.ent_kw.setClearButtonEnabled(True)
        self.ent_kw.addAction(widgets.svg_icon("search", theme.TEXT_MUTED, 16),
                              QLineEdit.LeadingPosition)
        self.ent_kw.editingFinished.connect(self._reload)
        lay.addWidget(self.ent_kw)

        # Hàng ô lọc dạng select 'nhãn nổi' — chọn 1 option là tìm luôn.
        filters = QHBoxLayout()
        filters.setSpacing(10)
        self.sel_pos = widgets.FilterSelect("Vị trí")
        self.sel_dept = widgets.FilterSelect("Bộ phận")
        self.sel_status = widgets.FilterSelect("Trạng thái")
        self.sel_batch = widgets.FilterSelect("Batch")
        self.sel_status.set_options(cv_schema.STATUS_CHOICES)
        for w in (self.sel_pos, self.sel_dept, self.sel_status, self.sel_batch):
            w.changed.connect(self._reload)
            filters.addWidget(w, 1)
        filters.addWidget(widgets.button(None, "Đặt lại", variant="neutral",
                                         icon="eraser", command=self._clear_filters), 0)
        lay.addLayout(filters)

    def _build_toolbar(self, lay):
        bar = QHBoxLayout()
        bar.setSpacing(6)
        B = widgets.button
        bar.addWidget(B(None, "Thêm mới", variant="success", icon="plus", command=self._add))
        bar.addWidget(B(None, "Xem chi tiết", variant="info", icon="sparkles",
                        command=self._show_details))
        bar.addWidget(B(None, "Nhập từ Excel", variant="primary", icon="download",
                        command=self._batch_import))
        self._btn_export = B(None, "Xuất Excel", variant="warning", icon="save",
                             command=self._export_excel)
        bar.addWidget(self._btn_export)
        bar.addStretch(1)
        bar.addWidget(B(None, "Tải lại", variant="neutral", icon="refresh", command=self._reload))
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
            f"Hiển thị {len(rows)} ứng viên · Tổng trong DB: {repo.count_candidates()}")

    def _clear_filters(self):
        self.ent_kw.clear()
        for w in (self.sel_pos, self.sel_dept, self.sel_status, self.sel_batch):
            w.clear()
        self._reload()

    def _selected_id(self):
        cid = self.table.selected_id()
        if cid is None:
            dialogs.info(self._root, "Chưa chọn", "Vui lòng chọn một ứng viên trong bảng.")
        return cid

    def _show_details(self):
        _CandidateDetailDialog(self._root, getattr(self, "_rows", [])).exec()

    # --------------------------------------------------- xuất Excel (các dòng đã tick)
    # Dùng lại đúng logic "Quét CV → Trích xuất Excel": ghi vào template có sẵn
    # (sheet "Candidates"). Chọn file MỚI → tạo theo template; chọn file CÓ SẴN
    # đúng mẫu → ghi nối tiếp. (Tương lai gỡ tool "Quét CV", tính năng nằm ở đây.)
    def _export_excel(self):
        if not _OPENPYXL_OK:
            dialogs.error(self._root, "Thiếu thư viện",
                          "Cần openpyxl để xuất Excel:\n  pip install openpyxl")
            return
        rows = self.table.checked_rows()
        if not rows:
            dialogs.info(self._root, "Chưa chọn",
                         "Hãy tick chọn ít nhất một ứng viên trong bảng để xuất.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self._root, "Xuất ứng viên đã chọn ra Excel", "Candidates.xlsx",
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
                mode = "nối tiếp"
            else:
                wb, ws = _open_template_workbook()
                mode = "mới"
            _write_candidates(ws, export_rows)
            wb.save(path)
        except PermissionError:
            dialogs.error(self._root, "Không ghi được file",
                          f"File Excel đang mở? Hãy đóng rồi thử lại:\n{path}")
            return
        except Exception as exc:
            dialogs.error(self._root, "Lỗi xuất Excel", str(exc))
            return
        finally:
            self._set_export_loading(False)

        if dialogs.confirm(
                self._root, "Hoàn tất",
                f"Đã xuất {len(export_rows)} ứng viên ({mode}) vào:\n{path}\n\nMở file ngay?",
                ok_label="Mở file", cancel_label="Đóng"):
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
            btn.setText("Đang xuất…")
        else:
            btn.setText(getattr(self, "_export_label", "Xuất Excel"))
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
            {"kind": "section", "label": "Thông tin cá nhân"},
            {"key": "full_name", "label": "Họ và tên (*)", "kind": "text", "required": True},
            {"key": "email", "label": "Email", "kind": "text"},
            {"key": "phone", "label": "Số điện thoại", "kind": "text"},
            {"key": "date_of_birth", "label": "Ngày sinh (dd/mm/yyyy)", "kind": "text"},
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
            {"key": "batch", "label": "Batch (đợt quét — chỉ số)", "kind": "int"},
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

        FormDialog(self._root, "Sửa ứng viên",
                   self._candidate_form_specs(), current,
                   on_save=_save, on_delete=lambda: self._delete(cid)).run()

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

    def _delete(self, cid):
        """Xóa ứng viên; trả về False nếu người dùng hủy xác nhận (giữ form mở)."""
        row = repo.get_candidate(cid)
        name = row["full_name"] if row else f"#{cid}"
        if not dialogs.confirm(self._root, "Xác nhận xóa",
                               f'Xóa ứng viên "{name}" khỏi DB?', ok_label="Xóa"):
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

        # Chỉ hỏi thư mục CV khi còn đường dẫn TƯƠNG ĐỐI cần ghép (file Excel cũ
        # chỉ có tên file). File mới từ tool quét AI đã ghi sẵn đường dẫn tuyệt
        # đối nên bỏ qua bước này.
        folder = ""
        if any(self._needs_cv_folder(r) for r in rows):
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
        lay.setContentsMargins(22, 20, 22, 18)   # padding thẻ→nội dung chuẩn
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
