"""Quản lý khóa học / đào tạo (SQLite) — bản PySide6.

Màn hình riêng trên sidebar (nằm ngay sau "Quản lý nhân viên"). Tra cứu các
LƯỢT GHI DANH trong bảng `course_employees`:
  • Tìm kiếm bằng 2 ô select (KHÔNG có ô nhập tự do, KHÔNG có nút):
      – Khóa học: mặc định luôn là khóa có course_id lớn nhất (mới nhất).
      – Trạng thái học: Not started / Completed.
  • Bảng liệt kê đầy đủ cột của course_employees + thông tin nhân viên/khóa học.
  • Double-click 1 dòng → modal chỉ cập nhật cột `status`.
Tầng dữ liệu dùng lại app.core.cv_repository (bảng `course_employees`).
"""
import os
import re

from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QLabel, QVBoxLayout

from app.core import cv_repository as repo
from app.core import cv_schema
from app.core import settings as app_settings
from app_qt import dialogs, widgets
from app_qt.base_tool import BaseTool
from app_qt.components.form_dialog import FormDialog
from app_qt.components.table import DataTable

try:
    import openpyxl  # noqa: F401 — chỉ để kiểm tra có thư viện trước khi xuất
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

# Cột bảng LƯỢT GHI DANH (course_employees): đầy đủ cột DB + thông tin nhân viên
# / khóa học (join) để bảng đọc được.
_COURSE_COLUMNS = [
    ("code",            "Mã NV",      100, "w"),
    ("full_name",       "Họ tên",     180, "w"),
    ("department_name", "Bộ phận",    150, "w"),
    ("status",          "Trạng thái", 120, "center"),
    ("note",            "Ghi chú",    220, "w"),
]


def _course_label(course):
    """Nhãn khóa học cho ô lọc: '#<id> — <tên>' (id đứng đầu → dễ nhận khóa mới
    nhất; kèm id nên KHÔNG trùng dù hai khóa cùng tên)."""
    return f"#{course['course_id']} — {course['title'] or 'Khóa học'}"


class CourseDbTool(BaseTool):
    name = "Quản lý khóa học"
    description = "Tra cứu lượt ghi danh & cập nhật trạng thái học (SQLite)."
    icon = "🎓"
    category = "Nhân sự"
    order = 20
    fills_height = True

    def build(self, parent=None):
        repo.init_db()
        card = widgets.Card(parent)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(22, 20, 22, 18)
        lay.setSpacing(10)
        self._root = card
        self._course_default_set = False

        widgets.section_label(card, "Tìm kiếm khóa học")
        self._build_search_bar(lay)

        self.table = DataTable(_COURSE_COLUMNS, pk="enrollment_id",
                               stretch_key="note", on_double=self._edit)
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
        filters = QHBoxLayout()
        filters.setSpacing(10)
        # Khóa học: allow_all=False → không có '— Tất cả —', luôn hiển thị đúng
        # dữ liệu của MỘT khóa học (mặc định là khóa có id lớn nhất). Nạp lại
        # danh sách mỗi khi bung dropdown → luôn lấy từ bảng courses (kể cả khi
        # trang đã bị cache và có khóa mới thêm sau đó).
        self.sel_course = widgets.FilterSelect("Khóa học", allow_all=False)
        self.sel_course.set_open_hook(self._load_course_options)
        self.sel_status = widgets.FilterSelect("Trạng thái")
        self.sel_status.set_options(cv_schema.COURSE_STATUS_CHOICES)
        for w in (self.sel_course, self.sel_status):
            w.changed.connect(self._reload)
            filters.addWidget(w, 1)
        filters.addStretch(1)
        self._btn_export = widgets.button(
            self._root, "Xuất Excel", variant="primary", icon="save",
            command=self._export_roster)
        filters.addWidget(self._btn_export)
        lay.addLayout(filters)

    # -------------------------------------------------------------- dữ liệu
    def _load_course_options(self):
        """Nạp danh sách khóa học TỪ BẢNG courses, sắp id GIẢM DẦN (mới nhất đầu).

        Chạy lúc dựng + mỗi lần bung dropdown. Giữ nguyên khóa đang chọn nếu còn;
        lần đầu tiên (chưa có mặc định) → chọn sẵn khóa mới nhất (id lớn nhất).
        """
        courses = sorted(repo.list_courses(),
                         key=lambda c: c["course_id"], reverse=True)
        self._course_opts = {_course_label(c): c["course_id"] for c in courses}
        self.sel_course.set_options(self._course_opts.keys())
        if not self._course_default_set and courses:
            self.sel_course.set_value(_course_label(courses[0]))
            self._course_default_set = True

    def _reload(self):
        self._load_course_options()
        course_id = self._course_opts.get(self.sel_course.value())
        rows = repo.search_course_employees(
            course_id=course_id, status=self.sel_status.value())
        self.table.set_rows(rows)
        self.count_lbl.setText(
            f"Hiển thị {len(rows)} lượt ghi danh · "
            f"Tổng trong DB: {repo.count_course_employees()}")

    # -------------------------------------------------------------- sửa
    def _edit(self, enrollment_id=None):
        if enrollment_id is None:
            return
        current = repo.get_enrollment(enrollment_id)
        specs = [
            {"key": "status", "label": "Trạng thái học", "kind": "choice",
             "choices": cv_schema.COURSE_STATUS_CHOICES, "allow_empty": True},
        ]

        def _save(data):
            repo.update_enrollment(enrollment_id, {"status": data.get("status")})
            self._reload()

        FormDialog(self._root, "Cập nhật trạng thái học",
                   specs, current, on_save=_save).run()

    # ---------------------------------------------------------- xuất Excel
    # Xuất roster khóa học theo template placeholder (file mẫu lấy ở màn Cài đặt).
    # CHỈ xuất các lượt ghi danh trạng thái "Not started" của khóa đang chọn — bỏ
    # qua ô lọc trạng thái trên UI (theo yêu cầu). Cột Dept dùng mã viết tắt bộ
    # phận (short_name), thiếu thì fallback về tên đầy đủ.
    def _export_roster(self):
        if not _OPENPYXL_OK:
            return dialogs.error(self._root, "Thiếu thư viện",
                                 "Cần openpyxl để xuất Excel:\n  pip install openpyxl")
        from app.core import excel_template

        tpl = (app_settings.get("course_template_path") or "").strip()
        if not tpl:
            return dialogs.error(
                self._root, "Chưa cấu hình template",
                "Hãy vào Cài đặt → Xuất Excel khóa học và chọn file template trước.")
        if not os.path.isfile(tpl):
            return dialogs.error(self._root, "Không tìm thấy template",
                                 f"File template không tồn tại:\n{tpl}")

        course_id = self._course_opts.get(self.sel_course.value())
        if not course_id:
            return dialogs.info(self._root, "Chưa chọn khóa học",
                                "Hãy chọn một khóa học để xuất.")

        not_started = cv_schema.COURSE_STATUS_CHOICES[0]   # "Not started"
        rows = repo.search_course_employees(course_id=course_id, status=not_started)
        if not rows:
            return dialogs.info(
                self._root, "Không có dữ liệu",
                f'Khóa học này không có nhân viên nào ở trạng thái "{not_started}".')

        course = repo.get_course(course_id)
        ct = course["course_type"]
        types = cv_schema.COURSE_TYPE_CHOICES
        context = {
            "course.id":       course["course_id"],
            "course.title":    course["title"] or "",
            "course.content":  course["content"] or "",
            "course.date":     course["date"] or "",
            "course.location": course["location"] or "",
            "course.type":     types[ct] if isinstance(ct, int) and 0 <= ct < len(types) else "",
            "t_inhouse":       "X" if ct == 0 else "",
            "t_external":      "X" if ct == 1 else "",
            "t_funded":        "X" if ct == 2 else "",
            "count":           len(rows),
        }
        # Mã NV toàn chữ số → xuất dạng SỐ để cột PERSON ID canh phải giống form gốc.
        def _code(v):
            v = (v or "").strip()
            return int(v) if v.isdigit() else v

        export_rows = [{
            "code":             _code(r["code"]),
            "full_name":        r["full_name"] or "",
            "department_short": (r["department_short"] or r["department_name"] or ""),
            "note":             r["note"] or "",
        } for r in rows]

        ext = ".xlsm" if tpl.lower().endswith(".xlsm") else ".xlsx"
        safe = re.sub(r'[\\/:*?"<>|]', "_", course["title"] or "roster").strip() or "roster"
        path, _ = QFileDialog.getSaveFileName(
            self._root, "Xuất roster khóa học", f"{safe}{ext}", "Excel (*.xlsx *.xlsm)")
        if not path:
            return
        if not path.lower().endswith((".xlsx", ".xlsm")):
            path += ext

        try:
            excel_template.render_template(
                tpl, path, context, export_rows,
                row_fields={"code", "full_name", "department_short", "note", "stt"})
        except PermissionError:
            return dialogs.error(self._root, "Không ghi được file",
                                 f"File Excel đang mở? Hãy đóng rồi thử lại:\n{path}")
        except Exception as exc:   # noqa: BLE001 — báo lại cho UI, không chặn luồng
            return dialogs.error(self._root, "Lỗi xuất Excel", str(exc))

        dialogs.success(
            self._root, "Hoàn tất",
            f'Đã xuất {len(export_rows)} nhân viên ("{not_started}") vào:\n{path}')
