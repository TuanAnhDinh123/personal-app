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
from app.core import signature_scan
from app_qt import dialogs, widgets
from app_qt.base_tool import BaseTool
from app_qt.components.form_dialog import FormDialog
from app_qt.components.modal import ModalDialog
from app_qt.components.progress_dialog import ProgressDialog
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
        self._btn_scan = widgets.button(
            self._root, "Quét chữ ký (AI)", variant="info", icon="sparkles",
            command=self._scan_signatures)
        filters.addWidget(self._btn_scan)
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

    # ---------------------------------------------------- quét chữ ký bằng AI
    # HR scan bảng điểm danh đã ký → gửi cho Gemini nhận diện ai đã ký → review →
    # cập nhật status sang "Completed". Chỉ đối chiếu nhóm "Not started" của khóa
    # đang chọn (đúng đối tượng cần ký, khớp với file đã xuất).
    def _scan_signatures(self):
        gen = app_settings.load()
        api_key = (gen.get("api_key") or "").strip()
        model = (gen.get("ai_model") or "").strip() or app_settings.DEFAULTS["ai_model"]
        if not api_key:
            return dialogs.error(
                self._root, "Chưa có API key",
                "Hãy vào Cài đặt → AI (Gemini) và nhập API key trước.")

        course_id = self._course_opts.get(self.sel_course.value())
        if not course_id:
            return dialogs.info(self._root, "Chưa chọn khóa học",
                                "Hãy chọn một khóa học để quét.")

        not_started = cv_schema.COURSE_STATUS_CHOICES[0]   # "Not started"
        roster = repo.search_course_employees(course_id=course_id, status=not_started)
        if not roster:
            return dialogs.info(
                self._root, "Không có dữ liệu",
                f'Khóa học này không có nhân viên nào ở trạng thái "{not_started}".')

        path, _ = QFileDialog.getOpenFileName(
            self._root, "Chọn file scan chữ ký (ảnh/PDF)", "",
            "Ảnh & PDF (*.png *.jpg *.jpeg *.webp *.pdf)")
        if not path:
            return
        if not signature_scan.is_supported(path):
            return dialogs.error(self._root, "Định dạng không hỗ trợ",
                                 "Chỉ hỗ trợ ảnh (PNG/JPG/WEBP) hoặc PDF.")
        try:
            with open(path, "rb") as f:
                file_bytes = f.read()
        except Exception as exc:   # noqa: BLE001
            return dialogs.error(self._root, "Lỗi đọc file", str(exc))
        mime = signature_scan.guess_mime(path)
        roster_min = [{"code": r["code"] or "", "full_name": r["full_name"] or ""}
                      for r in roster]

        def job(ctx):
            ctx.status(f"Đang gửi cho AI ({model})…")

            def on_retry(attempt, wait, reason):
                ctx.log(f"Thử lại lần {attempt} sau {wait}s ({reason})")

            data = signature_scan.detect_signatures(
                api_key, model, file_bytes, mime, roster_min,
                on_retry=on_retry, should_cancel=lambda: ctx.cancelled)
            ctx.step()
            return data

        def on_finish(dlg, result):
            dlg.close()
            rows = result.get("rows", []) if isinstance(result, dict) else []
            by_code = {str(r["code"] or "").strip(): r
                       for r in roster if (r["code"] or "").strip()}
            signed, unmatched, seen = [], [], set()
            for item in rows:
                if not item.get("signed"):
                    continue
                pid = str(item.get("person_id") or "").strip()
                if pid in by_code and pid not in seen:
                    signed.append(by_code[pid])
                    seen.add(pid)
                elif pid:
                    unmatched.append((pid, (item.get("name") or "").strip()))
            if not signed:
                note = ""
                if unmatched:
                    note = ("\n\nAI báo có ký nhưng không khớp mã NV nào: "
                            + ", ".join(p for p, _ in unmatched[:10]))
                return dialogs.info(
                    self._root, "Không có ai đã ký",
                    "AI không nhận thấy chữ ký khớp với nhân viên trong khóa." + note)
            self._review_signed(signed, unmatched)

        dlg = ProgressDialog(self._root, "Đang nhận diện chữ ký…", total=1,
                             subtitle=f"Phân tích {os.path.basename(path)} bằng {model}")
        dlg.start(job, on_finish)

    def _review_signed(self, signed_rows, unmatched):
        """Bảng review: HR bỏ tick người AI nhận nhầm rồi xác nhận → cập nhật status."""
        completed = cv_schema.COURSE_STATUS_CHOICES[-1]   # "Completed"
        dlg = ModalDialog(self._root, size="md")
        card, lay = dlg.build_shell("Xác nhận nhân viên đã ký")

        hint = QLabel(
            f"AI nhận diện {len(signed_rows)} nhân viên đã ký. Bỏ tick người bị nhận "
            f'nhầm, rồi bấm Xác nhận để cập nhật trạng thái sang "{completed}".')
        hint.setObjectName("Hint")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        cols = [
            ("code",            "Mã NV",   110, "w"),
            ("full_name",       "Họ tên",  200, "w"),
            ("department_name", "Bộ phận", 150, "w"),
        ]
        table = DataTable(cols, pk="enrollment_id", checkable=True)
        table.set_rows(signed_rows)
        table._model.set_all_checked(True)   # tick sẵn tất cả người AI báo đã ký
        dlg.set_grow_region(table)
        lay.addWidget(table, 1)

        if unmatched:
            warn = QLabel(
                "Không khớp mã NV (bỏ qua): "
                + ", ".join(f"{p} {n}".strip() for p, n in unmatched[:15]))
            warn.setObjectName("Hint")
            warn.setWordWrap(True)
            lay.addWidget(warn)

        def _confirm():
            chosen = table.checked_rows()
            for r in chosen:
                repo.update_enrollment(r["enrollment_id"], {"status": completed})
            dlg.accept()
            self._reload()
            dialogs.success(
                self._root, "Hoàn tất",
                f'Đã cập nhật {len(chosen)} nhân viên sang "{completed}".')

        foot = QHBoxLayout()
        foot.addStretch(1)
        foot.addWidget(widgets.button(card, "Hủy", variant="neutral", icon="x",
                                      command=dlg.reject))
        foot.addWidget(widgets.button(card, "Xác nhận", variant="success",
                                      icon="check", command=_confirm))
        lay.addLayout(foot)
        dlg.exec()
