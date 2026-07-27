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

from app.core import config
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
    ("code",            "Emp code",   100, "w"),
    ("full_name",       "Full name",  180, "w"),
    ("department_name", "Department", 150, "w"),
    ("status",          "Status",     120, "center"),
    ("note",            "Note",       220, "w"),
]

# Section config lưu "việc quét còn dở": các TRANG bị lỗi của lần quét trước, để
# lần sau chỉ quét tiếp mấy trang đó thay vì gửi lại cả tập 20 trang. Chỉ ghi khi
# HR đã bấm Xác nhận (đã cập nhật DB) — xem _review_attendance.
SECTION = "course_db"
_PENDING_KEY = "pending_scan"


def _course_label(course):
    """Nhãn khóa học cho ô lọc: '#<id> — <tên>' (id đứng đầu → dễ nhận khóa mới
    nhất; kèm id nên KHÔNG trùng dù hai khóa cùng tên)."""
    return f"#{course['course_id']} — {course['title'] or 'Course'}"


class CourseDbTool(BaseTool):
    name = "Course Manager"
    description = "Browse enrollments & update learning status (SQLite)."
    icon = "🎓"
    category = "Human Resources"
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

        widgets.section_label(card, "Search courses")
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
        self.sel_course = widgets.FilterSelect("Course", allow_all=False)
        self.sel_course.set_open_hook(self._load_course_options)
        self.sel_status = widgets.FilterSelect("Status")
        self.sel_status.set_options(cv_schema.COURSE_STATUS_CHOICES)
        for w in (self.sel_course, self.sel_status):
            w.changed.connect(self._reload)
            filters.addWidget(w, 1)
        filters.addStretch(1)
        self._btn_scan = widgets.button(
            self._root, "Scan signatures (AI)", variant="info", icon="sparkles",
            command=self._scan_signatures)
        filters.addWidget(self._btn_scan)
        self._btn_export = widgets.button(
            self._root, "Export to Excel", variant="primary", icon="save",
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
            f"Showing {len(rows)} enrollments · "
            f"Total in DB: {repo.count_course_employees()}")

    # -------------------------------------------------------------- sửa
    def _edit(self, enrollment_id=None):
        if enrollment_id is None:
            return
        current = repo.get_enrollment(enrollment_id)
        specs = [
            {"key": "status", "label": "Learning status", "kind": "choice",
             "choices": cv_schema.COURSE_STATUS_CHOICES, "allow_empty": True},
        ]

        def _save(data):
            repo.update_enrollment(enrollment_id, {"status": data.get("status")})
            self._reload()

        FormDialog(self._root, "Update learning status",
                   specs, current, on_save=_save).run()

    # ---------------------------------------------------------- xuất Excel
    # Xuất roster khóa học theo template placeholder (file mẫu lấy ở màn Cài đặt).
    # CHỈ xuất các lượt ghi danh trạng thái "Not started" của khóa đang chọn — bỏ
    # qua ô lọc trạng thái trên UI (theo yêu cầu). Cột Dept dùng mã viết tắt bộ
    # phận (short_name), thiếu thì fallback về tên đầy đủ.
    def _export_roster(self):
        if not _OPENPYXL_OK:
            return dialogs.error(self._root, "Missing library",
                                 "openpyxl is required to export Excel:\n  pip install openpyxl")
        from app.core import excel_template

        tpl = (app_settings.get("course_template_path") or "").strip()
        if not tpl:
            return dialogs.error(
                self._root, "No template configured",
                "Go to Settings → Course roster export and choose a template file first.")
        if not os.path.isfile(tpl):
            return dialogs.error(self._root, "Template not found",
                                 f"The template file doesn't exist:\n{tpl}")

        course_id = self._course_opts.get(self.sel_course.value())
        if not course_id:
            return dialogs.info(self._root, "No course selected",
                                "Please choose a course to export.")

        not_started = cv_schema.COURSE_STATUS_CHOICES[0]   # "Not started"
        rows = repo.search_course_employees(course_id=course_id, status=not_started)
        if not rows:
            return dialogs.info(
                self._root, "No data",
                f'This course has no employees with status "{not_started}".')

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
            self._root, "Export course roster", f"{safe}{ext}", "Excel (*.xlsx *.xlsm)")
        if not path:
            return
        if not path.lower().endswith((".xlsx", ".xlsm")):
            path += ext

        try:
            excel_template.render_template(
                tpl, path, context, export_rows,
                row_fields={"code", "full_name", "department_short", "note", "stt"})
        except PermissionError:
            return dialogs.error(self._root, "Can't write file",
                                 f"Is the Excel file open? Close it and retry:\n{path}")
        except Exception as exc:   # noqa: BLE001 — báo lại cho UI, không chặn luồng
            return dialogs.error(self._root, "Excel export error", str(exc))

        dialogs.success(
            self._root, "Done",
            f'Exported {len(export_rows)} employees ("{not_started}") to:\n{path}')

    # ---------------------------------------------------- quét chữ ký bằng AI
    # HR scan cả tập bảng điểm danh đã ký thành MỘT file PDF nhiều trang → chọn
    # đúng một lần, app tự tách từng trang, gửi lần lượt cho Gemini nhận diện ai
    # đã ký → review → cập nhật status sang "Completed". Chỉ đối chiếu nhóm
    # "Not started" của khóa đang chọn (đúng đối tượng cần ký, khớp file đã xuất).
    def _scan_signatures(self):
        gen = app_settings.load()
        api_key = (gen.get("api_key") or "").strip()
        model = (gen.get("ai_model") or "").strip() or app_settings.DEFAULTS["ai_model"]
        if not api_key:
            return dialogs.error(
                self._root, "No API key",
                "Go to Settings → AI (Gemini) and enter your API key first.")

        course_id = self._course_opts.get(self.sel_course.value())
        if not course_id:
            return dialogs.info(self._root, "No course selected",
                                "Please choose a course to scan.")

        not_started = cv_schema.COURSE_STATUS_CHOICES[0]   # "Not started"
        roster = repo.search_course_employees(course_id=course_id, status=not_started)
        if not roster:
            return dialogs.info(
                self._root, "No data",
                f'This course has no employees with status "{not_started}".')

        # Lần quét trước còn trang lỗi của ĐÚNG khóa này và file vẫn còn → cho
        # quét tiếp mấy trang đó thay vì gửi lại cả tập.
        path, only = "", None
        pending = self._pending_scan(course_id)
        if pending:
            todo = pending["pages"]
            choice = dialogs.AppDialog(
                self._root, "Pages left to scan",
                f"The previous scan of this course still has {len(todo)} failed pages "
                f"(pages {', '.join(str(n) for n in todo)}) in file:\n"
                f"{os.path.basename(pending['path'])}\n\n"
                "Continue with those pages, or choose another file to scan from scratch?",
                "question",
                buttons=[("Cancel", "neutral", 0),
                         ("Choose another file", "neutral", 2),
                         (f"Scan {len(todo)} more pages", "primary", 1)]).run()
            if not choice:
                return
            if choice == 1:
                path, only = pending["path"], todo

        if not path:
            path, _ = QFileDialog.getOpenFileName(
                self._root, "Choose the signature scan file (image / multi-page PDF)", "",
                "Images & PDF (*.png *.jpg *.jpeg *.webp *.pdf)")
            if not path:
                return
        if not signature_scan.is_supported(path):
            return dialogs.error(self._root, "Unsupported format",
                                 "Only images (PNG/JPG/WEBP) or PDF are supported.")
        try:
            with open(path, "rb") as f:
                file_bytes = f.read()
        except Exception as exc:   # noqa: BLE001
            return dialogs.error(self._root, "Read error", str(exc))

        # Tách file thành ẢNH từng trang. PDF gửi thẳng dễ bị model gán chữ ký
        # lệch lên dòng trên (xem docstring app.core.signature_scan).
        try:
            pages = signature_scan.render_pages(
                file_bytes, signature_scan.guess_mime(path), only=only)
        except Exception as exc:   # noqa: BLE001
            return dialogs.error(self._root, "Couldn't split pages", str(exc))

        roster_min = [{"code": r["code"] or "", "full_name": r["full_name"] or ""}
                      for r in roster]
        total = len(pages)

        def job(ctx):
            # Quét TUẦN TỰ từng trang. Một trang lỗi thì bỏ qua trang đó và đi
            # tiếp (tập 20 trang không nên mất sạch vì 1 trang lỗi); danh sách
            # trang lỗi được báo lại ở bước review.
            page_rows, failed = [], []
            for i, (no, blob, page_mime) in enumerate(pages, start=1):
                if ctx.cancelled:
                    return page_rows, failed, True
                ctx.status(f"Page {no} ({i}/{total}) — sending to AI ({model})…")

                def on_retry(attempt, wait, reason, n=no):
                    ctx.log(f"… page {n}: {reason} — retry {attempt} in {wait}s")

                try:
                    data = signature_scan.detect_signatures(
                        api_key, model, blob, page_mime, roster_min,
                        on_retry=on_retry, should_cancel=lambda: ctx.cancelled)
                except signature_scan.Cancelled:
                    ctx.log(f"✋ Cancelled while processing page {no}.")
                    return page_rows, failed, True
                except Exception as exc:   # noqa: BLE001
                    failed.append((no, str(exc)))
                    ctx.log(f"⛔ Page {no} failed: {exc}")
                    ctx.step()
                    continue
                rows = data.get("rows", []) or []
                page_rows.append(rows)
                n_signed = sum(1 for r in rows if r.get("signed"))
                ctx.log(f"✅ Page {no}: {len(rows)} rows, {n_signed} signed")
                ctx.step()
            return page_rows, failed, False

        def on_finish(dlg, result):
            page_rows, failed, cancelled = result
            if cancelled:
                dlg.set_final_status("Cancelled — no statuses updated.")
                return
            if failed and not page_rows:
                dlg.set_final_status(f"No pages could be scanned ({len(failed)} errors).")
                dlg.log("👉 Wait a moment, then click 'Scan signatures (AI)' to try again.")
                return

            # Gộp kết quả các trang rồi đối chiếu theo MÃ NV với roster.
            merged = signature_scan.merge_pages(page_rows)
            by_code = {str(r["code"] or "").strip(): r
                       for r in roster if (r["code"] or "").strip()}
            signed, unmatched = [], []
            for pid, info in merged.items():
                if not info["signed"]:
                    continue
                if pid in by_code:
                    signed.append(by_code[pid])
                else:
                    unmatched.append((pid, info["name"]))
            signed_codes = {str(r["code"] or "").strip() for r in signed}
            unsigned = [r for r in roster
                        if str(r["code"] or "").strip() not in signed_codes]
            dlg.close()
            self._review_attendance(signed, unsigned, unmatched,
                                    [no for no, _ in failed], course_id, path)

        what = (f"scanning pages {', '.join(str(n) for n, _, _ in pages)}"
                if only else f"{total} pages")
        dlg = ProgressDialog(self._root, "Detecting signatures…", total=total,
                             subtitle=f"{os.path.basename(path)} — {what}, "
                                      f"with {model}")
        dlg.start(job, on_finish)

    # ------------------------------------------- ghi nhớ trang quét lỗi (resume)
    def _pending_scan(self, course_id):
        """Trang lỗi còn nợ của khóa `course_id`, hoặc None nếu không còn/không dùng được."""
        p = config.load(SECTION, {}).get(_PENDING_KEY) or {}
        pages = [int(n) for n in p.get("pages") or []]
        if (p.get("course_id") != course_id or not pages
                or not os.path.isfile(p.get("path") or "")):
            return None
        return {"path": p["path"], "pages": sorted(pages)}

    def _save_pending_scan(self, course_id, path, pages):
        """Ghi (hoặc xóa nếu `pages` rỗng) danh sách trang cần quét tiếp."""
        saved = config.load(SECTION, {})
        if pages:
            saved[_PENDING_KEY] = {"course_id": course_id, "path": path,
                                   "pages": sorted(int(n) for n in pages)}
        else:
            saved.pop(_PENDING_KEY, None)
        config.save(SECTION, saved)

    def _review_attendance(self, signed_rows, unsigned_rows, unmatched,
                           failed_pages, course_id, path):
        """Bảng review liệt kê người CHƯA ký (không phải người đã ký).

        Thực tế phần lớn nhân viên đều ký, nên danh sách "chưa ký" ngắn hơn nhiều
        — HR chỉ cần soát nhóm đó. Tick người thực ra CÓ ký mà AI bỏ sót; khi xác
        nhận, cả nhóm AI báo đã ký lẫn nhóm được tick thêm đều sang "Completed",
        số còn lại giữ nguyên "Not started".

        `failed_pages` là các trang chưa quét được. Người thuộc mấy trang đó cũng
        nằm trong danh sách "chưa ký" (vì chưa có dữ liệu), nên phải nói rõ để HR
        đừng tick bừa — và chỉ khi HR Xác nhận mới ghi nhớ mấy trang đó để lần sau
        quét tiếp (Hủy = không lưu gì, lần sau quét lại cả file).
        """
        completed = cv_schema.COURSE_STATUS_CHOICES[-1]   # "Completed"
        total = len(signed_rows) + len(unsigned_rows)
        dlg = ModalDialog(self._root, size="md")
        card, lay = dlg.build_shell("Confirm attendance")

        hint = QLabel(
            f"The AI found {len(signed_rows)}/{total} employees signed "
            f'→ they will move to "{completed}".\n'
            + (f"The {len(unsigned_rows)} people below have NO signature. Tick anyone "
               "who actually signed to include them, then click Confirm."
               if unsigned_rows else
               "No one is missing a signature — click Confirm to update."))
        hint.setObjectName("Hint")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        table = None
        if unsigned_rows:
            cols = [
                ("code",            "Emp code",   110, "w"),
                ("full_name",       "Full name",  200, "w"),
                ("department_name", "Department", 150, "w"),
            ]
            table = DataTable(cols, pk="enrollment_id", checkable=True)
            table.set_rows(unsigned_rows)   # KHÔNG tick sẵn — mặc định là chưa ký
            dlg.set_grow_region(table)
            lay.addWidget(table, 1)

        pages_txt = ", ".join(str(n) for n in failed_pages)
        for text in (
            (f"⚠ Pages {pages_txt} couldn't be scanned. People on those pages are "
             "in the list above (no data yet) — don't tick them. Click Confirm to "
             "save what was scanned, then click 'Scan signatures (AI)' again and choose "
             f"'Scan {len(failed_pages)} more pages'." if failed_pages else ""),
            ("The AI found signatures for codes not in this course (ignored): "
             + ", ".join(f"{p} {n}".strip() for p, n in unmatched[:15]))
            if unmatched else "",
        ):
            if not text:
                continue
            warn = QLabel(text)
            warn.setObjectName("Hint")
            warn.setWordWrap(True)
            lay.addWidget(warn)

        def _confirm():
            extra = table.checked_rows() if table is not None else []
            rows = list(signed_rows) + list(extra)
            for r in rows:
                repo.update_enrollment(r["enrollment_id"], {"status": completed})
            # Ghi nhớ trang lỗi để lần sau quét tiếp. Phần vừa cập nhật đã nằm
            # trong DB nên lần quét tiếp không cần đọc lại các trang đã xong.
            self._save_pending_scan(course_id, path, failed_pages)
            dlg.accept()
            self._reload()
            dialogs.success(
                self._root, "Done",
                f'Updated {len(rows)} employees to "{completed}"'
                + (f", including {len(extra)} you ticked." if extra else ".")
                + (f"\n\nPages {pages_txt} still aren't scanned — click "
                   "'Scan signatures (AI)' to continue." if failed_pages else ""))

        def _cancel():
            # Không cập nhật gì → cũng đừng ghi nhớ trang lỗi, vì lần sau phải
            # quét lại CẢ file mới có đủ dữ liệu.
            self._save_pending_scan(course_id, path, [])
            dlg.reject()

        foot = QHBoxLayout()
        foot.addStretch(1)
        foot.addWidget(widgets.button(card, "Cancel", variant="neutral", icon="x",
                                      command=_cancel))
        foot.addWidget(widgets.button(card, "Confirm", variant="success",
                                      icon="check", command=_confirm))
        lay.addLayout(foot)
        dlg.exec()
