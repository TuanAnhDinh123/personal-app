"""Quét CV bằng AI (Gemini) → import thẳng vào bảng Candidates — bản PySide6.

JD dùng để chấm điểm KHÔNG chọn bằng tay nữa: người dùng chọn VỊ TRÍ tuyển dụng,
tool lấy file JD đã gắn cho vị trí đó (`positions.jd_file_path` — mỗi vị trí
đúng 1 JD, nhập ở Master Data → Vị trí tuyển dụng).

Logic gọi Gemini nằm ở app.core.ai_cv_scan. Quét TUẦN TỰ, mỗi CV thành công
được import thẳng vào DB (hoặc gom vào danh sách trùng nếu khớp email/SĐT một
ứng viên đã có) và chuyển sang folder '…_da_quet' ngay; gặp lỗi (giới hạn key
free) thì dừng, lần sau bấm lại sẽ quét tiếp phần còn lại. Ứng viên trùng được
hỏi lại sau khi quét xong: ghi đè bản ghi cũ, hoặc bỏ qua và xuất ra Excel.
Luồng nền chạy qua QThread trong ProgressDialog dùng chung.
"""
import os
import re
from datetime import date
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QLabel, QWidget

from app.core import config, cv_repository as repo, cv_schema, settings
from app.core.ai_cv_scan import (
    _call_gemini, _Cancelled, append_rows_to_excel, done_folder_for,
    flatten_result, move_to_done, read_jd_file, resolve_done_target,
)
from app_qt import theme, widgets
from app_qt.base_tool import BaseTool
from app_qt.components.cv_rename import RenameConfigDialog
from app_qt.components.modal import ModalDialog
from app_qt.components.progress_dialog import ProgressDialog
from app_qt.components.table import DataTable

try:
    import openpyxl  # noqa: F401
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

SECTION = "ai_scan_cv"
DEFAULTS = {"folder": "", "position_id": None, "extra_prompt": ""}


_SOURCE = "AI CV Scan"


def _read_cv_text(path):
    """Toàn văn CV để lưu vào DB — nguồn cho tìm kiếm & chấm lại sau này.

    Trích bằng thư viện, KHÔNG tốn thêm lượt gọi AI. Đọc hỏng (PDF ảnh scan,
    file lạ) thì trả về "" — chỉ mất tính năng tìm toàn văn, không chặn luồng.
    """
    try:
        from app.core.cv_scan import _extract_cv_text
        return _extract_cv_text(Path(path)) or ""
    except Exception:   # noqa: BLE001 — thiếu text không phải lỗi chặn quét
        return ""


def _num(value):
    """Chuỗi số/tiền do AI trả ("20,000,000 VND", "3.5 years") → float, hoặc None."""
    text = re.sub(r"[^\d.]", "", str(value or "").replace(",", ""))
    try:
        return float(text) if text else None
    except ValueError:
        return None


def _candidate_payload(profile, received_at):
    """Phần `profile` của AI → dict cột bảng `candidates`.

    Chỉ chứa dữ liệu TRUNG TÍNH (không dính JD nào). `experience_as_of` gắn
    bằng ngày nhận CV để về sau còn biết con số kinh nghiệm này đúng tại thời
    điểm nào.
    """
    salary_text = str(profile.get("expected_salary") or "").strip()
    available = repo.parse_date(profile.get("available_from"))
    return {
        "full_name":        profile.get("name") or "",
        "email":            profile.get("email") or "",
        "phone":            profile.get("phone") or "",
        "date_of_birth":    profile.get("dob") or "",
        "gender":           profile.get("gender") or "",
        "address":          profile.get("address") or "",
        "city":             profile.get("city") or "",
        "source":           _SOURCE,
        "current_title":    profile.get("current_title") or "",
        "industry":         profile.get("industry") or "",
        "years_experience": _num(profile.get("years_experience")),
        "experience_as_of": received_at,
        "education":        profile.get("education") or "",
        "major":            profile.get("major") or "",
        "languages":        profile.get("languages") or "",
        "skills_text":      "; ".join(profile.get("skills") or []),
        "profile_summary":  profile.get("profile_summary") or "",
        "expected_salary":  _num(salary_text),
        # Giữ nguyên chữ AI đọc được ("20tr gross", "thương lượng") — con số
        # bóc ra ở trên có thể mất ngữ cảnh.
        "salary_note":      salary_text,
        "available_from":   available.isoformat() if available else "",
    }


def _experience_rows(experiences, received_at):
    """Phần `experiences` của AI → list dict cột bảng `candidate_experiences`."""
    rows = []
    for i, item in enumerate(experiences or []):
        start = repo.parse_date(item.get("start_date"))
        if not start:
            continue          # không có mốc bắt đầu thì không tính được thời gian
        end = repo.parse_date(item.get("end_date"))
        rows.append({
            "company":     item.get("company") or "",
            "job_title":   item.get("job_title") or "",
            "industry":    item.get("industry") or "",
            "start_date":  start.isoformat(),
            "end_date":    end.isoformat() if end else "",
            "as_of_date":  received_at,
            "is_current":  1 if (item.get("is_current") or not end) else 0,
            "description": item.get("description") or "",
            "sort_order":  i,
        })
    return rows


def _save_result(data, position_id, model, extra, jd_hash, candidate_id=None):
    """Ghi MỘT kết quả quét vào DB — trả về candidate_id.

    `candidate_id` có giá trị = thêm BẢN CV MỚI cho người đã có trong pool (hồ
    sơ được cập nhật theo bản mới nhất, bản cũ vẫn giữ nguyên); None = tạo ứng
    viên mới.

    Một lượt quét chạm 6 bảng: candidates · candidate_cvs · candidate_experiences
    · candidate_skills · applications · candidate_evaluations.
    """
    profile = data.get("profile") or {}
    evaluation = data.get("evaluation") or {}
    received_at = date.today().isoformat()
    payload = _candidate_payload(profile, received_at)

    if candidate_id:
        # Người đã có: chỉ cập nhật những ô AI đọc được, không xóa dữ liệu cũ
        # bằng chuỗi rỗng (bản CV mới có thể thiếu thông tin bản cũ đã có).
        repo.update_candidate(candidate_id, {k: v for k, v in payload.items() if v})
    else:
        candidate_id = repo.insert_candidate(payload)

    cv_id = repo.add_candidate_cv(candidate_id, {
        "file_path":   data.get("cv_path") or "",
        "file_hash":   data.get("file_hash") or "",
        "received_at": received_at,
        "batch":       data.get("batch"),
        "source":      _SOURCE,
        "cv_text":     data.get("cv_text") or "",
        "scanned_at":  data.get("scanned_at"),
        "scan_model":  model,
        **{k: v for k, v in payload.items() if k in repo.PROFILE_SNAPSHOT_FIELDS},
    })
    repo.replace_candidate_experiences(
        candidate_id, cv_id, _experience_rows(data.get("experiences"), received_at))
    repo.replace_candidate_skills(candidate_id, cv_id, profile.get("skills") or [])

    application_id = repo.ensure_application(candidate_id, position_id, {
        "cv_id": cv_id, "origin": "Applied", "source": _SOURCE,
        "status": cv_schema.CANDIDATE_STATUS_DEFAULT,
    })
    repo.insert_evaluation({
        "candidate_id":   candidate_id,
        "position_id":    position_id,
        "application_id": application_id,
        "cv_id":          cv_id,
        "source":         "CV scan",
        "ai_score":       evaluation.get("fit_score"),
        "matched_skills": evaluation.get("matched_skills") or "",
        "missing_skills": evaluation.get("missing_skills") or "",
        "summary":        evaluation.get("fit_summary") or "",
        "strengths":      evaluation.get("strengths") or "",
        "weaknesses":     evaluation.get("weaknesses") or "",
        "model":          model,
        "jd_hash":        jd_hash,
        "extra_prompt":   extra,
    })
    return candidate_id


class _DuplicatesDialog(ModalDialog):
    """Modal xử lý ứng viên trùng sau khi quét AI (khớp email/SĐT với DB).

    `duplicates` = list các tuple (ai_data, existing_row).
    Trả về "attach" / "separate" / "export" / None (hủy) qua .run().
    """

    def __init__(self, parent, duplicates):
        super().__init__(parent, "md")
        self._result = None
        card, lay = self.build_shell(f"Already in the pool · {len(duplicates)}")

        desc = QLabel("These CVs belong to people who already share an email or "
                      "phone with someone in the talent pool:")
        desc.setObjectName("DialogMsg")
        desc.setWordWrap(True)
        lay.addWidget(desc)

        rows = []
        for data, existing in duplicates:
            flat = flatten_result(data)
            rows.append({
                "full_name": flat.get("name", ""),
                "email": flat.get("email", ""),
                "phone": flat.get("phone", ""),
                "existing": f"#{existing['candidate_id']} {existing['full_name'] or ''}",
            })
        table = DataTable([
            ("full_name", "From this CV", 190),
            ("email", "Email", 200),
            ("phone", "Phone", 110),
            ("existing", "Already in pool", 190),
        ])
        table.set_rows(rows)
        table.setMinimumHeight(min(320, self.modal_h))
        lay.addWidget(table, 1)
        self.set_grow_region(table)

        hint = QLabel(
            "Adding a new CV version keeps every earlier CV, AI assessment and "
            "contact history, and refreshes the profile from this newer CV — "
            "that is normally what you want when someone reapplies. Create "
            "separate candidates only if these are genuinely different people.")
        hint.setObjectName("Hint")
        hint.setWordWrap(True)
        lay.addWidget(hint)

        foot = QHBoxLayout()
        foot.addWidget(widgets.button(card, "Add as new CV version", variant="success",
                                      icon="check",
                                      command=lambda: self._choose("attach")))
        foot.addWidget(widgets.button(card, "Create separate candidates",
                                      variant="primary", icon="plus",
                                      command=lambda: self._choose("separate")))
        foot.addWidget(widgets.button(card, "Skip · export to Excel", variant="warning",
                                      icon="save",
                                      command=lambda: self._choose("export")))
        foot.addWidget(widgets.button(card, "Cancel", variant="neutral", icon="x",
                                      command=self.reject))
        foot.addStretch(1)
        lay.addLayout(foot)

    def _choose(self, value):
        self._result = value
        self.accept()

    def run(self):
        self.exec()
        return self._result


class _PositionCombo(widgets.ComboBox):
    """Ô chọn VỊ TRÍ tuyển dụng — file JD lấy theo vị trí (positions.jd_file_path).

    Danh sách nạp lại từ DB mỗi lần bấm mở, nên vị trí/JD vừa thêm ở trang
    "Vị trí tuyển dụng" là thấy ngay, không cần mở lại tool.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._rows = {}          # nhãn hiển thị → sqlite3.Row của vị trí
        self.on_open = self.reload   # hook của widgets.ComboBox: nạp lại khi bung
        self.reload()

    @staticmethod
    def _label(row):
        """'Backend Dev · IT — JD: jd_backend.pdf' (hoặc '— chưa có file JD')."""
        title = (row["position_title"] or f"#{row['position_id']}").strip()
        dept = (row["department_name"] or "").strip()
        jd = (row["jd_file_path"] or "").strip()
        left = f"{title} · {dept}" if dept else title
        return f"{left}  —  JD: {os.path.basename(jd)}" if jd \
            else f"{left}  —  ⚠ no JD file"

    def reload(self):
        keep = self.currentText()
        try:
            rows = repo.list_positions()
        except Exception:   # noqa: BLE001 — DB lỗi/chưa có: để list rỗng, báo khi chạy
            rows = []
        self._rows = {self._label(r): r for r in rows}
        self.blockSignals(True)
        self.clear()
        self.addItems(list(self._rows))
        self.blockSignals(False)
        self.select_text(keep)

    def select_text(self, text):
        i = self.findText(text) if text else -1
        if i >= 0:
            self.setCurrentIndex(i)

    def select_position(self, pos_id):
        """Chọn lại vị trí theo id (dùng khi mở tool: nhớ lựa chọn lần trước)."""
        if pos_id in (None, ""):
            return
        for label, row in self._rows.items():
            if str(row["position_id"]) == str(pos_id):
                self.select_text(label)
                return

    def current_row(self):
        """sqlite3.Row của vị trí đang chọn (None nếu chưa chọn / danh sách rỗng)."""
        return self._rows.get(self.currentText())


class AiScanCvTool(BaseTool):
    name = "AI CV Scan"
    description = "Use Gemini to read PDFs, score fit against a JD, and import results to Database."
    icon = "🤖"
    category = "Recruitment"
    order = 20
    action_label = "Scan CVs with AI"
    action_style = "success"
    action_icon = "sparkles"

    def build_body(self, parent):
        cfg = config.load(SECTION, DEFAULTS)
        gen = settings.load()

        # Hàng đầu: tiêu đề mục + chip model AI nép phải.
        head = QWidget(parent)
        h = QHBoxLayout(head)
        h.setContentsMargins(0, 2, 0, 2)
        lbl = QLabel("Input / output")
        lbl.setObjectName("SectionLabel")
        h.addWidget(lbl)
        h.addStretch(1)
        chip = QLabel(f"🤖  {gen['ai_model']}")
        chip.setStyleSheet(
            f"background: {theme.PALETTE['--accent-soft']}; color: {theme.ACCENT};"
            "border-radius: 9px; padding: 4px 10px; font-weight: 700; font-size: 11px;")
        h.addWidget(chip, 0, Qt.AlignRight)
        parent.layout().addWidget(head)

        self.var_folder = widgets.file_row(parent, "CV folder (PDF/DOCX)", mode="folder")
        self.var_folder.set(cfg["folder"])

        # JD gắn theo VỊ TRÍ (mỗi vị trí đúng 1 JD) → chỉ cần chọn vị trí, file JD
        # lấy từ positions.jd_file_path.
        widgets.section_label(parent, "Position")
        repo.init_db()
        block, v = widgets._field_block(parent, "Choose a position (JD comes from the position)")
        self.cbo_pos = _PositionCombo(block)
        v.addWidget(self.cbo_pos)
        self.cbo_pos.select_position(cfg["position_id"])
        widgets.hint(parent, "The AI reads this position's JD to score each CV's fit. "
                             "Add/edit positions & JD files under Master Data → Positions.")

        widgets.section_label(parent, "Extra instructions for the AI (optional)")
        self.extra_box = widgets.text_area(
            parent, "e.g. prefer English speakers, at least 2 years' experience…",
            value=cfg["extra_prompt"], height=7)

        widgets.hint(parent, "Each CV that scans successfully is imported into the Candidates "
                             "database and moved to a '…_scanned' folder, so if it stops "
                             "midway you can click again to continue with the rest.")
        widgets.hint(parent, "Tip: use “Normalize file names” to batch-rename the CV files "
                             "in the folder ({code}_{Candidate name}) before scanning.")

    def extra_actions(self, parent):
        """Nút phụ: mở hộp chuẩn hóa tên file CV trước khi quét."""
        return [widgets.button(parent, "Normalize file names", variant="neutral",
                               icon="pencil", command=self._open_rename)]

    def _open_rename(self):
        RenameConfigDialog(self._page, self.var_folder.get().strip()).exec()

    def run(self):
        # Bọc toàn bộ để mọi lỗi bất ngờ hiện ra hộp thoại thay vì "im lặng".
        # (App chạy bằng pythonw.exe — không có console, traceback sẽ mất tăm.)
        try:
            self._run_impl()
        except Exception as exc:
            import traceback
            self.error("CV scan error",
                       f"{exc}\n\n———\n{traceback.format_exc()}")

    def _run_impl(self):
        if not _OPENPYXL_OK:
            self.error("Missing library", "openpyxl is required to export duplicates:\n  pip install openpyxl")
            return

        folder = self.var_folder.get().strip()
        extra = self.extra_box.get()
        # Vị trí đang chọn → file JD của vị trí đó (positions.jd_file_path).
        pos = self.cbo_pos.current_row()
        pos_title = (pos["position_title"] or f"#{pos['position_id']}").strip() if pos else ""
        jd_file = (pos["jd_file_path"] or "").strip() if pos else ""

        gen = settings.load()
        api_key = gen.get("api_key", "").strip()
        model = gen.get("ai_model", "").strip() or settings.DEFAULTS["ai_model"]
        if not api_key:
            self.error("Missing API key",
                       "No Gemini API key configured.\n\nOpen ⚙️ Settings (bottom of the sidebar) to add one.")
            return
        if not folder or not os.path.isdir(folder):
            self.error("Missing folder", "Please choose the CV folder.")
            return
        if pos is None:
            self.error("No position selected",
                       "Please choose a position to pull its JD from.\n\n"
                       "No positions yet? Go to Master Data → Positions "
                       "to add one with a JD file.")
            return
        if not jd_file:
            self.error("Position has no JD",
                       f'The position "{pos_title}" has no JD file attached.\n\n'
                       "Go to Master Data → Positions, edit this position and "
                       "choose a JD file (PDF/DOCX/TXT).")
            return
        if not os.path.isfile(jd_file):
            self.error("JD file not found",
                       f'The JD file for "{pos_title}" is no longer at the saved path:\n'
                       f"{jd_file}\n\nIt may have been moved/renamed — go to "
                       "Master Data → Positions to re-select it.")
            return

        # Đọc nội dung JD từ file.
        try:
            jd = read_jd_file(jd_file)
        except Exception as exc:
            self.error("Can't read JD",
                       f"Error reading the JD file for \"{pos_title}\":\n{jd_file}\n\n{exc}")
            return
        if not jd.strip():
            self.error("Empty JD",
                       f'The JD file for "{pos_title}" has no text '
                       f"(scanned-image PDF?):\n{jd_file}")
            return

        files = sorted(p for p in Path(folder).iterdir()
                       if p.is_file() and p.suffix.lower() == ".pdf")
        if not files:
            self.info("No files",
                      "No CVs left to scan in this folder.\n\n"
                      "If you just finished, processed CVs were moved to the '…_scanned' folder.")
            return

        # Ghi NỐI vào section (không đè cả section) để không xóa mất các khóa
        # cũ như api_key/model từng lưu ở đây (settings đọc chung qua đó).
        saved = config.load(SECTION, {})
        saved.update({"folder": folder,
                      "position_id": pos["position_id"], "extra_prompt": extra})
        config.save(SECTION, saved)
        self._scan(files, api_key, model, jd, extra, pos["position_id"], folder, pos_title)

    def _scan(self, files, api_key, model, jd, extra, position_id, folder, pos_title=""):
        total = len(files)
        done_dir = done_folder_for(folder)
        # Băm nội dung JD → lưu kèm mỗi lượt chấm, sau này JD sửa thì biết ngay
        # điểm cũ đã không còn khớp bản JD hiện hành.
        jd_hash = repo.text_hash(jd)
        # 'batch' = SỐ bóc ra từ tên thư mục chứa CV (vd "batch1" → 1). Không có
        # số trong tên → để trống.
        m = re.search(r"\d+", os.path.basename(os.path.normpath(folder)))
        batch = int(m.group()) if m else None

        def job(ctx):
            # Quét TUẦN TỰ; mỗi CV thành công được import thẳng vào DB (hoặc gom
            # vào danh sách trùng) + chuyển sang folder 'đã quét' NGAY, nên khi
            # gặp lỗi (giới hạn key free) có thể DỪNG mà không mất tiến độ — lần
            # sau bấm lại sẽ quét tiếp phần còn lại.
            done, duplicates, errors, skipped = [], [], [], []
            stopped_at = None
            cancelled = False
            for i, p in enumerate(files, start=1):
                if ctx.cancelled:
                    cancelled = True
                    break
                ctx.status(f"({i}/{total}) {p.name}")

                # Cùng một file CV (kể cả đã đổi tên) từng quét rồi thì bỏ qua
                # ngay — không tốn thêm một lượt gọi AI nào.
                digest = repo.file_hash(p)
                if digest and repo.find_cv_by_hash(digest) is not None:
                    ctx.log(f"⏭ {p.name} — this exact CV was already scanned.")
                    skipped.append(p.name)
                    try:
                        move_to_done(p, done_dir)
                    except Exception as exc:
                        ctx.log(f"⚠ Couldn't move {p.name}: {exc}")
                    ctx.step()
                    continue

                def on_retry(attempt, wait, reason, name=p.name):
                    ctx.log(f"… {name}: {reason} — retry {attempt} in {wait}s")

                try:
                    data = _call_gemini(api_key, model, jd, p.read_bytes(),
                                        on_retry=on_retry, extra=extra,
                                        should_cancel=lambda: ctx.cancelled)
                except _Cancelled:
                    # Bấm Hủy giữa chừng — CV này CHƯA xong, để lại cho lần sau.
                    ctx.log(f"✋ Cancelled while processing {p.name} (not saved).")
                    cancelled = True
                    break
                except Exception as exc:
                    # Lỗi (thường là hết hạn mức) → DỪNG ngay tại đây.
                    errors.append(f"{p.name}: {exc}")
                    ctx.log(f"⛔ Stopped at {p.name}: {exc}")
                    stopped_at = p.name
                    break

                data["file"] = p.name
                data["batch"] = batch
                data["file_hash"] = digest
                data["scanned_at"] = repo.now_text()
                data["cv_text"] = _read_cv_text(p)
                # Tính TRƯỚC nơi file CV sẽ nằm sau khi quét (folder '…_scanned')
                # để ghi luôn đường dẫn đầy đủ vào DB — lúc sau không cần hỏi lại
                # thư mục CV nữa.
                target = resolve_done_target(p, done_dir)
                data["cv_path"] = str(target)
                profile = data.get("profile") or {}

                try:
                    existing = repo.find_duplicates(profile.get("email"),
                                                    profile.get("phone"))
                except Exception as exc:
                    errors.append(f"{p.name}: {exc}")
                    ctx.log(f"⛔ DB error while checking the pool: {exc}")
                    stopped_at = p.name
                    break

                if existing:
                    # Đã có người này trong pool → CHƯA ghi, gom lại để hỏi cuối
                    # lượt: thêm bản CV mới cho họ, hay tạo hồ sơ riêng.
                    duplicates.append((data, existing[0]))
                    ctx.log(f"⚠ {p.name} — already in the pool as "
                            f"#{existing[0]['candidate_id']}")
                else:
                    try:
                        _save_result(data, position_id, model, extra, jd_hash)
                    except Exception as exc:
                        errors.append(f"{p.name}: {exc}")
                        ctx.log(f"⛔ DB write error: {exc}")
                        stopped_at = p.name
                        break
                    score = (data.get("evaluation") or {}).get("fit_score", "?")
                    ctx.log(f"✅ {p.name} — score {score}")

                try:
                    move_to_done(p, done_dir, target=target)
                except Exception as exc:
                    ctx.log(f"⚠ Processed but couldn't move {p.name}: {exc}")
                done.append(p.name)
                ctx.step()
            return done, duplicates, errors, stopped_at, cancelled, skipped

        def on_finish(dlg, result):
            done, duplicates, errors, stopped_at, cancelled, skipped = result
            added = len(done) - len(duplicates)
            remaining = total - len(done) - len(skipped)
            if skipped:
                dlg.log(f"\n⏭ Skipped {len(skipped)} CV(s) already scanned before.")
            if done:
                dlg.log(f"\n✅ Imported {added} candidate(s) into the database.")
                if duplicates:
                    dlg.log(f"⚠ {len(duplicates)} CV(s) belong to people already in "
                            "the pool — you'll be asked how to handle them next.")
                dlg.log(f"📂 Scanned CVs moved to:\n{done_dir}")
            if stopped_at:
                dlg.set_final_status(
                    f"Stopped on error — scanned {len(done)}/{total}, {remaining} left.")
                dlg.log(f"\n⛔ Stopped at: {stopped_at}")
                dlg.log("👉 Wait a moment, then click 'Scan CVs with AI' "
                        "to continue with the remaining CVs.")
            elif cancelled:
                dlg.set_final_status(
                    f"Cancelled — scanned {len(done)}/{total}, {remaining} left.")
                dlg.log("\n✋ Cancelled. Scanned CVs were saved and moved "
                        "to the '…_scanned' folder.")
                if remaining:
                    dlg.log("👉 Click 'Scan CVs with AI' to continue with the rest.")
            elif not done:
                dlg.set_final_status("No CVs were processed.")
            else:
                dlg.set_final_status(f"Done — {len(done)}/{total} CVs processed.")
            if duplicates:
                self._handle_duplicates(duplicates, position_id, model, extra, jd_hash)

        subtitle = f"Scanning {total} CVs with {model}"
        if pos_title:
            subtitle += f" · position JD: {pos_title}"
        dlg = ProgressDialog(self._page, "Scanning CVs with AI…", total=total,
                             subtitle=subtitle)
        dlg.start(job, on_finish)

    # -------------------------------------------- xử lý ứng viên trùng sau khi quét
    def _handle_duplicates(self, duplicates, position_id, model, extra, jd_hash):
        choice = _DuplicatesDialog(self._page, duplicates).run()
        if choice in ("attach", "separate"):
            failed = []
            for data, existing in duplicates:
                cid = existing["candidate_id"] if choice == "attach" else None
                try:
                    _save_result(data, position_id, model, extra, jd_hash, cid)
                except Exception as exc:   # noqa: BLE001 — gom lỗi báo một lần
                    who = (data.get("profile") or {}).get("name") or "(no name)"
                    failed.append(f"{who} — {exc}")
            saved = len(duplicates) - len(failed)
            if failed:
                self.error("Some CVs weren't saved",
                           f"Saved {saved}/{len(duplicates)}.\n\n• " + "\n• ".join(failed))
            elif choice == "attach":
                self.info("Pool updated",
                          f"Added a new CV version to {saved} existing candidate(s). "
                          "Their earlier CVs and assessments are unchanged.")
            else:
                self.info("Added", f"Created {saved} separate candidate(s).")
        elif choice == "export":
            path, _ = QFileDialog.getSaveFileName(
                self._page, "Save these candidates to Excel",
                "Duplicate_candidates.xlsx", "Excel (*.xlsx)")
            if not path:
                self.info("Not saved", "Export cancelled — these results "
                                       "weren't saved anywhere.")
                return
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            try:
                append_rows_to_excel([flatten_result(d) for d, _existing in duplicates],
                                     path)
            except Exception as exc:
                self.error("Excel export error", str(exc))
                return
            self.info("Exported",
                      f"Saved {len(duplicates)} candidate(s) to:\n{path}")
