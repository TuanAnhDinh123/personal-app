"""Quét CV bằng AI (Gemini) → Excel — bản PySide6.

JD dùng để chấm điểm KHÔNG chọn bằng tay nữa: người dùng chọn VỊ TRÍ tuyển dụng,
tool lấy file JD đã gắn cho vị trí đó (`positions.jd_file_path` — mỗi vị trí
đúng 1 JD, nhập ở Master Data → Vị trí tuyển dụng).

Logic gọi Gemini + ghi Excel nằm ở app.core.ai_cv_scan. Quét TUẦN TỰ, mỗi CV
thành công được ghi nối tiếp vào Excel và chuyển sang folder '…_da_quet' ngay;
gặp lỗi (giới hạn key free) thì dừng, lần sau bấm lại sẽ quét tiếp phần còn lại.
Luồng nền chạy qua QThread trong ProgressDialog dùng chung.
"""
import os
import re
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QHBoxLayout, QLabel, QWidget

from app.core import config, cv_repository as repo, settings
from app.core.ai_cv_scan import (
    _call_gemini, _Cancelled, append_rows_to_excel, done_folder_for,
    move_to_done, read_jd_file, resolve_done_target,
)
from app_qt import dialogs, theme, widgets
from app_qt.base_tool import BaseTool
from app_qt.components.progress_dialog import ProgressDialog

try:
    import openpyxl  # noqa: F401
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

SECTION = "ai_scan_cv"
DEFAULTS = {"folder": "", "output": "", "position_id": None, "extra_prompt": ""}


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
        h.setContentsMargins(0, 2, 0, 4)
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
        self.var_output = widgets.file_row(parent, "Save result Excel to", mode="save")
        self.var_output.set(cfg["output"])
        widgets.hint(parent, "Pick a result .xlsx file. If it already exists, results are "
                             "appended (existing records aren't overwritten).")

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
            value=cfg["extra_prompt"], height=5)

        widgets.hint(parent, "Each CV that scans successfully is appended to the Excel and "
                             "moved to a '…_scanned' folder, so if it stops midway you can "
                             "click again to continue with the rest.")

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
            self.error("Missing library", "openpyxl is required to export Excel:\n  pip install openpyxl")
            return

        folder = self.var_folder.get().strip()
        out = self.var_output.get().strip()
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
        if not out:
            self.error("Missing path", "Please choose where to save the Excel file.")
            return
        if not out.lower().endswith(".xlsx"):
            out += ".xlsx"
        out_dir = os.path.dirname(out) or "."
        if not os.path.isdir(out_dir):
            self.error("Invalid path", f"Folder doesn't exist:\n{out_dir}")
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
        saved.update({"folder": folder, "output": out,
                      "position_id": pos["position_id"], "extra_prompt": extra})
        config.save(SECTION, saved)
        self._scan(files, api_key, model, jd, extra, out, folder, pos_title)

    def _scan(self, files, api_key, model, jd, extra, out, folder, pos_title=""):
        total = len(files)
        done_dir = done_folder_for(folder)
        # 'batch' = SỐ bóc ra từ tên thư mục chứa CV (vd "batch1" → 1). Không có
        # số trong tên → để trống.
        m = re.search(r"\d+", os.path.basename(os.path.normpath(folder)))
        batch = int(m.group()) if m else None

        def job(ctx):
            # Quét TUẦN TỰ; mỗi CV thành công được ghi Excel + chuyển sang folder
            # 'đã quét' ngay lập tức, nên khi gặp lỗi (giới hạn key free) có thể
            # DỪNG mà không mất tiến độ — lần sau bấm lại sẽ quét tiếp phần còn lại.
            done, errors = [], []
            stopped_at = None
            cancelled = False
            for i, p in enumerate(files, start=1):
                if ctx.cancelled:
                    cancelled = True
                    break
                ctx.status(f"({i}/{total}) {p.name}")

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
                # Tính TRƯỚC nơi file CV sẽ nằm sau khi quét (folder '…_da_quet')
                # để ghi luôn đường dẫn đầy đủ vào Excel — lúc import không cần
                # hỏi lại thư mục CV nữa.
                target = resolve_done_target(p, done_dir)
                data["cv_path"] = str(target)
                row = {k: (v if v is not None else "") for k, v in data.items()}
                try:
                    append_rows_to_excel([row], out)
                except PermissionError:
                    errors.append(f"{p.name}: Excel file is open")
                    ctx.log(f"⛔ Couldn't write Excel (open in Excel?): {out}")
                    stopped_at = p.name
                    break
                except Exception as exc:
                    errors.append(f"{p.name}: {exc}")
                    ctx.log(f"⛔ Excel write error: {exc}")
                    stopped_at = p.name
                    break

                try:
                    move_to_done(p, done_dir, target=target)
                except Exception as exc:
                    ctx.log(f"⚠ Wrote Excel but couldn't move {p.name}: {exc}")
                done.append(p.name)
                ctx.log(f"✅ {p.name} — score {data.get('fit_score', '?')}")
                ctx.step()
            return done, errors, stopped_at, cancelled

        def on_finish(dlg, result):
            done, errors, stopped_at, cancelled = result
            remaining = total - len(done)
            if done:
                dlg.log(f"\n✅ Saved {len(done)} CVs to:\n{out}")
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
                dlg.log("\n✋ Cancelled. Scanned CVs were saved to Excel and moved "
                        "to the '…_scanned' folder.")
                if remaining:
                    dlg.log("👉 Click 'Scan CVs with AI' to continue with the rest.")
            elif not done:
                dlg.set_final_status("No CVs were processed.")
            else:
                dlg.set_final_status(f"Done — {len(done)}/{total} CVs saved.")

        subtitle = f"Scanning {total} CVs with {model}"
        if pos_title:
            subtitle += f" · position JD: {pos_title}"
        dlg = ProgressDialog(self._page, "Scanning CVs with AI…", total=total,
                             subtitle=subtitle)
        dlg.start(job, on_finish)
