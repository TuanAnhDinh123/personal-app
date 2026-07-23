"""Quét CV bằng AI (Gemini) → Excel — bản PySide6.

Logic gọi Gemini + ghi Excel nằm ở app.core.ai_cv_scan. Quét TUẦN TỰ, mỗi CV
thành công được ghi nối tiếp vào Excel và chuyển sang folder '…_da_quet' ngay;
gặp lỗi (giới hạn key free) thì dừng, lần sau bấm lại sẽ quét tiếp phần còn lại.
Luồng nền chạy qua QThread trong ProgressDialog dùng chung.
"""
import os
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QHBoxLayout, QLabel, QWidget

from app.core import config, settings
from app.core.ai_cv_scan import (
    _call_gemini, append_rows_to_excel, done_folder_for, move_to_done,
    read_jd_file,
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
DEFAULTS = {"folder": "", "output": "", "jd_file": "", "extra_prompt": ""}


class AiScanCvTool(BaseTool):
    name = "Quét CV bằng AI"
    description = "Dùng Gemini đọc PDF, đánh giá độ phù hợp với JD và xuất Excel."
    icon = "🤖"
    category = "Tuyển dụng"
    order = 20
    action_label = "Quét CV bằng AI"
    action_style = "success"
    action_icon = "sparkles"

    def build_body(self, parent):
        cfg = config.load(SECTION, DEFAULTS)
        gen = settings.load()

        # Hàng đầu: tiêu đề mục + chip model AI nép phải.
        head = QWidget(parent)
        h = QHBoxLayout(head)
        h.setContentsMargins(0, 2, 0, 4)
        lbl = QLabel("Đầu vào / đầu ra")
        lbl.setObjectName("SectionLabel")
        h.addWidget(lbl)
        h.addStretch(1)
        chip = QLabel(f"🤖  {gen['ai_model']}")
        chip.setStyleSheet(
            f"background: {theme.PALETTE['--accent-soft']}; color: {theme.ACCENT};"
            "border-radius: 9px; padding: 4px 10px; font-weight: 700; font-size: 11px;")
        h.addWidget(chip, 0, Qt.AlignRight)
        parent.layout().addWidget(head)

        self.var_folder = widgets.file_row(parent, "Thư mục chứa CV (PDF)", mode="folder")
        self.var_folder.set(cfg["folder"])
        self.var_output = widgets.file_row(parent, "Lưu file Excel kết quả tại", mode="save")
        self.var_output.set(cfg["output"])
        widgets.hint(parent, "Chọn file .xlsx kết quả. Nếu file đã tồn tại, kết quả sẽ "
                             "được GHI NỐI TIẾP (không đè lên record cũ).")

        widgets.section_label(parent, "File mô tả công việc (JD)")
        self.var_jd = widgets.file_row(parent, "Chọn file JD (PDF / DOCX / TXT)", mode="file")
        self.var_jd.set(cfg["jd_file"])
        widgets.hint(parent, "AI đọc nội dung file JD này để chấm độ phù hợp của mỗi CV.")

        widgets.section_label(parent, "Yêu cầu bổ sung cho AI (tùy chọn)")
        self.extra_box = widgets.text_area(
            parent, "Ví dụ: ưu tiên ứng viên biết tiếng Nhật, tối thiểu 2 năm kinh nghiệm…",
            value=cfg["extra_prompt"], height=5)

        widgets.hint(parent, "Mẹo: key Gemini miễn phí có giới hạn nhịp — khi quét nhiều CV mà "
                             "gặp lỗi giới hạn, app sẽ dừng và chuyển các CV đã quét sang folder "
                             "'…_da_quet'. Bấm quét lại để tiếp tục phần còn lại.")

    def run(self):
        # Bọc toàn bộ để mọi lỗi bất ngờ hiện ra hộp thoại thay vì "im lặng".
        # (App chạy bằng pythonw.exe — không có console, traceback sẽ mất tăm.)
        try:
            self._run_impl()
        except Exception as exc:
            import traceback
            self.error("Lỗi khi quét CV",
                       f"{exc}\n\n———\n{traceback.format_exc()}")

    def _run_impl(self):
        if not _OPENPYXL_OK:
            self.error("Thiếu thư viện", "Cần cài openpyxl để xuất Excel:\n  pip install openpyxl")
            return

        folder = self.var_folder.get().strip()
        out = self.var_output.get().strip()
        jd_file = self.var_jd.get().strip()
        extra = self.extra_box.get()

        gen = settings.load()
        api_key = gen.get("api_key", "").strip()
        model = gen.get("ai_model", "").strip() or settings.DEFAULTS["ai_model"]
        if not api_key:
            self.error("Thiếu API key",
                       "Chưa cấu hình API key Gemini.\n\nVào ⚙️ Cài đặt (cuối thanh bên) để nhập.")
            return
        if not folder or not os.path.isdir(folder):
            self.error("Thiếu thư mục", "Vui lòng chọn thư mục chứa CV.")
            return
        if not jd_file or not os.path.isfile(jd_file):
            self.error("Thiếu file JD", "Vui lòng chọn file mô tả công việc (PDF/DOCX/TXT).")
            return
        if not out:
            self.error("Thiếu đường dẫn", "Vui lòng chọn nơi lưu file Excel.")
            return
        if not out.lower().endswith(".xlsx"):
            out += ".xlsx"
        out_dir = os.path.dirname(out) or "."
        if not os.path.isdir(out_dir):
            self.error("Đường dẫn không hợp lệ", f"Thư mục không tồn tại:\n{out_dir}")
            return

        # Đọc nội dung JD từ file.
        try:
            jd = read_jd_file(jd_file)
        except Exception as exc:
            self.error("Không đọc được JD", f"Lỗi khi đọc file JD:\n{exc}")
            return
        if not jd.strip():
            self.error("JD rỗng", "File JD không có nội dung text (PDF scan ảnh?).")
            return

        files = sorted(p for p in Path(folder).iterdir()
                       if p.is_file() and p.suffix.lower() == ".pdf")
        if not files:
            self.info("Không có file",
                      "Không còn CV nào để quét trong thư mục.\n\n"
                      "Nếu vừa quét xong, các CV đã xử lý đã được chuyển sang folder '…_da_quet'.")
            return

        # Ghi NỐI vào section (không đè cả section) để không xóa mất các khóa
        # cũ như api_key/model từng lưu ở đây (settings đọc chung qua đó).
        saved = config.load(SECTION, {})
        saved.update({"folder": folder, "output": out,
                      "jd_file": jd_file, "extra_prompt": extra})
        config.save(SECTION, saved)
        self._scan(files, api_key, model, jd, extra, out, folder)

    def _scan(self, files, api_key, model, jd, extra, out, folder):
        total = len(files)
        done_dir = done_folder_for(folder)

        def job(ctx):
            # Quét TUẦN TỰ; mỗi CV thành công được ghi Excel + chuyển sang folder
            # 'đã quét' ngay lập tức, nên khi gặp lỗi (giới hạn key free) có thể
            # DỪNG mà không mất tiến độ — lần sau bấm lại sẽ quét tiếp phần còn lại.
            done, errors = [], []
            stopped_at = None
            for i, p in enumerate(files, start=1):
                if ctx.cancelled:
                    break
                ctx.status(f"({i}/{total}) {p.name}")

                def on_retry(attempt, wait, reason, name=p.name):
                    ctx.log(f"… {name}: {reason} — thử lại lần {attempt} sau {wait}s")

                try:
                    data = _call_gemini(api_key, model, jd, p.read_bytes(),
                                        on_retry=on_retry, extra=extra)
                except Exception as exc:
                    # Lỗi (thường là hết hạn mức) → DỪNG ngay tại đây.
                    errors.append(f"{p.name}: {exc}")
                    ctx.log(f"⛔ Dừng tại {p.name}: {exc}")
                    stopped_at = p.name
                    break

                data["file"] = p.name
                row = {k: (v if v is not None else "") for k, v in data.items()}
                try:
                    append_rows_to_excel([row], out)
                except PermissionError:
                    errors.append(f"{p.name}: file Excel đang mở")
                    ctx.log(f"⛔ Không ghi được Excel (đang mở trong Excel?): {out}")
                    stopped_at = p.name
                    break
                except Exception as exc:
                    errors.append(f"{p.name}: {exc}")
                    ctx.log(f"⛔ Lỗi ghi Excel: {exc}")
                    stopped_at = p.name
                    break

                try:
                    move_to_done(p, done_dir)
                except Exception as exc:
                    ctx.log(f"⚠ Đã ghi Excel nhưng không chuyển được file {p.name}: {exc}")
                done.append(p.name)
                ctx.log(f"✅ {p.name} — điểm {data.get('fit_score', '?')}")
                ctx.step()
            return done, errors, stopped_at

        def on_finish(dlg, result):
            done, errors, stopped_at = result
            remaining = total - len(done)
            if done:
                dlg.log(f"\n✅ Đã lưu {len(done)} CV vào:\n{out}")
                dlg.log(f"📂 Các CV đã quét được chuyển sang:\n{done_dir}")
            if stopped_at:
                dlg.set_final_status(
                    f"Dừng do lỗi — đã quét {len(done)}/{total}, còn {remaining} CV.")
                dlg.log(f"\n⛔ Dừng tại: {stopped_at}")
                dlg.log("👉 Chờ ít phút cho hạn mức key hồi lại rồi bấm 'Quét CV bằng AI' "
                        "để quét tiếp các CV còn lại.")
            elif not done:
                dlg.set_final_status("Không có CV nào được xử lý.")
            else:
                dlg.set_final_status(f"Hoàn thành — {len(done)}/{total} CV đã lưu.")

        dlg = ProgressDialog(self._page, "Đang quét CV bằng AI…", total=total,
                             subtitle=f"Quét {total} CV bằng {model}")
        dlg.start(job, on_finish)
