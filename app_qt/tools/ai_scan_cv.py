"""Quét CV bằng AI (Gemini) → Excel — bản PySide6.

Logic gọi Gemini + ghi Excel (_call_gemini, _write_excel) dùng lại nguyên từ
module Tk cũ. Luồng nền chuyển sang QThread qua ProgressDialog dùng chung.
"""
import os
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QHBoxLayout, QLabel, QWidget

from app.core import config, settings
from app.core.ai_cv_scan import _call_gemini, _write_excel
from app_qt import dialogs, theme, widgets
from app_qt.base_tool import BaseTool
from app_qt.components.progress_dialog import ProgressDialog

try:
    import openpyxl  # noqa: F401
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

SECTION = "ai_scan_cv"
DEFAULTS = {"folder": "", "output": "", "jd": ""}


class AiScanCvTool(BaseTool):
    name = "Quét CV bằng AI"
    description = "Dùng Gemini đọc PDF, đánh giá độ phù hợp với JD và xuất Excel."
    icon = "🤖"
    category = "Tệp & Tài liệu"
    order = 11
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
        widgets.hint(parent, "Chọn đường dẫn file .xlsx sẽ tạo. Thiếu đuôi .xlsx sẽ tự thêm.")

        widgets.section_label(parent, "Yêu cầu công việc (JD)")
        self.jd_box = widgets.text_area(
            parent, "Dán mô tả công việc / tiêu chí tuyển dụng để AI chấm độ phù hợp:",
            value=cfg["jd"], height=8)

    def run(self):
        if not _OPENPYXL_OK:
            self.error("Thiếu thư viện", "Cần cài openpyxl để xuất Excel:\n  pip install openpyxl")
            return

        folder = self.var_folder.get().strip()
        out = self.var_output.get().strip()
        jd = self.jd_box.get()

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
        if not out:
            self.error("Thiếu đường dẫn", "Vui lòng chọn nơi lưu file Excel.")
            return
        if not out.lower().endswith(".xlsx"):
            out += ".xlsx"
        out_dir = os.path.dirname(out) or "."
        if not os.path.isdir(out_dir):
            self.error("Đường dẫn không hợp lệ", f"Thư mục không tồn tại:\n{out_dir}")
            return

        files = sorted(p for p in Path(folder).iterdir()
                       if p.is_file() and p.suffix.lower() == ".pdf")
        if not files:
            self.info("Không có file", "Không tìm thấy file PDF nào trong thư mục đã chọn.")
            return

        config.save(SECTION, {"folder": folder, "output": out, "jd": jd})
        self._scan(files, api_key, model, jd, out)

    def _scan(self, files, api_key, model, jd, out):
        total = len(files)

        def job(ctx):
            rows, errors = [], []
            for i, p in enumerate(files, start=1):
                if ctx.cancelled:
                    break
                ctx.status(f"({i}/{total}) {p.name}")

                def on_retry(attempt, wait, reason, name=p.name):
                    ctx.log(f"… {name}: {reason} — thử lại lần {attempt} sau {wait}s")

                try:
                    data = _call_gemini(api_key, model, jd, p.read_bytes(), on_retry=on_retry)
                    data["file"] = p.name
                    rows.append({k: (v if v is not None else "") for k, v in data.items()})
                    ctx.log(f"✅ {p.name} — điểm {data.get('fit_score', '?')}")
                except Exception as exc:
                    errors.append(f"{p.name}: {exc}")
                    ctx.log(f"⚠ {p.name}: {exc}")
                ctx.step()
            return rows, errors

        def on_finish(dlg, result):
            rows, errors = result
            if not rows:
                dlg.set_final_status("Không có CV nào được xử lý.")
                if errors:
                    dlg.log("— Không ghi Excel vì không có kết quả hợp lệ.")
                return
            try:
                _write_excel(rows, out)
            except PermissionError:
                dlg.log(f"⚠ Không ghi được file (đang mở trong Excel?): {out}")
                dlg.set_final_status("Lỗi ghi file.")
                return
            except Exception as exc:
                dlg.log(f"⚠ Lỗi ghi Excel: {exc}")
                dlg.set_final_status("Lỗi ghi file.")
                return
            dlg.set_final_status(f"Hoàn thành — {len(rows)}/{total} CV đã lưu.")
            dlg.log(f"\n✅ Đã lưu {len(rows)} CV vào:\n{out}")
            if errors:
                dlg.log(f"⚠ {len(errors)} file lỗi (xem nhật ký).")

        dlg = ProgressDialog(self._page, "Đang quét CV bằng AI…", total=total,
                             subtitle=f"Quét {total} CV bằng {model}")
        dlg.start(job, on_finish)
