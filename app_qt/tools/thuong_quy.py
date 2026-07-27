"""Tổng hợp lỗi chấm công theo quý (Thưởng Quý) — bản PySide6.

Logic điều khiển Excel qua COM (_aggregate_quarter) tách riêng ở
app.core.quarter_bonus — file này chỉ dựng giao diện bằng Qt.
"""
import os

from app.core import config
from app.core.quarter_bonus import QUARTER_MONTHS, _aggregate_quarter
from app_qt import dialogs, widgets
from app_qt.base_tool import BaseTool

_SECTION = "thuong_quy"

_DEFAULTS = {
    "source": "",
    "quarter": "Q4",
    "month_sheets": "T9, T10, T11",
    "year": "2024",
    "header_f": "Quên quét thẻ",
    "header_e": "Thiếudữ liệu chấm công",
    "header_num": "Đến trễ/ về sớm",
    "header_empty": "Nghỉ không lý do / không có dữ liệu ngày công",
}


class ThuongQuyTool(BaseTool):
    name = "Quarterly Bonus"
    description = "Aggregate a quarter's attendance issues into a summary sheet for bonus review."
    icon = "🏆"
    category = "Files & Documents"
    order = 16
    action_label = "Aggregate"

    def build_body(self, parent):
        cfg = config.load(_SECTION, _DEFAULTS)

        widgets.section_label(parent, "Source file")
        self._source = widgets.file_row(
            parent, "Attendance data file (with monthly sheets + a quarter sheet)", mode="file")
        self._source.set(cfg.get("source", ""))

        widgets.section_label(parent, "Settings")
        self._quarter = widgets.dropdown(
            parent, "Quarter to aggregate (also the target sheet name)", ["Q1", "Q2", "Q3", "Q4"])
        self._quarter.set(cfg.get("quarter", "Q4"))
        self._months = widgets.text_row(parent, "The quarter's 3 monthly sheets — auto-filled")
        self._months.set(cfg.get("month_sheets", ""))
        # Gắn trace SAU khi đã nạp giá trị đã lưu, để lần mở đầu không ghi đè.
        self._quarter.widget.currentTextChanged.connect(self._on_quarter_change)

        self._year = widgets.text_row(parent, "Year (to compute Sat/Sun for account 40)")
        self._year.set(cfg.get("year", ""))

        widgets.section_label(parent, "Target column names (match headers in row 3 of the quarter sheet)")
        self._hf = widgets.text_row(parent, "Column for codes starting 'F' (missed card scan)")
        self._hf.set(cfg.get("header_f", ""))
        self._he = widgets.text_row(parent, "Column for codes starting 'E'")
        self._he.set(cfg.get("header_e", ""))
        self._hn = widgets.text_row(parent, "Column for numeric hour cells (late/early leave)")
        self._hn.set(cfg.get("header_num", ""))
        self._hk = widgets.text_row(parent, "Column for blank / '0-8' / '0-12' cells (unexcused absence)")
        self._hk.set(cfg.get("header_empty", ""))

        widgets.hint(
            parent,
            "💡 Attendance file layout: date columns start at column I; employee ID in column B; "
            "account in column G; day counts in ROW 1; data from ROW 2; the end date column is "
            "just before the 'X' in row 1 (no 'X' → up to column AM).\n"
            "✍ The tool ONLY fills the result columns in the quarter sheet, then saves straight "
            "into the original file — CLOSE it before running.\n"
            "📌 Settings are saved each time you click Aggregate.")

    def _on_quarter_change(self, *_):
        months = QUARTER_MONTHS.get(self._quarter.get())
        if months:
            self._months.set(", ".join(months))

    def run(self):
        source = self._source.get().strip()
        months = [m.strip() for m in self._months.get().split(",") if m.strip()]
        target_sheet = self._quarter.get().strip()
        year_txt = self._year.get().strip()
        header_f = self._hf.get().strip()
        header_e = self._he.get().strip()
        header_num = self._hn.get().strip()
        header_empty = self._hk.get().strip()

        config.save(_SECTION, {
            "source": source, "quarter": target_sheet,
            "month_sheets": self._months.get().strip(), "year": year_txt,
            "header_f": header_f, "header_e": header_e,
            "header_num": header_num, "header_empty": header_empty,
        })

        if not source or not os.path.isfile(source):
            self.error("Error", "Please choose a valid attendance data file.")
            return
        if len(months) != 3:
            self.error("Error", "Exactly 3 monthly sheets are required, comma-separated.")
            return
        if not target_sheet:
            self.error("Error", "No quarter selected.")
            return
        if not all([header_f, header_e, header_num, header_empty]):
            self.error("Error", "All 4 target column names are required.")
            return
        try:
            year = int(year_txt)
        except ValueError:
            self.error("Error", "Year must be a number (e.g. 2024).")
            return
        try:
            import win32com.client  # noqa: F401
        except ImportError:
            self.error("Missing library", "pywin32 is required to control Excel:\n  pip install pywin32")
            return

        try:
            stats = _aggregate_quarter(
                source_path=source, month_sheets=months, target_sheet=target_sheet,
                year=year, header_f=header_f, header_e=header_e,
                header_num=header_num, header_empty=header_empty)
        except Exception as exc:
            self.error("Error", f"Error while aggregating:\n{exc}")
            return

        lines = [
            f"✅ Aggregated {target_sheet} into:\n   {os.path.basename(source)}", "",
            f"• Cells filled: {stats['cells_written']}",
            f"• Employees with attendance issues: {stats['employees_hit']}",
        ]
        if stats["not_found"]:
            nf = stats["not_found"]
            preview = ", ".join(nf[:8])
            extra = f" (+{len(nf) - 8} more)" if len(nf) > 8 else ""
            lines.append(f"⚠ IDs present in months but missing from the quarter sheet: {preview}{extra}")
        if stats["warnings"]:
            lines.append("")
            lines += [f"⚠ {w}" for w in stats["warnings"][:8]]
        self.info("Quarterly Bonus result", "\n".join(lines))
