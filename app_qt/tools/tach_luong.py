"""Tách bảng lương theo Nhà cung cấp — bản PySide6.

Logic điều khiển Excel qua COM (_split_payroll) tách riêng ở app.core.payroll_split
— file này chỉ dựng giao diện bằng Qt.
"""
import os

from app.core import config
from app.core.payroll_split import _split_payroll, month_sheets_from_name
from app_qt import dialogs, widgets
from app_qt.base_tool import BaseTool

_SECTION = "tach_luong"

_DEFAULTS = {
    "suppliers": "ANNK-HR\nPower Connect\nNhân Kiệt\nMEKONG SUBLABOR",
    "vendor_headers": "Vendor\nNCC\nAGENCY\nVendorName\nNhà cung cấp dịch vụ\nAgency",
    "stt_headers": "STT\nNo\nNo.",
    "delete_full": "HC\nCompare",
    "delete_rows": ("Att\nShift\nOT\nSat,Sun\nOff day working\nMeal\n"
                    "TRP\nPhep nam\nBHXH\nIncentive\nNVXS-NVTB\nReimburesement"),
    "source": "",
    "output_dir": "",
}


def _lines(text):
    return [ln.strip() for ln in (text or "").splitlines() if ln.strip()]


class TachBangLuongTool(BaseTool):
    name = "Split Payroll"
    description = "Split one master payroll file into separate files per vendor."
    icon = "💰"
    category = "Files & Documents"
    order = 15
    action_label = "Split files"

    def build_body(self, parent):
        cfg = config.load(_SECTION, _DEFAULTS)

        widgets.section_label(parent, "Source & target")
        self._source = widgets.file_row(parent, "Master payroll file (.xlsx)", mode="file")
        self._source.set(cfg.get("source", ""))
        self._output = widgets.file_row(
            parent, "Output folder (blank = same as source)", mode="folder")
        self._output.set(cfg.get("output_dir", ""))

        widgets.section_label(parent, "Settings")
        self._suppliers = widgets.text_area(
            parent, "Vendors — one per line (each vendor = one output file)",
            value=cfg.get("suppliers", ""), height=4)
        self._headers = widgets.text_area(
            parent, "Vendor-column header names — one per line (first match wins)",
            value=cfg.get("vendor_headers", ""), height=6)
        self._stt = widgets.text_area(
            parent, "Index-column (No.) header names — one per line (renumbered from 1 if found)",
            value=cfg.get("stt_headers", ""), height=3)
        self._del_full = widgets.text_area(
            parent, "Sheets to delete entirely — one sheet name per line",
            value=cfg.get("delete_full", ""), height=3)
        self._del_rows = widgets.text_area(
            parent, "Sheets to filter rows (keep only the vendor's rows) — one per line",
            value=cfg.get("delete_rows", ""), height=8)
        self._month = widgets.text_area(
            parent, "Monthly sheets (auto-detected from the source filename)", value="", height=2)

        # Tự điền 2 sheet theo tháng từ tên file nguồn, cập nhật khi đổi file.
        self._source.widget.textChanged.connect(self._refresh_month)
        self._refresh_month()

        widgets.hint(
            parent,
            "💡 For each vendor: reopen the original → save as \"<name>-<vendor>.xlsx\" → delete "
            "the sheets above → in each detail sheet, remove every row NOT belonging to that vendor.\n"
            "📅 The two monthly sheets are auto-detected from the source filename (still editable).\n"
            "🔢 If an index column is found, it's renumbered from 1.\n"
            "🔗 The tool breaks links and opens the original read-only, so you needn't close it.\n"
            "📌 Settings are saved each time you click Split files.")

    def _refresh_month(self):
        ms, ns = month_sheets_from_name(os.path.basename(self._source.get()))
        if ms:
            self._month.set(f"{ms}\n{ns}")

    def run(self):
        source = self._source.get().strip()
        output_dir = self._output.get().strip()
        suppliers = _lines(self._suppliers.get())
        vendor_headers = _lines(self._headers.get())
        stt_headers = _lines(self._stt.get())
        delete_full = _lines(self._del_full.get())
        delete_rows = _lines(self._del_rows.get()) + _lines(self._month.get())

        config.save(_SECTION, {
            "source": source,
            "output_dir": output_dir,
            "suppliers": self._suppliers.get(),
            "vendor_headers": self._headers.get(),
            "stt_headers": self._stt.get(),
            "delete_full": self._del_full.get(),
            "delete_rows": self._del_rows.get(),
        })

        if not source or not os.path.isfile(source):
            self.error("Error", "Please choose a valid master payroll file.")
            return
        if not suppliers:
            self.error("Error", "No vendors listed (empty list).")
            return
        if not vendor_headers:
            self.error("Error", "No vendor-column header names set.")
            return
        try:
            import win32com.client  # noqa: F401
        except ImportError:
            self.error("Missing library", "pywin32 is required to control Excel:\n  pip install pywin32")
            return

        out_dir = output_dir or os.path.dirname(source)
        if not os.path.isdir(out_dir):
            self.error("Error", f"Output folder doesn't exist:\n{out_dir}")
            return

        try:
            created, warnings = _split_payroll(
                source_path=source, output_dir=out_dir, suppliers=suppliers,
                vendor_headers=vendor_headers, stt_headers=stt_headers,
                delete_full=delete_full, delete_rows=delete_rows)
        except Exception as exc:
            self.error("Error", f"Error while splitting files:\n{exc}")
            return

        lines = [f"✅ Created {len(created)} files:"]
        lines += [f"   • {os.path.basename(p)}" for p in created]
        if warnings:
            lines += ["", "⚠ Warnings:"]
            lines += [f"   • {w}" for w in warnings[:12]]
            if len(warnings) > 12:
                lines.append(f"   • (+{len(warnings) - 12} more warnings)")
        self.info("Split result", "\n".join(lines))
