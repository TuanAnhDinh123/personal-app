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
    name = "Thưởng quý"
    description = "Tổng hợp lỗi chấm công cả quý vào sheet tổng hợp để xét thưởng."
    icon = "🏆"
    category = "Tệp & Tài liệu"
    order = 16
    action_label = "Tổng hợp"

    def build_body(self, parent):
        cfg = config.load(_SECTION, _DEFAULTS)

        widgets.section_label(parent, "File nguồn")
        self._source = widgets.file_row(
            parent, "File Dữ Liệu Chấm Công (chứa các sheet tháng + sheet quý)", mode="file")
        self._source.set(cfg.get("source", ""))

        widgets.section_label(parent, "Cấu hình")
        self._quarter = widgets.dropdown(
            parent, "Quý tổng hợp (cũng là tên sheet đích)", ["Q1", "Q2", "Q3", "Q4"])
        self._quarter.set(cfg.get("quarter", "Q4"))
        self._months = widgets.text_row(parent, "3 sheet tháng trong quý — tự sinh theo Quý")
        self._months.set(cfg.get("month_sheets", ""))
        # Gắn trace SAU khi đã nạp giá trị đã lưu, để lần mở đầu không ghi đè.
        self._quarter.widget.currentTextChanged.connect(self._on_quarter_change)

        self._year = widgets.text_row(parent, "Năm (để tính Thứ 7/CN cho account 40)")
        self._year.set(cfg.get("year", ""))

        widgets.section_label(parent, "Tên các cột đích (khớp tiêu đề ở dòng 3 sheet quý)")
        self._hf = widgets.text_row(parent, "Cột cho mã bắt đầu 'F' (quên quẹt thẻ)")
        self._hf.set(cfg.get("header_f", ""))
        self._he = widgets.text_row(parent, "Cột cho mã bắt đầu 'E'")
        self._he.set(cfg.get("header_e", ""))
        self._hn = widgets.text_row(parent, "Cột cho ô là SỐ giờ lẻ (đi trễ/về sớm)")
        self._hn.set(cfg.get("header_num", ""))
        self._hk = widgets.text_row(parent, "Cột cho ô trống / '0-8' / '0-12' (nghỉ không lý do)")
        self._hk.set(cfg.get("header_empty", ""))

        widgets.hint(
            parent,
            "💡 Quy ước file chấm công: cột ngày bắt đầu ở cột I; MSNV ở cột B; account "
            "ở cột G; số ngày ở DÒNG 1; dữ liệu từ DÒNG 2; cột ngày kết thúc ngay trước ô "
            "'X' ở dòng 1 (không có 'X' thì tới cột AM).\n"
            "✍ Tool CHỈ điền vào các cột kết quả trong sheet quý rồi lưu THẲNG vào file "
            "gốc — hãy ĐÓNG file trước khi chạy.\n"
            "📌 Cấu hình được lưu lại mỗi lần bấm Tổng hợp.")

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
            self.error("Lỗi", "Vui lòng chọn file Dữ Liệu Chấm Công hợp lệ.")
            return
        if len(months) != 3:
            self.error("Lỗi", "Cần đúng 3 sheet tháng, cách nhau bằng dấu phẩy.")
            return
        if not target_sheet:
            self.error("Lỗi", "Chưa chọn quý tổng hợp.")
            return
        if not all([header_f, header_e, header_num, header_empty]):
            self.error("Lỗi", "Chưa nhập đủ tên 4 cột đích.")
            return
        try:
            year = int(year_txt)
        except ValueError:
            self.error("Lỗi", "Năm phải là số (VD: 2024).")
            return
        try:
            import win32com.client  # noqa: F401
        except ImportError:
            self.error("Thiếu thư viện", "Cần pywin32 để điều khiển Excel:\n  pip install pywin32")
            return

        try:
            stats = _aggregate_quarter(
                source_path=source, month_sheets=months, target_sheet=target_sheet,
                year=year, header_f=header_f, header_e=header_e,
                header_num=header_num, header_empty=header_empty)
        except Exception as exc:
            self.error("Lỗi", f"Có lỗi xảy ra khi tổng hợp:\n{exc}")
            return

        lines = [
            f"✅ Đã tổng hợp {target_sheet} vào:\n   {os.path.basename(source)}", "",
            f"• Số ô đã điền: {stats['cells_written']}",
            f"• Nhân viên có lỗi chấm công: {stats['employees_hit']}",
        ]
        if stats["not_found"]:
            nf = stats["not_found"]
            preview = ", ".join(nf[:8])
            extra = f" (+{len(nf) - 8} nữa)" if len(nf) > 8 else ""
            lines.append(f"⚠ MSNV có trong tháng nhưng không thấy ở sheet quý: {preview}{extra}")
        if stats["warnings"]:
            lines.append("")
            lines += [f"⚠ {w}" for w in stats["warnings"][:8]]
        self.info("Kết quả Thưởng Quý", "\n".join(lines))
