"""Tách bảng lương theo Nhà cung cấp — bản PySide6.

Logic điều khiển Excel qua COM (_split_payroll) dùng lại nguyên từ module Tk cũ
(app.tools.tach_luong) — chỉ dựng lại giao diện bằng Qt.
"""
import os

from app.core import config
from app.tools.tach_luong import _split_payroll, month_sheets_from_name
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
    name = "Tách bảng lương"
    description = "Tách 1 file lương tổng thành nhiều file riêng cho từng Nhà cung cấp."
    icon = "💰"
    category = "Tệp & Tài liệu"
    order = 15
    action_label = "Tách file"

    def build_body(self, parent):
        cfg = config.load(_SECTION, _DEFAULTS)

        widgets.section_label(parent, "Nguồn & Đích")
        self._source = widgets.file_row(parent, "File bảng lương tổng (.xlsx)", mode="file")
        self._source.set(cfg.get("source", ""))
        self._output = widgets.file_row(
            parent, "Thư mục lưu kết quả (để trống = cùng thư mục file nguồn)", mode="folder")
        self._output.set(cfg.get("output_dir", ""))

        widgets.section_label(parent, "Cấu hình")
        self._suppliers = widgets.text_area(
            parent, "Nhà cung cấp — mỗi dòng 1 NCC (mỗi NCC = 1 file kết quả)",
            value=cfg.get("suppliers", ""), height=4)
        self._headers = widgets.text_area(
            parent, "Tên tiêu đề cột chứa NCC — mỗi dòng 1 tên (khớp cái nào trước dùng cái đó)",
            value=cfg.get("vendor_headers", ""), height=6)
        self._stt = widgets.text_area(
            parent, "Tên tiêu đề cột STT — mỗi dòng 1 tên (tìm thấy sẽ đánh lại số từ 1)",
            value=cfg.get("stt_headers", ""), height=3)
        self._del_full = widgets.text_area(
            parent, "Sheet xóa hoàn toàn — mỗi dòng 1 tên sheet",
            value=cfg.get("delete_full", ""), height=3)
        self._del_rows = widgets.text_area(
            parent, "Sheet xóa hàng (chỉ giữ hàng của NCC) — mỗi dòng 1 tên sheet",
            value=cfg.get("delete_rows", ""), height=8)
        self._month = widgets.text_area(
            parent, "Sheet theo tháng (tự nhận theo tên file nguồn)", value="", height=2)

        # Tự điền 2 sheet theo tháng từ tên file nguồn, cập nhật khi đổi file.
        self._source.widget.textChanged.connect(self._refresh_month)
        self._refresh_month()

        widgets.hint(
            parent,
            "💡 Với mỗi NCC: mở lại file gốc → lưu thành \"<tên gốc>-<NCC>.xlsx\" → xóa "
            "các sheet ở ô trên → ở mỗi sheet chi tiết, xóa mọi hàng KHÔNG thuộc NCC đó.\n"
            "📅 2 sheet đổi theo tháng tự nhận từ tên file nguồn (vẫn sửa tay được).\n"
            "🔢 Nếu tìm thấy cột STT thì tự đánh lại số thứ tự từ 1.\n"
            "🔗 Tool tự break link và mở file gốc chỉ-đọc, không cần đóng file gốc.\n"
            "📌 Cấu hình được lưu lại mỗi lần bấm Tách file.")

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
            self.error("Lỗi", "Vui lòng chọn file bảng lương tổng hợp lệ.")
            return
        if not suppliers:
            self.error("Lỗi", "Chưa có Nhà cung cấp nào (danh sách trống).")
            return
        if not vendor_headers:
            self.error("Lỗi", "Chưa khai báo tên tiêu đề cột chứa NCC.")
            return
        try:
            import win32com.client  # noqa: F401
        except ImportError:
            self.error("Thiếu thư viện", "Cần pywin32 để điều khiển Excel:\n  pip install pywin32")
            return

        out_dir = output_dir or os.path.dirname(source)
        if not os.path.isdir(out_dir):
            self.error("Lỗi", f"Thư mục lưu không tồn tại:\n{out_dir}")
            return

        try:
            created, warnings = _split_payroll(
                source_path=source, output_dir=out_dir, suppliers=suppliers,
                vendor_headers=vendor_headers, stt_headers=stt_headers,
                delete_full=delete_full, delete_rows=delete_rows)
        except Exception as exc:
            self.error("Lỗi", f"Có lỗi xảy ra khi tách file:\n{exc}")
            return

        lines = [f"✅ Đã tạo {len(created)} file:"]
        lines += [f"   • {os.path.basename(p)}" for p in created]
        if warnings:
            lines += ["", "⚠ Cảnh báo:"]
            lines += [f"   • {w}" for w in warnings[:12]]
            if len(warnings) > 12:
                lines.append(f"   • (+{len(warnings) - 12} cảnh báo nữa)")
        self.info("Kết quả tách bảng lương", "\n".join(lines))
