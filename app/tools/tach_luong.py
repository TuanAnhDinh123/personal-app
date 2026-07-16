"""Tách bảng lương theo Nhà cung cấp.

Port của macro VBA `file_split` (trong "Tach bang luong.xlsm") sang app.
Từ MỘT file bảng lương tổng, tạo ra nhiều file riêng cho từng Nhà cung cấp
(NCC): mỗi file chỉ giữ lại các hàng thuộc NCC đó ở những sheet chi tiết,
đồng thời xóa hẳn một số sheet không cần.

Engine dùng Excel thật qua COM (pywin32) — giống hệt macro gốc, nên giữ
nguyên 100% công thức, định dạng, biểu đồ, link. Cần máy có cài Excel
(Windows). Nếu không có Excel/pywin32, tool báo lỗi rõ ràng thay vì chạy sai.

Khác biệt nhỏ so với macro (đều là cải tiến an toàn):
  * File kết quả luôn lưu dạng .xlsx, đặt cạnh file nguồn (hoặc thư mục bạn
    chọn) — thay vì thư mục mặc định mơ hồ của Excel.
  * Tự động break hết link tới file ngoài (không cần thao tác tay), và mở file
    gốc ở chế độ chỉ-đọc nên không cần tắt file lương gốc trước khi chạy.
  * Sheet cấu hình không tồn tại thì bỏ qua kèm cảnh báo, thay vì dừng hẳn.
  * Đọc/ghi theo khối (block) nên nhanh hơn vòng lặp .Find từng ô của macro.
"""
import os
import re
import tkinter as tk
from tkinter import messagebox

from app.core import config
from app.core.base_tool import BaseTool
from app.ui import widgets

_SECTION = "tach_luong"

# --- Giá trị mặc định lấy đúng từ sheet DATA của file .xlsm gốc ---
_DEFAULTS = {
    # Cột B: tên các Nhà cung cấp (mỗi vòng lặp tạo 1 file)
    "suppliers": "ANNK-HR\nPower Connect\nNhân Kiệt\nMEKONG SUBLABOR",
    # Cột D: các tên tiêu đề cột chứa NCC (tìm cột nào khớp trước thì dùng)
    "vendor_headers": "Vendor\nNCC\nAGENCY\nVendorName\nNhà cung cấp dịch vụ\nAgency",
    # Cột H: các sheet bị xóa hoàn toàn
    "delete_full": "HC\nCompare",
    # Cột I: các sheet bị xóa hàng (chỉ giữ hàng thuộc NCC đang xử lý)
    "delete_rows": (
        "06-2026\nAtt\nShift\nOT\nSat,Sun\nOff day working\nMeal\n"
        "TRP\nPhep nam\nBHXH\nIncentive\nNVXS-NVTB\nReimburesement\nNghi T6"
    ),
    # Cột F2: đường dẫn file lương mặc định
    "source": (
        r"G:\HR\1- Personnel Management\16 - Datalogic\0 - OUTSOURCING PAYROLL"
        r"\2026\4. Apr\Seasonal_PaymentMonthly_Apr2026(All).xlsx"
    ),
    "output_dir": "",
}


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
        self._source_var = widgets.file_row(
            parent, "File bảng lương tổng (.xlsx/)", mode="file")
        self._source_var.set(cfg.get("source", ""))
        self._output_var = widgets.file_row(
            parent, "Thư mục lưu file kết quả (để trống = cùng thư mục file nguồn)",
            mode="folder")
        self._output_var.set(cfg.get("output_dir", ""))

        widgets.section_label(parent, "Cấu hình")
        self._suppliers_box = widgets.text_area(
            parent, "Nhà cung cấp — mỗi dòng 1 NCC (mỗi NCC = 1 file kết quả)",
            value=cfg.get("suppliers", ""), height=4)
        self._headers_box = widgets.text_area(
            parent, "Tên tiêu đề cột chứa NCC — mỗi dòng 1 tên (khớp cái nào trước dùng cái đó)",
            value=cfg.get("vendor_headers", ""), height=6)
        self._del_full_box = widgets.text_area(
            parent, "Sheet xóa hoàn toàn — mỗi dòng 1 tên sheet",
            value=cfg.get("delete_full", ""), height=3)
        self._del_rows_box = widgets.text_area(
            parent, "Sheet xóa hàng (chỉ giữ hàng của NCC) — mỗi dòng 1 tên sheet",
            value=cfg.get("delete_rows", ""), height=8)

        widgets.hint(
            parent,
            "💡 Với mỗi NCC: mở lại file lương gốc → lưu thành "
            "\"<tên gốc>-<NCC>.xlsx\" → xóa các sheet ở ô trên → ở mỗi sheet chi tiết, "
            "tìm cột NCC rồi xóa mọi hàng KHÔNG thuộc NCC đó (hàng trống được giữ).\n"
            "🔗 Tool tự động break hết link và mở file gốc ở chế độ chỉ-đọc, nên không "
            "cần tắt file lương gốc trước khi chạy.\n"
            "📌 Cấu hình sẽ được lưu lại mỗi lần bấm Tách file."
        )

    # ------------------------------------------------------------------
    def run(self):
        source = self._source_var.get().strip()
        output_dir = self._output_var.get().strip()
        suppliers = _lines(self._suppliers_box)
        vendor_headers = _lines(self._headers_box)
        delete_full = _lines(self._del_full_box)
        delete_rows = _lines(self._del_rows_box)

        # Lưu lại cấu hình hiện tại (kể cả khi báo lỗi bên dưới, để đỡ gõ lại)
        config.save(_SECTION, {
            "source": source,
            "output_dir": output_dir,
            "suppliers": _text(self._suppliers_box),
            "vendor_headers": _text(self._headers_box),
            "delete_full": _text(self._del_full_box),
            "delete_rows": _text(self._del_rows_box),
        })

        if not source or not os.path.isfile(source):
            messagebox.showerror("Lỗi", "Vui lòng chọn file bảng lương tổng hợp lệ.")
            return
        if not suppliers:
            messagebox.showerror("Lỗi", "Chưa có Nhà cung cấp nào (danh sách trống).")
            return
        if not vendor_headers:
            messagebox.showerror("Lỗi", "Chưa khai báo tên tiêu đề cột chứa NCC.")
            return

        try:
            import win32com.client  # noqa: F401
        except ImportError:
            messagebox.showerror(
                "Thiếu thư viện",
                "Cần pywin32 để điều khiển Excel:\n  pip install pywin32")
            return

        out_dir = output_dir or os.path.dirname(source)
        if not os.path.isdir(out_dir):
            messagebox.showerror("Lỗi", f"Thư mục lưu không tồn tại:\n{out_dir}")
            return

        try:
            created, warnings = _split_payroll(
                source_path=source,
                output_dir=out_dir,
                suppliers=suppliers,
                vendor_headers=vendor_headers,
                delete_full=delete_full,
                delete_rows=delete_rows,
            )
        except Exception as exc:
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra khi tách file:\n{exc}")
            return

        lines = [f"✅ Đã tạo {len(created)} file:"]
        lines += [f"   • {os.path.basename(p)}" for p in created]
        if warnings:
            lines.append("")
            lines.append("⚠ Cảnh báo:")
            lines += [f"   • {w}" for w in warnings[:12]]
            if len(warnings) > 12:
                lines.append(f"   • (+{len(warnings) - 12} cảnh báo nữa)")
        messagebox.showinfo("Kết quả tách bảng lương", "\n".join(lines))


# ---------------------------------------------------------------------------
# Helpers UI
# ---------------------------------------------------------------------------

def _text(box):
    """Toàn bộ nội dung một text_area."""
    return box.get("1.0", "end").rstrip("\n")


def _lines(box):
    """Danh sách dòng đã strip, bỏ dòng trống."""
    return [ln.strip() for ln in _text(box).splitlines() if ln.strip()]


# ---------------------------------------------------------------------------
# Core logic (độc lập với UI) — điều khiển Excel qua COM, bám sát macro VBA
# ---------------------------------------------------------------------------

# Hằng số Excel (tránh phải EnsureDispatch để lấy win32com.client.constants)
_XL_UP = -4162            # xlUp
_XL_LINK_EXCEL = 1        # xlLinkTypeExcelLinks (dùng cho BreakLink)
_HEADER_SCAN = "A1:DA20"  # vùng dò tiêu đề cột NCC, giống macro
_INVALID_FS = re.compile(r'[\\/:*?"<>|]')

# FileFormat theo phần mở rộng — LƯU ĐÚNG định dạng file nguồn (giống macro),
# tránh lỗi "SaveAs failed" khi nguồn là .xlsm mà ép lưu thành .xlsx.
_FMT_BY_EXT = {
    ".xlsx": 51,   # xlOpenXMLWorkbook
    ".xlsm": 52,   # xlOpenXMLWorkbookMacroEnabled
    ".xlsb": 50,   # xlExcel12
    ".xls": 56,    # xlExcel8
}


def _base_name(filename):
    """Tên gốc file lương, cắt tới hết CỤM SỐ ĐẦU TIÊN — port từ macro.

    VD: 'Seasonal_PaymentMonthly_Apr2026(All).xlsx' -> 'Seasonal_PaymentMonthly_Apr2026'
    (đồng thời bỏ luôn đuôi mở rộng). Nếu tên không có chữ số nào thì bỏ đuôi file.
    """
    n = len(filename)
    for i, ch in enumerate(filename):
        if ch.isdigit():
            nxt = filename[i + 1] if i + 1 < n else ""
            if not nxt.isdigit():
                return filename[:i + 1]
    return os.path.splitext(filename)[0]


def _find_header(top_block, vendor_headers):
    """Tìm ô tiêu đề cột NCC trong khối A1:DA20.

    Trả về (row, col) 1-based hoặc (None, None). Ưu tiên theo thứ tự
    vendor_headers (khớp cái nào trước dùng cái đó); so khớp KHÔNG phân biệt
    hoa thường (giống .Find MatchCase=False của macro).
    """
    for name in vendor_headers:
        target = name.strip().casefold()
        for r, row in enumerate(top_block):
            for c, val in enumerate(row):
                if val is not None and str(val).strip().casefold() == target:
                    return r + 1, c + 1
    return None, None


def _contiguous_groups(sorted_rows):
    """Gom các số hàng liên tiếp thành các cụm [start, end]."""
    groups = []
    for rn in sorted_rows:
        if groups and rn == groups[-1][1] + 1:
            groups[-1][1] = rn
        else:
            groups.append([rn, rn])
    return groups


def _split_payroll(source_path, output_dir, suppliers, vendor_headers,
                   delete_full, delete_rows):
    """Tạo 1 file kết quả cho mỗi NCC. Trả về (danh_sách_file, cảnh_báo)."""
    import pythoncom
    import shutil
    import tempfile
    import win32com.client as win32

    created = []
    warnings = []
    base = _base_name(os.path.basename(source_path))
    src_ext = os.path.splitext(source_path)[1].lower()
    save_fmt = _FMT_BY_EXT.get(src_ext)
    if save_fmt is None:          # đuôi lạ -> mặc định .xlsx
        src_ext, save_fmt = ".xlsx", 51

    # Excel SaveAs hay lỗi 1004 khi ghi THẲNG lên ổ ảo/đám mây (Google Drive
    # "G:\\", OneDrive) hay ổ mạng, do cách Excel ghi qua file tạm rồi đổi tên.
    # Vì vậy cho Excel lưu ra thư mục TẠM cục bộ, rồi Python copy sang đích.
    scratch = tempfile.mkdtemp(prefix="tach_luong_")

    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.AskToUpdateLinks = False
    try:
        for supplier in suppliers:
            safe = _INVALID_FS.sub("_", supplier)
            fname = f"{base}-{safe}{src_ext}"
            tmp_out = os.path.join(scratch, fname)       # Excel ghi ở đây (cục bộ)
            final_out = os.path.join(output_dir, fname)  # đích thật (có thể là G:\\)

            # Mở read-only: ta SaveAs sang file mới nên không cần ghi vào file
            # gốc; nhờ vậy chạy được kể cả khi file lương đang mở ở nơi khác.
            wb = excel.Workbooks.Open(source_path, UpdateLinks=0, ReadOnly=True)
            try:
                wb.SaveAs(tmp_out, FileFormat=save_fmt)

                # Break hết link tới file ngoài (biến công thức link thành giá
                # trị) — thay cho thao tác tay Data > Edit Links > Break Link.
                try:
                    links = wb.LinkSources(_XL_LINK_EXCEL)
                    for link in (links or []):
                        wb.BreakLink(Name=link, Type=_XL_LINK_EXCEL)
                except pythoncom.com_error as exc:
                    warnings.append(f"{supplier}: không break được link ({exc})")

                # 1) Xóa hẳn các sheet trong danh sách "xóa hoàn toàn"
                for sheet_name in delete_full:
                    try:
                        wb.Sheets(sheet_name).Delete()
                    except pythoncom.com_error:
                        warnings.append(
                            f"{supplier}: không tìm thấy sheet cần xóa '{sheet_name}'")

                # 2) Ở mỗi sheet "xóa hàng": chỉ giữ hàng thuộc NCC hiện tại
                supplier_key = supplier.strip()
                for sheet_name in delete_rows:
                    try:
                        ws = wb.Sheets(sheet_name)
                    except pythoncom.com_error:
                        warnings.append(
                            f"{supplier}: không tìm thấy sheet '{sheet_name}' (bỏ qua)")
                        continue

                    top = ws.Range(_HEADER_SCAN).Value  # 1 lần đọc cả khối
                    dest_row, dest_col = _find_header(top, vendor_headers)
                    if dest_col is None:
                        # Giống macro: đây là lỗi cấu hình cột -> dừng cả tiến trình
                        raise RuntimeError(
                            f"Sheet '{sheet_name}': không tìm thấy cột NCC "
                            f"(các tiêu đề đã thử: {', '.join(vendor_headers)}).")

                    last_row = ws.Cells(ws.Rows.Count, dest_col).End(_XL_UP).Row
                    if last_row <= dest_row:
                        continue  # không có dòng dữ liệu

                    rng = ws.Range(
                        ws.Cells(dest_row + 1, dest_col),
                        ws.Cells(last_row, dest_col)).Value
                    if isinstance(rng, tuple):
                        col_vals = [row[0] for row in rng]
                    else:
                        col_vals = [rng]  # vùng chỉ 1 ô -> Value trả scalar

                    to_delete = []
                    for idx, val in enumerate(col_vals):
                        name = "" if val is None else str(val).strip()
                        if name != "" and name != supplier_key:
                            to_delete.append(dest_row + 1 + idx)

                    # Xóa từ dưới lên để số hàng phía trên không bị lệch
                    for start, end in reversed(_contiguous_groups(to_delete)):
                        ws.Rows(f"{start}:{end}").Delete()

                wb.Close(SaveChanges=True)
            except Exception:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
                raise

            # Chuyển từ thư mục tạm sang đích thật bằng copy thường — ổn định
            # với Google Drive/OneDrive/ổ mạng (nơi Excel SaveAs hay thất bại).
            shutil.copy2(tmp_out, final_out)
            try:
                os.remove(tmp_out)
            except OSError:
                pass
            created.append(final_out)
    finally:
        excel.ScreenUpdating = True
        excel.DisplayAlerts = True
        try:
            excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()
        shutil.rmtree(scratch, ignore_errors=True)

    return created, warnings
