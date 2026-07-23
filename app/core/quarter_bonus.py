"""Tổng hợp lỗi chấm công theo quý (xét Thưởng Quý).

Port của macro VBA `thuong_quy` (trong "Tach bang luong_20241224.xlsm") sang app.
Đọc file "Dữ Liệu Chấm Công" (1 file chứa các sheet tháng và 1 sheet tổng hợp
quý, vd Q4). Với mỗi nhân viên (theo MSNV), duyệt từng ngày trong 3 tháng,
PHÂN LOẠI mã chấm công rồi ghi danh sách ngày vào đúng cột trong sheet quý:

  * Mã bắt đầu bằng 'F'                         -> cột "Quên quẹt thẻ"
  * Mã bắt đầu bằng 'E'                         -> cột (mã E)
  * Ô là SỐ giờ lẻ (đi trễ/về sớm)             -> cột (số giờ lẻ)
        - account 48-12: 0 < giờ < 12 và ≠ 8
        - account khác : 0 < giờ < 8
  * Ô TRỐNG hoặc '0-8'/'0-12' (nghỉ/không có dữ liệu) -> cột "Nghỉ không lý do"
        - account 48-12: ô trống hoặc '0-12'
        - account 48-08: ô trống hoặc '0-8'
        - account khác (40): ô trống hoặc '0-8' VÀ chỉ tính ngày trong tuần
          (bỏ Thứ 7 & Chủ Nhật — cần đúng "Năm" để tính thứ trong tuần)

Mỗi tháng ghi vào 1 cột liền kề: cột-gốc + 0/1/2 cho tháng 1/2/3 của quý.

Engine dùng Excel thật qua COM (pywin32) — giống hệt macro gốc. Cần máy có cài
Excel (Windows). Tool CHỈ ghi vào các cột kết quả (không xóa/không đụng dữ liệu
khác) nên ghi thẳng vào file gốc & lưu lại, như macro. Hãy đóng file trước khi chạy.
"""
import os

from app.core import config

_SECTION = "thuong_quy"

# Quý -> 3 sheet tháng (lịch tài chính lệch 1 tháng: Q1 bắt đầu từ T12)
QUARTER_MONTHS = {
    "Q1": ["T12", "T1", "T2"],
    "Q2": ["T3", "T4", "T5"],
    "Q3": ["T6", "T7", "T8"],
    "Q4": ["T9", "T10", "T11"],
}

# --- Giá trị mặc định (tên cột lấy đúng từ sheet Thuong_Quy của file .xlsm) ---
_DEFAULTS = {
    "source": "",
    "quarter": "Q4",
    "month_sheets": "T9, T10, T11",
    "year": "2024",                   # năm (dùng để tính thứ trong tuần cho acc 40)
    "header_f": "Quên quét thẻ",      # cột cho mã 'F...'
    "header_e": "Thiếudữ liệu chấm công",  # cột cho mã 'E...'
    "header_num": "Đến trễ/ về sớm",  # cột cho ô là SỐ giờ lẻ
    "header_empty": "Nghỉ không lý do / không có dữ liệu ngày công",  # ô trống/0-x
}


# ---------------------------------------------------------------------------
# Core logic (điều khiển Excel qua COM, bám sát macro VBA)
# ---------------------------------------------------------------------------

# Hằng số Excel
_XL_UP = -4162        # xlUp
_XL_TO_LEFT = -4159   # xlToLeft

# Quy ước cấu trúc file chấm công (giống macro)
_DAY_START_COL = 9        # cột I: cột ngày đầu tiên
_XMARK_LAST_COL = 39      # dò 'X' trong A1:AM1 (cột 39 = AM)
_MONTH_MSNV_COL = 2       # cột B: MSNV ở sheet tháng
_MONTH_ACCT_COL = 7       # cột G: loại account
_DAYNUM_ROW = 1           # dòng 1: số ngày
_MONTH_DATA_ROW = 2       # dòng 2: bắt đầu dữ liệu nhân viên
_TARGET_MSNV_COL = 2      # cột B: MSNV ở sheet quý
_TARGET_HEADER_ROW = 3    # dòng 3: tiêu đề cột ở sheet quý


def _num_str(v):
    """Chuỗi hóa số kiểu Excel: 5.0 -> '5' (để ghép 'ngày/tháng')."""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _msnv_key(v):
    """Chuẩn hóa MSNV thành khóa so khớp, KHÔNG phụ thuộc kiểu ô Excel.

    Excel/COM trả `.Value` theo định dạng ô: ô Number -> float ('15001174.0'),
    ô Text -> str ('15001174'). Nếu sheet tháng lưu số còn sheet quý lưu text
    (hoặc ngược lại) thì str() thô sẽ tạo 2 khóa khác nhau và không khớp được.
    Ép float nguyên về int để '15001174.0' và '15001174' cho cùng một khóa."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _find_col(header_vals, text):
    """Tìm cột (1-based) trong dòng tiêu đề mà ô CHỨA `text` (không phân biệt
    hoa thường, đã bỏ khoảng trắng đầu/cuối) — mô phỏng .Find LookAt=xlPart."""
    needle = text.strip().casefold()
    for c, val in enumerate(header_vals, start=1):
        if val is not None and needle in str(val).strip().casefold():
            return c
    return None


def _row_values(ws, row, first_col, last_col):
    """Đọc 1 dòng thành list (chuẩn hóa kết quả .Value của COM)."""
    rng = ws.Range(ws.Cells(row, first_col), ws.Cells(row, last_col)).Value
    if not isinstance(rng, tuple):
        return [rng]
    return list(rng[0])


def _column_values(ws, col, first_row, last_row):
    """Đọc 1 cột thành list (chuẩn hóa .Value của COM)."""
    rng = ws.Range(ws.Cells(first_row, col), ws.Cells(last_row, col)).Value
    if not isinstance(rng, tuple):
        return [rng]
    return [r[0] for r in rng]


def _add(results, row, col, label):
    results.setdefault((row, col), []).append(label)


def _is_weekday(year, month, day_val):
    """True nếu (year, month, day) là ngày trong tuần (không phải T7/CN).

    Trả về True khi không dựng được ngày hợp lệ để không âm thầm bỏ sót."""
    import datetime
    if month is None:
        return True
    try:
        d = datetime.date(int(year), int(month), int(float(day_val)))
    except (ValueError, TypeError):
        return True
    return d.weekday() < 5  # 0..4 = T2..T6


def _aggregate_quarter(source_path, month_sheets, target_sheet, year,
                       header_f, header_e, header_num, header_empty):
    """Điền lỗi chấm công vào sheet quý & lưu thẳng vào file gốc. Trả về stats."""
    import pythoncom
    import win32com.client as win32

    warnings = []
    not_found = {}  # msnv -> "sheet!dòng" (nơi xuất hiện ĐẦU TIÊN ở sheet tháng)
    results = {}  # (row, col) -> [labels]

    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.AskToUpdateLinks = False
    try:
        wb = excel.Workbooks.Open(source_path, UpdateLinks=0, ReadOnly=False)
        try:
            tgt = wb.Sheets(target_sheet)

            # 1) Tiêu đề dòng 3 -> cột gốc cho từng loại lỗi
            last_hdr_col = tgt.Cells(_TARGET_HEADER_ROW, tgt.Columns.Count).End(_XL_TO_LEFT).Column
            hdr = _row_values(tgt, _TARGET_HEADER_ROW, 1, last_hdr_col)
            col_f = _find_col(hdr, header_f)
            col_e = _find_col(hdr, header_e)
            col_num = _find_col(hdr, header_num)
            col_empty = _find_col(hdr, header_empty)
            missing = [name for name, col in [
                (header_f, col_f), (header_e, col_e),
                (header_num, col_num), (header_empty, col_empty)] if col is None]
            if missing:
                raise RuntimeError(
                    f"Sheet '{target_sheet}' dòng {_TARGET_HEADER_ROW} thiếu cột: "
                    + "; ".join(missing))

            # 2) Bản đồ MSNV -> dòng ở sheet quý (lấy lần xuất hiện đầu tiên)
            last_tgt_row = tgt.Cells(tgt.Rows.Count, _TARGET_MSNV_COL).End(_XL_UP).Row
            msnv_col_vals = _column_values(tgt, _TARGET_MSNV_COL, 1, last_tgt_row)
            msnv_to_row = {}
            for idx, val in enumerate(msnv_col_vals):
                if val is None:
                    continue
                key = _msnv_key(val)
                if key and key not in msnv_to_row:
                    msnv_to_row[key] = idx + 1

            base_cols = {"F": col_f, "E": col_e, "NUM": col_num, "EMPTY": col_empty}

            # 3) Duyệt 3 tháng
            for i, sheet_name in enumerate(month_sheets):
                try:
                    ws = wb.Sheets(sheet_name)
                except pythoncom.com_error:
                    warnings.append(f"Không tìm thấy sheet tháng '{sheet_name}' (bỏ qua)")
                    continue

                end_row = ws.Cells(ws.Rows.Count, _MONTH_MSNV_COL).End(_XL_UP).Row
                if end_row < _MONTH_DATA_ROW:
                    continue

                # Cột ngày kết thúc: ngay trước ô 'X' ở dòng 1 (A1:AM1)
                row1 = _row_values(ws, _DAYNUM_ROW, 1, _XMARK_LAST_COL)
                end_col = _XMARK_LAST_COL
                for c, val in enumerate(row1, start=1):
                    if val is not None and str(val).strip() == "X":
                        end_col = c - 1
                        break
                if end_col < _DAY_START_COL:
                    continue

                month_suffix = sheet_name.replace("T", "/")   # "T9" -> "/9"
                try:
                    month_num = int(sheet_name.replace("T", ""))
                except ValueError:
                    month_num = None

                # Đọc cả khối 1 lần (dòng 1..end_row, cột 1..end_col)
                block = ws.Range(ws.Cells(1, 1), ws.Cells(end_row, end_col)).Value
                day_row = block[0]  # dòng 1 = số ngày

                for j in range(_MONTH_DATA_ROW, end_row + 1):
                    row = block[j - 1]
                    msnv_val = row[_MONTH_MSNV_COL - 1]
                    msnv = _msnv_key(msnv_val)
                    if not msnv:
                        continue
                    trow = msnv_to_row.get(msnv)
                    if trow is None:
                        not_found.setdefault(msnv, f"{sheet_name}!{j}")
                        continue

                    acct_val = row[_MONTH_ACCT_COL - 1]
                    acct = "" if acct_val is None else str(acct_val).strip()

                    for h in range(_DAY_START_COL, end_col + 1):
                        cell = row[h - 1]
                        day_val = day_row[h - 1]
                        if day_val is None:
                            continue
                        label = _num_str(day_val) + month_suffix

                        # F*
                        if isinstance(cell, str) and cell.startswith("F"):
                            _add(results, trow, base_cols["F"] + i, label)
                        # E*
                        if isinstance(cell, str) and cell.startswith("E"):
                            _add(results, trow, base_cols["E"] + i, label)
                        # Số giờ lẻ
                        if isinstance(cell, (int, float)) and not isinstance(cell, bool):
                            if acct == "48-12":
                                if 0 < cell < 12 and cell != 8:
                                    _add(results, trow, base_cols["NUM"] + i, label)
                            else:
                                if 0 < cell < 8:
                                    _add(results, trow, base_cols["NUM"] + i, label)
                        # Ô trống / 0-x (nghỉ không lý do)
                        is_blank = cell is None or cell == ""
                        if acct == "48-12":
                            if is_blank or cell == "0-12":
                                _add(results, trow, base_cols["EMPTY"] + i, label)
                        elif acct == "48-08":
                            if is_blank or cell == "0-8":
                                _add(results, trow, base_cols["EMPTY"] + i, label)
                        else:  # account 40 (mặc định): bỏ Thứ 7 & Chủ Nhật
                            if is_blank or cell == "0-8":
                                if _is_weekday(year, month_num, day_val):
                                    _add(results, trow, base_cols["EMPTY"] + i, label)

            # 4) Ghi kết quả (chỉ ghi ô có dữ liệu — giống macro). Ép định dạng
            # Text để Excel không tự đổi nhãn "1/9" thành ngày tháng.
            for (r, c), labels in results.items():
                cell = tgt.Cells(r, c)
                cell.NumberFormat = "@"
                cell.Value = ", ".join(labels)

            employees_hit = len({r for (r, _c) in results})
            wb.Save()
            wb.Close(SaveChanges=True)
        except Exception:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
            raise
    finally:
        excel.ScreenUpdating = True
        excel.DisplayAlerts = True
        try:
            excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    return {
        "cells_written": len(results),
        "employees_hit": employees_hit,
        # Kèm vị trí "sheet!dòng" để dễ tra ngược trong file gốc
        "not_found": [f"{m} ({loc})" for m, loc in sorted(not_found.items())],
        "warnings": warnings,
    }
