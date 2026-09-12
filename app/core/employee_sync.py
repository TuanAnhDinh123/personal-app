"""Đồng bộ bảng `employees` với file Excel "Personnel Data" của HR.

Khớp theo **MÃ NHÂN VIÊN**: mỗi dòng Excel tìm đúng một nhân viên trong DB rồi
so từng cột; kết quả là danh sách thay đổi để giao diện hỏi lại người dùng
TRƯỚC khi ghi (module này không tự ghi gì khi so sánh). Chiều ngược lại: nhân
viên đang làm việc mà file không còn dòng nào mang mã của họ được gom riêng để
người dùng điền ngày nghỉ việc.

Module KHÔNG biết gì về bố cục file Excel (tên sheet, dòng header, tên cột) —
bên gọi đọc file rồi truyền vào các bản ghi đã chuẩn hóa cùng danh sách cột cần
so (`fields`), nên logic ở đây chỉ làm việc trên dict.

Hai quy ước quan trọng:

* **Ô TRỐNG LÀ MỘT GIÁ TRỊ, nhưng chỉ khi FILE CÓ CỘT ĐÓ.** Bên gọi truyền vào
  `present_fields` — tập field mà dòng header của file có cột tương ứng — nên
  phân biệt được hai chuyện khác hẳn nhau: *cột file không quản* (giữ nguyên dữ
  liệu app, vd cột chỉ có trong đơn dự tuyển) và *ô đã bị xóa trong file* (đề
  nghị xóa luôn bên app). Không có `present_fields` thì không đề nghị xóa gì.
* **So sánh KHOAN DUNG** (xem `same_value`): bỏ qua khác biệt về khoảng trắng,
  hoa/thường, cách viết ngày và cách viết số. Hai bên cùng một dữ liệu mà hiện
  lên thành "thay đổi" thì danh sách duyệt đầy nhiễu, người dùng sẽ tick bừa.
"""
import datetime
import unicodedata

from app.core import cv_repository as repo

# Các dạng ngày gặp trong dữ liệu: DB lưu "dd/mm/yyyy" (kiểu file Excel), còn ô
# chọn ngày trên giao diện trả "yyyy-mm-dd".
_DATE_FORMATS = ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y")

# Dạng ngày LƯU XUỐNG DB — bám theo dữ liệu import từ Excel (xem
# employee_db._cell_str) để mọi cột ngày trong bảng cùng một kiểu viết.
DB_DATE_FORMAT = "%d/%m/%Y"

# Chữ thay cho ô trống trên bảng duyệt (giá trị lưu xuống vẫn là chuỗi rỗng).
EMPTY_TEXT = "(blank)"


def _text(value) -> str:
    """Giá trị bất kỳ → chuỗi đã gộp khoảng trắng, chuẩn hóa Unicode NFC.

    NFC là bắt buộc khi so chữ tiếng Việt: cùng một cái tên, file Excel có thể
    lưu dạng tổ hợp (NFD) còn DB lưu dạng dựng sẵn, so chuỗi thô sẽ lệch.
    """
    if value is None:
        return ""
    return " ".join(unicodedata.normalize("NFC", str(value)).split())


def _as_date(text: str):
    """Chuỗi ngày → `datetime.date`, hoặc None nếu không phải ngày."""
    for fmt in _DATE_FORMATS:
        try:
            return datetime.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _as_number(text: str):
    """Chuỗi số → float, hoặc None nếu không phải số. Nhận cả "1.0" lẫn "1,0"."""
    try:
        return float(text.replace(",", "."))
    except ValueError:
        return None


def same_value(app_value, excel_value) -> bool:
    """Hai giá trị có coi như GIỐNG NHAU không (để không báo thay đổi giả)?

    Giống nhau khi: cùng rỗng · cùng một ngày dù viết khác kiểu · cùng một số
    dù viết khác kiểu (0 ≡ "0" ≡ "0.0") · cùng chuỗi khi bỏ qua hoa/thường và
    khoảng trắng thừa.
    """
    a, b = _text(app_value), _text(excel_value)
    if a == b:
        return True
    if not a or not b:
        return False
    da, db = _as_date(a), _as_date(b)
    if da and db:
        return da == db
    na, nb = _as_number(a), _as_number(b)
    if na is not None and nb is not None:
        return na == nb
    return a.lower() == b.lower()


def _code_key(value) -> str:
    """Mã NV chuẩn hóa để tra bảng — bỏ khoảng trắng, không phân biệt hoa/thường."""
    return _text(value).upper()


def build_changes(excel_rows, db_rows, fields, present_fields=None,
                  key_field="code") -> dict:
    """So mọi dòng Excel với dữ liệu trong DB → gói kết quả cho giao diện.

    `fields` = [(khóa trong bản ghi Excel, khóa trong dòng DB, nhãn hiển thị)].
    Hai khóa tách nhau vì bốn cột danh mục (bộ phận · cost center · loại NV ·
    cấp bậc) lưu **id** trong DB: bản ghi Excel giữ text dưới một khóa sentinel,
    còn dòng DB đọc text ra từ cột nối bảng (`department_short_name`…).

    `present_fields` = tập field mà FILE CÓ CỘT (dù cột đó trống trơn). Ô trống
    ở một cột như vậy là lệnh XÓA; field không nằm trong tập này thì file không
    quản, dữ liệu app giữ nguyên.

    Trả về dict:
        changes        – [dict] mỗi ô lệch một dòng (đã đánh số `_id` để tick)
        clears         – số dòng trong `changes` là lệnh xóa
        new_codes      – mã NV có trong file mà app chưa có
        duplicate_codes– mã xuất hiện nhiều lần trong file (chỉ lấy dòng đầu)
        rows_no_code   – số dòng Excel không có mã NV (không khớp được ai)
        matched        – số nhân viên khớp được theo mã
    """
    by_code = {}
    for row in db_rows:
        key = _code_key(row["code"])
        if key:
            by_code.setdefault(key, row)

    present = set(present_fields or ())
    changes, new_codes, duplicate_codes = [], [], []
    rows_no_code = clears = 0
    seen_codes = set()
    for rec in excel_rows:
        key = _code_key(rec.get(key_field))
        if not key:
            rows_no_code += 1
            continue
        if key in seen_codes:
            duplicate_codes.append(_text(rec.get(key_field)))
            continue
        seen_codes.add(key)
        row = by_code.get(key)
        if row is None:
            new_codes.append(_text(rec.get(key_field)))
            continue
        for rec_key, row_key, label in fields:
            excel_value = _text(rec.get(rec_key))
            # Ô trống: là lệnh XÓA nếu file có cột đó, còn cột file không quản
            # thì bỏ qua — app giữ nguyên dữ liệu đến từ nguồn khác.
            if not excel_value and rec_key not in present:
                continue
            app_value = _row_value(row, row_key)
            if same_value(app_value, excel_value):
                continue
            clearing = not excel_value
            clears += clearing
            changes.append({
                "_id": len(changes) + 1,
                "employee_id": row["employee_id"],
                "code": _text(row["code"]),
                "full_name": _text(row["full_name"]),
                "field": rec_key,
                "label": label,
                "old": _text(app_value),
                "new": excel_value,
                # Cột hiển thị riêng: ô trống mà để trắng trên bảng thì nhìn
                # không ra là "xóa giá trị đang có" hay "không có gì xảy ra".
                "old_text": _text(app_value) or EMPTY_TEXT,
                "new_text": excel_value or EMPTY_TEXT,
                "clears": clearing,
            })
    return {
        "changes": changes,
        "clears": clears,
        "new_codes": new_codes,
        "duplicate_codes": duplicate_codes,
        "rows_no_code": rows_no_code,
        "matched": len(seen_codes) - len(new_codes),
    }


def _row_value(row, key):
    """Đọc row[key] an toàn cho cả sqlite3.Row lẫn dict; thiếu/NULL → ""."""
    try:
        value = row[key]
    except (KeyError, IndexError):
        return ""
    return "" if value is None else value


def missing_from_excel(db_rows, excel_rows, key_field="code"):
    """(nhân viên đang làm việc mà file KHÔNG nhắc tới, số người chưa có mã NV).

    Chỉ xét người **đang làm việc**: người đã nghỉ thì file không còn liệt kê là
    chuyện đương nhiên. Người chưa có mã NV (mới nhập tay từ đơn dự tuyển, HR
    chưa cấp mã) không khớp được theo mã nên để riêng — coi họ là "đã nghỉ" chỉ
    vì thiếu mã là sai.
    """
    in_excel = {_code_key(rec.get(key_field)) for rec in excel_rows}
    in_excel.discard("")
    missing, no_code = [], 0
    for row in db_rows:
        if _text(_row_value(row, "work_status")) != "Working":
            continue
        key = _code_key(_row_value(row, "code"))
        if not key:
            no_code += 1
        elif key not in in_excel:
            missing.append(row)
    return missing, no_code


def apply_changes(changes, resolve=None) -> int:
    """Ghi các thay đổi người dùng đã duyệt → số nhân viên được cập nhật.

    Gom theo nhân viên rồi ghi MỘT câu UPDATE cho mỗi người (thay vì mỗi ô một
    câu). `resolve(field, value) -> (field DB, giá trị)` để bên gọi đổi text của
    cột danh mục sang id; trả về field None thì bỏ qua ô đó.
    """
    by_employee = {}
    for change in changes:
        field, value = change["field"], change["new"]
        if resolve is not None:
            field, value = resolve(field, value)
        if field is None:
            continue
        by_employee.setdefault(change["employee_id"], {})[field] = value
    for employee_id, data in by_employee.items():
        repo.update_employee(employee_id, data)
    return len(by_employee)


def apply_resignations(entries) -> int:
    """Ghi ngày nghỉ việc + lý do → số nhân viên được đánh dấu đã nghỉ.

    `entries` = [(employee_id, ngày nghỉ 'yyyy-mm-dd', lý do)]. Dòng KHÔNG có
    ngày bị bỏ qua hoàn toàn: trạng thái làm việc suy ra từ `termination_date`
    nên ghi lý do mà thiếu ngày thì người đó vẫn "đang làm việc", chỉ tổ để lại
    một dòng lý do lơ lửng.
    """
    saved = 0
    for employee_id, date_iso, reason in entries:
        date_text = to_db_date(date_iso)
        if not date_text:
            continue
        repo.update_employee(employee_id, {
            "termination_date": date_text,
            "leaving_reason": _text(reason),
        })
        saved += 1
    return saved


def to_db_date(value) -> str:
    """Ngày từ ô chọn ('yyyy-mm-dd') → dạng lưu trong DB ('dd/mm/yyyy').

    Chuỗi rỗng/không phải ngày → "" (bên gọi hiểu là chưa nhập).
    """
    date = _as_date(_text(value))
    return date.strftime(DB_DATE_FORMAT) if date else ""
