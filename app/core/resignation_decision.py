"""Xuất QUYẾT ĐỊNH THÔI VIỆC (.docx) từ bảng `employees`.

Thay cho lối làm cũ: file Excel "Quick update resignation" + tính năng Mail
Merge của Word. Dữ liệu giờ lấy thẳng từ DB, hai ngày suy ra bằng công thức nên
không còn ô nào phải gõ tay — và không còn chỗ để lệch nhau (file Excel đang
dùng có ba dòng ghi tháng BHXH khác nhau cho cùng một ngày làm việc cuối).

File mẫu (Word) và thư mục xuất đặt ở ⚙️ Settings; phần điền trường nằm ở
[docx_merge.py](docx_merge.py) — module này chỉ lo phần NGHIỆP VỤ:

    Last date  = Termination date − 1 ngày, lùi tiếp nếu rơi vào T7/CN
    Tháng BHXH = tháng của Last date, nhưng Last date trước ngày 15 thì LÙI
                 1 tháng (tháng đó công ty không đóng cho nhân viên nữa)
    No.Decision= "<năm>-<số thứ tự>", số tự tăng sau mỗi file (xem settings)
"""
import datetime
import os
import re
from typing import NamedTuple

from app.core import docx_merge

# ── Tên trường trong file mẫu Word ───────────────────────────────────────────
# Word sinh tên trường từ tiêu đề cột Excel (bỏ ký tự không hợp lệ, khoảng
# trắng thành "_") nên tên có dấu tiếng Việt và dấu gạch dưới đôi là bình
# thường, không phải viết nhầm. So khớp qua `docx_merge.norm_name`.
FIELD_DECISION_NO = "NoDecision"
FIELD_FULL_NAME = "Full_name"
FIELD_ANH_CHI = "Anhchị"
FIELD_MR_MS = "MrMs"
FIELD_DEPT_VN = "Dept__VN"
FIELD_DEPT_EN = "Dept__EN"
FIELD_JOINING_DATE = "Joining_date"
FIELD_TERMINATION_DATE = "Termination_date"
FIELD_LAST_DATE = "Last_Date"
FIELD_INSURANCE_MONTH = "Tháng_còn_đóng_BHXH"

# Xưng hô suy ra từ giới tính. DB lưu nhãn đầy đủ (cv_schema.GENDER_CHOICES),
# nhưng file Excel HC ghi "M"/"F" nên nhận cả hai cách viết.
_SALUTATIONS = {
    "male": ("Anh", "Mr"),
    "m": ("Anh", "Mr"),
    "female": ("Chị", "Ms"),
    "f": ("Chị", "Ms"),
}

# Ngày trong bảng `employees` là CHUỖI, phần lớn "dd/mm/yyyy" (xem
# employee_db._cell_str) nhưng bản ghi cũ có thể còn dạng ISO.
_DATE_FORMATS = ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y")

# Ngày trong tháng mà từ đó trở đi công ty còn đóng BHXH tháng ấy.
INSURANCE_CUTOFF_DAY = 15

# Ký tự Windows không cho đặt trong tên file.
_BAD_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


class MissingData(Exception):
    """Thiếu dữ liệu BẮT BUỘC (ngày nghỉ việc) — không dựng được quyết định."""


def parse_date(value):
    """Ô ngày của bảng `employees` → `datetime.date`; None nếu trống/không đọc được."""
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    text = str(value or "").strip()
    if not text:
        return None
    for fmt in _DATE_FORMATS:
        try:
            return datetime.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def last_working_day(termination):
    """Ngày làm việc cuối cùng = ngày nghỉ việc − 1, LÙI QUA thứ 7 & chủ nhật.

    Ngày nghỉ việc là ngày ĐẦU TIÊN không còn là nhân viên, nên ngày làm việc
    cuối là ngày liền trước — và ngày đó phải là ngày đi làm thật: terminate
    thứ 2 thì ngày làm việc cuối là thứ 6 tuần trước, không phải chủ nhật.

    KHÔNG xét ngày lễ: app không giữ lịch nghỉ lễ, đoán bừa thì sai còn tệ hơn.
    Rơi đúng dịp lễ thì sửa lại trong file Word xuất ra.
    """
    day = termination - datetime.timedelta(days=1)
    while day.weekday() >= 5:              # 5 = thứ 7, 6 = chủ nhật
        day -= datetime.timedelta(days=1)
    return day


def insurance_month(last_day):
    """Tháng công ty còn đóng BHXH, trả về ngày 01 của tháng đó.

    Làm tới ngày 15 trở đi thì công ty đóng cả tháng ấy; nghỉ từ ngày 14 trở về
    trước thì tháng ấy không đóng nữa nên lùi lại một tháng.
    """
    if last_day.day >= INSURANCE_CUTOFF_DAY:
        return last_day.replace(day=1)
    first = last_day.replace(day=1)
    return (first - datetime.timedelta(days=1)).replace(day=1)


def decision_no(year: int, number: int) -> str:
    """Số quyết định: "2026-20". Số một chữ số thêm 0 ở đầu cho khớp file cũ."""
    return f"{year}-{number:02d}"


def _row(row, key, default=""):
    """Đọc row[key] an toàn cho cả sqlite3.Row lẫn dict."""
    try:
        value = row[key]
    except (KeyError, IndexError):
        return default
    return default if value is None else value


def output_filename(row) -> str:
    """Tên file kết quả: "MSNV_Họ và tên.docx" (bỏ ký tự Windows không cho phép)."""
    code = str(_row(row, "code")).strip()
    name = " ".join(str(_row(row, "full_name")).split())
    stem = "_".join(p for p in (code, name) if p) or "resignation-decision"
    return f"{_BAD_FILENAME_CHARS.sub('', stem).strip().rstrip('.')}.docx"


def missing_labels(row) -> list[str]:
    """Các ô CÓ MẶT trong quyết định mà nhân viên này còn trống.

    Không chặn việc xuất file (chỉ `termination_date` mới chặn): thiếu một ô
    thì chỗ đó trong file để trắng, người dùng thấy cảnh báo rồi điền tay. Chặn
    cả lượt vì một ô trống là bắt HR quay lại sửa Excel rồi đồng bộ lại, trong
    khi họ chỉ cần một tờ quyết định.
    """
    checks = [
        ("Employee code", _row(row, "code")),
        ("Full name", _row(row, "full_name")),
        ("Gender", _row(row, "gender")),
        ("Department", _row(row, "department_name")),
        ("Department (Vietnamese)", _row(row, "department_name_vn")),
        ("Date of employment", _row(row, "date_of_employment")),
    ]
    missing = [label for label, value in checks if not str(value).strip()]
    if str(_row(row, "gender")).strip().lower() not in _SALUTATIONS:
        if "Gender" not in missing:
            missing.append("Gender")
    return missing


def build_values(row, number: int, year: int) -> dict:
    """Giá trị cho từng trường MERGEFIELD của một nhân viên.

    Ngày trả về dạng `datetime.date` — cách hiển thị (dd/MM/yyyy hay MM/yyyy)
    do khóa `\\@` trong chính file mẫu quyết định, xem `docx_merge.render`.
    """
    termination = parse_date(_row(row, "termination_date", None))
    if termination is None:
        raise MissingData("No termination date")
    last_day = last_working_day(termination)
    anh_chi, mr_ms = _SALUTATIONS.get(
        str(_row(row, "gender")).strip().lower(), ("", ""))
    return {
        FIELD_DECISION_NO: decision_no(year, number),
        FIELD_FULL_NAME: str(_row(row, "full_name")).strip(),
        FIELD_ANH_CHI: anh_chi,
        FIELD_MR_MS: mr_ms,
        FIELD_DEPT_VN: str(_row(row, "department_name_vn")).strip(),
        FIELD_DEPT_EN: str(_row(row, "department_name")).strip(),
        FIELD_JOINING_DATE: parse_date(_row(row, "date_of_employment", None)),
        FIELD_TERMINATION_DATE: termination,
        FIELD_LAST_DATE: last_day,
        FIELD_INSURANCE_MONTH: insurance_month(last_day),
    }


class Exported(NamedTuple):
    """Một quyết định đã ghi ra đĩa."""

    row: object
    number: int
    path: str
    missing: list


class Result(NamedTuple):
    exported: list      # [Exported]
    skipped: list       # [(tên nhân viên, lý do)]
    next_number: int    # số quyết định cho lượt xuất SAU
    unfilled: set       # trường có trong file mẫu mà app không biết điền


def export(rows, template_path, out_dir, start_number: int, year: int,
           overwrite: bool = True) -> Result:
    """Xuất mỗi nhân viên một file .docx, số quyết định tăng dần theo thứ tự.

    SỐ CHỈ TĂNG KHI FILE ĐÃ GHI XONG: bỏ qua ai (trùng file, thiếu ngày nghỉ
    việc) thì người kế tiếp lấy chính số đó, không để thủng số trong sổ quyết
    định.
    """
    exported, skipped, unfilled = [], [], set()
    number = start_number
    for row in rows:
        name = str(_row(row, "full_name")).strip() or str(_row(row, "code")).strip()
        try:
            values = build_values(row, number, year)
        except MissingData:
            skipped.append((name, "no termination date"))
            continue
        path = os.path.join(out_dir, output_filename(row))
        if os.path.exists(path) and not overwrite:
            skipped.append((name, "file already exists"))
            continue
        unfilled |= docx_merge.render(template_path, path, values)
        exported.append(Exported(row, number, path, missing_labels(row)))
        number += 1
    return Result(exported, skipped, number, unfilled)
