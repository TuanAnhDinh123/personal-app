"""Đọc "DLVN Application Form" (đơn dự tuyển nhân viên tự điền) bằng AI (Gemini).

Nghiệp vụ: HR gửi file mẫu `DLVN Application Form.xlsx` cho nhân viên mới. Nhân
viên nộp lại theo MỘT trong hai cách, module này nhận cả hai:

  • **Điền thẳng vào file Excel** → HR nhận lại đúng file mẫu, phần nhân viên
    điền là chữ MÀU (mẫu gốc dùng màu xanh). Module dựng lại lưới ô của mọi
    sheet thành TEXT (kèm tọa độ ô, giá trị chữ màu bọc trong «…») rồi gửi text
    đó cho Gemini — không gửi file .xlsx vì Gemini không đọc được định dạng này.
  • **In ra điền tay** → HR scan thành 1 file PDF nhiều trang. Mỗi trang được
    RASTERIZE thành ảnh PNG rồi gửi kèm trong cùng một request (xem lý do phải
    đổi sang ảnh ở docstring `app.core.signature_scan`: gửi thẳng
    `application/pdf` cho bản scan viết tay đọc kém chính xác hơn hẳn).

Kết quả trả về là dict THEO ĐÚNG TÊN CỘT của bảng `employees` nên UI chỉ việc
đổ vào form xem lại rồi `repo.insert_employee`. Đơn dự tuyển có nhiều thông tin
hơn bảng `employees` (kinh nghiệm làm việc, tình trạng sức khỏe, lương mong
đợi…) — những phần KHÔNG có cột tương ứng trong DB thì bỏ qua, không lưu.

Gọi API chỉ bằng thư viện chuẩn (urllib) — thống nhất với app.core.ai_cv_scan.
Cần API key Gemini (màn hình Settings) + kết nối mạng; đọc .xlsx cần openpyxl,
đọc .pdf cần PyMuPDF.
"""
import base64
import datetime
import json
import os
import re
import time
import urllib.error
import urllib.request

# Dùng lại bộ tách trang PDF → ảnh PNG của tool quét chữ ký (hàm dùng chung cho
# mọi file scan, không riêng bảng điểm danh) thay vì viết bản sao thứ hai.
from app.core.signature_scan import render_pages

_API_URL = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    "{model}:generateContent"
)

# Các mã HTTP tạm thời (quá tải / giới hạn nhịp) — nên thử lại.
_RETRY_STATUS = {429, 500, 502, 503, 504}
_MAX_RETRIES = 4
_RETRY_BASE_DELAY = 4      # giây; chờ tăng dần 4s, 8s, 16s…

# DPI khi đổi trang PDF scan → ảnh gửi cho AI. 150 đủ nét để đọc chữ viết tay mà
# payload còn vừa giới hạn inline_data (~20MB cho cả request, đơn thường 4 trang).
_RENDER_DPI = 150

SUPPORTED_EXTS = (".xlsx", ".xlsm", ".pdf")

# ─────────────────────────────────────────────────────────────────────────
#  Các cột bảng `employees` mà đơn dự tuyển có thể điền được. Mọi field khai
#  STRING (kể cả số) để model trả "" khi không tìm thấy — `_clean` đổi kiểu
#  và chuẩn hóa lại sau. Thêm/bớt field ở đây là đủ: `_clean` chỉ đụng tới các
#  field cần chuẩn hóa, phần còn lại đi thẳng vào dict kết quả.
# ─────────────────────────────────────────────────────────────────────────
_FIELDS = {
    "code":              "Employee code / local code, only if the form already has one",
    "global_code":       "Global employee code, only if the form already has one",
    "full_name":         "Full name of the applicant",
    "gender":            "Male or Female (form may say NAM / NỮ)",
    "date_of_birth":     "Date of birth, dd/mm/yyyy",
    "place_of_birth":    "Place of birth",
    "nationality":       "Nationality",
    "religion":          "Religion",
    "marital_status":    "One of: Single, Married, Divorced, Widowed",
    "spouse_name":       "Name of the spouse, taken from the family members table "
                         "(relation wife/husband — vợ/chồng). Empty if none",
    "children_count":    "Number of children, digits only",
    "children_names":    "Names of the children from the family members table "
                         "(relation son/daughter — con), one per line",
    "phone":             "Phone numbers of the applicant (mobile and home), "
                         'separated by "; "',
    "email":             "Personal email address",
    "permanent_address": "Permanent address (địa chỉ thường trú)",
    "temporary_address": "Temporary address (địa chỉ tạm trú)",
    "emergency_contact_name": "Name of the person to notify in case of emergency",
    "emergency_contact_relationship":
                         "Relationship of that emergency contact to the applicant",
    "education":         "Highest education level, e.g. University, College, "
                         "High school. Leave empty unless it is clear",
    "school_name":       "School / university of the highest qualification",
    "major":             "Major of that highest qualification",
    "graduation_year":   "Graduation year of that qualification, yyyy",
    "id_no":             "ID card number (CMND/CCCD)",
    "id_issued_date":    "Issue date of the ID card, dd/mm/yyyy",
    "id_issued_place":   "Issue place of the ID card",
    "bank_account_no":   "Bank account number, if the check list page has one",
    "bank_address":      "Bank name / branch, if the check list page has one",
    "tax_code":          "Personal tax code, if the check list page has one",
    "dependants":        "Number of tax dependants (số người phụ thuộc), digits only",
    "insurance_book_no": "Social insurance book number (số sổ BHXH)",
    "job_title":         "Position applied for (vị trí dự tuyển)",
}

_RESPONSE_SCHEMA = {
    "type": "OBJECT",
    "properties": {k: {"type": "STRING", "description": v} for k, v in _FIELDS.items()},
    "required": list(_FIELDS),
}

_PROMPT = (
    "You are an HR officer. The attached document is a job application form of "
    "Datalogic Vietnam (\"DLVN Application Form\") that a new employee filled in "
    "about THEMSELVES. Labels are printed in both English and Vietnamese; the "
    "answers are what the employee wrote.\n\n"
    "Extract the applicant's personal data into the given JSON schema.\n\n"
    "Rules:\n"
    "- Extract ONLY the applicant's own data. The form also lists family members, "
    "referees and previous employers — never use those people as the applicant.\n"
    "- Leave a field as an empty string when the form does not answer it. Never "
    "guess or invent a value.\n"
    "- Keep the values exactly as written (including Vietnamese diacritics); do "
    "not translate names, addresses or schools.\n"
    "- All dates must be formatted dd/mm/yyyy.\n"
    "- Ignore everything that is not in the schema (work experience, salary, "
    "health record, references, declarations…).\n"
    "Return only the JSON object matching the given schema."
)

_XLSX_NOTE = (
    "\n\nThe document is given below as a text dump of the Excel form: one line "
    "per spreadsheet row, each cell as [coordinate] value. Values wrapped in "
    "«…» are written in a non-default font colour, which is how the employee's "
    "own answers are marked in this form — but also accept an unmarked value "
    "when it clearly answers a nearby printed label. Use the coordinates to tell "
    'which label an answer or an "X" tick belongs to.\n\n'
)


class TransientError(Exception):
    """Lỗi tạm thời (nên thử lại): quá tải, giới hạn nhịp, mất kết nối."""


class Cancelled(Exception):
    """Người dùng bấm Hủy giữa chừng (kể cả khi đang chờ thử lại)."""


# ───────────────────────────── đọc file Excel ────────────────────────────
# Màu chữ mặc định của Excel (đen / tự động) — chữ mang các màu này coi như phần
# IN SẴN của biểu mẫu, khác đi là phần nhân viên tự điền.
_DEFAULT_RGB = {"FF000000", "00000000"}
_DEFAULT_THEME = {0, 1}          # theme 0/1 = background1/text1 (trắng/đen)


def _is_filled_in(cell) -> bool:
    """Ô có phải do nhân viên điền không? — suy từ MÀU CHỮ.

    Mẫu gốc in đen, phần nhân viên điền để màu xanh. Không so đúng mã màu xanh
    mà nhận mọi màu khác đen: nhân viên có thể đổi màu khác, và email tự thành
    hyperlink (màu theo theme) vẫn phải tính là đã điền.
    """
    color = cell.font.color if cell.font else None
    if color is None:
        return False
    # openpyxl trả về chuỗi mô tả lỗi thay vì None khi thuộc tính không đúng
    # kiểu (vd .rgb của màu theo theme) → phải kiểm tra `type` trước.
    if color.type == "rgb":
        return isinstance(color.rgb, str) and color.rgb.upper() not in _DEFAULT_RGB
    if color.type == "theme":
        return isinstance(color.theme, int) and color.theme not in _DEFAULT_THEME
    return color.type == "indexed" and color.indexed not in (8, 64)


def _cell_text(value) -> str:
    """Đổi giá trị 1 ô → chuỗi gọn; ngày/giờ về dd/mm/yyyy cho khớp prompt."""
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime("%d/%m/%Y")
    return " ".join(str(value).split())


def workbook_to_text(path: str) -> str:
    """Dựng lại toàn bộ workbook thành TEXT để gửi cho AI.

    Mỗi dòng có ô không rỗng → một dòng text `[A9] nhãn  [E9] «giá trị»`; ô do
    nhân viên điền (chữ khác màu đen — xem `_is_filled_in`) được bọc trong «…».
    Giữ tọa độ ô để AI gán được câu trả lời / dấu "X" vào đúng nhãn bên cạnh.
    """
    try:
        import openpyxl
    except ImportError:
        raise RuntimeError(
            "openpyxl is required to read the Excel form:\n"
            "  pip install openpyxl") from None
    # data_only=True: lấy KẾT QUẢ công thức (giá trị nhìn thấy), không lấy công thức.
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        lines = []
        for ws in wb.worksheets:
            lines.append(f"=== SHEET: {ws.title} ===")
            for row in ws.iter_rows():
                cells = []
                for cell in row:
                    if cell.value is None or str(cell.value).strip() == "":
                        continue
                    text = _cell_text(cell.value)
                    if _is_filled_in(cell):
                        text = f"«{text}»"
                    cells.append(f"[{cell.coordinate}] {text}")
                if cells:
                    lines.append("  ".join(cells))
            lines.append("")
        return "\n".join(lines)
    finally:
        wb.close()


# ─────────────────────────── chuẩn hóa kết quả AI ────────────────────────
_GENDER_MAP = {
    "nam": "Male", "male": "Male", "m": "Male", "man": "Male",
    "nữ": "Female", "nu": "Female", "female": "Female", "f": "Female",
    "woman": "Female",
}

_MARITAL_MAP = {
    "single": "Single", "độc thân": "Single", "doc than": "Single",
    "married": "Married", "có gia đình": "Married", "co gia dinh": "Married",
    "đã kết hôn": "Married", "da ket hon": "Married",
    "divorced": "Divorced", "ly dị": "Divorced", "ly di": "Divorced",
    "widowed": "Widowed", "góa": "Widowed", "goa": "Widowed",
}

# Ngày trong đơn có thể viết nhiều kiểu (nhân viên tự gõ / AI đọc từ chữ viết
# tay) — thử lần lượt các dạng này rồi chuẩn hóa hết về dd/mm/yyyy.
_DATE_FORMATS = ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y")
_DATE_FIELDS = ("date_of_birth", "id_issued_date")
_INT_FIELDS = ("children_count", "dependants")

_PHONE_SPLIT_RE = re.compile(r"[,;/\n]+")
_DIGITS_RE = re.compile(r"\d+")


def _normalize_date(text: str) -> str:
    """Đưa 1 chuỗi ngày về dd/mm/yyyy; không nhận dạng được thì GIỮ NGUYÊN.

    Giữ nguyên thay vì bỏ đi để HR còn nhìn thấy mà tự sửa ở form xem lại (vd
    chữ viết tay chỉ đọc được "10/1993").
    """
    s = " ".join(text.split())
    for fmt in _DATE_FORMATS:
        try:
            return datetime.datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return s


def _normalize_phones(text: str) -> str:
    """Gộp nhiều số điện thoại về một chuỗi ngăn bởi "; " (bỏ số trùng).

    Cùng quy ước với cột `employees.phone` khi import từ "Master HC file.xlsx".
    """
    out = []
    for part in _PHONE_SPLIT_RE.split(text):
        p = " ".join(part.split())
        if p and p not in out:
            out.append(p)
    return "; ".join(out)


def _split_name(full_name: str) -> dict:
    """Tách họ tên tiếng Việt → surname / middle_name / name.

    Theo thứ tự tiếng Việt: từ đầu là HỌ, từ cuối là TÊN, phần giữa là tên đệm
    ("Lê Phương Di" → Lê · Phương · Di). Chỉ có 1 từ thì coi như tên.
    """
    parts = full_name.split()
    if len(parts) < 2:
        return {"name": full_name} if parts else {}
    return {
        "surname": parts[0],
        "middle_name": " ".join(parts[1:-1]),
        "name": parts[-1],
    }


def normalize_name(name: str) -> str:
    """Viết hoa chữ cái đầu mỗi từ, GIỮ dấu ("LÊ PHƯƠNG DI" → "Lê Phương Di")."""
    return " ".join(w[:1].upper() + w[1:].lower() for w in name.split())


def _clean(data: dict) -> dict:
    """Kết quả thô của AI → dict đúng cột bảng `employees` để ghi thẳng vào DB.

    Bỏ field rỗng, chuẩn hóa ngày · số điện thoại · giới tính · tình trạng hôn
    nhân, đổi số sang int, tách họ tên và suy ra `marriage_status` (Y/N) từ
    tình trạng hôn nhân. Giá trị không khớp danh mục cho sẵn (vd "Separated")
    được BỎ TRỐNG để HR tự chọn ở form xem lại, tránh ghi vào DB một nhãn lạ.
    """
    rec = {}
    for key in _FIELDS:
        value = " ".join(str(data.get(key) or "").split())
        if value:
            rec[key] = value

    for key in _DATE_FIELDS:
        if key in rec:
            rec[key] = _normalize_date(rec[key])
    if "phone" in rec:
        rec["phone"] = _normalize_phones(rec["phone"])
    if "gender" in rec:
        rec["gender"] = _GENDER_MAP.get(rec["gender"].lower(), "")
    if "marital_status" in rec:
        rec["marital_status"] = _MARITAL_MAP.get(rec["marital_status"].lower(), "")
        # Cột "Marriage status (Yes)" trong file HC chỉ là Y/N — suy từ trạng thái.
        if rec["marital_status"]:
            rec["marriage_status"] = "Y" if rec["marital_status"] == "Married" else "N"
    for key in _INT_FIELDS:
        if key in rec:
            m = _DIGITS_RE.search(rec[key])
            rec[key] = int(m.group()) if m else None
    if "children_names" in rec:
        # Schema trả 1 chuỗi; cột DB quy ước mỗi tên một dòng.
        rec["children_names"] = "\n".join(
            n.strip() for n in re.split(r"[\n,;]+", rec["children_names"]) if n.strip())
    if "full_name" in rec:
        rec["full_name"] = normalize_name(rec["full_name"])
        rec.update(_split_name(rec["full_name"]))

    return {k: v for k, v in rec.items() if v not in (None, "")}


# ───────────────────────────── gọi Gemini ────────────────────────────────
def _build_parts(path: str) -> list:
    """Dựng phần 'parts' của request theo loại file (Excel → text, PDF → ảnh)."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        dump = workbook_to_text(path)
        if not dump.strip():
            raise RuntimeError("The Excel form is empty.")
        return [{"text": _PROMPT + _XLSX_NOTE + dump}]
    if ext == ".pdf":
        with open(path, "rb") as fh:
            pages = render_pages(fh.read(), "application/pdf", dpi=_RENDER_DPI)
        # Cả request (ảnh đã mã hóa base64, nở ~4/3) phải dưới giới hạn 20MB của
        # inline_data — đơn dự tuyển chỉ vài trang nên chạm ngưỡng này gần như
        # chắc chắn là scan nhầm cả tập nhiều người vào một file.
        size = sum(len(img) for _no, img, _mime in pages)
        if size > 14 * 1024 * 1024:
            raise RuntimeError(
                f"This PDF is too large to send ({len(pages)} pages). "
                "Please scan one application form per file.")
        parts = [{"text": _PROMPT + f"\n\nThe form is attached as {len(pages)} "
                                    "scanned page images, in order."}]
        parts += [{"inline_data": {"mime_type": mime,
                                   "data": base64.b64encode(img).decode("ascii")}}
                  for _no, img, mime in pages]
        return parts
    raise RuntimeError(
        f"Unsupported file type '{ext}' — the application form must be "
        ".xlsx, .xlsm or .pdf.")


def _interruptible_sleep(seconds: float, should_cancel=None) -> None:
    """Ngủ `seconds` giây nhưng cứ 0.2s lại kiểm tra Hủy; nếu Hủy → Cancelled."""
    remaining = seconds
    while remaining > 0:
        if should_cancel and should_cancel():
            raise Cancelled()
        chunk = 0.2 if remaining > 0.2 else remaining
        time.sleep(chunk)
        remaining -= chunk


def extract(api_key: str, model: str, path: str, timeout: int = 240,
            on_retry=None, should_cancel=None) -> dict:
    """Đọc 1 đơn dự tuyển (.xlsx/.xlsm/.pdf) → dict cột bảng `employees`.

    Tự thử lại khi lỗi tạm thời (429/5xx, mất mạng) với thời gian chờ tăng dần;
    `on_retry(attempt, wait, reason)` (nếu có) được gọi trước mỗi lần chờ.
    `should_cancel()` được kiểm tra thường xuyên (kể cả trong lúc chờ) → bấm Hủy
    có hiệu lực ngay, ném Cancelled. Ném RuntimeError nếu vẫn thất bại.
    """
    api_key = (api_key or "").strip()
    if not api_key:
        raise ValueError("No API key — please enter your Gemini API key in Settings.")
    parts = _build_parts(path)

    last_error = ""
    for attempt in range(1, _MAX_RETRIES + 1):
        if should_cancel and should_cancel():
            raise Cancelled()
        try:
            return _clean(_extract_once(api_key, model, parts, timeout))
        except TransientError as exc:
            last_error = str(exc)
            if attempt == _MAX_RETRIES:
                break
            wait = _RETRY_BASE_DELAY * (2 ** (attempt - 1))
            if on_retry:
                on_retry(attempt, wait, last_error)
            _interruptible_sleep(wait, should_cancel)
    raise RuntimeError(f"Failed after {_MAX_RETRIES} attempts — {last_error}")


def _extract_once(api_key: str, model: str, parts: list, timeout: int) -> dict:
    """Một lần gọi API. Ném TransientError nếu lỗi có thể thử lại."""
    body = {
        "contents": [{"parts": parts}],
        "generationConfig": {
            "temperature": 0.0,
            "responseMimeType": "application/json",
            "responseSchema": _RESPONSE_SCHEMA,
        },
    }
    req = urllib.request.Request(
        _API_URL.format(model=model),
        data=json.dumps(body).encode("utf-8"),
        headers={"Content-Type": "application/json", "x-goog-api-key": api_key},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", "ignore")
        msg = detail
        try:
            msg = json.loads(detail).get("error", {}).get("message", detail)
        except ValueError:
            pass
        if exc.code in _RETRY_STATUS:
            raise TransientError(f"HTTP {exc.code}: {msg}") from exc
        raise RuntimeError(f"HTTP {exc.code}: {msg}") from exc
    except urllib.error.URLError as exc:
        raise TransientError(f"Network error: {exc.reason}") from exc

    try:
        text = payload["candidates"][0]["content"]["parts"][0]["text"]
    except (KeyError, IndexError):
        reason = payload.get("promptFeedback", {}).get("blockReason")
        raise RuntimeError(
            "No content returned by the model"
            + (f" (blocked: {reason})" if reason else ""))
    try:
        data = json.loads(text)
    except ValueError:
        raise RuntimeError("The model returned invalid JSON.")
    if not isinstance(data, dict):
        raise RuntimeError("The model returned an unexpected format.")
    return data
