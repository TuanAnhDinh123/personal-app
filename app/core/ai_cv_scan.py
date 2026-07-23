"""Quét CV bằng AI (Google Gemini) → xuất bảng đánh giá ứng viên ra Excel.

Khác với tool "Quét CV" (chỉ dùng regex để tách Email/SĐT), tool này gửi
NGUYÊN file PDF cho mô hình Gemini để đọc hiểu và trả về:
    • Họ tên, ngày sinh, email, số điện thoại
    • Mức độ phù hợp với JD (điểm 0–100 + nhận xét)
    • Ưu điểm / nhược điểm của ứng viên

Cách hoạt động:
    1. Đọc từng file PDF trong thư mục → mã hóa base64.
    2. Gọi Gemini REST API (generateContent) kèm JD làm ngữ cảnh, yêu cầu
       trả về JSON có cấu trúc (responseSchema).
    3. Gom kết quả tất cả CV → ghi ra một file Excel (openpyxl).

Chỉ dùng thư viện chuẩn để gọi API (urllib) nên KHÔNG cần cài thêm gói.
Cần: một API key Gemini (https://aistudio.google.com/apikey) và kết nối mạng.
"""
import base64
import json
import os
import shutil
import threading
import time
import urllib.error
import urllib.request
from pathlib import Path


from app.core import config, settings

try:
    import openpyxl
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

SECTION = "ai_scan_cv"
# API key & model nay là thiết lập CHUNG (màn hình Cài đặt) — xem app/core/settings.py.
# Tool này chỉ còn giữ đường dẫn vào/ra và JD của riêng nó.
DEFAULTS = {
    "folder":       "",
    "output":       "",
    "jd_file":      "",   # đường dẫn tới file JD (PDF/DOCX/TXT) — thay cho ô dán JD cũ
    "extra_prompt": "",   # yêu cầu bổ sung người dùng gửi kèm cho AI
}

# Tên folder (ngang cấp) chứa các CV đã quét xong, để lần sau không quét lại.
_DONE_SUFFIX = "_da_quet"

_API_URL = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    "{model}:generateContent"
)

# Schema JSON mà Gemini phải tuân theo khi trả kết quả cho mỗi CV.
_RESPONSE_SCHEMA = {
    "type": "OBJECT",
    "properties": {
        "name":       {"type": "STRING", "description": "Họ và tên ứng viên"},
        "dob":        {"type": "STRING", "description": "Ngày/tháng/năm sinh (dd/mm/yyyy nếu có)"},
        "email":      {"type": "STRING"},
        "phone":      {"type": "STRING", "description": "Số điện thoại"},
        "fit_score":  {"type": "INTEGER", "description": "Mức độ phù hợp với JD, 0-100"},
        "fit_summary": {"type": "STRING", "description": "Nhận xét ngắn vì sao phù hợp / chưa phù hợp"},
        "strengths":  {"type": "STRING", "description": "Ưu điểm nổi bật (gạch đầu dòng ngắn)"},
        "weaknesses": {"type": "STRING", "description": "Nhược điểm / điểm còn thiếu so với JD"},
    },
    "required": [
        "name", "dob", "email", "phone",
        "fit_score", "fit_summary", "strengths", "weaknesses",
    ],
}

# Cột xuất ra Excel: (khóa trong JSON, tiêu đề hiển thị, độ rộng cột).
_COLUMNS = [
    ("file",        "Tên file",         28),
    ("name",        "Họ tên",           22),
    ("dob",         "Ngày sinh",        14),
    ("email",       "Email",            28),
    ("phone",       "Số điện thoại",    16),
    ("fit_score",   "Điểm phù hợp",     13),
    ("fit_summary", "Đánh giá phù hợp", 46),
    ("strengths",   "Ưu điểm",          46),
    ("weaknesses",  "Nhược điểm",       46),
]


def read_jd_file(path: str) -> str:
    """Đọc nội dung JD từ file PDF / DOCX / TXT thành text thuần.

    Dùng lại bộ trích text của tool "Quét CV" cho PDF/DOCX; .txt đọc trực tiếp.
    """
    p = Path(path)
    if not p.is_file():
        raise FileNotFoundError(f"Không tìm thấy file JD: {path}")
    suffix = p.suffix.lower()
    if suffix == ".txt":
        return p.read_text(encoding="utf-8", errors="ignore")
    # PDF/DOCX → dùng chung logic với app.core.cv_scan (tránh phụ thuộc vòng).
    from app.core.cv_scan import _extract_cv_text
    return _extract_cv_text(p)


def _build_prompt(jd: str, extra: str = "") -> str:
    """Câu lệnh gửi kèm mỗi CV.

    jd    — nội dung mô tả công việc (đọc từ file JD người dùng chọn).
    extra — yêu cầu bổ sung tùy ý người dùng nhập thêm cho AI.
    """
    jd = jd.strip() or "(Không có mô tả công việc cụ thể — hãy đánh giá tổng quát.)"
    extra = extra.strip()
    extra_block = ""
    if extra:
        extra_block = (
            "\n=== YÊU CẦU BỔ SUNG TỪ NGƯỜI DÙNG ===\n"
            f"{extra}\n"
            "=== HẾT YÊU CẦU BỔ SUNG ===\n"
        )
    return (
        "Bạn là chuyên viên tuyển dụng. Hãy đọc kỹ CV trong file PDF đính kèm "
        "và trích xuất thông tin ứng viên, đồng thời đánh giá mức độ phù hợp với "
        "mô tả công việc (JD) dưới đây.\n\n"
        "=== MÔ TẢ CÔNG VIỆC (JD) ===\n"
        f"{jd}\n"
        "=== HẾT JD ===\n"
        f"{extra_block}\n"
        "Yêu cầu:\n"
        "- Trả lời hoàn toàn bằng tiếng Việt.\n"
        "- 'fit_score' là số nguyên 0-100 thể hiện độ khớp giữa CV và JD.\n"
        "- Nếu thông tin nào không tìm thấy trong CV thì để chuỗi rỗng.\n"
        "- 'strengths' và 'weaknesses' viết mỗi ý một dòng, ngắn gọn.\n"
        "Chỉ trả về đúng đối tượng JSON theo schema đã cho."
    )


# Các mã lỗi HTTP mang tính tạm thời (quá tải / giới hạn nhịp) — nên thử lại.
_RETRY_STATUS = {429, 500, 502, 503, 504}
_MAX_RETRIES = 4          # số lần thử lại tối đa cho mỗi CV
_RETRY_BASE_DELAY = 4     # giây; chờ tăng dần: 4s, 8s, 16s…


def _call_gemini(api_key: str, model: str, jd: str, pdf_bytes: bytes,
                 timeout: int = 180, on_retry=None, extra: str = "",
                 should_cancel=None) -> dict:
    """Gửi 1 file PDF cho Gemini, trả về dict theo _RESPONSE_SCHEMA.

    Tự động thử lại khi gặp lỗi tạm thời (429/5xx, ví dụ HTTP 503 "high
    demand"), chờ tăng dần giữa các lần. `on_retry(attempt, wait, reason)`
    (nếu có) được gọi trước mỗi lần chờ để báo tiến trình. Ném RuntimeError
    với thông báo dễ hiểu nếu vẫn thất bại.

    `should_cancel()` (nếu có) được kiểm tra thường xuyên — kể cả trong lúc
    đang chờ giữa các lần thử lại — để bấm Hủy có hiệu lực NGAY, ném _Cancelled.
    """
    last_error = ""
    for attempt in range(1, _MAX_RETRIES + 1):
        if should_cancel and should_cancel():
            raise _Cancelled()
        try:
            return _call_gemini_once(api_key, model, jd, pdf_bytes, timeout, extra)
        except _TransientError as exc:
            last_error = str(exc)
            if attempt == _MAX_RETRIES:
                break
            wait = _RETRY_BASE_DELAY * (2 ** (attempt - 1))
            if on_retry:
                on_retry(attempt, wait, last_error)
            _interruptible_sleep(wait, should_cancel)
    raise RuntimeError(f"Thất bại sau {_MAX_RETRIES} lần thử — {last_error}")


class _TransientError(Exception):
    """Lỗi tạm thời (nên thử lại): quá tải, giới hạn nhịp, mất kết nối."""


class _Cancelled(Exception):
    """Người dùng bấm Hủy giữa chừng (kể cả khi đang chờ thử lại)."""


def _interruptible_sleep(seconds: float, should_cancel=None) -> None:
    """Ngủ `seconds` giây nhưng cứ 0.2s lại kiểm tra Hủy; nếu Hủy → _Cancelled."""
    remaining = seconds
    while remaining > 0:
        if should_cancel and should_cancel():
            raise _Cancelled()
        chunk = 0.2 if remaining > 0.2 else remaining
        time.sleep(chunk)
        remaining -= chunk


def _call_gemini_once(api_key: str, model: str, jd: str, pdf_bytes: bytes,
                      timeout: int, extra: str = "") -> dict:
    """Một lần gọi API. Ném _TransientError nếu lỗi có thể thử lại."""
    body = {
        "contents": [{
            "parts": [
                {"text": _build_prompt(jd, extra)},
                {"inline_data": {
                    "mime_type": "application/pdf",
                    "data": base64.b64encode(pdf_bytes).decode("ascii"),
                }},
            ],
        }],
        "generationConfig": {
            "temperature": 0.2,
            "responseMimeType": "application/json",
            "responseSchema": _RESPONSE_SCHEMA,
        },
    }
    url = _API_URL.format(model=model)
    req = urllib.request.Request(
        url,
        data=json.dumps(body).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "x-goog-api-key": api_key,
        },
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
            raise _TransientError(f"HTTP {exc.code}: {msg}") from exc
        raise RuntimeError(f"HTTP {exc.code}: {msg}") from exc
    except urllib.error.URLError as exc:
        # Timeout / mất mạng thường là tạm thời → cho thử lại.
        raise _TransientError(f"Lỗi kết nối mạng: {exc.reason}") from exc

    try:
        text = payload["candidates"][0]["content"]["parts"][0]["text"]
    except (KeyError, IndexError):
        # Có thể bị chặn do safety hoặc phản hồi rỗng.
        reason = payload.get("promptFeedback", {}).get("blockReason")
        raise RuntimeError(
            f"Không nhận được nội dung từ mô hình"
            + (f" (bị chặn: {reason})" if reason else ""))
    try:
        return json.loads(text)
    except ValueError:
        raise RuntimeError("Mô hình trả về không đúng JSON.")


_SHEET_TITLE = "AI CV Scan"


def _write_header(ws) -> None:
    """Kẻ hàng tiêu đề in đậm + đặt độ rộng cột."""
    from openpyxl.styles import Alignment, Font, PatternFill

    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="6366F1")
    for col, (_, title, width) in enumerate(_COLUMNS, start=1):
        c = ws.cell(row=1, column=col, value=title)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(vertical="center", horizontal="center")
        ws.column_dimensions[c.column_letter].width = width
    ws.freeze_panes = "A2"


def append_rows_to_excel(rows: list[dict], out_path: str) -> None:
    """Ghi NỐI TIẾP các dòng kết quả vào file .xlsx.

    • Nếu file chưa tồn tại → tạo mới kèm hàng tiêu đề.
    • Nếu đã tồn tại → mở ra và ghi tiếp phía dưới các record cũ, nhờ vậy
      nhiều lần quét (mỗi lần dừng do hết hạn mức key free) vẫn dồn chung một
      file thay vì đè lên nhau.
    """
    from openpyxl.styles import Alignment, Font

    if os.path.exists(out_path):
        wb = openpyxl.load_workbook(out_path)
        ws = wb[_SHEET_TITLE] if _SHEET_TITLE in wb.sheetnames else wb.active
        if ws.max_row < 1 or ws.cell(row=1, column=1).value is None:
            _write_header(ws)
        start = ws.max_row + 1
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = _SHEET_TITLE
        _write_header(ws)
        start = 2

    cell_font = Font(name="Segoe UI", size=11)
    wrap      = Alignment(vertical="top", wrap_text=True)
    for r, row in enumerate(rows, start=start):
        for col, (key, _, _) in enumerate(_COLUMNS, start=1):
            c = ws.cell(row=r, column=col, value=row.get(key, ""))
            c.font = cell_font
            c.alignment = wrap
    wb.save(out_path)


# Giữ tên cũ cho tương thích; mặc định giờ là ghi nối tiếp.
_write_excel = append_rows_to_excel


# --------------------------------------------------------------- folder "đã quét"
def done_folder_for(folder: str) -> Path:
    """Trả về đường dẫn folder (ngang cấp) chứa CV đã quét — chưa tạo trên đĩa."""
    src = Path(folder)
    return src.parent / f"{src.name}{_DONE_SUFFIX}"


def move_to_done(path: Path, done_dir: Path) -> Path:
    """Chuyển 1 file CV đã quét xong sang folder 'đã quét'.

    Tự tạo folder đích nếu chưa có; nếu trùng tên thì thêm hậu tố _1, _2…
    Trả về đường dẫn mới của file.
    """
    done_dir.mkdir(parents=True, exist_ok=True)
    target = done_dir / path.name
    i = 1
    while target.exists():
        target = done_dir / f"{path.stem}_{i}{path.suffix}"
        i += 1
    shutil.move(str(path), str(target))
    return target
