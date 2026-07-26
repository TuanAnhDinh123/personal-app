"""Nhận diện CHỮ KÝ trên ảnh/PDF scan bằng Gemini (REST, urllib).

Luồng nghiệp vụ: HR in bảng điểm danh khóa học, phát cho nhân viên ký tay; sau
buổi học scan lại — thường là cả tập giấy thành MỘT file PDF nhiều trang. Module
này tách file thành từng trang, gửi lần lượt kèm danh sách nhân viên đã ghi danh
cho Gemini và hỏi: dòng nào có chữ ký tay ở ô SIGNATURE? Trả về danh sách dòng
``{person_id, name, signed}`` để UI đối chiếu theo MÃ NHÂN VIÊN rồi cập nhật
trạng thái học sang "Completed".

QUAN TRỌNG — vì sao phải RASTERIZE PDF trước khi gửi (`render_pages`): gửi thẳng
``application/pdf`` thì một số model (đã gặp với gemini-3.6-flash) gán chữ ký
LỆCH LÊN DÒNG TRÊN — dòng trống bị báo đã ký còn người ký thật thì bị bỏ sót.
Cùng model, cùng prompt, gửi ảnh raster của đúng trang đó thì đọc chính xác (đã
kiểm ở mọi mức 100–300 dpi). Nên PDF luôn được đổi sang PNG từng trang.

Gọi API chỉ dùng thư viện chuẩn (urllib) — thống nhất với app.core.ai_cv_scan.
Riêng bước tách trang PDF cần PyMuPDF (``pip install pymupdf``, app đã dùng ở
app.core.pdf_text). Cần API key Gemini (màn hình Cài đặt) + kết nối mạng.
"""
import base64
import json
import os
import time
import urllib.error
import urllib.request

_API_URL = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    "{model}:generateContent"
)

# Các mã HTTP tạm thời (quá tải / giới hạn nhịp) — nên thử lại.
_RETRY_STATUS = {429, 500, 502, 503, 504}
_MAX_RETRIES = 4
_RETRY_BASE_DELAY = 4     # giây; chờ tăng dần 4s, 8s, 16s…

# DPI khi đổi trang PDF → ảnh PNG. 150 đủ nét để phân biệt ô có/không có chữ ký
# (đã thử: 100 dpi cũng đọc đúng) mà payload còn nhỏ; 300 dpi cho ảnh ~6MB/trang,
# quét cả tập 20 trang sẽ rất chậm mà không đọc chính xác hơn.
_RENDER_DPI = 150

# Đuôi file → mime type Gemini chấp nhận cho inline_data.
_MIME_BY_EXT = {
    ".png":  "image/png",
    ".jpg":  "image/jpeg",
    ".jpeg": "image/jpeg",
    ".webp": "image/webp",
    ".heic": "image/heic",
    ".heif": "image/heif",
    ".pdf":  "application/pdf",
}

# Schema JSON mà Gemini phải tuân theo: mảng 'rows', mỗi dòng 1 nhân viên.
_RESPONSE_SCHEMA = {
    "type": "OBJECT",
    "properties": {
        "rows": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "properties": {
                    "person_id": {"type": "STRING",
                                  "description": "Mã nhân viên đọc ở cột PERSON ID của dòng"},
                    "name":      {"type": "STRING", "description": "Họ tên ở cột NAME"},
                    "signed":    {"type": "BOOLEAN",
                                  "description": "true nếu ô SIGNATURE của CHÍNH dòng đó có nét viết tay; false nếu ô trống"},
                },
                "required": ["person_id", "signed"],
            },
        },
    },
    "required": ["rows"],
}


class TransientError(Exception):
    """Lỗi tạm thời (nên thử lại): quá tải, giới hạn nhịp, mất kết nối."""


class Cancelled(Exception):
    """Người dùng bấm Hủy giữa chừng (kể cả khi đang chờ thử lại)."""


def guess_mime(path: str) -> str:
    """Đoán mime type theo đuôi file (mặc định octet-stream nếu lạ)."""
    return _MIME_BY_EXT.get(os.path.splitext(path)[1].lower(), "application/octet-stream")


def is_supported(path: str) -> bool:
    """File có đuôi mà Gemini đọc được (ảnh/PDF) không?"""
    return os.path.splitext(path)[1].lower() in _MIME_BY_EXT


def render_pages(file_bytes: bytes, mime_type: str, dpi: int = _RENDER_DPI,
                 only=None):
    """Tách file scan thành từng TRANG ẢNH để gửi riêng từng request.

    Trả về list ``[(số_trang, bytes, mime_type), …]``:
      • PDF (kể cả tập 20 trang từ máy scan) → mỗi trang một ảnh PNG. BẮT BUỘC
        phải qua bước này, xem lý do ở docstring đầu module.
      • Ảnh đơn (PNG/JPG/…) → giữ nguyên, coi như 1 trang.

    `only` (nếu có): tập số trang (đếm từ 1) cần lấy — dùng khi QUÉT TIẾP mấy
    trang bị lỗi lần trước, khỏi phải gửi lại cả tập.

    Ném RuntimeError nếu là PDF mà chưa cài PyMuPDF, PDF hỏng/không mở được, hoặc
    `only` không trỏ tới trang nào có thật.
    """
    want = set(only) if only else None
    if mime_type != "application/pdf":
        return [(1, file_bytes, mime_type)] if want is None or 1 in want else []
    try:
        import fitz          # PyMuPDF
    except ImportError:
        raise RuntimeError(
            "Cần cài PyMuPDF để tách trang PDF trước khi gửi cho AI:\n"
            "  pip install pymupdf") from None
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
    except Exception as exc:                                      # noqa: BLE001
        raise RuntimeError(f"Không mở được file PDF: {exc}") from exc
    try:
        pages = [(no, page.get_pixmap(dpi=dpi).tobytes("png"), "image/png")
                 for no, page in enumerate(doc, start=1)
                 if want is None or no in want]
        n_total = doc.page_count
    finally:
        doc.close()
    if not pages:
        raise RuntimeError(
            f"File PDF này có {n_total} trang, không có trang nào trong số cần quét "
            f"({', '.join(str(n) for n in sorted(want))})."
            if want else "File PDF không có trang nào.")
    return pages


def merge_pages(page_rows) -> dict:
    """Gộp kết quả nhiều trang → ``{person_id: {"name": …, "signed": bool}}``.

    `page_rows` là list các mảng 'rows' (một mảng / trang). Một mã nhân viên nếu
    xuất hiện ở nhiều trang thì CÓ chữ ký ở bất kỳ trang nào cũng tính là đã ký.
    """
    merged: dict = {}
    for rows in page_rows:
        for item in rows or ():
            pid = str(item.get("person_id") or "").strip()
            if not pid:
                continue
            cur = merged.setdefault(pid, {"name": "", "signed": False})
            if item.get("signed"):
                cur["signed"] = True
            if not cur["name"]:
                cur["name"] = (item.get("name") or "").strip()
    return merged


def _build_prompt(roster) -> str:
    """Dựng prompt kèm danh sách nhân viên (mã | họ tên) để AI đọc đúng person_id.

    Prompt tả MỘT TRANG (mỗi request chỉ gửi 1 trang ảnh). Nhấn hai điểm hay sai:
    phải gán chữ ký vào ĐÚNG dòng theo đường kẻ ngang, và KHÔNG cần chữ ký khớp
    tên — chỉ cần ô có nét viết tay.
    """
    lines = "\n".join(
        f"- {(r.get('code') or '').strip()} | {(r.get('full_name') or '').strip()}"
        for r in roster) or "- (không có)"
    return (
        "Bạn được cung cấp ẢNH SCAN MỘT TRANG của BẢNG ĐIỂM DANH khóa học mà nhân "
        "viên đã ký tên tay sau khi tham dự.\n"
        "Bảng danh sách tham dự (thường có tiêu đề PARTICIPANTS) gồm các cột: "
        "No. | PERSON ID | NAME | DEPARTMENT | SIGNATURE.\n\n"
        "Nhiệm vụ: đọc TỪNG dòng nhân viên trong bảng đó và xác định ô cột SIGNATURE "
        "của CHÍNH dòng đó có nét viết tay hay không.\n"
        "  • signed = true  nếu ô chữ ký của dòng đó có nét VIẾT TAY. Chữ ký thường "
        "viết to, chiếm gần hết ô.\n"
        "  • signed = false nếu ô chữ ký của dòng đó TRỐNG (chỉ có nền giấy trắng).\n"
        "  • Hãy căn theo các ĐƯỜNG KẺ NGANG của bảng để gán chữ ký vào ĐÚNG dòng, "
        "và kiểm tra lại bằng cách đếm số dòng từ trên xuống. Tuyệt đối không đẩy "
        "chữ ký sang dòng liền trên hoặc liền dưới.\n"
        "  • KHÔNG cần đọc nội dung chữ ký, KHÔNG cần chữ ký khớp với tên ở cột NAME. "
        "Chỉ cần ô có nét viết tay là signed = true.\n"
        "  • BỎ QUA bảng giảng viên/trainer ở đầu trang (nếu có) và mọi bảng không "
        "phải danh sách nhân viên tham dự. Nếu trang này không chứa bảng điểm danh "
        "nào (ví dụ trang lịch học, trang bìa) thì trả về rows = [].\n\n"
        "Với mỗi dòng, đọc 'person_id' (cột PERSON ID) và 'name' (cột NAME) để định "
        "danh. Nếu chữ ở cột PERSON ID khó đọc, hãy đối chiếu với danh sách bên dưới.\n\n"
        "Danh sách nhân viên đã ghi danh (mã | họ tên):\n"
        f"{lines}\n\n"
        "Chỉ trả về đúng JSON theo schema đã cho (mảng 'rows'), gồm TẤT CẢ các dòng "
        "nhân viên nhìn thấy trên TRANG NÀY, kèm cờ signed của từng dòng."
    )


def _interruptible_sleep(seconds: float, should_cancel=None) -> None:
    """Ngủ `seconds` giây nhưng cứ 0.2s lại kiểm tra Hủy; nếu Hủy → Cancelled."""
    remaining = seconds
    while remaining > 0:
        if should_cancel and should_cancel():
            raise Cancelled()
        chunk = 0.2 if remaining > 0.2 else remaining
        time.sleep(chunk)
        remaining -= chunk


def detect_signatures(api_key: str, model: str, file_bytes: bytes, mime_type: str,
                      roster, timeout: int = 180, on_retry=None,
                      should_cancel=None) -> dict:
    """Gửi 1 file scan cho Gemini → trả về dict {"rows": [{person_id, name, signed}]}.

    Tự thử lại khi lỗi tạm thời (429/5xx, mất mạng) với backoff tăng dần.
    `on_retry(attempt, wait, reason)` (nếu có) báo trước mỗi lần chờ.
    `should_cancel()` (nếu có) được kiểm tra thường xuyên → bấm Hủy có hiệu lực
    ngay, ném Cancelled. Ném RuntimeError với thông báo dễ hiểu nếu vẫn thất bại.
    """
    api_key = (api_key or "").strip()
    if not api_key:
        raise ValueError("Chưa có API key — hãy nhập API key Gemini ở Cài đặt.")
    prompt = _build_prompt(roster)

    last_error = ""
    for attempt in range(1, _MAX_RETRIES + 1):
        if should_cancel and should_cancel():
            raise Cancelled()
        try:
            return _detect_once(api_key, model, file_bytes, mime_type, prompt, timeout)
        except TransientError as exc:
            last_error = str(exc)
            if attempt == _MAX_RETRIES:
                break
            wait = _RETRY_BASE_DELAY * (2 ** (attempt - 1))
            if on_retry:
                on_retry(attempt, wait, last_error)
            _interruptible_sleep(wait, should_cancel)
    raise RuntimeError(f"Thất bại sau {_MAX_RETRIES} lần thử — {last_error}")


def _detect_once(api_key: str, model: str, file_bytes: bytes, mime_type: str,
                 prompt: str, timeout: int) -> dict:
    """Một lần gọi API. Ném TransientError nếu lỗi có thể thử lại."""
    body = {
        "contents": [{
            "parts": [
                {"text": prompt},
                {"inline_data": {
                    "mime_type": mime_type,
                    "data": base64.b64encode(file_bytes).decode("ascii"),
                }},
            ],
        }],
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
        raise TransientError(f"Lỗi kết nối mạng: {exc.reason}") from exc

    try:
        text = payload["candidates"][0]["content"]["parts"][0]["text"]
    except (KeyError, IndexError):
        reason = payload.get("promptFeedback", {}).get("blockReason")
        raise RuntimeError(
            "Không nhận được nội dung từ mô hình"
            + (f" (bị chặn: {reason})" if reason else ""))
    try:
        data = json.loads(text)
    except ValueError:
        raise RuntimeError("Mô hình trả về không đúng JSON.")
    if not isinstance(data, dict):
        raise RuntimeError("Mô hình trả về không đúng định dạng mong đợi.")
    return data
