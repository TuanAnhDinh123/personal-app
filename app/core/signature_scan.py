"""Nhận diện CHỮ KÝ trên ảnh/PDF scan bằng Gemini (REST, urllib).

Luồng nghiệp vụ: HR in bảng điểm danh khóa học, phát cho nhân viên ký tay; sau
buổi học scan lại thành ảnh/PDF. Module này gửi file scan + danh sách nhân viên
đã ghi danh cho Gemini và hỏi: những AI đã KÝ (ô cột SIGNATURE có chữ ký tay)?
Trả về danh sách dòng ``{person_id, name, signed}`` để UI đối chiếu theo MÃ NHÂN
VIÊN rồi cập nhật trạng thái học sang "Completed".

Chỉ dùng thư viện chuẩn (urllib) — thống nhất với app.core.ai_cv_scan, KHÔNG cần
cài thêm gói. Cần API key Gemini (màn hình Cài đặt) + kết nối mạng. Hỗ trợ ảnh
(PNG/JPG/WEBP…) và PDF (gửi inline base64 như ai_cv_scan).
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
                                  "description": "true nếu ô cột SIGNATURE có chữ ký tay; false nếu trống"},
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


def _build_prompt(roster) -> str:
    """Dựng prompt kèm danh sách nhân viên (mã | họ tên) để AI đọc đúng person_id."""
    lines = "\n".join(
        f"- {(r.get('code') or '').strip()} | {(r.get('full_name') or '').strip()}"
        for r in roster)
    return (
        "Bạn được cung cấp ẢNH hoặc PDF SCAN của một BẢNG ĐIỂM DANH khóa học mà "
        "nhân viên đã ký tên tay sau khi tham dự.\n"
        "Bảng thường có các cột: No. | PERSON ID | NAME | Dept | SIGNATURE | notes.\n\n"
        "Nhiệm vụ: đọc TỪNG dòng nhân viên và xác định ô ở cột SIGNATURE có CHỮ KÝ "
        "TAY hay không.\n"
        "  • signed = true  nếu ô chữ ký có nét viết tay (chữ ký, ký nháy, tên viết tay…).\n"
        "  • signed = false nếu ô chữ ký để TRỐNG.\n"
        "Với mỗi dòng, đọc 'person_id' (cột PERSON ID) và 'name' (cột NAME) để định danh. "
        "Nếu chữ ở cột PERSON ID khó đọc, hãy đối chiếu với danh sách bên dưới.\n\n"
        "Danh sách nhân viên đã ghi danh (mã | họ tên):\n"
        f"{lines}\n\n"
        "Chỉ trả về đúng JSON theo schema đã cho (mảng 'rows'). Trả về TẤT CẢ các "
        "dòng nhìn thấy trên bảng, kèm cờ signed của từng dòng."
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
