"""Sửa ảnh bằng AI (Gemini — dòng model ảnh "Nano Banana").

Người dùng đưa 1 ảnh + mô tả thay đổi bằng lời ("xóa người bên trái", "chỉnh
tông màu ấm hơn"…) → model trả về ẢNH MỚI đã sửa. Gọi REST `generateContent`
(chỉ dùng urllib chuẩn, như các tính năng AI khác của app) với
`responseModalities = ["TEXT", "IMAGE"]`.

Mỗi lượt sửa là một lời gọi ĐỘC LẬP (ảnh + yêu cầu), không giữ hội thoại phía
server: muốn sửa tiếp trên kết quả thì giao diện đưa ảnh kết quả vào làm ảnh
gốc của lượt sau. Cách này đơn giản, không phải lưu thought signature, và mỗi
lượt chỉ tốn tiền đúng một ảnh.

Lời người dùng gõ được BỌC thêm một khung chỉ dẫn (`build_prompt`) nhắc model
chỉ sửa đúng phần được yêu cầu, giữ nguyên phần còn lại — lỗi hay gặp nhất của
model ảnh là tiện tay "vẽ lại" cả bức (đổi mặt người, đổi bố cục).

Cần API key Gemini (màn hình Settings) đã bật thanh toán, và kết nối mạng.
"""
import base64
import json
import time
import urllib.error
import urllib.request
from dataclasses import dataclass

from app.core import settings

_API_URL = ("https://generativelanguage.googleapis.com/v1beta/models/"
            "{model}:generateContent")

# Model ảnh đã biết, theo thứ tự gợi ý ở ô chọn (model mới chưa có trong list
# vẫn hiện, xếp sau). Danh sách thật lấy từ ListModels theo API key.
KNOWN_MODELS = [
    "gemini-nano-banana-2.1",        # khuyên dùng: chất lượng tốt, giá vừa
    "gemini-3.1-flash-lite-image",   # rẻ + nhanh nhất, chỉ ra ảnh 1K
    "gemini-3-pro-image",            # sửa khó tốt nhất, đắt gấp mấy lần
    "gemini-3.1-flash-image",        # thế hệ trước của 2.1
    "gemini-2.5-flash-image",        # bản cũ, Google khuyên chuyển sang Lite
]

IMAGE_SIZES = ("1K", "2K", "4K")

# Định dạng gửi thẳng được lên API (đuôi file → mime). Định dạng khác (BMP,
# GIF…) giao diện tự chuyển sang PNG/JPEG trước khi gửi.
UPLOAD_MIME = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".webp": "image/webp",
}

# Trần dung lượng ảnh gửi đi. Request inline tối đa ~20 MB SAU base64 (phình
# thêm 1/3) → giữ ảnh gốc dưới 10 MB cho an toàn; lớn hơn thì giao diện thu nhỏ.
MAX_UPLOAD_BYTES = 10 * 1024 * 1024

# Prompt mẫu cho người mới: (nhãn chip, câu mẫu). Phần trong [ngoặc vuông] là
# chỗ người dùng điền — giao diện tự bôi đen ngoặc đầu tiên để gõ đè luôn.
# Câu mẫu theo cấu trúc Google khuyên cho model ảnh: nói rõ sửa gì + cách hòa
# vào ảnh + phần nào phải giữ nguyên.
PROMPT_TEMPLATES = [
    ("Remove an object",
     "Remove the [object] from the image and fill the area naturally so it "
     "blends with the surrounding background."),
    ("Replace the background",
     "Replace the background with [new background]. Keep the main subject "
     "exactly as it is, with clean natural edges and lighting that matches "
     "the new background."),
    ("Change one thing",
     "Change only the [element] to [new look]. Keep everything else in the "
     "image exactly the same."),
    ("Add something",
     "Add [element] to [where in the image], matching the style, lighting and "
     "perspective of the scene."),
    ("Colour tone",
     "Make the colours [warmer / cooler / more vivid / softer]. Do not change "
     "any content, details or composition."),
    ("Fix the lighting",
     "Brighten the dark areas and balance the exposure so the photo looks "
     "evenly and naturally lit. Do not change any content."),
    ("Enhance quality",
     "Improve the overall quality: sharpen details, reduce noise and improve "
     "clarity, without changing any content."),
]


@dataclass
class EditResult:
    image: bytes        # ảnh kết quả (đúng byte model trả về, không nén lại)
    mime_type: str      # vd "image/png"
    note: str = ""      # lời model kèm theo ảnh (nếu có)


class Cancelled(Exception):
    """Người dùng bấm Hủy trong lúc đang chờ thử lại."""


class _TransientError(Exception):
    """Lỗi tạm thời (quá tải, mất kết nối) — nên thử lại."""


_RETRY_STATUS = {500, 502, 503, 504}
_MAX_ATTEMPTS = 3
_RETRY_BASE_DELAY = 3     # giây; chờ tăng dần 3s, 6s


def list_image_models(api_key, timeout=30):
    """Model ảnh mà API key gọi được → cho ô chọn ở màn hình Cài đặt.

    Lọc bớt cho người dùng đỡ rối: chỉ giữ model ảnh dùng được với
    generateContent, bỏ các bản `-preview` (đã có bản chính thức cùng tên).
    """
    names = [n for n in settings.fetch_models(api_key, timeout)
             if settings.is_image_model(n) and "preview" not in n]

    def rank(name):
        try:
            return (KNOWN_MODELS.index(name), name)
        except ValueError:
            return (len(KNOWN_MODELS), name)

    return sorted(set(names), key=rank)


def supports_image_size(model):
    """Model có nhận tham số `imageSize` không. Bản Lite chỉ ra 1K và dòng 2.5
    không có tham số này — gửi kèm thì API báo lỗi."""
    return "lite" not in model and not model.startswith("gemini-2.5")


def build_prompt(request):
    """Bọc yêu cầu của người dùng trong khung chỉ dẫn giữ nguyên phần còn lại."""
    return (
        "Edit the attached image. Apply ONLY the change requested below and "
        "keep everything else exactly as it is — people, faces, expressions, "
        "poses, composition, framing, lighting, colours, textures and any "
        "text — unless the request explicitly asks to change them. Keep the "
        "same aspect ratio and the same visual style as the original. The "
        "result must look natural, with no visible seams, smudges or "
        "artefacts. The request may be written in any language.\n\n"
        f"Requested change: {request.strip()}"
    )


def extension_for(mime_type):
    """Đuôi file hợp với mime ảnh kết quả (mặc định .png)."""
    return {"image/jpeg": ".jpg", "image/webp": ".webp"}.get(mime_type, ".png")


def edit_image(api_key, model, image, mime_type, request, image_size="",
               timeout=300, on_retry=None, should_cancel=None):
    """Gửi ảnh + yêu cầu sửa cho Gemini → EditResult.

    `image_size` ("1K"/"2K"/"4K") chỉ gửi kèm khi model hỗ trợ. Tự thử lại khi
    gặp lỗi tạm thời (5xx, mất mạng), gọi `on_retry(attempt, wait, reason)`
    trước mỗi lần chờ. Ném ValueError khi thiếu đầu vào, RuntimeError với
    thông báo dễ hiểu khi API từ chối / không trả ảnh, Cancelled khi Hủy.
    """
    api_key = (api_key or "").strip()
    if not api_key:
        raise ValueError("No API key — please enter your Gemini API key in Settings.")
    if not (request or "").strip():
        raise ValueError("Please describe the change you want.")
    model = (model or "").strip() or settings.DEFAULTS["image_model"]

    gen_config = {"responseModalities": ["TEXT", "IMAGE"]}
    if image_size in IMAGE_SIZES and supports_image_size(model):
        gen_config["imageConfig"] = {"imageSize": image_size}
    body = {
        "contents": [{
            "role": "user",
            "parts": [
                {"inline_data": {
                    "mime_type": mime_type,
                    "data": base64.b64encode(image).decode("ascii"),
                }},
                {"text": build_prompt(request)},
            ],
        }],
        "generationConfig": gen_config,
    }
    data = json.dumps(body).encode("utf-8")

    last_error = ""
    for attempt in range(1, _MAX_ATTEMPTS + 1):
        if should_cancel and should_cancel():
            raise Cancelled()
        try:
            return _parse(_post(api_key, model, data, timeout))
        except _TransientError as exc:
            last_error = str(exc)
            if attempt == _MAX_ATTEMPTS:
                break
            wait = _RETRY_BASE_DELAY * (2 ** (attempt - 1))
            if on_retry:
                on_retry(attempt, wait, last_error)
            _sleep(wait, should_cancel)
    raise RuntimeError(f"The AI service is busy right now — please try again "
                       f"in a minute.\n\n({last_error})")


def _sleep(seconds, should_cancel):
    """Ngủ nhưng cứ 0.2s kiểm tra Hủy một lần để bấm Hủy có hiệu lực ngay."""
    end = time.monotonic() + seconds
    while time.monotonic() < end:
        if should_cancel and should_cancel():
            raise Cancelled()
        time.sleep(0.2)


def _post(api_key, model, data, timeout):
    req = urllib.request.Request(
        _API_URL.format(model=model),
        data=data,
        headers={"Content-Type": "application/json", "x-goog-api-key": api_key},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", "ignore")
        msg = detail
        try:
            msg = json.loads(detail).get("error", {}).get("message", detail)
        except ValueError:
            pass
        if exc.code in _RETRY_STATUS:
            raise _TransientError(f"HTTP {exc.code}: {msg}") from exc
        raise RuntimeError(_friendly_http_error(exc.code, msg, model)) from exc
    except urllib.error.URLError as exc:
        raise _TransientError(f"Network error: {exc.reason}") from exc
    except TimeoutError as exc:
        raise _TransientError("The request timed out.") from exc


def _friendly_http_error(code, msg, model):
    """Lỗi HTTP không thử lại được → câu giải thích người dùng tự xử lý được."""
    if code == 429:
        return ("Your Gemini API key has hit its rate limit or quota (HTTP 429). "
                "Image editing needs billing enabled on the key — check your "
                "plan at aistudio.google.com, or wait a minute and try again."
                f"\n\n({msg})")
    if code == 404:
        return (f"The model \"{model}\" isn't available for this API key. "
                f"Pick another image model in Settings.\n\n({msg})")
    if code in (401, 403) or "API key" in msg:
        return (f"Gemini rejected the API key (HTTP {code}). Check the key in "
                f"Settings.\n\n({msg})")
    return f"HTTP {code}: {msg}"


# finishReason → lời giải thích khi model không trả ảnh.
_FINISH_REASONS = {
    "SAFETY": "The request was blocked by the AI's safety filters.",
    "IMAGE_SAFETY": "The AI declined to create this image for safety reasons.",
    "PROHIBITED_CONTENT": "The AI declined this request (prohibited content).",
    "IMAGE_PROHIBITED_CONTENT": "The AI declined this request (prohibited content).",
    "IMAGE_RECITATION": "The AI declined because the result would copy "
                        "protected content too closely.",
    "RECITATION": "The AI declined because the result would copy protected "
                  "content too closely.",
    "NO_IMAGE": "The AI answered without producing an image.",
}


def _parse(payload):
    """Rút ảnh + lời nhắn từ phản hồi generateContent.

    Model ảnh có "nghĩ" trước khi vẽ: các part `thought: true` là ảnh/chữ nháp
    → bỏ qua; ảnh cuối cùng KHÔNG phải nháp mới là kết quả.
    """
    candidates = payload.get("candidates") or []
    if not candidates:
        reason = (payload.get("promptFeedback") or {}).get("blockReason")
        raise RuntimeError(
            "The AI returned nothing"
            + (f" — the request was blocked ({reason}). Try rephrasing it."
               if reason else ". Please try again."))

    cand = candidates[0]
    image = mime = None
    notes = []
    for part in (cand.get("content") or {}).get("parts") or []:
        if part.get("thought"):
            continue
        blob = part.get("inlineData") or part.get("inline_data")
        if blob and blob.get("data"):
            image = base64.b64decode(blob["data"])
            mime = blob.get("mimeType") or blob.get("mime_type") or "image/png"
        elif part.get("text"):
            notes.append(part["text"].strip())
    note = "\n".join(n for n in notes if n)

    if image is None:
        reason = cand.get("finishReason", "")
        msg = _FINISH_REASONS.get(reason, "The AI didn't return an image.")
        msg += " Try rephrasing the request."
        if note:
            msg += f"\n\nThe AI said: {note}"
        raise RuntimeError(msg)
    return EditResult(image=image, mime_type=mime, note=note)
