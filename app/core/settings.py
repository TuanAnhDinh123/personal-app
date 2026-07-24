"""Cấu hình CHUNG của toàn app (màn hình "Cài đặt").

Khác với cấu hình riêng của từng tool (mỗi tool giữ một section riêng trong
`config.py`), đây là các thiết lập DÙNG CHUNG cho nhiều tính năng: API key AI,
model mặc định… Người dùng chỉnh ở màn hình **Cài đặt**; mọi tool đọc lại qua
module này thay vì tự hỏi API key.

Muốn thêm một thiết lập chung mới: thêm khóa vào `DEFAULTS`, rồi thêm ô nhập
tương ứng ở `app/ui/settings_page.py`.
"""
import json
import urllib.error
import urllib.request

from app.core import config

SECTION = "general"

# Endpoint liệt kê model của Gemini (dùng cho ô chọn Model ở màn hình Cài đặt).
_MODELS_URL = "https://generativelanguage.googleapis.com/v1beta/models"

DEFAULTS = {
    "api_key":  "",                 # API key Gemini, dùng chung cho các tính năng AI
    "ai_model": "gemini-3.6-flash",  # model mặc định khi gọi AI
}

# Section cũ (tool "Quét CV bằng AI" từng lưu API key/model ở đây). Dùng để tự
# chuyển cấu hình cũ sang cấu hình chung cho người đã dùng bản trước.
_LEGACY_SECTION = "ai_scan_cv"


def load():
    """Đọc toàn bộ cấu hình chung (gộp lên trên DEFAULTS).

    Nếu người dùng chưa từng lưu ở Cài đặt mà đã có API key/model cũ trong tool
    Quét CV bằng AI thì tự kế thừa sang, tránh phải nhập lại.
    """
    data = config.load(SECTION, DEFAULTS)
    if not data.get("api_key") or not data.get("ai_model"):
        legacy = config.load(_LEGACY_SECTION, {})
        if not data.get("api_key") and legacy.get("api_key"):
            data["api_key"] = legacy["api_key"]
        if not data.get("ai_model") and legacy.get("model"):
            data["ai_model"] = legacy["model"]
    return data


def save(values):
    """Ghi đè toàn bộ cấu hình chung rồi lưu xuống đĩa."""
    config.save(SECTION, values)


def get(key, default=""):
    """Lấy nhanh một thiết lập chung."""
    return load().get(key, default)


def list_models(api_key, timeout=30):
    """Gọi Gemini ListModels API → trả về list tên model hỗ trợ generateContent.

    Dùng cho ô chọn Model ở màn hình Cài đặt: bấm vào để lấy danh sách model
    hiện có ứng với API key. Chỉ dùng thư viện chuẩn (urllib) — không cần cài
    thêm gói. Ném ValueError nếu thiếu key, RuntimeError nếu API/mạng lỗi.
    """
    api_key = (api_key or "").strip()
    if not api_key:
        raise ValueError("Chưa có API key — hãy nhập API key Gemini trước.")

    req = urllib.request.Request(
        f"{_MODELS_URL}?pageSize=200",
        headers={"x-goog-api-key": api_key},
        method="GET",
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
        raise RuntimeError(f"HTTP {exc.code}: {msg}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Lỗi kết nối mạng: {exc.reason}") from exc

    models = []
    for m in payload.get("models", []):
        if "generateContent" not in m.get("supportedGenerationMethods", []):
            continue
        name = m.get("name", "")
        if name.startswith("models/"):
            name = name[len("models/"):]
        # Chỉ giữ dòng flash / flash-lite / pro (các model khác không dùng đến,
        # bỏ đi cho gọn). flash-lite đã nằm trong "flash".
        if name and ("flash" in name or "pro" in name):
            models.append(name)
    models.sort()
    return models
