"""Cấu hình CHUNG của toàn app (màn hình "Cài đặt").

Khác với cấu hình riêng của từng tool (mỗi tool giữ một section riêng trong
`config.py`), đây là các thiết lập DÙNG CHUNG cho nhiều tính năng: API key AI,
model mặc định… Người dùng chỉnh ở màn hình **Cài đặt**; mọi tool đọc lại qua
module này thay vì tự hỏi API key.

Muốn thêm một thiết lập chung mới: thêm khóa vào `DEFAULTS`, rồi thêm ô nhập
tương ứng ở `app/ui/settings_page.py`.
"""
from app.core import config

SECTION = "general"

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
