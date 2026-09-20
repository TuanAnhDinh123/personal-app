"""Cấu hình CHUNG của toàn app (màn hình "Cài đặt").

Khác với cấu hình riêng của từng tool (mỗi tool giữ một section riêng trong
`config.py`), đây là các thiết lập DÙNG CHUNG cho nhiều tính năng: API key AI,
model mặc định… Người dùng chỉnh ở màn hình **Cài đặt**; mọi tool đọc lại qua
module này thay vì tự hỏi API key.

Muốn thêm một thiết lập chung mới: thêm khóa vào `DEFAULTS`, rồi thêm ô nhập
tương ứng ở `app/ui/settings_page.py`.
"""
import datetime
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
    "course_template_path": "",     # file Excel mẫu để xuất roster khóa học
    # File Excel "Personnel Data" của HR — nguồn để đồng bộ bảng `employees`
    # (nút Sync with Excel ở màn hình Employees).
    "hc_excel_path": "",
    "birthday_images_folder": "",   # thư mục ảnh thiệp sinh nhật (Canva Bulk Create, tên file = mã NV)
    "birthday_from_account": "",    # địa chỉ Outlook dùng để gửi mail chúc mừng sinh nhật (khác tài khoản mặc định)
    # Tiêu đề mail sinh nhật. Phải lưu ở đây (không để trong widget như trước):
    # app tự gửi lúc mở máy, khi người dùng chưa từng mở trang tool.
    "birthday_subject": "Happy Birthday, {name}!",
    # Số ngày còn GỬI BÙ khi đúng hôm sinh nhật không ai mở app (nghỉ phép, cuối
    # tuần, máy tắt). Quá hạn thì mail bị đánh "Missed", không tự gửi nữa.
    "birthday_catchup_days": 3,
    # ── Quyết định thôi việc (app/core/resignation_decision.py) ──
    "resignation_template_path": "",   # file Word mẫu (chứa các trường MERGEFIELD)
    "resignation_output_folder": "",   # thư mục nhận các file .docx xuất ra
    # Số thứ tự cho quyết định TIẾP THEO + năm mà số đó thuộc về. Hai khóa đi
    # cùng nhau: sang năm mới, `resignation_decision_year` khác năm hiện tại là
    # dấu hiệu để đánh số lại từ 1 (xem `next_decision_number`).
    "resignation_decision_no": 1,
    "resignation_decision_year": 0,
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


def update(**changes):
    """Sửa vài thiết lập chung, GIỮ NGUYÊN các khóa còn lại.

    Khác `save()` (ghi đè cả section) — dùng khi một tool tự lưu lựa chọn của
    nó (vd giờ hẹn gửi mail sinh nhật) mà không được đụng tới cấu hình khác.
    """
    values = load()
    values.update(changes)
    save(values)


def get(key, default=""):
    """Lấy nhanh một thiết lập chung."""
    return load().get(key, default)


def next_decision_number(data=None, today=None) -> tuple:
    """Số quyết định thôi việc dùng cho lần xuất tới → (năm, số thứ tự).

    Sổ quyết định đánh số lại từ 1 mỗi năm, nên số đã lưu chỉ còn giá trị khi
    nó thuộc về NĂM NAY; năm đã sang thì trả về (năm nay, 1). Việc đánh số lại
    chỉ được GHI XUỐNG khi thực sự xuất file (`save_decision_number`) — mở màn
    hình Cài đặt xem rồi thoát ra thì không đụng gì tới cấu hình.
    """
    data = load() if data is None else data
    year = (today or datetime.date.today()).year
    try:
        saved_year = int(data.get("resignation_decision_year") or 0)
        number = int(data.get("resignation_decision_no") or 1)
    except (TypeError, ValueError):
        saved_year, number = 0, 1
    if saved_year != year:
        return year, 1
    return year, max(1, number)


def save_decision_number(year: int, number: int) -> None:
    """Ghi lại số quyết định cho lần xuất SAU (kèm năm mà số đó thuộc về)."""
    update(resignation_decision_year=int(year), resignation_decision_no=int(number))


def list_models(api_key, timeout=30):
    """Gọi Gemini ListModels API → trả về list tên model hỗ trợ generateContent.

    Dùng cho ô chọn Model ở màn hình Cài đặt: bấm vào để lấy danh sách model
    hiện có ứng với API key. Chỉ dùng thư viện chuẩn (urllib) — không cần cài
    thêm gói. Ném ValueError nếu thiếu key, RuntimeError nếu API/mạng lỗi.
    """
    api_key = (api_key or "").strip()
    if not api_key:
        raise ValueError("No API key — please enter your Gemini API key first.")

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
        raise RuntimeError(f"Network error: {exc.reason}") from exc

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
