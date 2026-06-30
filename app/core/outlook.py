"""Cầu nối tới Outlook trên Windows (đọc lịch + gửi mail) qua COM.

Chỉ chạy được trên Windows có cài Outlook. Trên môi trường khác (ví dụ
lúc dev trên Linux) các hàm sẽ báo lỗi rõ ràng thay vì làm sập app —
nhờ `available()` kiểm tra trước.

Tham chiếu hằng số Outlook:
    olFolderCalendar = 9      (thư mục Lịch)
    olAppointmentItem = 26    (Class của một sự kiện lịch)
    olMailItem = 0            (loại item khi tạo mail mới)
"""
import datetime

_OL_FOLDER_CALENDAR = 9
_OL_APPOINTMENT = 26
_OL_MAIL_ITEM = 0


def available():
    """True nếu đang ở Windows và import được pywin32."""
    try:
        import win32com.client  # noqa: F401
        import pythoncom        # noqa: F401
        return True
    except Exception:
        return False


def _to_datetime(value):
    """Đổi giá trị thời gian của Outlook (pywintypes time) sang datetime."""
    try:
        return datetime.datetime(
            value.year, value.month, value.day,
            value.hour, value.minute, value.second,
        )
    except Exception:
        return None


def today_appointments(log=None):
    """Trả về danh sách sự kiện lịch của HÔM NAY.

    Mỗi phần tử là dict: subject, start, end, location, organizer, categories.
    log: callback(str) tùy chọn để ghi debug.
    """
    import pythoncom
    import win32com.client

    def _log(msg):
        if log:
            log(msg)

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        calendar = ns.GetDefaultFolder(_OL_FOLDER_CALENDAR)

        _log(f"[DEBUG] Calendar folder: {calendar.Name} | Items.Count={calendar.Items.Count}")

        items = calendar.Items
        # IncludeRecurrences + Restrict không tương thích (Restrict trả về 2147483647).
        # Giải pháp: sort theo Start rồi duyệt thủ công, break khi qua ngày hôm nay.
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        today = datetime.date.today()
        _log(f"[DEBUG] Đang lọc ngày: {today}")

        result = []
        skipped_before = 0
        skipped_class = 0
        for item in items:
            try:
                t_start = _to_datetime(getattr(item, "Start", None))
                if t_start is None:
                    continue
                if t_start.date() < today:
                    skipped_before += 1
                    continue
                if t_start.date() > today:
                    break   # đã qua hôm nay, dừng (danh sách đã sort)
                item_class = item.Class
                subj = str(getattr(item, "Subject", "") or "")
                _log(f"[DEBUG]   {t_start.strftime('%H:%M')} | Class={item_class} | {subj!r}")
                if item_class != _OL_APPOINTMENT:
                    skipped_class += 1
                    _log(f"[DEBUG]   → bỏ qua (Class={item_class}, cần {_OL_APPOINTMENT})")
                    continue
                result.append({
                    "subject": subj,
                    "start": t_start,
                    "end": _to_datetime(getattr(item, "End", None)),
                    "location": str(getattr(item, "Location", "") or ""),
                    "organizer": str(getattr(item, "Organizer", "") or ""),
                    "categories": str(getattr(item, "Categories", "") or ""),
                })
            except Exception as e:
                _log(f"[DEBUG]   item lỗi: {e}")
                continue

        _log(f"[DEBUG] Kết quả: {len(result)} appointment(s) | bỏ qua trước hôm nay={skipped_before} | sai class={skipped_class}")
        return result
    finally:
        pythoncom.CoUninitialize()


def send_mail(to, subject, body, cc=""):
    """Tạo và GỬI một mail qua Outlook (dùng tài khoản mặc định)."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(_OL_MAIL_ITEM)
        mail.To = to
        if cc:
            mail.CC = cc
        mail.Subject = subject
        mail.Body = body
        mail.Send()
    finally:
        pythoncom.CoUninitialize()
