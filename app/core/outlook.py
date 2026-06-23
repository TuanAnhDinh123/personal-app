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


def today_appointments():
    """Trả về danh sách sự kiện lịch của HÔM NAY.

    Mỗi phần tử là dict: subject, start, end, location, organizer, categories.
    """
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        calendar = ns.GetDefaultFolder(_OL_FOLDER_CALENDAR)

        items = calendar.Items
        items.IncludeRecurrences = True      # nở các lịch lặp lại
        items.Sort("[Start]")

        today = datetime.date.today()
        start = datetime.datetime(today.year, today.month, today.day)
        end = start + datetime.timedelta(days=1)
        fmt = "%m/%d/%Y %I:%M %p"             # định dạng Outlook Restrict hiểu
        restriction = (
            f"[Start] >= '{start.strftime(fmt)}' "
            f"AND [Start] < '{end.strftime(fmt)}'"
        )

        result = []
        for item in items.Restrict(restriction):
            try:
                if item.Class != _OL_APPOINTMENT:
                    continue
            except Exception:
                pass
            result.append({
                "subject": str(getattr(item, "Subject", "") or ""),
                "start": _to_datetime(getattr(item, "Start", None)),
                "end": _to_datetime(getattr(item, "End", None)),
                "location": str(getattr(item, "Location", "") or ""),
                "organizer": str(getattr(item, "Organizer", "") or ""),
                "categories": str(getattr(item, "Categories", "") or ""),
            })
        result.sort(key=lambda a: a["start"] or datetime.datetime.max)
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
