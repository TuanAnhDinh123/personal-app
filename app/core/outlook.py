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
_OL_APPOINTMENT = 26          # Class của item lịch đã tồn tại
_OL_APPOINTMENT_ITEM = 1      # loại item khi tạo sự kiện lịch mới (olAppointmentItem)
_OL_MAIL_ITEM = 0

# PR_SMTP_ADDRESS — dùng để lấy địa chỉ SMTP thật từ một địa chỉ Exchange (X.500)
_PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"


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
        # IncludeRecurrences + Restrict không tương thích — duyệt thủ công thay thế.
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        today = datetime.date.today()
        result = []
        for item in items:
            try:
                t_start = _to_datetime(getattr(item, "Start", None))
                if t_start is None:
                    continue
                if t_start.date() < today:
                    continue
                if t_start.date() > today:
                    break
                if item.Class != _OL_APPOINTMENT:
                    continue
                result.append({
                    "subject": str(getattr(item, "Subject", "") or ""),
                    "start": t_start,
                    "end": _to_datetime(getattr(item, "End", None)),
                    "location": str(getattr(item, "Location", "") or ""),
                    "organizer": str(getattr(item, "Organizer", "") or ""),
                    "categories": str(getattr(item, "Categories", "") or ""),
                })
            except Exception:
                continue

        return result
    finally:
        pythoncom.CoUninitialize()


def _resolve_smtp(recipient):
    """Lấy email SMTP thật của một Recipient (kể cả khi là địa chỉ Exchange)."""
    try:
        addr = str(getattr(recipient, "Address", "") or "")
    except Exception:
        addr = ""
    if "@" in addr:
        return addr
    # Địa chỉ Exchange (X.500) -> tra ra SMTP qua ExchangeUser
    try:
        entry = recipient.AddressEntry
        exch = entry.GetExchangeUser()
        if exch is not None and exch.PrimarySmtpAddress:
            return str(exch.PrimarySmtpAddress)
    except Exception:
        pass
    # Dự phòng: đọc thẳng thuộc tính PR_SMTP_ADDRESS
    try:
        smtp = recipient.AddressEntry.PropertyAccessor.GetProperty(_PR_SMTP_ADDRESS)
        if smtp and "@" in str(smtp):
            return str(smtp)
    except Exception:
        pass
    return addr


def _attendee_emails(item):
    """Danh sách email của người tham dự một sự kiện lịch (đã khử trùng lặp)."""
    emails = []
    try:
        for r in item.Recipients:
            addr = _resolve_smtp(r)
            if addr and addr not in emails:
                emails.append(addr)
    except Exception:
        pass
    return emails


def appointments_between(start_date, end_date):
    """Trả về các sự kiện lịch có ngày bắt đầu trong [start_date, end_date].

    `start_date`, `end_date` là datetime.date. Mỗi phần tử là dict giống
    today_appointments() nhưng có thêm:
        attendees  – list email người tham dự
        entry_id   – ID duy nhất của sự kiện (để nhớ đã xử lý)
    """
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        calendar = ns.GetDefaultFolder(_OL_FOLDER_CALENDAR)

        items = calendar.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        # Giới hạn theo khoảng ngày để KHÔNG phải duyệt toàn bộ lịch (nhanh hơn
        # rất nhiều — lịch lặp có thể nở ra hàng nghìn instance). Nếu Restrict
        # lỗi (định dạng ngày theo locale…) thì lùi về duyệt toàn bộ.
        start_dt = datetime.datetime.combine(start_date, datetime.time.min)
        end_dt = datetime.datetime.combine(end_date, datetime.time.max)
        fmt = "%m/%d/%Y %I:%M %p"
        restriction = (
            "[Start] <= '" + end_dt.strftime(fmt) + "' AND "
            "[End] >= '" + start_dt.strftime(fmt) + "'"
        )
        try:
            items = items.Restrict(restriction)
        except Exception:
            pass

        result = []
        for item in items:
            try:
                t_start = _to_datetime(getattr(item, "Start", None))
                if t_start is None:
                    continue
                if t_start.date() < start_date:
                    continue
                if t_start.date() > end_date:
                    break            # đã sắp xếp tăng dần -> qua khỏi khoảng thì dừng
                if item.Class != _OL_APPOINTMENT:
                    continue
                result.append({
                    "subject": str(getattr(item, "Subject", "") or ""),
                    "start": t_start,
                    "end": _to_datetime(getattr(item, "End", None)),
                    "location": str(getattr(item, "Location", "") or ""),
                    "organizer": str(getattr(item, "Organizer", "") or ""),
                    "categories": str(getattr(item, "Categories", "") or ""),
                    "attendees": _attendee_emails(item),
                    "entry_id": str(getattr(item, "EntryID", "") or ""),
                })
            except Exception:
                continue

        return result
    finally:
        pythoncom.CoUninitialize()


def create_appointment(subject, start, duration_minutes=30, body="",
                       reminder_minutes=15):
    """Tạo một sự kiện lịch (kèm lời nhắc) trên Outlook. `start` là datetime."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        appt = outlook.CreateItem(_OL_APPOINTMENT_ITEM)
        appt.Subject = subject
        appt.Start = start
        appt.Duration = duration_minutes
        if body:
            appt.Body = body
        appt.ReminderSet = True
        appt.ReminderMinutesBeforeStart = reminder_minutes
        appt.Save()
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
