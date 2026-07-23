"""Nhắc phản hồi kết quả phỏng vấn cho ứng viên.

Luồng hoạt động:
  • Bấm "🔄 Quét lịch" để quét lịch Outlook từ HÔM NAY lùi về 1 tháng, lấy các
    sự kiện có tiêu đề chứa từ khóa phỏng vấn ("interview", "interview
    invitation"…). Cấu hình (từ khóa, mẫu mail) giống tính năng "Gửi mail theo
    lịch". Không quét tự động khi mở app (tránh chặn/đơ lúc khởi động).
  • Hiển thị các lịch đó thành 1 bảng: cột 1 là tiêu đề (subject) của lịch,
    cột 2 có 2 nút Yes / No cho câu hỏi "đã phản hồi ứng viên hay chưa?".
        - Bấm YES  → coi như đã phản hồi: xóa khỏi bảng, lần sau không hiện nữa.
        - Bấm NO   → hiện popup xác nhận "Gửi phản hồi ngay bây giờ?"
              · Yes → mở popup SOẠN MAIL (gửi xong thì xóa record như bấm Yes).
              · No  → mở popup ĐẶT LỊCH NHẮC khác.
  • Người nhận mail lấy từ chính lịch đó (loại bỏ email nội bộ, vd datalogic.com).

Các lịch đã xử lý được nhớ trong config (theo EntryID) nên không hiện lại.
"""
import calendar
import datetime
import re


from app.core import config, outlook

SECTION = "reminder"

DEFAULTS = {
    "keywords": "interview, interview invitation, phỏng vấn, pv",
    "exclude_domain": "datalogic.com",
    "subject": "Interview Result — {position}",
    "body": (
        "Dear {name},\n\n"
        "Thank you for taking the time to interview for the {position} "
        "position with us. We truly appreciate your interest and the effort "
        "you put into the process.\n\n"
        "After careful consideration, we would like to inform you of the "
        "result of your interview:\n\n"
        "…\n\n"
        "Should you have any questions, please feel free to contact us.\n\n"
        "Best regards,\n"
    ),
    "dismissed": [],      # EntryID các lịch đã xử lý (đã phản hồi / đã gửi mail)
}


def _month_ago(d):
    """Trả về ngày cùng 'ngày' nhưng lùi lại 1 tháng (kẹp cuối tháng nếu cần)."""
    month = d.month - 1 or 12
    year = d.year - (1 if d.month == 1 else 0)
    day = min(d.day, calendar.monthrange(year, month)[1])
    return d.replace(year=year, month=month, day=day)


def _extract_name(subject):
    """Lấy tên ứng viên: phần sau 'Mr.'/'Ms.'/'Mrs.' nếu có, không thì cả subject."""
    m = re.search(r'\b(?:Mr|Ms|Mrs)\.\s*(.+)', subject, re.IGNORECASE)
    return m.group(1).strip() if m else subject.strip()


def _extract_position(subject):
    """Lấy vị trí ứng tuyển từ tiêu đề dạng 'Interview invitation for <vị trí> - Mr. ...'.

    Bỏ phần tên (từ 'Mr.'/'Ms.'/'Mrs.' trở đi) và các dấu ngăn cách ở cuối.
    """
    m = re.search(r'interview\s+(?:invitation\s+)?for\s+(.+)', subject,
                  re.IGNORECASE)
    if not m:
        return ""
    rest = re.split(r'\b(?:Mr|Ms|Mrs)\.', m.group(1), maxsplit=1,
                    flags=re.IGNORECASE)[0]
    return rest.strip().strip("-–—").strip()


def _fill_template(text, appt):
    """Thay các placeholder {name}/{position}/{subject}/{date}/{time} trong mẫu."""
    start = appt.get("start")
    return (
        text.replace("{name}", _extract_name(appt["subject"]))
            .replace("{position}", _extract_position(appt["subject"]))
            .replace("{subject}", appt["subject"])
            .replace("{date}", start.strftime("%d/%m/%Y") if start else "")
            .replace("{time}", start.strftime("%H:%M") if start else "")
    )
