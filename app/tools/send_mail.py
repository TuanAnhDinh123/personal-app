from app.core.base_tool import BaseTool
from app.ui import widgets


class SendMailTool(BaseTool):
    name = "Gửi mail theo lịch"
    description = "Đọc danh sách từ Excel và tạo email gửi theo lịch qua Outlook."
    icon = "📧"
    category = "Văn phòng"
    order = 10
    action_label = "Tạo mail"

    def build_body(self, parent):
        widgets.file_row(parent, "File Excel danh sách", mode="file")
        widgets.hint(
            parent,
            "File cần 3 cột: email, time (giờ gửi), Linked (đường dẫn file đính kèm).",
        )
        widgets.text_row(parent, "Tiêu đề mail", placeholder="VD: Thông báo tháng 6")
