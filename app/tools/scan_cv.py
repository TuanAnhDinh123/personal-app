from app.core.base_tool import BaseTool
from app.ui import widgets


class ScanCvTool(BaseTool):
    name = "Quét CV (AI)"
    description = "Đọc nhiều CV (PDF) bằng AI và xuất bảng tổng hợp ứng viên ra Excel."
    icon = "🤖"
    category = "Trí tuệ nhân tạo"
    order = 10
    action_label = "Bắt đầu quét"
    action_style = "success"

    def build_body(self, parent):
        widgets.file_row(parent, "Thư mục chứa CV", mode="folder")
        widgets.file_row(parent, "Xuất bảng tổng hợp ra", mode="save")
        widgets.section_label(parent, "Thông tin cần trích")
        widgets.checkbox(parent, "Họ tên, Email, Số điện thoại")
        widgets.checkbox(parent, "Số năm kinh nghiệm")
        widgets.checkbox(parent, "Kỹ năng & Học vấn")
        widgets.hint(
            parent,
            "⚠️ Tính năng này dùng Claude API — cần cấu hình API key trước khi chạy thật.",
        )
