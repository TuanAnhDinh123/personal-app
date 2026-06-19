from app.core.base_tool import BaseTool
from app.ui import widgets


class BatchRenameTool(BaseTool):
    name = "Đổi tên hàng loạt"
    description = "Đổi tên toàn bộ file trong một thư mục theo quy tắc."
    icon = "🏷️"
    category = "Tệp & Tài liệu"
    order = 30
    action_label = "Đổi tên"

    def build_body(self, parent):
        widgets.file_row(parent, "Thư mục chứa file", mode="folder")
        widgets.text_row(parent, "Tiền tố thêm vào", placeholder="VD: 2026_")
        widgets.text_row(parent, "Thay chuỗi (cũ -> mới)", placeholder="VD: ' ' -> '_'")
        widgets.checkbox(parent, "Đánh số thứ tự (001, 002, ...)", checked=True)
