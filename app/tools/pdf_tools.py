from app.core.base_tool import BaseTool
from app.ui import widgets


class PdfTool(BaseTool):
    name = "Gộp / Tách PDF"
    description = "Gộp nhiều PDF thành một, hoặc tách một PDF thành nhiều trang."
    icon = "📄"
    category = "Tệp & Tài liệu"
    order = 20
    action_label = "Xử lý PDF"
    show_on_home = False   # chưa gắn logic → tạm ẩn ở Trang chủ

    def build_body(self, parent):
        widgets.dropdown(parent, "Chế độ", ["Gộp nhiều PDF", "Tách theo trang"])
        widgets.file_row(parent, "File / thư mục PDF", mode="file")
        widgets.file_row(parent, "Lưu vào", mode="folder")
