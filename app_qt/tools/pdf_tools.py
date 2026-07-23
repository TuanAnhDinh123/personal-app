"""Gộp / Tách PDF (stub — chưa gắn logic) — bản PySide6."""
from app_qt import widgets
from app_qt.base_tool import BaseTool


class PdfTool(BaseTool):
    name = "Gộp / Tách PDF"
    description = "Gộp nhiều PDF thành một, hoặc tách một PDF thành nhiều trang."
    icon = "📄"
    category = "Tệp & Tài liệu"
    order = 20
    action_label = "Xử lý PDF"
    show_on_home = False

    def build_body(self, parent):
        widgets.dropdown(parent, "Chế độ", ["Gộp nhiều PDF", "Tách theo trang"])
        widgets.file_row(parent, "File / thư mục PDF", mode="file")
        widgets.file_row(parent, "Lưu vào", mode="folder")
