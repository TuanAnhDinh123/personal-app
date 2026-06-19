from app.core.base_tool import BaseTool
from app.ui import widgets


class MergeExcelTool(BaseTool):
    name = "Gộp file Excel"
    description = "Gộp nhiều file Excel (hoặc nhiều sheet) thành một file duy nhất."
    icon = "📊"
    category = "Tệp & Tài liệu"
    order = 10
    action_label = "Gộp file"

    def build_body(self, parent):
        widgets.section_label(parent, "Nguồn & đích")
        widgets.file_row(parent, "File Excel nguồn", mode="file")
        widgets.file_row(parent, "File Excel nguồn thứ 2", mode="file")
        widgets.file_row(parent, "Lưu kết quả vào", mode="save")
        widgets.dropdown(
            parent, "Kiểu gộp",
            ["Nối thêm dòng (concat)", "Ghép theo cột chung (merge)"],
        )
