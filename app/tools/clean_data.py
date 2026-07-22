from app.core.base_tool import BaseTool
from app.ui import widgets


class CleanDataTool(BaseTool):
    name = "Làm sạch dữ liệu"
    description = "Chuẩn hóa file Excel: xóa dòng trùng, cắt khoảng trắng, chuẩn hóa SĐT/email."
    icon = "🧹"
    category = "Dữ liệu"
    order = 10
    action_label = "Làm sạch"
    show_on_home = False   # chưa gắn logic → tạm ẩn ở Trang chủ

    def build_body(self, parent):
        widgets.file_row(parent, "File cần làm sạch", mode="file")
        widgets.file_row(parent, "Lưu kết quả vào", mode="save")
        widgets.section_label(parent, "Tùy chọn")
        widgets.checkbox(parent, "Xóa dòng trùng lặp")
        widgets.checkbox(parent, "Cắt khoảng trắng thừa")
        widgets.checkbox(parent, "Chuẩn hóa số điện thoại", checked=False)
        widgets.checkbox(parent, "Chuẩn hóa email về chữ thường", checked=False)
