# Personal Toolbox

Ứng dụng desktop "hộp đồ nghề cá nhân" — gom nhiều tác vụ hằng ngày
(xử lý file, dữ liệu, AI...) vào một app với giao diện sidebar.

## Chạy ứng dụng

```bash
pip install -r requirements.txt
python main.py
```

## Cấu trúc dự án

```
personal-app/
├── main.py                  # Điểm khởi động
├── requirements.txt
├── icon_app.ico
└── app/
    ├── core/
    │   ├── base_tool.py     # Lớp cha của mọi tool
    │   └── registry.py      # Tự động phát hiện tool
    ├── ui/
    │   ├── theme.py         # Màu sắc & font (đổi giao diện ở đây)
    │   ├── widgets.py       # Widget dựng sẵn (ô chọn file, checkbox...)
    │   └── main_window.py   # Cửa sổ chính: sidebar + nội dung
    └── tools/               # MỖI FILE = 1 TÁC VỤ
        ├── merge_excel.py
        ├── pdf_tools.py
        ├── batch_rename.py
        ├── clean_data.py
        ├── scan_cv.py
        └── send_mail.py
```

## Thêm một tác vụ mới

Tạo 1 file trong `app/tools/`, ví dụ `app/tools/my_tool.py`:

```python
from app.core.base_tool import BaseTool
from app.ui import widgets


class MyTool(BaseTool):
    name = "Tên công cụ"
    description = "Mô tả ngắn."
    icon = "✨"
    category = "Nhóm hiển thị"
    order = 10
    action_label = "Thực hiện"

    def build_body(self, parent):
        widgets.file_row(parent, "Chọn file", mode="file")
        # ... thêm ô nhập tùy ý

    def run(self):
        # Gắn logic thật ở đây. Mặc định chỉ hiện thông báo hoàn thành.
        super().run()
```

App sẽ **tự động** nhận tool mới và hiện trong sidebar — không cần sửa file nào khác.

## Đóng gói thành .exe

```bash
pip install pyinstaller
pyinstaller --onefile --noconsole --icon=icon_app.ico --name personal_app --collect-submodules app.tools main.py
```

> `--collect-submodules app.tools` là bắt buộc, vì tool được nạp động;
> thiếu nó thì bản .exe sẽ không thấy tool nào.
