# git command
## Đẩy code lên git
git add .
git commit -m "description"
git push origin main

## Lấy code về
git pull origin main


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
    │   ├── base_tool.py     # Lớp cha của mọi tool (+ hook startup tự chạy)
    │   ├── config.py        # Lưu/đọc cấu hình tool (JSON dùng chung)
    │   ├── outlook.py       # Đọc lịch & gửi mail qua Outlook (Windows)
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

### Tool chạy tự động khi mở app

Đặt `auto_startup = True` và ghi đè `startup(self, window)` trong class tool.
Mỗi lần mở app, `MainWindow` sẽ gọi `startup()` của các tool bật cờ này (sau
khi cửa sổ đã hiện). Dùng `app/core/config.py` để lưu trạng thái giữa các lần
chạy (ví dụ "đã quét hôm nay chưa"). Xem mẫu ở `app/tools/interview_gate.py`.

## Tool: Mở cổng lịch phỏng vấn 🛂

`app/tools/interview_gate.py` — mỗi sáng khi mở app sẽ tự quét **lịch Outlook
hôm nay**, tìm sự kiện có tiêu đề chứa từ khóa phỏng vấn (cấu hình được), rồi
**soạn sẵn mail** nhờ team Security mở cổng và **hiện ra cho bạn xem/sửa trước
khi gửi**. Chỉ quét 1 lần/ngày; vẫn có nút bấm tay để quét bất cứ lúc nào.

> Cần Windows + Outlook + `pywin32`. Trên môi trường khác app vẫn chạy, chỉ
> báo là tính năng quét/gửi mail không khả dụng.

## Đóng gói thành .exe

```bash
pip install pyinstaller
pyinstaller --onefile --noconsole --icon=icon_app.ico --name personal_app --collect-submodules app.tools --add-data "icon_app.ico;." main.py
```

> - `--collect-submodules app.tools`: nhét code của tất cả tool vào .exe (vì tool
>   được nạp động). Thiếu nó thì code tool không có trong bản đóng gói.
> - `--add-data "icon_app.ico;."`: gói kèm icon để cửa sổ hiển thị đúng icon.
>   (Trên Linux/macOS dùng dấu `:` thay vì `;`.)
>
> Lưu ý: `registry.py` đã xử lý cả trường hợp `--onefile` (module nằm trong
> archive, không trên đĩa). Nếu sửa cách quét tool, nhớ giữ logic đọc `toc` của
> PyInstaller, nếu không bản .exe sẽ hiện thiếu tool dù build báo thành công.
