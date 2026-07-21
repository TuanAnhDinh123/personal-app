# git command
## Đẩy code lên git
git add .
git commit -m "description"
git push origin main

## Lấy code về
git pull origin main


# Personal Toolbox

Ứng dụng desktop "hộp đồ nghề cá nhân" — gom nhiều tác vụ hằng ngày
(xử lý file, dữ liệu, văn phòng...) vào một app với giao diện sidebar.

## Chạy ứng dụng

```bash
pip install -r requirements.txt
python main.py
```

```bash
./Personal_Tool.bat
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
        ├── clean_data.py
        ├── scan_cv.py
        ├── ai_scan_cv.py       # Quét CV bằng AI (Gemini) → Excel đánh giá
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

## Tool: Quét CV bằng AI 🤖

`app/tools/ai_scan_cv.py` — gửi **nguyên file PDF** cho mô hình **Google Gemini**
để đọc hiểu và chấm điểm ứng viên theo JD (khác tool "Quét CV" chỉ dùng regex).

Giao diện gồm: ô nhập **API key**, ô chọn **model** (mặc định `gemini-3.6-flash`),
ô chọn **thư mục chứa CV (PDF)**, ô chọn **đường dẫn lưu file Excel**, và một
**text area nhập JD** (mô tả công việc). Kết quả xuất ra Excel gồm: họ tên, ngày
sinh, email, SĐT, **điểm phù hợp (0–100)** + nhận xét, **ưu điểm / nhược điểm**.

> - Lấy API key miễn phí tại <https://aistudio.google.com/apikey>.
> - Gọi API bằng thư viện chuẩn (`urllib`) nên **không cần cài thêm gói**; chỉ
>   cần `openpyxl` để ghi Excel và có kết nối mạng.
> - Việc quét chạy trong **luồng nền** kèm thanh tiến trình + nhật ký, không treo
>   giao diện. Mỗi CV là một lần gọi API (tốn quota theo số file).
> - Mặc định dùng `gemini-3.6-flash`. Free tier giới hạn theo TỪNG model (vd 5
>   request/phút, 20/ngày); chạm trần sẽ báo 429/503 — đổi sang model khác
>   (`gemini-3.5-flash`, `gemini-2.5-flash`) để dùng hạn ngạch riêng, hoặc bật
>   billing để nâng giới hạn. Tool tự thử lại tối đa 4 lần khi gặp 429/5xx.

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
