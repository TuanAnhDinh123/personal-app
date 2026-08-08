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

## Giao diện (PySide6)

Giao diện dựng bằng **PySide6 (Qt)** — phong cách dashboard sáng, sidebar tối làm
điểm nhấn, style bằng **QSS** (gần giống CSS).

- **Đổi giao diện toàn app**: sửa `app_qt/theme.py` (bảng màu) hoặc
  `app_qt/theme.qss` ("CSS" của app). Icon line (SVG) ở `app_qt/assets/icons/`.
- **Mỗi tool = 1 file** trong `app_qt/tools/` kế thừa `app_qt.base_tool.BaseTool`;
  `app_qt/registry.py` tự phát hiện — thêm 1 file là có 1 tool.
- **Tách bạch UI ↔ logic**: giao diện ở `app_qt/`, còn **logic nghiệp vụ thuần
  Python** (openpyxl / Excel COM / Gemini / SQLite / Outlook / OCR) nằm ở
  `app/core/` — hoàn toàn không phụ thuộc giao diện. Tool Qt chỉ là lớp vỏ gọi
  vào `app/core`.
- **Component dùng chung** ở `app_qt/components/`: `table` (bảng), `form_dialog`
  (form nhập liệu), `crud_panel` (CRUD master data), `progress_dialog` &
  `task` (chạy nền QThread), `dialog_base` (khung hộp thoại).

## Cấu trúc dự án

```
personal-app/
├── main.py                  # Điểm khởi động (PySide6)
├── requirements.txt
├── icon_app.ico
├── app_qt/                  # GIAO DIỆN (PySide6)
│   ├── theme.py / theme.qss # Bảng màu + "CSS" (QSS)
│   ├── widgets.py           # Widget dựng sẵn (API .get()/.set()) + icon SVG
│   ├── base_tool.py         # Lớp cha tool
│   ├── registry.py          # Tự phát hiện tool
│   ├── main_window.py       # Cửa sổ chính (frameless): sidebar + nội dung
│   ├── dialogs.py           # Hộp thoại tùy biến (info/error/confirm)
│   ├── settings_page.py     # Trang Cài đặt
│   ├── icons.py             # Map emoji → tên icon line
│   ├── richtext.py          # Ô soạn thảo rich text (QTextEdit)
│   ├── assets/icons/        # Bộ icon line (SVG)
│   ├── components/          # table · form_dialog · crud_panel · progress_dialog · task · dialog_base · cv_rename
│   └── tools/               # MỖI FILE = 1 TÁC VỤ (giao diện)
└── app/core/                # LOGIC NGHIỆP VỤ (thuần Python, KHÔNG dính UI)
    ├── config.py · settings.py          # cấu hình
    ├── cv_repository.py · cv_schema.py   # SQLite quản lý CV ứng viên
    ├── outlook.py                        # Outlook COM (đọc lịch + gửi mail)
    ├── payroll_split.py                  # Tách bảng lương (Excel COM)
    ├── quarter_bonus.py                  # Thưởng quý (Excel COM)
    ├── ai_cv_scan.py                     # Quét CV bằng Gemini
    ├── cv_scan.py                        # Chuẩn hóa tên file CV + template Excel
    ├── pdf_text.py                       # PDF → Text (+ OCR Tesseract)
    └── reminder_logic.py                 # Helper nhắc phản hồi phỏng vấn
```

## Thêm một tác vụ mới

Tạo 1 file trong `app_qt/tools/`, ví dụ `app_qt/tools/my_tool.py`:

```python
from app_qt import widgets
from app_qt.base_tool import BaseTool


class MyTool(BaseTool):
    name = "Tên công cụ"
    description = "Mô tả ngắn."
    icon = "✨"
    category = "Nhóm hiển thị"
    order = 10
    action_label = "Thực hiện"

    def build_body(self, parent):
        self.file = widgets.file_row(parent, "Chọn file", mode="file")
        # ... thêm ô nhập tùy ý (widgets.text_row / text_area / dropdown / checkbox)

    def run(self):
        # Gắn logic thật ở đây. Đọc giá trị bằng .get(); báo kết quả:
        #   self.info("Xong", "...")  /  self.error("Lỗi", "...")
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
để đọc hiểu và chấm điểm ứng viên theo JD.

Giao diện gồm: ô chọn **thư mục chứa CV (PDF)**, ô chọn **vị trí tuyển dụng** và
ô **yêu cầu bổ sung cho AI**. (API key + model đặt ở ⚙️ Cài đặt, dùng chung cho
cả app.) Mỗi CV quét xong được ghi thẳng vào bảng **Candidates** gồm: họ tên,
ngày sinh, email, SĐT, **điểm phù hợp (0–100)** + nhận xét, **ưu điểm / nhược
điểm**; ứng viên trùng email/SĐT thì hỏi *ghi đè* hay *xuất ra Excel*.

**Nút “Normalize file names”** (cạnh nút quét) mở hộp **chuẩn hóa tên file CV**.
Trong hộp: chọn **thư mục CV**, **prefix code**
+ **start code** và danh sách **từ nhiễu** cần bỏ khỏi tên ứng viên; bấm *Preview
& rename* để xem bảng đối chiếu (sửa tay được cột tên) rồi đổi tên hàng loạt
theo dạng `{prefix}{startcode}_{Tên ứng viên}.pdf`. Cấu hình lưu ở section
`scan_cv`; giao diện ở
[app_qt/components/cv_rename.py](app_qt/components/cv_rename.py), logic ở
[app/core/cv_scan.py](app/core/cv_scan.py).

> **JD lấy theo vị trí, không chọn file bằng tay**: chọn 1 **vị trí tuyển dụng**
> thì tool đọc file JD đã gắn cho vị trí đó (`positions.jd_file_path`). Ô chọn
> hiển thị luôn tên file JD, vị trí nào chưa gắn thì ghi *“⚠ chưa có file JD”*;
> danh sách tự nạp lại mỗi lần bung nên vị trí/JD vừa thêm là thấy ngay. Thêm JD
> ở **Master Data → Vị trí tuyển dụng**.

> - Lấy API key miễn phí tại <https://aistudio.google.com/apikey>.
> - Gọi API bằng thư viện chuẩn (`urllib`) nên **không cần cài thêm gói**; chỉ
>   cần `openpyxl` để ghi Excel và có kết nối mạng.
> - Việc quét chạy trong **luồng nền** kèm thanh tiến trình + nhật ký, không treo
>   giao diện. Mỗi CV là một lần gọi API (tốn quota theo số file).
> - Mặc định dùng `gemini-3.6-flash`. Free tier giới hạn theo TỪNG model (vd 5
>   request/phút, 20/ngày); chạm trần sẽ báo 429/503 — đổi sang model khác
>   (`gemini-3.5-flash`, `gemini-2.5-flash`) để dùng hạn ngạch riêng, hoặc bật
>   billing để nâng giới hạn. Tool tự thử lại tối đa 4 lần khi gặp 429/5xx.

## Tool: Quản lý CV ứng viên 🗂️

`app/tools/candidate_db.py` — quản lý hồ sơ ứng viên + danh mục tuyển dụng,
lưu bằng **SQLite** ngay trên máy (`%APPDATA%\PersonalToolbox\candidates.sqlite`).

Màn hình chính (ỨNG VIÊN): **thanh tìm kiếm** (tên / email / SĐT + lọc vị trí +
trạng thái), **bảng kết quả** (double-click để sửa), và các nút **Thêm mới ·
Sửa · Xóa · Mở CV · Nhập từ Excel · Tải lại**.

- **Chống trùng**: khi Thêm mới (hoặc Sửa) ứng viên, nếu **trùng email hoặc SĐT**
  với người đã có, tool cảnh báo và cho quyết định vẫn lưu hay không. Khi *Nhập
  từ Excel* có thể chọn **bỏ qua các bản trùng** (email/SĐT — cả trong file lẫn
  so với DB).
- **Master data** tách thành nhóm **Master Data** riêng ở sidebar, gồm 6 trang:
  **Departments · Employee types · Levels · Cost centers · Positions · Courses**
  — mỗi trang là một bảng + thanh *Add / Edit / Delete / Reload* (dùng chung
  `CrudTablePanel`), thêm/sửa/xóa riêng. (Các trang này không hiện thẻ ở Trang
  chủ.) Bốn trang danh mục nhân sự đã có **dữ liệu nạp sẵn** từ `Code.xlsx`:
  20 bộ phận · 6 loại nhân viên · 12 cấp bậc · 42 cost center.
  Muốn thêm/bớt cột hay ô nhập của một trang → sửa `_master_specs()` trong
  [app_qt/tools/candidate_db.py](app_qt/tools/candidate_db.py).
- **JD nằm trong vị trí**: mỗi vị trí chỉ có **đúng 1 mô tả công việc**, nên
  *file JD* nhập ngay trong form của trang **Vị trí tuyển dụng**
  (cột `positions.jd_file_path`); tiêu đề JD luôn lấy theo **tên vị trí**. Không
  còn bảng `job_descriptions` lẫn trang master "Mô tả công việc (JD)" riêng.
- Nút **📥 Nhập từ Excel** đọc thẳng file kết quả do tool *Quét CV bằng AI* xuất
  ra (tự khớp cột theo tiêu đề), ghi hàng loạt vào DB. Có thể chọn thư mục chứa
  CV để lưu **đường dẫn đầy đủ** vào cột `cv_file_path`.
- Nút **📂 Mở CV** mở file CV của ứng viên đang chọn. Nếu file đã bị **di chuyển
  / đổi tên** (đường dẫn trong DB không còn đúng), tool mời chọn lại vị trí và
  **tự lưu đường dẫn mới** vào DB để lần sau khỏi hỏi.

**Đường dẫn file** lưu thẳng vào cột `candidates.cv_file_path` và
`positions.jd_file_path` (không dùng bảng riêng — file thực tế đã nằm sẵn trên
máy). Xem thảo luận về xử lý đường dẫn bị lệch ở cuối mục.

> Cờ ở `BaseTool`: `show_on_home=False` để ẩn thẻ khỏi Trang chủ (vẫn hiện ở
> sidebar), `fills_height=True` để trang chiếm full chiều cao khi phóng to cửa
> sổ. Tool chưa gắn logic (`clean_data`, `pdf_tools`) đã đặt `show_on_home=False`.

Thiết kế cơ sở dữ liệu tách riêng để dễ chỉnh:

| File | Vai trò |
|------|---------|
| `app/core/cv_schema.py` | **Thiết kế DB** — toàn bộ bảng dưới dạng SQL (`SCHEMA_SQL`) kèm chú thích. Sửa cấu trúc DB ở đây; có sẵn mục `MIGRATIONS` để thêm cột an toàn cho DB đã có dữ liệu. |
| `app/core/cv_repository.py` | **Tầng truy cập dữ liệu** — kết nối SQLite + CRUD generic cho mọi bảng. Giao diện chỉ gọi hàm, không đụng SQL. |
| `app/tools/candidate_db.py` | **Giao diện** tool + form nhập liệu tổng quát. |

**Các bảng** (quan hệ mềm, không dùng khóa ngoại; mọi cột cho phép NULL trừ PK):
`departments` (phòng ban) → `positions` (vị trí, **kèm JD**: `jd_file_path`, và
mẫu mail mời PV) → `candidates` (ứng viên, có `cv_file_path`);
ngoài ra `employees` (nhân viên) và `courses` ↔ `course_employees` (đào tạo).
Đường dẫn file lưu thẳng vào cột — không có bảng file riêng.

**Bảng `employees` bám sát `Master HC file.xlsx`** (sheet *Master file*) — gần
như mỗi cột trong file có một cột tương ứng trong DB; chú thích
`-- <Tiêu đề Excel>` ghi ngay cạnh từng cột trong `app/core/cv_schema.py`.

- **Import khớp theo TÊN cột, KHÔNG theo thứ tự cột** (`_EXCEL_HEADER_MAP` ở
  [app_qt/tools/employee_db.py](app_qt/tools/employee_db.py)): tiêu đề được
  chuẩn hóa (chữ thường · gộp khoảng trắng · bỏ ký hiệu font Wingdings · NFC)
  nên đổi chỗ cột hay thêm cột lạ đều không làm vỡ import. Cột **trùng tên**
  ("Issued date", "Changing date") khai báo bằng *tuple* → lần xuất hiện thứ n
  lấy field thứ n.
- **Trạng thái làm việc suy ra từ `termination_date`** (không có cột `status`):
  có ngày (khác rỗng) = đã nghỉ việc. Truy vấn danh sách trả thêm cột
  `work_status` ("Working"/"Resigned") để hiển thị; mặc định ẩn người đã nghỉ
  (tick *Include resigned employees* để xem cả).
- **4 cột text được tra sang bảng danh mục**: "Function (Common)" → `departments`
  (short_name) · "New Cost center" → `cost_centers` (code) · "IBC/DBC/WC" →
  `employee_types` (code) · "Job level" → `levels` (level_name). Không khớp thì
  để trống liên kết và báo lại danh sách cuối lần import.
- Các cột **file Excel tự tính** (Age, Age range, Year of service, Length of
  service, Year of birthday) được lưu như **ảnh chụp lúc import** — app không
  tính lại. Các cột "Legal Entity (Company)", "Position status", "Business Unit
  (Department)", "BC/WC", "STT", "Birthday"… không lưu vì đã có nguồn khác hoặc
  chỉ là cột phụ trợ trong file (xem chú thích trong schema).
- Bảng ~90 cột → giao diện chỉ hiện vài cột mặc định, bật/tắt thêm ở modal
  **Columns** — cột gom theo nhóm (Identity / Contact / Education…), lựa chọn
  được ghi nhớ. Sửa nhóm & cột mặc định ở `_EMP_COLUMN_GROUPS` /
  `_EMP_DEFAULT_COLUMNS` trong `app_qt/tools/employee_db.py`.

**Bảng danh mục (master data)** — nạp sẵn dữ liệu từ file `Code.xlsx`:
`departments` (tên + mã viết tắt), `employee_types` (WC/WCA/IBC/IBCA/DBC/DBCA +
nhóm Blue/White Collar), `cost_centers` (mã VN1001… + Group Function
VNPlant/Corporate/R&D — dùng để gom nhân viên theo nhóm khi tính chi phí vận
hành), `levels` (Director, Manager, Officer…).

> **Dữ liệu khởi tạo**: khai báo ở `SEED_DATA` (cv_schema.py), nạp bởi
> `_seed_master_data()` — mỗi khối chỉ chạy **một lần** cho mỗi file .db (đánh
> dấu ở bảng `app_meta`), và bỏ qua dòng đã tồn tại nên không tạo bản ghi trùng.
> Muốn nạp lại: xóa khóa `seed:<bảng>:v1` trong `app_meta`. Bổ sung danh mục về
> sau: thêm dòng vào `rows` rồi tăng `version`.

> **Bảng `job_descriptions` đã bị bỏ**: JD nằm trong `positions`. Khi mở tool
> trên DB cũ, `init_db()` **xóa hẳn** bảng đó — dữ liệu JD cũ **không** được
> chuyển sang, cột `jd_file_path` để trống, nhập lại ở form vị trí.

> Vì thiết kế cố tình **không dùng khóa ngoại**, các cột `*_id` chỉ là tham
> chiếu mềm — ứng dụng tự đảm bảo liên kết. `init_db()` tự tạo bảng khi mở tool;
> nếu bảng cũ đang **trống** mà lệch cấu trúc, nó tự dựng lại (không mất dữ liệu).

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
