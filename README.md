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
    ├── outlook.py                        # Outlook COM (lịch · gửi mail · thư mời họp)
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
chạy (ví dụ "đã quét hôm nay chưa"). Xem mẫu ở `app_qt/tools/interview_gate.py`.

## Tool: Mở cổng lịch phỏng vấn 🛂

`app_qt/tools/interview_gate.py` — mỗi sáng khi mở app sẽ tự quét **lịch Outlook
hôm nay**, tìm sự kiện có tiêu đề chứa từ khóa phỏng vấn (cấu hình được), rồi
**soạn sẵn mail** nhờ team Security mở cổng và **hiện ra cho bạn xem/sửa trước
khi gửi**. Chỉ quét 1 lần/ngày; vẫn có nút bấm tay để quét bất cứ lúc nào.

> Cần Windows + Outlook + `pywin32`. Trên môi trường khác app vẫn chạy, chỉ
> báo là tính năng quét/gửi mail không khả dụng.

## Tool: Quét CV bằng AI 🤖

`app_qt/tools/ai_scan_cv.py` — gửi **nguyên file PDF** cho mô hình **Google Gemini**
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

[app_qt/tools/candidate_db.py](app_qt/tools/candidate_db.py) — quản lý hồ sơ ứng
viên + danh mục tuyển dụng, lưu bằng **SQLite** ngay trên máy
(`%APPDATA%\PersonalToolbox\candidates.sqlite`).

Màn hình chính (ỨNG VIÊN) gồm **ô tìm kiếm toàn văn**, hàng lọc *Position ·
Department · Status · Batch*, **bảng kết quả** có cột tick chọn, và toolbar:

| Nút | Cần tick hồ sơ? | Việc |
|-----|-----------------|------|
| **View details** | có (1 hoặc nhiều) | mở modal xem chi tiết từng hồ sơ |
| **Update status** | có | đổi trạng thái hàng loạt (xem bên dưới) |
| **Send email** | có | mời phỏng vấn / gửi thư cảm ơn qua Outlook (xem bên dưới) |
| **Export to Excel** | có | xuất các hồ sơ đã tick ra `.xlsx` (file đã có thì **ghi nối thêm**) |
| **Add** | không | thêm hồ sơ nhập tay |
| **Reload** | không | tải lại bảng |

*Add* nằm ở **cụm bên phải cạnh Reload**, tông neutral: toolbar chia hai vùng —
trái là thao tác trên **các hồ sơ đang tick**, phải là thao tác **cấp trang**.
Hồ sơ giờ chủ yếu vào DB qua tool *Quét CV bằng AI*, nhập tay chỉ còn là trường
hợp lẻ. **Sửa** một hồ sơ: **double-click vào dòng**; **xóa**: mở form sửa rồi
bấm *Delete* trong đó — cả hai không có nút riêng trên toolbar.

- **Chống trùng**: khi thêm mới (hoặc sửa) ứng viên, nếu **trùng email hoặc SĐT**
  với người đã có, tool cảnh báo và cho quyết định vẫn lưu hay không.
- **Master data** tách thành nhóm **Master Data** riêng ở sidebar, gồm 7 trang:
  **Departments · Employee types · Levels · Cost centers · Positions ·
  Mail templates · Courses** — mỗi trang là một bảng + thanh
  *Add / Edit / Delete / Reload* (dùng chung `CrudTablePanel`), thêm/sửa/xóa
  riêng. (Các trang này không hiện thẻ ở Trang chủ.) Bốn trang danh mục nhân sự
  đã có **dữ liệu nạp sẵn** từ `Code.xlsx`:
  20 bộ phận · 6 loại nhân viên · 12 cấp bậc · 42 cost center.
  Muốn thêm/bớt cột hay ô nhập của một trang → sửa `_master_specs()` trong
  [app_qt/tools/candidate_db.py](app_qt/tools/candidate_db.py).
- **Mẫu mail dùng chung**: trang **Mail templates** (bảng `mail_templates`) giữ
  mọi mẫu mail — *tên · loại · CC · tiêu đề · nội dung (rich text)*. Loại lấy từ
  `MAIL_TEMPLATE_TYPE_CHOICES` (Interview Round 1/2/3 · Application Thank You ·
  Notification · Offer · Rejection) và chủ yếu để phân nhóm cho dễ tìm, không
  ràng buộc — **trừ `Application Thank You`**: nút *Send email* lọc đúng loại này
  cho ô chọn thư cảm ơn (hằng `cv_schema.MAIL_TEMPLATE_TYPE_THANK_YOU`).
  Ngoài CRUD, trang này có thêm nút **Duplicate**: chọn 1 dòng → tạo bản
  sao y hệt, tên thêm hậu tố `_copy` (trùng nữa thì `_copy2`, `_copy3`…) rồi mở
  luôn form sửa bản mới.
- **Mỗi vị trí gán 3 mẫu mail** cho **3 vòng phỏng vấn**: form của trang
  *Positions* có 3 ô chọn (Interview Round 1/2/3) trỏ tới `mail_templates` qua
  `positions.mail_template_r1_id / _r2_id / _r3_id`. Danh sách vòng khai báo ở
  `cv_schema.INTERVIEW_ROUNDS` (kèm trạng thái ứng viên gợi ý sau khi gửi thư mời
  vòng đó).
- **JD nằm trong vị trí**: mỗi vị trí chỉ có **đúng 1 mô tả công việc**, nên
  *file JD* nhập ngay trong form của trang **Vị trí tuyển dụng**
  (cột `positions.jd_file_path`); tiêu đề JD luôn lấy theo **tên vị trí**. Không
  còn bảng `job_descriptions` lẫn trang master "Mô tả công việc (JD)" riêng.
- Nút **Update status** (đổi trạng thái hàng loạt): tick **một hoặc nhiều** ứng
  viên đang ở **CÙNG một trạng thái** → modal hiện trạng thái hiện tại và ô
  *Move to* điền sẵn **bước kế tiếp** trong luồng (sửa được, chọn bất kỳ nhãn
  nào trong `CANDIDATE_STATUS_CHOICES`); bấm *OK* mới ghi xuống DB. Nếu các hồ
  sơ đang ở trạng thái khác nhau thì app **báo lỗi kèm danh sách từng nhóm** và
  không đổi gì.
- Nút **Send email**: tick **một hoặc nhiều** ứng viên đang ở
  **CÙNG một trạng thái** (lệch nhau thì báo lỗi kèm danh sách từng nhóm, giống
  *Update status*) → hộp thoại chọn **loại mail muốn gửi**: 3 **vòng phỏng vấn**
  (mẫu lấy theo vị trí ứng tuyển) hoặc **Application Thank You** (thư cảm ơn đã
  ứng tuyển — mẫu chọn thẳng trong hộp thoại vì không gắn với vị trí nào). Hộp
  thoại hiện **trạng thái hiện tại** của các hồ sơ và **chọn sẵn vòng suy ra từ
  trạng thái đó** (*Short List* → vòng 1, *First Interview* →
  vòng 2, *Second Interview* → vòng 3, còn lại → vòng 1; bảng tra ở
  `cv_schema.INTERVIEW_ROUND_BY_STATUS`), vẫn đổi tay được.
- **Chọn một vòng phỏng vấn** → hộp thoại chọn
  ngày/giờ cho **từng người** (người sau mặc định nối tiếp ngay
  sau người trước, nút *Skip* để bỏ qua một người) → app mở bấy nhiêu **cửa sổ
  Meeting của Outlook** đã điền sẵn người nhận, CC, giờ và nội dung theo mẫu mail
  của vòng đó, **kèm file CV** của ứng viên (bỏ qua nếu chưa có CV hoặc
  đường dẫn không còn đúng). Ứng viên mà vị trí **chưa gán mẫu cho vòng đang mời**
  thì bị bỏ qua và liệt kê lại cuối lượt. Người dùng chỉ việc duyệt từng cửa sổ,
  thêm phòng họp nếu cần rồi bấm **Send** — Outlook lo cả ba việc: gửi mail mời,
  tạo lịch, đặt phòng. Mở xong, app hiện modal **cập nhật trạng thái** cho đúng
  những ứng viên đó (điền sẵn trạng thái của vòng vừa mời — vòng 1 →
  *First Interview*, vòng 2 → *Second Interview*…, đổi được từng người); bấm
  *Update* mới ghi xuống DB. Placeholder trong mẫu (`{name} {possion} {date}
  {time_start} {time_end}`) vẫn nhận đúng kể cả khi bị bôi đậm/đổi màu, và định
  dạng đó được giữ cho giá trị thay vào.
- **Chọn Application Thank You** → **không hỏi giờ, không phải thư mời họp**: app
  mở thẳng **cửa sổ mail thường của Outlook** cho từng ứng viên, điền sẵn người
  nhận · CC · tiêu đề · nội dung từ mẫu (`outlook.create_mail`, không đính kèm
  CV), người dùng xem lại rồi bấm **Send**. **Trạng thái ứng viên giữ nguyên** —
  không hiện modal cập nhật trạng thái. Chỉ thay `{name} {possion} {position}`;
  `{date}`/`{time…}` (nếu lỡ có trong mẫu) được **giữ nguyên** để nhìn thấy mà
  sửa trước khi gửi. Chưa có mẫu nào loại này thì app báo và mời tạo ở trang
  **Mail templates**.
- **Mở CV**: click thẳng vào tên file ở cột *CV file* trong bảng (cột này hiển
  thị dạng link, không có nút riêng). Ứng viên **chưa gắn CV**, hoặc file đã bị
  **di chuyển / đổi tên** (đường dẫn trong DB không còn đúng) → tool mời chọn lại
  file và **tự lưu đường dẫn mới** vào DB để lần sau khỏi hỏi.

**Đường dẫn file** lưu thẳng vào cột `candidates.cv_file_path` và
`positions.jd_file_path` (không dùng bảng riêng — file thực tế đã nằm sẵn trên
máy). Xem thảo luận về xử lý đường dẫn bị lệch ở cuối mục.

> Cờ ở `BaseTool`: `show_on_home=False` để ẩn thẻ khỏi Trang chủ (vẫn hiện ở
> sidebar), `fills_height=True` để trang chiếm full chiều cao khi phóng to cửa
> sổ.

Thiết kế cơ sở dữ liệu tách riêng để dễ chỉnh:

| File | Vai trò |
|------|---------|
| `app/core/cv_schema.py` | **Thiết kế DB** — toàn bộ bảng dưới dạng SQL (`SCHEMA_SQL`) kèm chú thích. Sửa cấu trúc DB ở đây; có sẵn `MIGRATIONS` để thêm cột an toàn cho DB đã có dữ liệu, và `DATA_MIGRATIONS` để sửa dữ liệu sẵn có (chạy **một lần** mỗi file .db, đánh dấu ở `app_meta` với khóa `data:<tên lượt>`). |
| `app/core/cv_repository.py` | **Tầng truy cập dữ liệu** — kết nối SQLite + CRUD generic cho mọi bảng. Giao diện chỉ gọi hàm, không đụng SQL. |
| `app_qt/tools/candidate_db.py` | **Giao diện** tool + form nhập liệu tổng quát. |

**Các bảng** (quan hệ mềm, không dùng khóa ngoại; mọi cột cho phép NULL trừ PK):
`departments` (phòng ban) → `positions` (vị trí, **kèm JD**: `jd_file_path`) →
`candidates` (ứng viên, có `cv_file_path`); `mail_templates` (mẫu mail dùng
chung) → `positions` qua 3 cột `mail_template_r1_id / _r2_id / _r3_id`;
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

> **3 cột `positions.mail_cc / mail_subject / mail_body` đã bị bỏ**: mẫu mail
> chuyển sang bảng dùng chung `mail_templates`. Khi mở tool trên DB cũ,
> `MIGRATIONS` **xóa hẳn** 3 cột đó — nội dung cũ **không** được chuyển sang;
> soạn lại ở trang **Mail templates** rồi gán cho từng vị trí ở form vị trí.

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
