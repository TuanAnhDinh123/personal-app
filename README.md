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
- **Component dùng chung** ở `app_qt/components/`: `table` (bảng, có menu chuột phải tùy biến), `form_dialog`
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
│   ├── widgets.py           # Widget dựng sẵn (API .get()/.set()) + DateEdit + icon SVG
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
    ├── employee_sync.py                  # Đối chiếu bảng employees ↔ file Excel HC
    ├── outlook.py                        # Outlook COM (lịch · gửi mail · thư mời họp)
    ├── payroll_split.py                  # Tách bảng lương (Excel COM)
    ├── quarter_bonus.py                  # Thưởng quý (Excel COM)
    ├── ai_cv_scan.py                     # Quét CV bằng Gemini
    ├── cv_scan.py                        # Chuẩn hóa tên file CV + trích text CV
    ├── candidate_export.py               # Xuất sheet "Candidates" ra Excel
    ├── pdf_text.py                       # PDF → Markdown (+ OCR bằng Gemini)
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

## Tool: PDF → Text 📝

[app_qt/tools/pdf_to_text.py](app_qt/tools/pdf_to_text.py) — đọc PDF thành văn bản
**giữ tương đối bố cục gốc** (tiêu đề, danh sách, bảng) để copy sang chỗ khác.
Logic ở [app/core/pdf_text.py](app/core/pdf_text.py).

Khung đọc chiếm trọn bề ngang, **chỉ đọc**, luôn hiển thị bản đã dựng định dạng.
Dải nút chọn-một (nút đang bật tô đậm) đổi phạm vi **Whole document** /
**Page by page**; ảnh trang gốc chỉ hiện khi bấm *Show original*. Nút *Copy* đặt
lên clipboard cả bản HTML lẫn bản chữ trần nên dán vào Word giữ được bố cục, dán
vào ô text thường vẫn sạch. Xuất ra `.md` hoặc `.txt`.

> Không có chế độ sửa tay: tool để LẤY text ra, cho sửa mà không có chỗ lưu bản
> sửa thì chuyển trang là mất. Trên giao diện cũng tránh chữ "Markdown" và tránh
> kiểu một nút bấm-đổi-nhãn (nhìn nhãn không rõ đó là thứ đang xem hay thứ sắp
> chuyển sang).

### Cách đọc

| Loại trang | Cách xử lý | Cần gì |
|---|---|---|
| Có lớp text (PDF xuất từ Word) | `PyMuPDF` dựng lại Markdown từ toạ độ chữ: tiêu đề theo cỡ chữ, bảng qua `find_tables`, tự tách 2 cột và bỏ header/footer lặp | Không cần gì thêm, chạy offline |
| Ảnh scan | Render trang ra PNG rồi nhờ **Gemini** đọc thành Markdown | API key Gemini ở màn hình **Settings** + mạng |

Ô **Mode** chọn: `Auto` (chỉ OCR trang scan), `Text layer only` (không gọi AI),
`Force OCR` (ép OCR cả trang đã có lớp text — dùng khi lớp text bị lỗi font).

> Cố ý **không dùng Tesseract**: nó là chương trình native phải cài riêng kèm gói
> ngôn ngữ, máy không có quyền admin sẽ tắc. Cả tính năng này giờ chỉ cần
> `pip install pymupdf` (wheel thuần) + `urllib` chuẩn.

## Tool: Mở cổng lịch phỏng vấn 🛂

`app_qt/tools/interview_gate.py` — mỗi sáng khi mở app sẽ tự quét **lịch Outlook
hôm nay**, tìm sự kiện có tiêu đề chứa từ khóa phỏng vấn (cấu hình được), rồi
**soạn sẵn mail** nhờ team Security mở cổng và **hiện ra cho bạn xem/sửa trước
khi gửi**. Chỉ quét 1 lần/ngày; vẫn có nút bấm tay để quét bất cứ lúc nào.

> Cần Windows + Outlook + `pywin32`. Trên môi trường khác app vẫn chạy, chỉ
> báo là tính năng quét/gửi mail không khả dụng.

## Tool: Mail chúc mừng sinh nhật 🎂

[app_qt/tools/birthday_email.py](app_qt/tools/birthday_email.py) — gửi **thiệp
sinh nhật** (ảnh xuất từ Canva Bulk Create) cho nhân viên qua Outlook. Logic ở
[app/core/birthday_mail.py](app/core/birthday_mail.py), hàng chờ nằm ở bảng
`birthday_emails`.

**LỊCH GỬI DO APP GIỮ, không nhờ Outlook.** Đây là điểm cốt lõi: trước đây cả
tháng mail được đẩy sang Outlook một lượt, mỗi mail gắn `DeferredDeliveryTime`
để nằm ở Outbox tới đúng ngày. Cách đó chỉ chạy khi tài khoản gửi **đã đăng
nhập** trong Outlook; gửi từ **hộp thư dùng chung** (chỉ được IT share, không
login được) thì Outlook bỏ qua giờ hẹn và **gửi ngay** — cả tháng nhận mail cùng
một lúc. Nên giờ chia làm ba nhịp, và **người dùng duyệt ở cả hai đầu**:

| Nhịp | Việc |
|---|---|
| **Enqueue** (người dùng bấm) | Modal *Birthdays this month* liệt kê người sinh nhật tháng này để **duyệt**; bấm *Enqueue* thì ghi vào hàng chờ. **Không gửi gì cả.** |
| **Mở app** → `prepare_due()` | Dọn hàng chờ rồi hiện modal **xác nhận** *Birthday emails to send* — liệt kê ai sắp được gửi, tick/bỏ tick từng người. **Vẫn chưa gửi.** |
| **Bấm *Send now*** → `send_rows()` | Gửi đúng những dòng đã tick, ở luồng nền. |

- **App KHÔNG BAO GIỜ tự gửi mail mà không hỏi.** Modal xác nhận chỉ hiện **một
  lần mỗi ngày** (mở app 5 lần không bị hỏi 5 lần — dấu ngày ở `config.json`
  section `birthday_email`, khóa `last_prompt`). Khác Gate-Open Mail một điểm:
  chỗ đó đóng dấu `last_scan` **trước** khi làm việc nên Outlook lỗi là mất luôn
  ngày đó; đây chỉ đóng dấu khi modal **đã thực sự hiện ra**.
- **Bấm Cancel thì không mất gì**: mọi dòng nằm nguyên trong hàng chờ. Đổi ý thì
  bấm nút **Send due emails** trên màn hình Queue — nút này *không* bị chặn bởi
  dấu ngày, nên đó là đường vào lại bất cứ lúc nào.

- **Trigger là MỞ APP, không phải mở máy.** Repo không có đăng ký autostart nào
  (Gate-Open Mail cũng vậy) — hôm nào không mở Personal Toolbox thì hôm đó không
  gửi. Bù lại bằng **hạn gửi bù** `birthday_catchup_days` (mặc định 3 ngày, sửa ở
  ⚙️ Settings): mở app muộn vẫn gửi, quá hạn thì dòng đó thành `Missed` và không
  tự gửi nữa. Muốn đúng nghĩa "mở máy là gửi" thì phải tự thêm shortcut vào
  Startup folder của Windows.
- **Không còn ô "Delivery time"**: gửi ở lần mở app đầu tiên trong ngày sinh
  nhật, không hẹn giờ trong ngày (app không chạy 24/7 nên hẹn giờ là hứa hẹn
  không giữ được). Setting `birthday_send_time` đã bỏ.
- **Bảng Queue** ngay trên trang tool: *Employee · Email · Birthday · Status ·
  Sent at · Error*, lọc theo trạng thái. Hai đường gửi tay, đừng lẫn:
  **nút *Send due emails*** chạy đúng lượt mà app hỏi lúc mở máy (tìm dòng tới
  hạn → modal xác nhận → gửi), còn **chuột phải → *Send now*** gửi đúng những
  dòng đang chọn, **bất kể** còn bao lâu tới sinh nhật (đây là đường gửi tay cho
  dòng `Missed`). Chuột phải còn có **Remove from queue**. Tháng nào chưa bấm
  *Enqueue* thì trang hiện một dòng nhắc, kẻo quên là cả tháng không ai được gửi.
- **Hàng chờ chỉ giữ khóa + trạng thái**, không chụp lại email/tên/đường dẫn
  thiệp — tất cả tra lại từ `employees` + thư mục thiệp **lúc gửi**. Nhờ vậy
  thiệp bị xóa, đổi mail công ty hay nhân viên nghỉ việc sau khi xếp hàng đều
  không sinh ra dữ liệu sai: người đã nghỉ thành `Cancelled`, thiếu thiệp/mail
  thành `Failed` (còn thử lại được). Xem
  [docs/db_design.md](docs/db_design.md#birthday_emails--hàng-chờ-mail-chúc-mừng-sinh-nhật).
- **Thư mục thiệp không truy cập được** (ổ mạng/OneDrive chưa mount lúc mới mở
  máy) thì **bỏ cả lượt, giữ nguyên mọi dòng `Pending`** và không tính vào hạn
  gửi bù — không được đánh `Failed` hàng loạt chỉ vì chưa thấy thư mục.
- **Thiệp khớp theo MÃ NV**: tên file là mã NV, nhận cả dạng Canva xuất ra
  (`1-20170456.png`) lẫn tên trần (`20170456.png`). Nút **Export CSV (missing
  cards)** xuất danh sách **toàn bộ** nhân viên chưa có thiệp (không lọc theo
  tháng) để làm một lượt cho cả năm trong Canva Bulk Create — trong đó đặt
  *"Name each page using" → code* thì tên file tải về khớp mã NV luôn.
- **Mail không có nội dung**: thiệp được nhúng inline (`cid:`) nên hiện to ngay
  khi mở mail, không phải icon file đính kèm. Chỉ có tiêu đề, đặt ở ⚙️ Settings
  (`{name}` = tên nhân viên).
- **Chống gửi trùng** ở hai lớp: unique `(employee_id, due_date)` chặn xếp hàng
  trùng, còn lúc gửi thì `claim_birthday_email()` giành dòng bằng một câu
  `UPDATE` có điều kiện — mở app hai lần cùng lúc cũng chỉ một bên gửi được.

> Cần Windows + Outlook + `pywin32`. Hộp thư dùng chung phải đã được cấp quyền
> Send As / Send on Behalf, và chọn ở ⚙️ Settings → *Send birthday emails from
> account* (gõ thẳng địa chỉ, hộp thư dùng chung không hiện trong danh sách).

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
Department · Status · Pool · Batch*, **bảng kết quả** có cột tick chọn, và
toolbar chỉ còn **nút ⋮** ở mép phải — hai việc **cấp trang** nằm trong đó:
**Add** (thêm hồ sơ nhập tay) và **Reload** (tải lại bảng). Bố cục này giống hệt
màn hình *Employees*: ô lọc rộng cố định 180px, nút *Reset* và nút **⋮** cùng
một mốc mép phải, các hàng cách đều 10px.

Mọi thao tác **trên hồ sơ** vào bằng **chuột phải trên bảng** (ngoài *Copy* /
*Copy row* dùng chung): chúng luôn cần biết chạy trên dòng nào, đặt ngay tại
dòng thì phạm vi rõ hơn nút rời trên toolbar.

Menu chia **ba nhóm** ngăn bằng đường kẻ — *sửa dữ liệu hồ sơ* · *gửi ra ngoài*
· *chỉ xem*:

| Mục chuột phải | Việc |
|----------------|------|
| **Update feedback** | nhập nhận xét phỏng vấn — **chỉ 1 dòng** (xem ① bên dưới) |
| **Update status** | đổi trạng thái hàng loạt (xem bên dưới) |
| **Update source** | điền sàn cung cấp CV (xem ② bên dưới) |
| — | |
| **Send email** | mời phỏng vấn / gửi thư cảm ơn qua Outlook (xem bên dưới) |
| **Export to Excel** | xuất ra `.xlsx` (tên file trùng thì **hỏi nối thêm hay ghi đè**) |
| — | |
| **View details** | mở modal xem chi tiết từng hồ sơ |

**Phạm vi thao tác** — luật chung `DataTable._target_rows`: chuột phải vào dòng
**chưa tick** thì chạy đúng dòng đó (không cần tick gì); vào dòng **đang tick**
thì chạy **cả nhóm tick**. Nhãn menu luôn kèm số dòng (*Update status (3 rows)*)
để thấy phạm vi TRƯỚC khi bấm; mục chỉ làm được trên 1 dòng thì để mờ kèm lý do
khi đang trỏ nhiều dòng.

**① Update feedback** — chỉ làm trên **1 dòng**. Đây là chỗ **nhập nhận xét
phỏng vấn**: form sửa hồ sơ đầy đủ có ~30 ô
(thông tin cá nhân · hồ sơ nghề nghiệp · nguyện vọng · đơn ứng tuyển) nên quá dài
cho việc làm thường xuyên nhất sau mỗi buổi phỏng vấn. Hộp này chỉ giữ **đúng các
ô có trong file Excel xuất ra**:

- **Application** — *Applying for* (**chỉ để xem** — đổi vị trí là đổi cả luồng
  tuyển dụng: JD, mẫu mail, các vòng đã có, nên chỉ làm ở form đầy đủ) · Status ·
  Result · Phone screen date & note
- **1st / 2nd / 3rd interview** — Date · Final result, kèm danh sách
  **Interviewer feedback**: bấm *Add interviewer* để thêm bao nhiêu người tuỳ ý,
  mỗi người một dòng *(tên · kết luận · nhận xét)*, bấm ✕ để bỏ. Mỗi dòng ứng với
  một bản ghi `interview_feedbacks` — đúng thứ mà cột *INTERVIEW EVALUATION* của
  file Excel ghép lại, nên nhập tách từng người ở đây thì lúc xuất mới khớp đúng
  tên với nhận xét tương ứng.

Ô người phỏng vấn là **danh sách nhân viên** (`repo.list_interviewers()`) và lưu
luôn `employee_id` — chức danh & phòng ban tra thẳng từ `employees`, không hỏi
lại trong form. Tên đã lưu mà nay không còn trong danh sách (khách mời, người đã
nghỉ) vẫn hiện đúng, chỉ không nối được vào `employees`.

Mọi ô ngày dùng **`widgets.DateEdit`** — lịch bung ra, hiển thị `dd/MM/yyyy`, lưu
`yyyy-mm-dd`. Ô **để trống được** (vòng chưa diễn ra thì không điền sẵn ngày vô
nghĩa): nhấn `Delete`/`Backspace` để xoá trắng. Lăn chuột không đổi ngày, để cuộn
modal không vô tình sửa dữ liệu.

> Nhận xét chỉ tồn tại ở **cấp từng người phỏng vấn**. Hai cột cũ ở cấp buổi —
> `interviews.note` (HR note) và `interviews.summary` (Panel summary) — đã bỏ ở
> lượt migration `0002_drop_interview_notes`.

Vòng nào để trống hoàn toàn thì **không tạo dòng rỗng** trong DB. Bỏ một người ra
khỏi form là xoá luôn nhận xét của họ (`repo.save_interview_feedbacks()` đồng bộ
hai chiều). Các cột do máy sinh — *Batch* · *ID* (bóc từ tên file CV) · *Score* &
*AI Evaluation* (AI chấm) — **không cho sửa tay**, sửa là hỏng lịch sử đánh giá.

**② Update source** — mở popup nhỏ, chọn **sàn cung cấp CV** rồi ghi một lượt
cho **mọi hồ sơ trong phạm vi**. Ô chọn **gõ tay được** nên sàn mới chưa có trong
`cv_schema.CANDIDATE_SOURCE_CHOICES` vẫn điền thẳng; các hồ sơ đang cùng một
nguồn thì popup điền sẵn nguồn đó.

> **Vì sao phải điền tay**: tool *AI CV Scan* quét cả thư mục nên **không thể
> biết từng file CV lấy từ sàn nào** — nó chỉ đóng dấu
> `cv_schema.CANDIDATE_SOURCE_AUTO` (`"AI CV Scan"`, nghĩa là *hồ sơ vào app bằng
> đường quét tự động*). Cột **Source** của bảng hiện **trống** ở những dòng còn
> mang dấu đó, nên nhìn là biết ngay hồ sơ nào chưa gán sàn.
>
> Cột `source` có ở **ba bảng** — `candidates` (biết đến từ đâu) · `applications`
> (đơn này từ đâu) · `candidate_cvs` (bản CV này lấy ở đâu) — nên
> `repo.set_candidate_source()` ghi **cả ba** trong một lượt (đơn đang hiển thị
> trên dòng + bản CV `latest_cv_id`), tránh mỗi màn hình đọc ra một giá trị khác.

*Add* nằm trong **nút ⋮** cạnh *Reload*: hồ sơ giờ chủ yếu vào DB qua tool *Quét
CV bằng AI*, nhập tay chỉ còn là trường hợp lẻ nên không đáng chiếm một nút
riêng. **Sửa** một hồ sơ: **double-click vào dòng**; **xóa**: mở form sửa rồi
bấm *Delete* trong đó — cả hai không có nút riêng trên toolbar.

- **Chống trùng**: khi thêm mới (hoặc sửa) ứng viên, nếu **trùng email hoặc SĐT**
  với người đã có, tool cảnh báo và cho quyết định vẫn lưu hay không.
### Xuất Excel (chuột phải → *Export to Excel*)

Sheet **Candidates** được **dựng thẳng bằng code**
([app/core/candidate_export.py](app/core/candidate_export.py)), **không đọc file
`.xlsx` mẫu nào** — trước đây phải mở file mẫu mang theo ~90 liên kết ngoài với
15 MB XML cache, riêng bước mở đã mất ~9 giây; giờ xuất 166 hồ sơ hết dưới 1
giây và file kết quả chỉ ~56 KB. Mọi ô là **chữ / số / ngày thuần** — không công
thức, không liên kết ngoài.

25 cột `A→Y`, lấy dữ liệu từ **đơn ứng tuyển mới nhất** của mỗi ứng viên:

| Cột | Nguồn |
|-----|-------|
| `A` Batch · `B` ID · `C` NAME | `candidate_cvs.batch` · mã bóc từ tên file CV (không có thì `candidate_id`) · `full_name` |
| `D` APPLYING FOR · `E` SOURCE | `positions.position_title` (chưa gắn vị trí → tên bộ phận) · **nơi cung cấp CV** — nguồn của đơn, thiếu thì nguồn của ứng viên |
| `F` EMAIL · `G` PHONE | `email` · `phone` (định dạng text, giữ số 0 đầu) |
| `H` Score · `I` AI Evaluation | `ai_score` · **tóm tắt độ phù hợp** của lượt AI chấm mới nhất (chỉ `summary`, không kèm Strengths/Weaknesses) |
| `J` STATUS · `K` Results | `applications.status` · `applications.final_status` |
| `L` `M` PHONE SCREEN | `phone_screen_date` · ghi chú của HR trên đơn |
| `N…Y` — 3 vòng phỏng vấn | mỗi vòng 4 cột: **ngày · người phỏng vấn · nhận xét · kết luận** |

**Nhiều người phỏng vấn trong một vòng thì gộp vào MỘT ô**: cột *ASSIGNED
INTERVIEWER* nối tên bằng dấu phẩy, cột *INTERVIEW EVALUATION* xếp mỗi người một
đoạn `Tên (vai trò · kết luận): nhận xét` (kèm Strengths/Weaknesses nếu có), cách
nhau một dòng trống, phía trên là *HR note* và *Panel summary* của buổi đó.

> **SOURCE là SÀN cung cấp CV** (Itviec · VietnamWorks · LinkedIn · TopCV ·
> Referral · headhunt…), không phải cách hồ sơ vào app. Hồ sơ do tool *AI CV
> Scan* nạp mang dấu `cv_schema.CANDIDATE_SOURCE_AUTO` (`"AI CV Scan"`) ở cột
> `source` — quét cả thư mục thì chưa biết CV lấy từ sàn nào — nên khi xuất
> Excel ô này **để trống** cho HR chọn lại từ dropdown. Thêm sàn mới: sửa
> `cv_schema.CANDIDATE_SOURCE_CHOICES`.

Sheet còn có sẵn **ô chọn** cho các cột APPLYING FOR / SOURCE / STATUS / Results
/ người phỏng vấn / kết luận từng vòng (danh sách nằm ở sheet ẩn `_Lists`, lấy từ
DB và `cv_schema`; vẫn gõ tay được giá trị ngoài danh sách), **tô màu** trạng
thái & kết luận, và **bôi đỏ email trùng**.

**Trùng tên file thì được hỏi**, không âm thầm nối thêm nữa:

| Nút | Việc |
|-----|------|
| **Append** | ghi tiếp phía dưới các dòng đã có; vùng ô chọn/tô màu tự nới theo |
| **Overwrite** | thay hẳn file bằng đúng các hồ sơ vừa tick — **mất sạch nội dung cũ** |
| **Cancel** | không đụng tới file |

Hộp thoại của Windows chỉ hỏi được có/không nên tool truyền
`DontConfirmOverwrite` rồi tự hỏi bằng `dialogs.choose()` — helper nhiều lựa chọn
dùng chung, đóng bằng ✕/Esc coi như Cancel.

- **Master data** tách thành nhóm **Master Data** riêng ở sidebar, gồm 9 trang:
  **Departments · Employee types · Levels · Cost centers · Code lists ·
  Positions · Mail templates · Skills · Courses** — mỗi trang là một bảng +
  thanh *Add / Edit / Delete / Reload* (dùng chung `CrudTablePanel`),
  thêm/sửa/xóa riêng. (Các trang này không hiện thẻ ở Trang chủ.) Các trang
  danh mục nhân sự đã có **dữ liệu nạp sẵn** từ sheet *Code* của file HC:
  20 bộ phận · 33 nhóm chức năng · 6 loại nhân viên · 12 cấp bậc ·
  42 cost center · 19 giá trị code list.
  Muốn thêm/bớt cột hay ô nhập của một trang → sửa `_master_specs()` trong
  [app_qt/tools/candidate_db.py](app_qt/tools/candidate_db.py).
- **Nhóm chức năng (`functions`) sửa ngay trong form Departments**, không có
  trang riêng: bảng có cột *Functions* liệt kê các nhóm của phòng ban (ngăn bởi
  dấu phẩy), form có ô nhập nhiều dòng — thêm tên là tạo, bỏ tên là xóa, tên
  giữ nguyên thì giữ nguyên bản ghi (`repo.replace_department_functions`).
- **Trang Code lists** gộp NHIỀU danh mục "một cột giá trị" vào một bảng
  (`code_lists`), phân biệt bằng `type`; ô chọn phía trên quyết định đang
  xem/sửa danh mục nào, và dòng thêm mới được điền sẵn type đó. Hiện có
  *Qualification · Qualification (VN) · ID issued place*
  (`cv_schema.CODE_LIST_TYPES`) — thêm danh mục một cột về sau chỉ cần thêm một
  `type`, không phải thêm bảng và trang. Ô chọn này là tính năng chung của
  `CrudTablePanel` (khóa `"filter"` trong spec).
- **Mẫu mail dùng chung**: trang **Mail templates** (bảng `mail_templates`) giữ
  mọi mẫu mail — *tên · loại · CC · tiêu đề · nội dung (rich text)*. Loại lấy từ
  `MAIL_TEMPLATE_TYPE_CHOICES` (Interview Round 1/2/3 · Application Thank You ·
  Notification · Offer · Rejection) và chủ yếu để phân nhóm cho dễ tìm, không
  ràng buộc — **trừ `Application Thank You`**: mục *Send email* lọc đúng loại này
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
- **Update status** (đổi trạng thái hàng loạt): các hồ sơ trong phạm vi phải
  đang ở **CÙNG một trạng thái** → modal hiện trạng thái hiện tại và ô
  *Move to* điền sẵn **bước kế tiếp** trong luồng (sửa được, chọn bất kỳ nhãn
  nào trong `CANDIDATE_STATUS_CHOICES`); bấm *OK* mới ghi xuống DB. Nếu các hồ
  sơ đang ở trạng thái khác nhau thì app **báo lỗi kèm danh sách từng nhóm** và
  không đổi gì.
- **Send email**: các hồ sơ trong phạm vi cũng phải đang ở **CÙNG một trạng
  thái** (lệch nhau thì báo lỗi kèm danh sách từng nhóm, giống *Update status*)
  → hộp thoại chọn **loại mail muốn gửi**: 3 **vòng phỏng vấn**
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

**Bảng `employees` bám sát file Excel "Personnel Data"** (HR export, sheet
*Personnel Data-new*) — gần
như mỗi cột trong file có một cột tương ứng trong DB; chú thích
`-- <Tiêu đề Excel>` ghi ngay cạnh từng cột trong `app/core/cv_schema.py`.

- **Import khớp theo TÊN cột, KHÔNG theo thứ tự cột** (`_EXCEL_HEADER_MAP` ở
  [app_qt/tools/employee_db.py](app_qt/tools/employee_db.py)): tiêu đề được
  chuẩn hóa (chữ thường · gộp khoảng trắng · bỏ ký hiệu font Wingdings · NFC)
  nên đổi chỗ cột hay thêm cột lạ đều không làm vỡ import. Cột **trùng tên**
  ("Issued date", "Changing date") khai báo bằng *tuple* → lần xuất hiện thứ n
  lấy field thứ n.
- **Ô "Emergency Contact Name" tự tách làm 2 cột KHI IMPORT**: file gốc gộp cả
  tên lẫn số điện thoại trong một ô, đủ kiểu ngăn cách (`tên ⏎ số`, `tên (số)`,
  `số - tên`) → tách sang `emergency_contact_name` (chỉ họ tên) và
  `emergency_contact_phone` bằng `cv_repository.split_contact_name_phone`
  (không tìm thấy số thì giữ nguyên cả ô làm tên). File nào đã có sẵn cột SĐT
  riêng thì lấy thẳng theo tên cột, không tách nữa. Cả hai cột đều có mặt ở
  bảng (nhóm *Contact* trong modal **Columns**) và ở form Add/Edit nhân viên.
  Việc tách **chỉ chạy lúc import** — dữ liệu đã nằm trong DB không bị đụng tới.
- **Trạng thái làm việc suy ra từ `termination_date`** (không có cột `status`):
  có ngày (khác rỗng) = đã nghỉ việc. Truy vấn danh sách trả thêm cột
  `work_status` ("Working"/"Resigned") để hiển thị; mặc định chỉ hiện người đang
  làm việc (tick *Only resigned employees* để đảo sang danh sách người đã nghỉ —
  hai danh sách tách riêng, không có chế độ xem gộp).
- **THỨ TỰ CỘT PHẢI BÁM ĐÚNG FILE EXCEL.** `_EMP_COLUMN_SPECS`
  (`app_qt/tools/employee_db.py`) xếp theo đúng thứ tự cột của file "Personnel
  Data"; thêm/đổi cột thì chèn vào **đúng vị trí của cột đó trong file**, không
  nối vào cuối. Ba nơi bám theo danh sách này: thứ tự cột trên bảng, nhóm cột
  trong modal **Columns** (`_EMP_COLUMN_SECTIONS` cắt nhóm ngay từ danh sách
  này nên đọc dọc modal luôn ra đúng thứ tự bảng) và **Copy row**
  (`_EMP_COPY_COLUMNS` — bố cục cột của file Excel đích).
- **4 cột text được tra sang bảng danh mục** (lưu **id**): "Department (short)"
  → `departments` (short_name) · "New Cost center" → `cost_centers` (code) ·
  "IBC/DBC/WC" → `employee_types` (code) · "Job level" → `levels` (level_name).
  Xem `_MASTER_LOOKUPS`.
- **4 cột text khớp danh mục nhưng lưu NGUYÊN VĂN** (`_TEXT_LOOKUPS`):
  "Function (Common)" → `functions` · "Qualification" · "Qualification
  (Việt Nam)" · "Issued Place" → `code_lists` theo `type` tương ứng. Danh mục ở
  đây chỉ đóng vai **người gác cổng** để mọi nhân viên viết giống nhau, giá trị
  vẫn lưu dạng text. Danh mục rỗng thì bỏ qua kiểm tra.
- Cả ba nhóm (thêm `_CHOICE_FIELDS` — hằng cứng trong code như Gender, Marital
  status) đều **chặn import** nếu có ô không khớp, báo lại từng ô sai kèm số
  dòng Excel và KHÔNG ghi gì cả. So khớp **không phân biệt hoa/thường**, nhưng
  giá trị lưu xuống lấy đúng cách viết của danh mục.
- **Copy row** (chuột phải → *Copy row*) trải dòng thành **82 ô TSV** đúng bố
  cục cột A → CD của file Excel (*Legal Entity* → *Birthday*, cột hợp lệ cuối
  cùng), chừa ô trống cho cột file có mà DB không lưu → dán thẳng vào file, bắt
  đầu từ cột A của một dòng trống (thao tác này dùng khi THÊM nhân viên mới).
  Cột "Department (short)" dán **mã viết tắt** (`department_short_name`), không
  phải tên bộ phận đầy đủ. Mặc định dán **text** (kể cả ô mà file vốn để công
  thức — Surname/Name/Middle Name, Department (short), Full name of manager:
  dữ liệu trong DB đã là kết quả cuối). Chỉ dán **công thức** ở ô mà giá trị
  phụ thuộc ngày hôm nay (4 cột *Figures*) hoặc app không có dữ liệu (*Count*,
  *Birthday*); công thức viết bằng tham chiếu có cấu trúc `[@[Tên cột]]` nên dán
  dòng nào cũng đúng. Ô *Emergency Contact Name* dán **tên ⏎ số ĐT gộp lại** —
  file chỉ có một ô cho cả hai, app tách ra hai cột lúc import nên copy phải nối
  ngược lại (`_emergency_contact_cell`).
- Các cột **file Excel tự tính** (Age, Age range, Year of service, Length of
  service) **không lưu trong DB** — app tính lại mỗi lần truy vấn bằng đúng
  công thức của file (`cv_repository.EMPLOYEE_COMPUTED_SQL`) nên không bao giờ
  cũ. Cột *Year of birthday (year)* **bỏ khỏi cả app lẫn file Excel** (công
  thức trùng hệt cột Age). Các cột "Legal Entity (Company)", "Position status", "Business Unit
  (Department)", "BC/WC", "STT", "Birthday"… không lưu vì đã có nguồn khác hoặc
  chỉ là cột phụ trợ trong file (xem chú thích trong schema).
- **Thanh nút** chỉ để lộ việc làm thường xuyên — *Import application form* (nhập
  đơn dự tuyển) và *Sync with Excel* (đồng bộ với file HC, xem mục riêng bên
  dưới). Ba việc thi thoảng mới dùng — **Add · Bulk Import · Reload** —
  nằm trong nút **⋮** ở góc phải (`_build_more_button` / `_show_more_menu`).
  *Bulk Import* chính là nhập hàng loạt từ file "Personnel Data" nói ở trên.
  *Enroll to course* (ghi danh khóa học) vào bằng **chuột phải trên bảng**
  (`menu_actions` của `DataTable`): thao tác luôn gắn với dòng nên phạm vi là
  dòng chuột phải, hoặc cả nhóm tick nếu dòng đó đang được tick.
- Bảng ~90 cột → giao diện chỉ hiện vài cột mặc định, bật/tắt thêm ở modal
  **Columns** — cột gom theo nhóm (Identity / Personal info / Organization…),
  mỗi nhóm có checkbox bật/tắt cả nhóm, nút *Clear all* bỏ hết chỉ chừa Emp code
  + Full name; lựa chọn được ghi nhớ. Sửa nhóm & cột mặc định ở
  `_EMP_COLUMN_SECTIONS` / `_EMP_DEFAULT_COLUMNS` trong
  `app_qt/tools/employee_db.py`.

### Đồng bộ với file Excel của HR 🔄

Nút **Sync with Excel** (màn hình *Employees*) đối chiếu bảng `employees` với
chính file **"Personnel Data"** mà HR đang giữ, rồi **hỏi lại ở hai bước** — app
**không ghi gì xuống DB trước khi người dùng bấm xác nhận**. Đường dẫn file đặt
một lần ở ⚙️ **Settings → Employee data** (khóa `hc_excel_path`); logic so sánh
& ghi ở [app/core/employee_sync.py](app/core/employee_sync.py), bố cục cột lấy
thẳng từ `_EXCEL_HEADER_MAP` nên thêm cột vào map là lượt đồng bộ tự so thêm cột
đó.

Dùng lại **đúng bộ đọc file của *Bulk Import*** (`_read_excel`): tự dò **sheet
dữ liệu** (ưu tiên sheet có tên chứa "personnel data" — cũng là sheet đầu tiên
của file HC) và **dòng tiêu đề**, khớp **theo tên cột**. Khóa khớp hai bên là
**mã nhân viên** (`code`), và lượt đồng bộ nhìn **cả người đã nghỉ**
(`repo.list_all_employees` — ngoại lệ duy nhất của global scope) để họ không bị
coi là "mã lạ" mỗi lần chạy.

| Bước | Modal | Việc |
|---|---|---|
| 1 | *Sync from Excel · N changes* | Mỗi dòng là **MỘT Ô lệch**: mã NV · tên · tên cột · giá trị trong app · giá trị trong file. **Tick sẵn tất cả** (bấm ô tick ở header để tick/bỏ cả bảng); bỏ tick dòng nào thì ô đó **giữ nguyên** giá trị trong app. Bấm *Apply ticked changes* mới ghi — mỗi nhân viên gom thành một câu `UPDATE`. |
| 2 | *Employees missing from the Excel file · N* | Người **đang làm việc** trong app mà file không còn dòng nào mang mã của họ. Mỗi người một dòng kèm **ô chọn ngày nghỉ việc** + **ô lý do**; *Save resignations* ghi `termination_date` + `leaving_reason` (có ngày = đã nghỉ việc). |

- **Chia theo Ô, không theo NGƯỜI**, ở modal 1: cùng một người có thể đúng ở cột
  này mà sai ở cột kia; gộp cả người thành một dòng thì muốn giữ lại một ô là
  phải bỏ luôn những ô khác.
- **Ô TRỐNG LÀ MỘT GIÁ TRỊ — nhưng chỉ khi FILE CÓ CỘT ĐÓ.** Bố cục file cố
  định nên dòng header nói đủ: `_read_excel` trả thêm **tập field mà file có
  cột** (kể cả cột trống trơn), nhờ đó phân biệt được hai chuyện khác hẳn nhau
  — *cột file không quản* (giữ nguyên dữ liệu app, vd cột chỉ có trong đơn dự
  tuyển) và *ô đã bị xóa trong file* (đề nghị xóa luôn bên app). Dòng xóa hiện
  giá trị `(blank)` ở cột *In Excel*, đếm riêng thành một dòng cảnh báo trong
  modal, và tick sẵn như mọi dòng khác — bỏ tick nếu muốn giữ giá trị trong
  app. Cột danh mục bị xóa thì ghi **NULL** (gỡ liên kết), không phải chuỗi
  rỗng. Ô "Emergency Contact Name" của file gộp tên + SĐT nên có cột đó là file
  quản **cả hai** field app tách ra.
- **So sánh KHOAN DUNG** (`employee_sync.same_value`): bỏ qua khác biệt về
  khoảng trắng, hoa/thường, **cách viết ngày** (`10/09/1993` ≡ `1993-09-10`) và
  **cách viết số** (`0` ≡ `"0.0"`), Unicode chuẩn hóa NFC trước khi so (tên
  tiếng Việt trong file có thể ở dạng tổ hợp). Hai bên cùng một dữ liệu mà hiện
  lên thành "thay đổi" thì danh sách duyệt đầy nhiễu, người dùng sẽ tick bừa.
- **Điền ngày CHÍNH LÀ cách chọn** ở modal 2 (không có ô tick): để trống = chưa
  kết luận gì, người đó vẫn "đang làm việc" như cũ — vắng mặt trong file còn có
  thể vì HR chưa cập nhật hay người đó mới vào. Người **chưa có mã NV** không
  khớp được theo mã nên không bị liệt vào đây, chỉ báo lại số lượng.
- **Không tự thêm người mới**: mã có trong file mà app chưa có được **liệt kê
  lại** kèm lời nhắc dùng *Bulk Import* (nút ⋮) — chỗ đó có sẵn bước kiểm tra
  danh mục đầy đủ và tự bỏ qua mã đã tồn tại, nên nhập cả file lần nữa chỉ thêm
  đúng những người còn thiếu.
- **Ô cột danh mục không tra được** (bộ phận/cost center/loại NV/cấp bậc ghi sai
  chính tả) thì **bỏ riêng ô đó** và báo lại, không chặn cả lượt như *Bulk
  Import*: ở đây mỗi ô là một thay đổi độc lập trên người đã có sẵn, bỏ ô sai
  vẫn ghi đúng được các ô còn lại.
- **Hủy modal 1 KHÔNG bỏ luôn modal 2**: "không ghi đè mấy ô này" và "ai đã nghỉ
  việc" là hai quyết định rời nhau. Các cảnh báo (mã mới · ô không tra được · mã
  trùng · dòng thiếu mã · cột lạ) hiện ở modal 1, modal đó không hiện thì dồn
  xuống hộp tổng kết.
- Đọc file (vài MB) + dựng danh sách lệch chạy ở **luồng nền** kèm ProgressDialog
  — làm thẳng ở luồng giao diện thì app đứng hình vài giây, nhìn như treo.
- Danh sách "vắng mặt trong file" dài bất thường (>150 người) thì **hỏi lại
  trước khi mở**: cả công ty biến mất khỏi file gần như luôn là chọn nhầm file,
  và mỗi dòng một ô chọn ngày nên danh sách dài cũng mất vài giây mới dựng xong.

### Nhập nhân viên từ ĐƠN DỰ TUYỂN (AI đọc form) 📄

Nút **Import application form** (màn hình *Employees*) đọc file
**DLVN Application Form** do nhân viên mới tự điền rồi đổ vào form nhập liệu để
HR duyệt trước khi ghi vào bảng `employees`. Chọn được **nhiều đơn một lượt**;
logic ở [app/core/application_form.py](app/core/application_form.py).

| Nhân viên nộp kiểu gì | App gửi gì cho AI |
|---|---|
| Điền thẳng vào file mẫu (`.xlsx`/`.xlsm`) — phần họ điền là **chữ màu** | Dựng lưới ô của mọi sheet thành **text** kèm tọa độ ô, giá trị chữ màu bọc trong `«…»` (Gemini không đọc được file .xlsx) |
| In ra viết tay, HR **scan thành PDF** nhiều trang | **Rasterize từng trang thành PNG** (150 dpi) rồi gửi hết trong một request |

- **Luôn có bước người duyệt**: mỗi đơn đọc xong mở một form *Review employee
  n/N · tên file* điền sẵn — AI đọc chữ viết tay không phải lúc nào cũng đúng.
  Bấm *Save* mới ghi xuống DB, *Cancel* là bỏ qua đơn đó.
- **Cảnh báo trùng người** trước khi ghi: khớp **mã NV · số CMND/CCCD · họ tên**
  (`repo.find_employees_by_identity`) → hiện danh sách người đã có, vẫn cho ghi
  nếu HR xác nhận.
- **Đơn có nhiều thông tin hơn DB** (kinh nghiệm làm việc, sức khỏe, lương mong
  đợi, người tham khảo…) → **chỉ lấy phần có cột trong `employees`**, khai ở
  `_FIELDS` trong `application_form.py` (thêm/bớt field chỉ cần sửa dict này) —
  gồm cả **SĐT người báo tin khẩn cấp** (`emergency_contact_phone`, đọc ở ô
  *Tel.* mục *In case of emergency* trang 3).
  Vài chỗ suy ra thêm: họ tên tách sang `surname`/`middle_name`/`name` theo thứ
  tự tiếng Việt, `marriage_status` (Y/N) suy từ tình trạng hôn nhân, nhiều số
  điện thoại gộp bằng `"; "`, ngày chuẩn hóa về `dd/mm/yyyy`.
- Ô chọn (Gender, Marital status) chỉ nhận đúng nhãn trong danh mục — giá trị lạ
  (vd *Separated*) để **trống** cho HR tự chọn, tránh ghi nhãn lạ vào DB.
- Dùng chung **API key + model ở ⚙️ Settings** như các tính năng AI khác; không
  cần cài thêm gói (đã có `openpyxl` + `pymupdf`).

**Bảng danh mục (master data)** — nạp sẵn dữ liệu từ sheet *Code* của file HC:
`departments` (tên + mã viết tắt), `functions` (nhóm chức năng, gắn vào phòng
ban), `employee_types` (WC/WCA/IBC/IBCA/DBC/DBCA + nhóm Blue/White Collar),
`cost_centers` (mã VN1001… + Group Function VNPlant/Corporate/R&D — dùng để gom
nhân viên theo nhóm khi tính chi phí vận hành), `levels` (Director, Manager,
Officer…), `code_lists` (Qualification · Qualification (VN) · ID issued place).

> **Dữ liệu khởi tạo**: khai báo ở `SEED_DATA` (cv_schema.py), nạp bởi
> `_seed_master_data()` — mỗi khối chỉ chạy **một lần** cho mỗi file .db (đánh
> dấu ở bảng `app_meta`), và bỏ qua dòng đã tồn tại nên không tạo bản ghi trùng.
> Muốn nạp lại: xóa khóa `seed:<bảng>:v1` trong `app_meta`. Bổ sung danh mục về
> sau: thêm dòng vào `rows` rồi tăng `version`. Khối có khóa `lookup` (vd
> `functions`) thì ô khóa ngoại ghi bằng **tên**, seeder tra sang id lúc nạp —
> bảng được tra phải khai **trước** trong `SEED_DATA`.

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
