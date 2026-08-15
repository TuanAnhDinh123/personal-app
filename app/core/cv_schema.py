"""THIẾT KẾ CƠ SỞ DỮ LIỆU — Quản lý tuyển dụng / CV ứng viên (SQLite).

⭐ ĐÂY LÀ FILE ĐỂ BẠN THIẾT KẾ / CHỈNH SỬA CẤU TRÚC DB ⭐

Toàn bộ bảng, cột được mô tả bằng SQL trong hằng `SCHEMA_SQL` bên dưới. Sửa
cấu trúc DB ở đây rồi mở lại tool — `cv_repository.init_db()` tự chạy lại các
câu `CREATE TABLE IF NOT EXISTS` để tạo phần còn thiếu.

QUY ƯỚC THIẾT KẾ (theo yêu cầu):
  • KHÔNG dùng ràng buộc khóa ngoại (FOREIGN KEY). Các cột *_id chỉ là số
    tham chiếu "mềm" — ứng dụng tự đảm bảo liên kết, DB không ép buộc.
  • Mọi cột đều CHO PHÉP NULL, trừ khóa chính (PK) tự tăng.
  • Kiểu dữ liệu ghi theo sơ đồ (VARCHAR/INT/TEXT/DATE/DATETIME/DECIMAL).
    SQLite dùng "type affinity" nên chấp nhận các tên kiểu này bình thường.

────────────────────────────────────────────────────────────────────────────
SƠ ĐỒ QUAN HỆ (mềm, không ràng buộc FK)

  departments ──1:N──> positions (+ jd_file_path)
                          │
                          │ 1:N
                          ▼
                      candidates (+ cv_file_path)

  mail_templates ──1:N──> positions  (qua 3 cột mail_template_r1/r2/r3_id)

  • departments (phòng ban)  1—N  positions (vị trí)
  • positions   (vị trí)     1—N  candidates
  • MỖI VỊ TRÍ CHỈ CÓ 1 JD → đường dẫn file JD nằm THẲNG trong bảng positions
    (tiêu đề JD = luôn dùng tên vị trí); KHÔNG còn bảng job_descriptions riêng.
  • Đường dẫn file lưu thẳng: candidates.cv_file_path, positions.jd_file_path
    (không còn bảng document_files — file thực tế đã nằm sẵn trên máy).
  • MẪU MAIL nằm ở bảng dùng chung `mail_templates` (không còn nhúng trong
    positions). Luồng tuyển dụng có 3 VÒNG phỏng vấn nên mỗi vị trí trỏ tới 3
    mẫu khác nhau — xem INTERVIEW_ROUNDS.

  departments ──1:N──> employees
  levels · cost_centers · employee_types ──1:N──> employees
  • employees soi theo file "Master HC file.xlsx" (gần như 1 cột Excel = 1 cột
    DB, khớp THEO TÊN CỘT khi import). Trạng thái làm việc suy ra từ
    `termination_date`: có ngày = đã nghỉ việc.
  • 4 cột text trong file được tra sang id của bảng danh mục:
    "Function (Common)" → department_id · "New Cost center" → cost_center_id ·
    "IBC/DBC/WC" → employee_type_id · "Job level" → level_id.
  courses  ──N:M──  employees   (qua bảng trung gian course_employees)
  • Một khóa học (courses) có NHIỀU nhân viên; một nhân viên tham gia NHIỀU
    khóa học. Mỗi dòng trong course_employees = 1 lượt ghi danh (course_id +
    employee_id, unique để không ghi danh trùng).

  DANH MỤC DÙNG CHUNG (master data — nạp sẵn từ file Code.xlsx):
  • departments    (cột B + C)  — tên bộ phận + mã viết tắt
  • employee_types (cột D + E)  — loại nhân viên (WC/WCA/IBC/…) + nhóm Blue/White Collar
  • cost_centers   (cột F + G)  — mã trung tâm chi phí (VN1001…) + Group Function;
    gán cho nhân viên để GOM NHÓM tính chi phí vận hành của từng team
  • levels         (cột H)      — cấp bậc (Director, Manager, Officer…)
  Dữ liệu khởi tạo nằm trong `SEED_DATA` bên dưới, được
  `cv_repository._seed_master_data()` nạp MỘT LẦN cho mỗi DB (xem bảng app_meta).
────────────────────────────────────────────────────────────────────────────

LƯU Ý khi sửa về sau:
  • `CREATE TABLE IF NOT EXISTS` chỉ tạo bảng khi CHƯA có; KHÔNG tự thêm cột
    vào bảng đã tồn tại. Muốn thêm cột cho bảng đang có dữ liệu → dùng
    `MIGRATIONS` (ALTER TABLE, chạy an toàn nhiều lần).
  • Đừng đổi tên/xóa cột đang có dữ liệu nếu chưa sao lưu file .db.
"""

SCHEMA_SQL = """
-- ───────────────────────── MASTER: PHÒNG BAN ─────────────────────────
CREATE TABLE IF NOT EXISTS departments (
    department_id   INTEGER PRIMARY KEY AUTOINCREMENT,
    department_name VARCHAR,
    short_name      VARCHAR,                 -- mã viết tắt (vd FIN, IT, R&D) — khớp cột
                                              -- "Function (Common)" khi import nhân viên
    manager_name    VARCHAR,
    description     TEXT,
    created_at      DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at      DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- ─────────────────── MASTER: LOẠI NHÂN VIÊN (Employee type) ──────────
-- Mã loại nhân viên dùng khi khai báo hồ sơ: WC / WCA / IBC / IBCA / DBC / DBCA.
-- `collar` = nhóm lao động (Blue Collar / White Collar) — xem COLLAR_CHOICES.
CREATE TABLE IF NOT EXISTS employee_types (
    employee_type_id INTEGER PRIMARY KEY AUTOINCREMENT,
    code             VARCHAR,             -- mã loại (WC, WCA, IBC…)
    collar           VARCHAR,             -- Blue Collar / White Collar
    description      TEXT,
    created_at       DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at       DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- ──────────────── MASTER: TRUNG TÂM CHI PHÍ (Cost center) ─────────────
-- Gán cho nhân viên để gom thành từng nhóm khi tính chi phí vận hành của team.
-- `group_function` = nhóm chức năng cấp trên (VNPlant / Corporate / R&D).
CREATE TABLE IF NOT EXISTS cost_centers (
    cost_center_id   INTEGER PRIMARY KEY AUTOINCREMENT,
    code             VARCHAR,             -- mã cost center (VN1001, VN3000…)
    group_function   VARCHAR,             -- VNPlant / Corporate / R&D
    name             VARCHAR,             -- tên gọi (tùy chọn, tự đặt thêm)
    description      TEXT,
    created_at       DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at       DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- ───────────────────────── MASTER: CẤP BẬC (Level) ────────────────────
CREATE TABLE IF NOT EXISTS levels (
    level_id    INTEGER PRIMARY KEY AUTOINCREMENT,
    level_name  VARCHAR,                  -- Director, Manager, Officer…
    sort_order  INT,                      -- thứ tự hiển thị (nhỏ → trước)
    description TEXT,
    created_at  DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at  DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- ─────────────── KEY-VALUE nội bộ (đánh dấu đã nạp seed, v.v.) ────────
CREATE TABLE IF NOT EXISTS app_meta (
    key   VARCHAR PRIMARY KEY,
    value VARCHAR
);

-- ───────────────────── MASTER: MẪU MAIL (mail_templates) ──────────────
-- Kho mẫu mail dùng chung cho cả tuyển dụng: thư mời phỏng vấn từng vòng, thư
-- cảm ơn đã ứng tuyển, thư thông báo… Vị trí tuyển dụng chỉ TRỎ tới mẫu qua id.
-- Nội dung hỗ trợ placeholder {name} {possion} {position} {date} {time_start}
-- {time_end} — điền lúc gửi (xem _fill_template trong app_qt/tools/candidate_db.py).
CREATE TABLE IF NOT EXISTS mail_templates (
    mail_template_id INTEGER PRIMARY KEY AUTOINCREMENT,
    name             VARCHAR,           -- tên mẫu (để nhận diện khi chọn)
    type             VARCHAR,           -- loại mẫu (xem MAIL_TEMPLATE_TYPE_CHOICES)
    mail_cc          VARCHAR,           -- CC mặc định (nhiều email ngăn bởi ";")
    mail_subject     VARCHAR,           -- tiêu đề mẫu mail
    mail_body        TEXT,              -- nội dung mẫu mail (HTML)
    created_at       DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at       DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- ────────── MASTER: VỊ TRÍ TUYỂN DỤNG (kèm luôn JD của vị trí đó) ──────
-- Mỗi vị trí chỉ có ĐÚNG 1 mô tả công việc (JD) → cột jd_file_path nằm ngay đây.
-- 3 cột mail_template_r*_id = mẫu mail dùng cho 3 vòng phỏng vấn (xem
-- INTERVIEW_ROUNDS); tham chiếu mềm → mail_templates.mail_template_id.
CREATE TABLE IF NOT EXISTS positions (
    position_id        INTEGER PRIMARY KEY AUTOINCREMENT,
    department_id      INT,             -- tham chiếu mềm → departments.department_id
    position_code      VARCHAR,
    position_title     VARCHAR,
    level              VARCHAR,         -- cấp bậc (Junior/Senior/Lead…)
    headcount          INT,             -- số lượng cần tuyển
    status             VARCHAR,         -- Đang tuyển / Tạm dừng / Đã đóng
    jd_file_path       VARCHAR,         -- đường dẫn file JD trên máy
    mail_template_r1_id INT,            -- mẫu mail vòng 1
    mail_template_r2_id INT,            -- mẫu mail vòng 2
    mail_template_r3_id INT,            -- mẫu mail vòng 3
    created_at         DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at         DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- ───────────────────────── ỨNG VIÊN ─────────────────────────────────
CREATE TABLE IF NOT EXISTS candidates (
    candidate_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    full_name        VARCHAR,
    email            VARCHAR,
    phone            VARCHAR,
    date_of_birth    DATE,
    address          VARCHAR,
    position_id      INT,               -- tham chiếu mềm → positions.position_id
    years_experience INT,
    education        VARCHAR,
    applied_at       DATETIME,          -- ngày nộp CV
    status           VARCHAR,           -- giai đoạn tuyển dụng (xem CANDIDATE_STATUS_CHOICES)
    source           VARCHAR,           -- nguồn CV
    batch            INT,               -- đợt/lô quét CV (chỉ lưu SỐ: 1, 2, 3… từ tên thư mục batch1…)
    fit_score        DECIMAL,           -- điểm phù hợp (0–100)
    fit_summary      TEXT,
    strengths        TEXT,
    weaknesses       TEXT,
    cv_file_path     VARCHAR,           -- đường dẫn file CV trên máy
    note             TEXT,
    created_at       DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at       DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- ───────────────────────── NHÂN VIÊN ────────────────────────────────
-- Bảng soi theo file "Master HC file.xlsx" (sheet "Master file"): gần như mỗi
-- cột trong file có một cột tương ứng ở đây, chú thích `-- <Tiêu đề Excel>` ghi
-- rõ cột nguồn. Import khớp THEO TÊN CỘT, không theo thứ tự cột (xem
-- _EXCEL_HEADER_MAP trong app_qt/tools/employee_db.py).
--
-- Trạng thái làm việc suy ra từ `termination_date`: có ngày (khác NULL/rỗng)
-- = ĐÃ NGHỈ VIỆC (xem cv_repository.EMPLOYEE_WORK_STATUS_SQL).
--
-- Các cột trong file Excel không lưu ở đây vì đã có nguồn khác:
--   • "Business Unit (Department)" → lấy qua departments (department_id)
--   • "BC/WC"                      → lấy qua employee_types.collar
--   • "Position status"            → suy ra từ termination_date
--   • "Legal Entity (Company)", "STT", "Birthday", "Level",
--     "Job Title with level (no use)", "(Old) Phone Number" → cột cố định /
--     phụ trợ / công thức trong file.
CREATE TABLE IF NOT EXISTS employees (
    employee_id       INTEGER PRIMARY KEY AUTOINCREMENT,
    -- ── Định danh ──
    code              VARCHAR,           -- EC (mã nhân viên nội bộ)
    global_code       VARCHAR,           -- GlobalEmpCode
    -- ── Họ tên ──
    full_name         VARCHAR,           -- Full name
    surname           VARCHAR,           -- Surname (họ)
    name              VARCHAR,           -- Name (tên)
    middle_name       VARCHAR,           -- Middle Name (only for vietnam)
    -- ── Thông tin cá nhân ──
    date_of_birth     DATE,              -- Date of Birth
    gender            VARCHAR,           -- Gender
    place_of_birth    VARCHAR,           -- Place of birth
    native_place      VARCHAR,           -- Native country (nguyên quán)
    nationality       VARCHAR,           -- Nationality
    religion          VARCHAR,           -- Religion
    marriage_status   VARCHAR,           -- Marriage status (Yes) — Y/N
    marital_status    VARCHAR,           -- Marital Status (Single/Married…)
    spouse_name       VARCHAR,           -- Spouse Name
    spouse_dob        DATE,              -- Spouse date (ngày sinh vợ/chồng)
    children_count    INT,               -- Number of children
    children_names    TEXT,              -- Children's name (nhiều tên, mỗi dòng 1)
    children_birthdays TEXT,             -- Children's birthday
    -- ── Liên hệ ──
    phone             VARCHAR,           -- Phone Number (nhiều số, ngăn bởi "; ")
    email             VARCHAR,           -- Personal Email address
    company_email     VARCHAR,           -- Company email
    address           VARCHAR,           -- Street (address)
    city              VARCHAR,           -- City (address)
    country           VARCHAR,           -- Country (address)
    permanent_address VARCHAR,           -- Địa chỉ thường trú
    temporary_address VARCHAR,           -- Địa chỉ tạm trú
    emergency_contact_name VARCHAR,      -- Emergency Contact Name (CHỈ họ tên)
    emergency_contact_phone VARCHAR,     -- SĐT người báo tin khẩn cấp. File Excel gộp
                                          -- chung ô "Emergency Contact Name" dạng
                                          -- "tên ⏎ số ĐT" → lúc import tách sang đây
                                          -- (cv_repository.split_contact_name_phone)
    emergency_contact_relationship VARCHAR,  -- Relationship
    -- ── Học vấn ──
    education         VARCHAR,           -- Education Level
    education_field   VARCHAR,           -- Trình độ theo lĩnh vực
    major             VARCHAR,           -- Major (nhiều chuyên ngành, mỗi dòng 1)
    graduation_year   VARCHAR,           -- Year of graduated
    school_name       VARCHAR,           -- School name
    qualification     VARCHAR,           -- Qualification
    qualification_code VARCHAR,          -- Qualification code
    -- ── Giấy tờ · ngân hàng · thuế · bảo hiểm ──
    id_no             VARCHAR,           -- ID no. (CCCD/CMND)
    id_issued_date    DATE,              -- Issued date (của ID no.)
    id_issued_place   VARCHAR,           -- Issued Place
    passport_no       VARCHAR,           -- Passport No.
    passport_issued_date DATE,           -- Issued date (của Passport)
    bank_account_no   VARCHAR,           -- Bank account no.
    bank_address      VARCHAR,           -- Bank address
    tax_code          VARCHAR,           -- Personal Tax Code
    dependants        INT,               -- Dependance (số người phụ thuộc)
    insurance_book_no VARCHAR,           -- Insurance Book No.
    -- ── Tổ chức & công việc ──
    department_id     INT,               -- Function (Common) → departments.short_name
    cost_center_id    INT,               -- New Cost center  → cost_centers.code
    employee_type_id  INT,               -- IBC/DBC/WC       → employee_types.code
    level_id          INT,               -- Job level        → levels.level_name
    manager_name      VARCHAR,           -- Full name of manager
    job_title         VARCHAR,           -- Job Title (Description)
    current_position  VARCHAR,           -- Current Position
    time_in_position  VARCHAR,           -- Time in Position
    facility_country  VARCHAR,           -- Country of facility
    facility_town     VARCHAR,           -- Town of facility
    local_function    VARCHAR,           -- Function (for local only)
    by_group          VARCHAR,           -- BY GROUP
    labor_type        VARCHAR,           -- Type of labor
    production_line   VARCHAR,           -- Production Line (Internal)
    operator_skill    VARCHAR,           -- Operator skill
    driving_forklift  VARCHAR,           -- Driving forklift
    working_hours_per_week VARCHAR,      -- Working hour/week
    smart_working_eligible VARCHAR,      -- Eligible -Smart Working Policy Eligible
    er_jrf            VARCHAR,           -- #ER/ JRF
    -- ── Hợp đồng & thời gian làm việc ──
    date_of_employment DATE,             -- Date of Employment
    seniority_date    DATE,              -- Cột tính thâm niên
    contract_permanency VARCHAR,         -- Permanent/Temporary contract
    work_time_type    VARCHAR,           -- Full Time/ Part Time (FT/PT)
    working_time_pct  VARCHAR,           -- % working time
    direct_indirect   VARCHAR,           -- Direct/Indirect
    contract_type     VARCHAR,           -- Type of contract
    contract_start_date DATE,            -- Starting date of contract
    contract_end_date DATE,              -- Ending date of contract
    changing_date     DATE,              -- Changing date (đổi vị trí/hợp đồng)
    termination_date  DATE,              -- Termination Date — CÓ giá trị = ĐÃ NGHỈ VIỆC
    leaving_reason    VARCHAR,           -- Reason for leaving
    -- ── Số liệu file Excel tự tính (lưu lại ảnh chụp lúc import) ──
    years_of_service  VARCHAR,           -- Year of service
    length_of_service VARCHAR,           -- Length of service
    birth_year        VARCHAR,           -- Year of birthday (year)
    age               VARCHAR,           -- Age
    age_range         VARCHAR,           -- Age range
    -- ── Ghi chú ──
    changing_notes    TEXT,              -- Changing notes
    changing_dates    TEXT,              -- Changing date (lịch sử, nhiều dòng)
    updated_changing_date DATE,          -- Updated changing date
    note              TEXT,              -- Note
    created_at        DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at        DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- ───────────────────────── KHÓA HỌC / ĐÀO TẠO ───────────────────────
CREATE TABLE IF NOT EXISTS courses (
    course_id   INTEGER PRIMARY KEY AUTOINCREMENT,
    title       VARCHAR,                 -- tên khóa học
    content     TEXT,                    -- nội dung
    date        DATE,                    -- ngày tổ chức
    location    VARCHAR,                 -- địa điểm
    course_type INT,                     -- loại: 0=inhouse, 1=external, 2=funded (xem COURSE_TYPE_CHOICES)
    created_at  DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at  DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- ─────────────── LIÊN KẾT KHÓA HỌC ↔ NHÂN VIÊN (nhiều-nhiều) ─────────
-- Một khóa học có nhiều nhân viên; một nhân viên tham gia nhiều khóa học.
-- Tham chiếu mềm (không FK): course_id → courses, employee_id → employees.
CREATE TABLE IF NOT EXISTS course_employees (
    enrollment_id INTEGER PRIMARY KEY AUTOINCREMENT,
    course_id     INT,                   -- tham chiếu mềm → courses.course_id
    employee_id   INT,                   -- tham chiếu mềm → employees.employee_id
    status        VARCHAR,               -- trạng thái học (xem COURSE_STATUS_CHOICES)
    note          TEXT,
    created_at    DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at    DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- ───────────────────────── CHỈ MỤC (tăng tốc tìm kiếm) ───────────────
CREATE INDEX IF NOT EXISTS idx_candidates_name   ON candidates(full_name);
CREATE INDEX IF NOT EXISTS idx_candidates_email  ON candidates(email);
CREATE INDEX IF NOT EXISTS idx_candidates_phone  ON candidates(phone);
CREATE INDEX IF NOT EXISTS idx_candidates_pos    ON candidates(position_id);
CREATE INDEX IF NOT EXISTS idx_candidates_status ON candidates(status);
CREATE INDEX IF NOT EXISTS idx_candidates_batch  ON candidates(batch);
CREATE INDEX IF NOT EXISTS idx_positions_dept    ON positions(department_id);
CREATE INDEX IF NOT EXISTS idx_employees_name    ON employees(full_name);
CREATE INDEX IF NOT EXISTS idx_employees_code     ON employees(code);
CREATE INDEX IF NOT EXISTS idx_employees_dept     ON employees(department_id);
CREATE INDEX IF NOT EXISTS idx_employees_level     ON employees(level_id);
CREATE INDEX IF NOT EXISTS idx_employees_costcenter ON employees(cost_center_id);
CREATE INDEX IF NOT EXISTS idx_employees_emptype   ON employees(employee_type_id);
CREATE INDEX IF NOT EXISTS idx_employees_termination ON employees(termination_date);
CREATE INDEX IF NOT EXISTS idx_ce_course           ON course_employees(course_id);
CREATE INDEX IF NOT EXISTS idx_ce_employee         ON course_employees(employee_id);
CREATE UNIQUE INDEX IF NOT EXISTS idx_ce_unique    ON course_employees(course_id, employee_id);
CREATE INDEX IF NOT EXISTS idx_emptype_code        ON employee_types(code);
CREATE INDEX IF NOT EXISTS idx_costcenter_code     ON cost_centers(code);
CREATE INDEX IF NOT EXISTS idx_levels_name         ON levels(level_name);
CREATE INDEX IF NOT EXISTS idx_mail_templates_type ON mail_templates(type);
"""

# =============================================================================
#  MIGRATIONS — câu SQL chạy SAU khi tạo bảng, để thêm cột / chỉ mục cho file DB
#  ĐÃ tồn tại (SCHEMA_SQL chỉ tạo bảng khi CHƯA có, không tự thêm cột).
#  Mỗi câu chạy độc lập, lỗi được bỏ qua nên an toàn khi chạy lại nhiều lần.
#  Ví dụ:  "ALTER TABLE candidates ADD COLUMN linkedin VARCHAR",
#
#  Chỉ mục cho cột thêm bằng ALTER cũng đặt ở đây, KHÔNG đặt trong SCHEMA_SQL:
#  SCHEMA_SQL chạy trước, index trỏ vào cột chưa có sẽ làm vỡ cả script.
# =============================================================================
MIGRATIONS: list[str] = [
    # Mẫu mail tách khỏi bảng positions → chỉ còn 3 cột trỏ tới mail_templates.
    "ALTER TABLE positions ADD COLUMN mail_template_r1_id INT",
    "ALTER TABLE positions ADD COLUMN mail_template_r2_id INT",
    "ALTER TABLE positions ADD COLUMN mail_template_r3_id INT",
    # Bỏ hẳn 3 cột mẫu mail cũ (nội dung nhập lại ở màn hình Mail templates).
    "ALTER TABLE positions DROP COLUMN mail_cc",
    "ALTER TABLE positions DROP COLUMN mail_subject",
    "ALTER TABLE positions DROP COLUMN mail_body",
    # SĐT người báo tin khẩn cấp tách khỏi cột tên (xem DATA_MIGRATIONS bên dưới).
    "ALTER TABLE employees ADD COLUMN emergency_contact_phone VARCHAR",
]

# =============================================================================
#  DATA_MIGRATIONS — sửa DỮ LIỆU (không phải cấu trúc), chạy đúng MỘT LẦN cho
#  mỗi file DB. `cv_repository._run_data_migrations()` gọi trong init_db() và
#  ghi dấu vết vào app_meta với khóa "data:<tên lượt>" → lần mở sau bỏ qua hẳn,
#  không quét lại bảng. Muốn chạy lại: xóa dòng đó trong app_meta.
#
#  Khác với MIGRATIONS ở trên (ALTER TABLE — chạy lại mỗi lần mở, vô hại vì chỉ
#  báo lỗi "duplicate column" rồi bỏ qua).
#
#  Mỗi lượt là một danh sách câu SQL chạy theo thứ tự. Đặt tên kèm version để
#  sau này cần sửa tiếp thì thêm lượt mới, KHÔNG sửa lượt cũ (DB đã chạy rồi).
# =============================================================================
DATA_MIGRATIONS: dict[str, list[str]] = {
    # Đưa trạng thái ứng viên của DB cũ về bộ nhãn trong CANDIDATE_STATUS_CHOICES.
    # So khớp bỏ qua hoa/thường & khoảng trắng thừa.
    "candidate_status:v1": [
        "UPDATE candidates SET status = 'New Application' "
        "WHERE LOWER(TRIM(COALESCE(status, ''))) = 'new'",
        "UPDATE candidates SET status = 'Screening' "
        "WHERE LOWER(TRIM(COALESCE(status, ''))) IN ('contacted', 'on hold')",
        "UPDATE candidates SET status = 'First Interview' "
        "WHERE LOWER(TRIM(COALESCE(status, ''))) = 'interview'",
        "UPDATE candidates SET status = 'Ready To Hire' "
        "WHERE LOWER(TRIM(COALESCE(status, ''))) = 'passed'",
        "UPDATE candidates SET status = 'Not Proceed' "
        "WHERE LOWER(TRIM(COALESCE(status, ''))) = 'rejected'",
    ],
}

# =============================================================================
#  SEED — DỮ LIỆU KHỞI TẠO cho các bảng danh mục (nguồn: file Code.xlsx).
#
#  Cách chạy: `cv_repository._seed_master_data()` gọi trong init_db().
#    • Mỗi khối chỉ nạp MỘT LẦN cho mỗi file .db — dấu vết ghi ở bảng app_meta
#      với khóa "seed:<bảng>:v<version>". Xóa dòng đó nếu muốn nạp lại.
#    • Trong một lần nạp, dòng đã có sẵn (trùng theo cột ở `match`) sẽ bị bỏ qua
#      → chạy lại không tạo bản ghi trùng, và KHÔNG đụng vào dữ liệu bạn tự nhập.
#    • Muốn bổ sung danh mục về sau: thêm dòng vào `rows` rồi TĂNG `version`.
# =============================================================================
SEED_DATA: dict[str, dict] = {
    # ── Cột B (Department) + C (Department (2)) của Code.xlsx ───────────
    # Bỏ 1 dòng trùng khít ở cuối file gốc: "Global Procurement / Gpr".
    "departments": {
        "version": 1,
        "columns": ("department_name", "short_name"),
        "match": ("department_name",),     # đã có tên này rồi → bỏ qua
        "rows": [
            ("Facilities Control",               "FC"),
            ("Sales SEA",                        "Sale"),
            ("Global Supply Quality",            "Gpr"),
            ("Finance",                          "FIN"),
            ("Global Operations",                "BOM"),
            ("Global Procurement",               "Gpr"),
            ("Production",                       "PD"),
            ("Manufacturing Engineering",        "ME"),
            ("Production Planning",              "PL"),
            ("Logistics",                        "LG"),
            ("Quality Assurance",                "QA"),
            ("Human Resource & Administration",  "HR"),
            ("Research & Development",           "R&D"),
            ("Global Planning",                  "GP"),
            ("Global Logistics",                 "GL"),
            ("Materials Management",             "MM"),
            ("Information Technology",           "IT"),
            ("Local COM",                        "COM"),
            ("Global COM",                       "GCOM"),
            ("Purchasing/Indirect Material",     "Gpr"),
        ],
    },
    # ── Cột D (Employee type (2)) + E (BC/WC) ───────────────────────────
    # LƯU Ý: trong file gốc cột E chỉ có 3 ô và KHÔNG thẳng hàng với cột D
    # (D2=WC nhưng E2=Blue Collar). Ở đây suy ra theo nghĩa của mã:
    #   WC/WCA = White Collar · IBC/IBCA (Indirect) & DBC/DBCA (Direct) = Blue Collar.
    # Nếu quy ước công ty khác → sửa lại cột `collar` ngay bên dưới.
    "employee_types": {
        "version": 1,
        "columns": ("code", "collar"),
        "match": ("code",),
        "rows": [
            ("WC",   "White Collar"),
            ("WCA",  "White Collar"),
            ("IBC",  "Blue Collar"),
            ("IBCA", "Blue Collar"),
            ("DBC",  "Blue Collar"),
            ("DBCA", "Blue Collar"),
        ],
    },
    # ── Cột F (Cost center) + G (Group Function) ────────────────────────
    "cost_centers": {
        "version": 1,
        "columns": ("code", "group_function"),
        "match": ("code",),
        "rows": [
            ("VN1001", "VNPlant"),
            ("VN1002", "VNPlant"),
            ("VN1003", "VNPlant"),
            ("VN1004", "VNPlant"),
            ("VN1005", "VNPlant"),
            ("VN1006", "VNPlant"),
            ("VN1007", "VNPlant"),
            ("VN1008", "VNPlant"),
            ("VN1011", "VNPlant"),
            ("VN1012", "VNPlant"),
            ("VN1021", "VNPlant"),
            ("VN1023", "VNPlant"),
            ("VN1024", "VNPlant"),
            ("VN1031", "Corporate"),
            ("VN1032", "Corporate"),
            ("VN1033", "VNPlant"),
            ("VN1035", "Corporate"),
            ("VN1041", "Corporate"),
            ("VN1042", "VNPlant"),
            ("VN1051", "Corporate"),
            ("VN1052", "Corporate"),
            ("VN1054", "VNPlant"),
            ("VN1071", "Corporate"),
            ("VN1072", "Corporate"),
            ("VN3000", "R&D"),
            ("VN3001", "R&D"),
            ("VN3007", "Corporate"),
            ("VN3017", "R&D"),
            ("VN3024", "R&D"),
            ("VN3034", "R&D"),
            ("VN3041", "R&D"),
            ("VN3051", "R&D"),
            ("VN3061", "R&D"),
            ("VN4112", "Corporate"),
            ("VN4120", "Corporate"),
            ("VN4211", "Corporate"),
            ("VN6021", "Corporate"),
            ("VN7010", "Corporate"),
            ("VN7031", "Corporate"),
            ("VN7032", "Corporate"),
            ("VN7040", "Corporate"),
            ("VN7043", "Corporate"),
        ],
    },
    # ── Cột H (Tên Cấp) — giữ nguyên thứ tự trong file (sort_order) ─────
    "levels": {
        "version": 1,
        "columns": ("level_name", "sort_order"),
        "match": ("level_name",),
        "rows": [
            ("Director",          1),
            ("Manager",           2),
            ("Officer",           3),
            ("Engineer",          4),
            ("Supervisor",        5),
            ("Technician",        6),
            ("Clerk",             7),
            ("Lead Operator",     8),
            ("Assistant Manager", 9),
            ("Operator",          10),
            ("Technician Lead",   11),
            ("Team Leader",       12),
        ],
    },
}

# Trạng thái ứng viên (cột candidates.status) — LƯU DẠNG TEXT đúng bằng nhãn
# dưới đây. Danh sách xếp theo đúng luồng tuyển dụng: chỉ số trong list = thứ tự
# giai đoạn, dùng khi cần sắp xếp/so sánh tiến độ (xem candidate_status_order).
# "Not Proceed" là nhánh dừng: hồ sơ không đi tiếp sau khi lọc short list.
CANDIDATE_STATUS_CHOICES = [
    "New Application",
    "Screening",
    "Short List",
    "Not Proceed",
    "First Interview",
    "Second Interview",
    "Third Interview",
    "Offer Approval",
    "Ready To Hire",
    "Rejected Offer",
    "Fail Probation Period",
]

# Trạng thái mặc định của hồ sơ vừa vào DB (nhập tay hoặc quét CV bằng AI).
CANDIDATE_STATUS_DEFAULT = CANDIDATE_STATUS_CHOICES[0]

# Bước kế tiếp trong luồng — dùng làm gợi ý sẵn khi đổi trạng thái hàng loạt.
# Nhãn KHÔNG có mặt ở đây là điểm dừng của luồng (Not Proceed · Rejected Offer ·
# Fail Probation Period): hồ sơ đến đó thì không còn bước kế tiếp mặc định.
CANDIDATE_STATUS_NEXT = {
    "New Application":  "Screening",
    "Screening":        "Short List",
    "Short List":       "First Interview",
    "First Interview":  "Second Interview",
    "Second Interview": "Third Interview",
    "Third Interview":  "Offer Approval",
    "Offer Approval":   "Ready To Hire",
}


def candidate_status_order(status: str) -> int:
    """Thứ tự giai đoạn của một trạng thái (-1 nếu là nhãn lạ/rỗng)."""
    target = (status or "").strip().lower()
    for i, label in enumerate(CANDIDATE_STATUS_CHOICES):
        if label.lower() == target:
            return i
    return -1


def candidate_next_status(status: str) -> str:
    """Bước kế tiếp gợi ý của một trạng thái.

    Hồ sơ chưa có trạng thái → bắt đầu từ đầu luồng. Trả về chuỗi rỗng khi
    trạng thái là điểm dừng hoặc là nhãn lạ (không có bước kế tiếp).
    """
    target = (status or "").strip().lower()
    if not target:
        return CANDIDATE_STATUS_DEFAULT
    for label, nxt in CANDIDATE_STATUS_NEXT.items():
        if label.lower() == target:
            return nxt
    return ""


POSITION_STATUS_CHOICES = ["Open", "Paused", "Closed"]

# Loại mẫu mail (cột mail_templates.type) — chỉ để phân nhóm/lọc cho dễ tìm,
# KHÔNG ràng buộc: vị trí tuyển dụng vẫn chọn được bất kỳ mẫu nào cho mỗi vòng.
#
# Riêng "Application Thank You" có ý nghĩa với luồng gửi mail: mẫu loại này KHÔNG
# gắn vào vị trí mà chọn thẳng lúc gửi, và gửi dạng MAIL THƯỜNG (không giờ giấc,
# không phải thư mời họp) — xem MAIL_TEMPLATE_TYPE_THANK_YOU.
MAIL_TEMPLATE_TYPE_THANK_YOU = "Application Thank You"

MAIL_TEMPLATE_TYPE_CHOICES = [
    "Interview Round 1",
    "Interview Round 2",
    "Interview Round 3",
    MAIL_TEMPLATE_TYPE_THANK_YOU,
    "Notification",
    "Offer",
    "Rejection",
]

# 3 VÒNG PHỎNG VẤN của một vị trí: (cột trong positions, nhãn hiển thị, trạng
# thái ứng viên gợi ý sau khi gửi thư mời vòng đó).
INTERVIEW_ROUNDS = [
    ("mail_template_r1_id", "Interview Round 1", "First Interview"),
    ("mail_template_r2_id", "Interview Round 2", "Second Interview"),
    ("mail_template_r3_id", "Interview Round 3", "Third Interview"),
]

# Trạng thái HIỆN TẠI của hồ sơ → vòng phỏng vấn gợi ý sẵn khi bấm Send email
# (chỉ số trong INTERVIEW_ROUNDS). Hồ sơ vừa lọt short list thì mời vòng 1; mời
# xong vòng 1 trạng thái thành "First Interview" nên lần sau gợi ý vòng 2…
INTERVIEW_ROUND_BY_STATUS = {
    "Short List":       0,
    "First Interview":  1,
    "Second Interview": 2,
}


def interview_round_for_status(status: str) -> int:
    """Vòng phỏng vấn gợi ý cho một trạng thái (mặc định vòng 1 nếu không khớp)."""
    target = (status or "").strip().lower()
    for label, idx in INTERVIEW_ROUND_BY_STATUS.items():
        if label.lower() == target:
            return idx
    return 0

# Nhân viên: giới tính — dùng cho ô lọc + form nhập. Cấp bậc (level) tham chiếu
# bảng danh mục `levels` qua level_id.
GENDER_CHOICES = ["Male", "Female", "Other"]

# Nhãn trạng thái làm việc để hiển thị/lọc. Đây là giá trị SUY RA từ
# `employees.termination_date` (rỗng = Working, có ngày = Resigned — xem
# cv_repository.EMPLOYEE_WORK_STATUS_SQL), không phải cột trong DB.
EMPLOYEE_STATUS_CHOICES = ["Working", "Resigned"]

# Các ô chọn của form nhân viên — giá trị lấy đúng như trong "Master HC file.xlsx".
CONTRACT_PERMANENCY_CHOICES = ["Permanent", "Temporary"]
WORK_TIME_TYPE_CHOICES = ["FT", "PT"]           # Full Time / Part Time
DIRECT_INDIRECT_CHOICES = ["Direct", "Indirect"]
YES_NO_CHOICES = ["Y", "N"]                     # Marriage status (Yes)
MARITAL_STATUS_CHOICES = ["Single", "Married", "Divorced", "Widowed"]

# Danh mục nhân viên: nhóm lao động & nhóm chức năng của cost center.
COLLAR_CHOICES = ["Blue Collar", "White Collar"]
GROUP_FUNCTION_CHOICES = ["VNPlant", "Corporate", "R&D"]

# Loại khóa học — lưu dạng INT trong cột courses.course_type (chỉ số = giá trị lưu).
COURSE_TYPE_CHOICES = ["inhouse", "external", "funded"]  # 0, 1, 2

# Trạng thái học của nhân viên trong 1 khóa (cột course_employees.status).
# "Not started" = chưa học · "Completed" = đã học xong.
COURSE_STATUS_CHOICES = ["Not started", "Completed"]

# Danh sách bảng do init_db quản lý — dùng khi cần dựng lại bảng trống lệch schema.
_MANAGED_TABLES = [
    "departments", "positions", "candidates",
    "employees", "courses", "course_employees",
    "employee_types", "cost_centers", "levels", "mail_templates",
]
