"""THIẾT KẾ CƠ SỞ DỮ LIỆU — Tuyển dụng & Nhân sự (SQLite).

⭐ ĐÂY LÀ FILE ĐỂ THIẾT KẾ / CHỈNH SỬA CẤU TRÚC DB ⭐
Tài liệu mô tả đầy đủ: `docs/db_design.md`.

Toàn bộ bảng, cột nằm trong hằng `SCHEMA_SQL` bên dưới.

QUY ƯỚC THIẾT KẾ
  • KHÔNG dùng ràng buộc khóa ngoại (FOREIGN KEY). Các cột *_id chỉ là số tham
    chiếu "mềm" — ứng dụng tự đảm bảo liên kết, DB không ép buộc.
  • Mọi cột đều CHO PHÉP NULL, trừ khóa chính (PK) tự tăng.
  • Kiểu dữ liệu ghi theo sơ đồ (VARCHAR/INT/TEXT/DATE/DATETIME/DECIMAL).
  • Danh sách nhiều giá trị trong một ô: ngăn bởi dấu ";".
  • `note` trong mọi bảng là ghi chú của HR (HR là người vận hành hệ thống).

CÁCH SỬA CẤU TRÚC VỀ SAU — chỉ có MỘT đường duy nhất: `MIGRATIONS`.
  Đó là danh sách các lượt SQL xếp theo thứ tự, mỗi lượt chạy ĐÚNG MỘT LẦN cho
  mỗi file .db (dấu vết ghi ở bảng `app_meta`):
     • Máy chưa có DB   → chạy TẤT CẢ các lượt, từ lượt tạo bảng đầu tiên.
     • Máy đã có DB rồi → chỉ chạy những lượt CÒN THIẾU.
  Thêm cột thì thêm một lượt `ALTER TABLE ...` vào cuối danh sách, tắt mở lại
  app là xong — dữ liệu đang có không suy suyển. Xem chú thích ở `MIGRATIONS`.

────────────────────────────────────────────────────────────────────────────
SƠ ĐỒ QUAN HỆ (mềm, không ràng buộc FK)

  DANH MỤC        departments ──1:N──> functions  (nhóm nhỏ trong phòng ban)
                  levels · employee_types · cost_centers
                  code_lists  (nhiều danh mục một cột, phân biệt bằng `type`)
                  skills · mail_templates · app_meta

  TUYỂN DỤNG      positions  ──1:N──> position_requirements  (JD đã bóc tách)
                  positions  ──1:N──> applications
                  candidates ──1:N──> candidate_cvs          (CV theo thời gian)
                  candidates ──1:N──> candidate_experiences  (dòng thời gian việc)
                  candidates ──1:N──> candidate_skills ──N:1── skills
                  candidates ──1:N──> applications ──1:N──> interviews
                                                   ──1:N──> candidate_activities
                  interviews ──1:N──> interview_feedbacks ──N:1── employees
                  candidates ──1:N──> candidate_evaluations  (lịch sử AI chấm)
                  candidates ──1:1──> candidates_fts         (tìm toàn văn)

  NHÂN SỰ         employees ──N:M── courses  (qua course_employees)

BA NGUYÊN TẮC CỦA THIẾT KẾ
  1. `candidates` mô tả CON NGƯỜI — trung tính, không dính JD nào. Trạng thái
     tuyển dụng nằm ở `applications`, điểm AI nằm ở `candidate_evaluations`.
  2. `candidate_evaluations` CHỈ GHI THÊM, không ghi đè: một ứng viên được chấm
     nhiều lần (nhiều vị trí, nhiều thời điểm, nhiều bản CV) → giữ đủ lịch sử.
  3. Kinh nghiệm lưu thành DÒNG THỜI GIAN (`candidate_experiences`) chứ không
     phải một con số chết, nên CV nhận năm 2023 xem lại năm 2026 vẫn tính đúng.
────────────────────────────────────────────────────────────────────────────
"""

# Cấu trúc ban đầu — nội dung của lượt migration đầu tiên (xem `MIGRATIONS`).
# ĐÃ PHÁT HÀNH thì KHÔNG sửa hằng này nữa: máy nào chạy qua lượt "0001" rồi sẽ
# không chạy lại, nên sửa ở đây chỉ có tác dụng với máy cài mới → hai máy lệch
# cấu trúc nhau. Muốn đổi cấu trúc: thêm một lượt mới ở cuối `MIGRATIONS`.
SCHEMA_SQL = """
-- ═══════════════════════ DANH MỤC DÙNG CHUNG ════════════════════════════

-- Phòng ban. `short_name` khớp cột "Function (Common)" khi import nhân viên.
CREATE TABLE IF NOT EXISTS departments (
    department_id   INTEGER PRIMARY KEY AUTOINCREMENT,
    department_name VARCHAR,
    short_name      VARCHAR,
    manager_name    VARCHAR,
    description     TEXT,
    created_at      DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at      DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- Loại nhân viên (WC/WCA/IBC/IBCA/DBC/DBCA) + nhóm lao động (collar).
CREATE TABLE IF NOT EXISTS employee_types (
    employee_type_id INTEGER PRIMARY KEY AUTOINCREMENT,
    code             VARCHAR,
    collar           VARCHAR,
    description      TEXT,
    created_at       DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at       DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- Trung tâm chi phí — gán cho nhân viên để gom nhóm tính chi phí vận hành.
CREATE TABLE IF NOT EXISTS cost_centers (
    cost_center_id   INTEGER PRIMARY KEY AUTOINCREMENT,
    code             VARCHAR,
    group_function   VARCHAR,
    name             VARCHAR,
    description      TEXT,
    created_at       DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at       DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- Cấp bậc (Director, Manager, Officer…).
CREATE TABLE IF NOT EXISTS levels (
    level_id    INTEGER PRIMARY KEY AUTOINCREMENT,
    level_name  VARCHAR,
    sort_order  INT,
    description TEXT,
    created_at  DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at  DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- Kỹ năng chuẩn hóa. `aliases` gom các cách viết khác nhau (ngăn bởi ";") để
-- "JS" trong CV khớp được với "JavaScript" trong JD.
CREATE TABLE IF NOT EXISTS skills (
    skill_id    INTEGER PRIMARY KEY AUTOINCREMENT,
    name        VARCHAR,
    aliases     VARCHAR,
    category    VARCHAR,
    description TEXT,
    created_at  DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at  DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- Kho mẫu mail dùng chung. Nội dung hỗ trợ placeholder {name} {possion}
-- {position} {date} {time_start} {time_end} — điền lúc gửi.
CREATE TABLE IF NOT EXISTS mail_templates (
    mail_template_id INTEGER PRIMARY KEY AUTOINCREMENT,
    name             VARCHAR,
    type             VARCHAR,
    mail_cc          VARCHAR,
    mail_subject     VARCHAR,
    mail_body        TEXT,
    created_at       DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at       DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- Key-value nội bộ: phiên bản schema, dấu vết các lượt nạp dữ liệu khởi tạo.
CREATE TABLE IF NOT EXISTS app_meta (
    key   VARCHAR PRIMARY KEY,
    value VARCHAR
);


-- ═════════════════════════════ TUYỂN DỤNG ═══════════════════════════════

-- Vị trí tuyển dụng. Mỗi vị trí có ĐÚNG 1 JD → đường dẫn file nằm ngay đây.
-- 3 cột mail_template_r*_id = mẫu mail cho 3 vòng phỏng vấn (xem INTERVIEW_ROUNDS).
CREATE TABLE IF NOT EXISTS positions (
    position_id         INTEGER PRIMARY KEY AUTOINCREMENT,
    department_id       INT,             -- → departments.department_id
    position_code       VARCHAR,
    jrf_code            VARCHAR,         -- mã yêu cầu tuyển dụng (Job Requisition Form)
    position_title      VARCHAR,         -- dùng luôn làm tiêu đề JD
    description         TEXT,
    required_experience VARCHAR,
    salary_level        VARCHAR,         -- dải lương đã duyệt cho vị trí
    starting_date       DATE,            -- ngày cần người vào làm
    level               VARCHAR,         -- Junior/Senior/Lead…
    headcount           INT,
    status              VARCHAR,         -- xem POSITION_STATUS_CHOICES
    jd_file_path        VARCHAR,
    mail_template_r1_id INT,             -- → mail_templates
    mail_template_r2_id INT,
    mail_template_r3_id INT,
    note                TEXT,
    created_at          DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at          DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- Yêu cầu BÓC TÁCH TỪ JD bằng AI. Bóc một lần, dùng lại cho mọi lượt tìm kiếm.
-- `jd_hash` = băm nội dung file JD: JD sửa → hash đổi → bóc lại, dòng cũ giữ nguyên.
CREATE TABLE IF NOT EXISTS position_requirements (
    requirement_id INTEGER PRIMARY KEY AUTOINCREMENT,
    position_id    INT,              -- → positions.position_id
    jd_hash        VARCHAR,
    must_have      TEXT,             -- kỹ năng bắt buộc, ngăn bởi ";"
    nice_to_have   TEXT,             -- kỹ năng ưu tiên, ngăn bởi ";"
    min_years      DECIMAL,
    education      VARCHAR,
    major          VARCHAR,
    language_req   VARCHAR,
    level          VARCHAR,
    summary        TEXT,
    model          VARCHAR,          -- model AI đã bóc
    parsed_at      DATETIME,
    created_at     DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at     DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- ỨNG VIÊN — hồ sơ một CON NGƯỜI trong pool.
-- Nhóm "ảnh chụp hồ sơ" là bản CHÉP từ bản CV mới nhất (candidate_cvs), cố ý dư
-- thừa để lọc/sắp xếp nhanh; bản gốc theo từng thời điểm nằm ở candidate_cvs.
CREATE TABLE IF NOT EXISTS candidates (
    candidate_id       INTEGER PRIMARY KEY AUTOINCREMENT,
    -- ── Định danh & liên hệ ──
    full_name          VARCHAR,
    email              VARCHAR,
    phone              VARCHAR,
    date_of_birth      DATE,
    gender             VARCHAR,
    address            VARCHAR,
    city               VARCHAR,
    -- ── Vòng đời trong pool ──
    pool_status        VARCHAR,      -- xem POOL_STATUS_CHOICES
    first_seen_at      DATETIME,     -- lần đầu vào pool
    last_contacted_at  DATETIME,     -- chép từ candidate_activities mới nhất
    latest_cv_id       INT,          -- → candidate_cvs.cv_id (bản CV mới nhất)
    source             VARCHAR,      -- nguồn biết đến lần đầu
    note               TEXT,
    -- ── Ảnh chụp hồ sơ nghề nghiệp (từ bản CV mới nhất) ──
    current_title      VARCHAR,
    industry           VARCHAR,
    years_experience   DECIMAL,      -- số năm TẠI experience_as_of, không phải hôm nay
    experience_as_of   DATE,
    education          VARCHAR,
    major              VARCHAR,
    languages          VARCHAR,
    skills_text        TEXT,         -- kỹ năng thô như trong CV, ngăn bởi ";"
    profile_summary    TEXT,         -- 2–3 dòng trung tính, không nhắc JD
    -- ── Nguyện vọng (tiêu chí lọc khi tìm trong pool) ──
    expected_salary    DECIMAL,
    salary_note        VARCHAR,      -- Gross/Net/khoảng thương lượng
    available_from     DATE,
    willing_to_relocate VARCHAR,     -- xem RELOCATE_CHOICES
    preferred_location VARCHAR,
    created_at         DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at         DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- CÁC BẢN CV THEO THỜI GIAN — mỗi dòng là ẢNH CHỤP BẤT BIẾN, ghi rồi không sửa.
-- `file_hash` băm nội dung file → cùng một CV thả vào lần hai (kể cả đã đổi tên)
-- vẫn nhận ra, không tốn tiền gọi AI đọc lại.
CREATE TABLE IF NOT EXISTS candidate_cvs (
    cv_id            INTEGER PRIMARY KEY AUTOINCREMENT,
    candidate_id     INT,           -- → candidates.candidate_id
    file_path        VARCHAR,
    file_hash        VARCHAR,
    received_at      DATE,          -- MỐC THỜI GIAN của mọi số liệu trong bản này
    batch            INT,           -- số đợt quét, bóc từ tên thư mục batch1, batch2…
    source           VARCHAR,
    cv_text          TEXT,          -- toàn văn CV → chấm lại khỏi mở lại PDF
    scanned_at       DATETIME,
    scan_model       VARCHAR,
    -- ── Ảnh chụp hồ sơ tại thời điểm này (đóng băng) ──
    current_title    VARCHAR,
    industry         VARCHAR,
    years_experience DECIMAL,
    experience_as_of DATE,
    education        VARCHAR,
    major            VARCHAR,
    languages        VARCHAR,
    skills_text      TEXT,
    profile_summary  TEXT,
    created_at       DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at       DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- DÒNG THỜI GIAN CÔNG VIỆC — thứ cho phép tính lại số năm kinh nghiệm ở bất kỳ
-- thời điểm nào. `end_date` rỗng = còn làm TÍNH ĐẾN `as_of_date` (ngày nhận CV),
-- từ đó tới hôm nay chỉ là suy đoán (xem experience_years trong repository).
CREATE TABLE IF NOT EXISTS candidate_experiences (
    experience_id INTEGER PRIMARY KEY AUTOINCREMENT,
    candidate_id  INT,              -- → candidates.candidate_id
    cv_id         INT,              -- bóc từ bản CV nào → candidate_cvs.cv_id
    company       VARCHAR,
    job_title     VARCHAR,
    industry      VARCHAR,
    start_date    DATE,             -- chỉ có tháng/năm thì lấy ngày 01
    end_date      DATE,             -- RỖNG = còn làm tính đến as_of_date
    as_of_date    DATE,             -- chép từ candidate_cvs.received_at
    is_current    INT,              -- 1 = lúc viết CV vẫn đang làm ở đây
    description   TEXT,
    sort_order    INT,              -- mới nhất trước
    created_at    DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at    DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- KỸ NĂNG ĐÃ CHUẨN HÓA — cầu nối giữa chữ trong CV và danh mục `skills`.
CREATE TABLE IF NOT EXISTS candidate_skills (
    candidate_skill_id INTEGER PRIMARY KEY AUTOINCREMENT,
    candidate_id       INT,         -- → candidates.candidate_id
    skill_id           INT,         -- → skills.skill_id (rỗng = chưa tra được)
    cv_id              INT,         -- đọc từ bản CV nào
    raw_name           VARCHAR,     -- đúng chữ trong CV
    years              DECIMAL,
    level              VARCHAR,     -- xem SKILL_LEVEL_CHOICES
    source             VARCHAR,     -- xem SKILL_SOURCE_CHOICES
    created_at         DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at         DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- ĐƠN ỨNG TUYỂN (ứng viên × vị trí). TRẠNG THÁI TUYỂN DỤNG NẰM Ở ĐÂY, không
-- nằm ở candidates — nhờ vậy một người ứng tuyển nhiều vị trí, hoặc ứng tuyển
-- lại cùng vị trí sau vài năm, đều giữ nguyên lịch sử cũ.
-- CỐ Ý không đặt unique (candidate_id, position_id): cho phép nộp lại đợt sau.
CREATE TABLE IF NOT EXISTS applications (
    application_id    INTEGER PRIMARY KEY AUTOINCREMENT,
    candidate_id      INT,          -- → candidates.candidate_id
    position_id       INT,          -- → positions.position_id
    cv_id             INT,          -- nộp bằng bản CV nào
    origin            VARCHAR,      -- xem APPLICATION_ORIGIN_CHOICES
    source            VARCHAR,      -- nguồn của lần ứng tuyển này
    status            VARCHAR,      -- ĐANG Ở ĐÂU — xem CANDIDATE_STATUS_CHOICES
    final_status      VARCHAR,      -- KẾT CỤC RA SAO — xem FINAL_STATUS_CHOICES
    phone_screen_date DATE,
    applied_at        DATETIME,
    status_changed_at DATETIME,
    closed_at         DATETIME,
    note              TEXT,         -- nhận xét của HR về đơn này
    created_at        DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at        DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- BUỔI PHỎNG VẤN TỪNG VÒNG. Một dòng = một vòng của một đơn → thêm vòng 4,
-- vòng 5 không phải sửa cấu trúc. Nhận xét của từng người phỏng vấn nằm ở
-- interview_feedbacks.
--
-- Hai cột `note` (nhận xét của HR) và `summary` (tổng hợp hội đồng) ĐÃ BỎ ở lượt
-- "0002_drop_interview_notes". Chúng vẫn nằm trong lượt 0001 vì lượt cũ không
-- được sửa (máy khác đã chạy qua rồi) — máy mới tạo ra rồi 0002 xóa đi ngay.
CREATE TABLE IF NOT EXISTS interviews (
    interview_id     INTEGER PRIMARY KEY AUTOINCREMENT,
    application_id   INT,           -- → applications.application_id
    candidate_id     INT,           -- chép lại để truy vấn nhanh theo ứng viên
    round            INT,           -- 1 · 2 · 3…
    interview_date   DATETIME,
    duration_minutes INT,
    mode             VARCHAR,       -- xem INTERVIEW_MODE_CHOICES
    location         VARCHAR,       -- phòng họp hoặc link online
    overall_score    VARCHAR,       -- kết luận chung — xem INTERVIEW_SCORE_CHOICES
    note             TEXT,          -- (bỏ ở lượt 0002)
    summary          TEXT,          -- (bỏ ở lượt 0002)
    next_step        VARCHAR,
    status           VARCHAR,       -- xem INTERVIEW_STATUS_CHOICES
    mail_activity_id INT,           -- → candidate_activities (thư mời đã gửi)
    created_at       DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at       DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- NHẬN XÉT CỦA TỪNG NGƯỜI PHỎNG VẤN. Người phỏng vấn là NHÂN VIÊN công ty →
-- trỏ thẳng vào employees, không có danh mục riêng. `interviewer_name` dùng cho
-- khách mời ngoài công ty hoặc người chưa có trong employees.
CREATE TABLE IF NOT EXISTS interview_feedbacks (
    feedback_id      INTEGER PRIMARY KEY AUTOINCREMENT,
    interview_id     INT,           -- → interviews.interview_id
    employee_id      INT,           -- → employees.employee_id
    interviewer_name VARCHAR,
    role             VARCHAR,       -- xem INTERVIEWER_ROLE_CHOICES
    score            VARCHAR,       -- xem INTERVIEW_SCORE_CHOICES
    rating           DECIMAL,       -- 1–5 nếu muốn chấm định lượng
    feedback         TEXT,
    strengths        TEXT,
    weaknesses       TEXT,
    submitted_at     DATETIME,
    created_at       DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at       DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- LỊCH SỬ AI CHẤM ĐIỂM — CHỈ GHI THÊM, KHÔNG BAO GIỜ GHI ĐÈ.
-- Chấm lại lần thứ ba thì có ba dòng: xem được điểm đổi ra sao, model nào,
-- dựa trên bản CV nào. Điểm "hiện hành" = dòng có evaluated_at lớn nhất.
--   • position_id rỗng    = chấm chung, chưa gắn vị trí.
--   • application_id rỗng = chấm khi TÌM TRONG POOL (ứng viên chưa nộp đơn).
CREATE TABLE IF NOT EXISTS candidate_evaluations (
    evaluation_id  INTEGER PRIMARY KEY AUTOINCREMENT,
    candidate_id   INT,             -- → candidates.candidate_id
    position_id    INT,             -- → positions.position_id
    application_id INT,             -- → applications.application_id
    cv_id          INT,             -- CHẤM DỰA TRÊN BẢN CV NÀO
    source         VARCHAR,         -- xem EVALUATION_SOURCE_CHOICES
    rule_score     DECIMAL,         -- 0–100, tính bằng công thức, KHÔNG gọi AI
    ai_score       DECIMAL,         -- 0–100, AI chấm theo JD đầy đủ
    matched_skills TEXT,            -- JD yêu cầu VÀ ứng viên có
    missing_skills TEXT,            -- JD yêu cầu nhưng ứng viên thiếu
    summary        TEXT,
    strengths      TEXT,
    weaknesses     TEXT,
    model          VARCHAR,
    jd_hash        VARCHAR,         -- JD sửa thì biết điểm đã cũ
    extra_prompt   TEXT,            -- yêu cầu bổ sung người dùng nhập lúc chấm
    evaluated_at   DATETIME,
    created_at     DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- LỊCH SỬ LIÊN HỆ & SỰ KIỆN — trả lời "người này đã từng được liên hệ chưa,
-- khi nào, chuyện gì". Kết quả phỏng vấn KHÔNG nằm ở đây mà ở `interviews`.
CREATE TABLE IF NOT EXISTS candidate_activities (
    activity_id      INTEGER PRIMARY KEY AUTOINCREMENT,
    candidate_id     INT,           -- → candidates.candidate_id
    application_id   INT,           -- rỗng = việc chung, không gắn đơn
    type             VARCHAR,       -- xem ACTIVITY_TYPE_CHOICES
    round            INT,           -- thư mời cho vòng mấy
    occurred_at      DATETIME,
    scheduled_at     DATETIME,      -- giờ hẹn ghi trong thư mời
    subject          VARCHAR,
    content          TEXT,
    mail_template_id INT,           -- → mail_templates
    mail_to          VARCHAR,
    mail_cc          VARCHAR,
    result           VARCHAR,       -- xem ACTIVITY_RESULT_CHOICES
    from_status      VARCHAR,
    to_status        VARCHAR,
    created_at       DATETIME DEFAULT (datetime('now', 'localtime'))
);


-- ═══════════════════════ NHÂN SỰ & ĐÀO TẠO ══════════════════════════════

-- Bảng soi theo file "Master HC file.xlsx" (sheet "Master file"): mỗi cột trong
-- file có một cột tương ứng ở đây, chú thích `-- <Tiêu đề Excel>` ghi rõ cột
-- nguồn. Import khớp THEO TÊN CỘT, không theo thứ tự cột.
--
-- Trạng thái làm việc suy ra từ `termination_date`: có ngày = ĐÃ NGHỈ VIỆC
-- (xem cv_repository.EMPLOYEE_WORK_STATUS_SQL).
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
    emergency_contact_name VARCHAR,      -- Emergency Contact Name (phần tên)
    emergency_contact_phone VARCHAR,     -- Emergency Contact Name (phần số ĐT, tách khi import)
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
    is_interviewer    INT,               -- 1 = có tham gia phỏng vấn ứng viên
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

CREATE TABLE IF NOT EXISTS courses (
    course_id   INTEGER PRIMARY KEY AUTOINCREMENT,
    title       VARCHAR,
    content     TEXT,
    date        DATE,
    location    VARCHAR,
    course_type INT,                     -- 0=inhouse, 1=external, 2=funded
    created_at  DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at  DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- Ghi danh (nhiều–nhiều): một khóa có nhiều nhân viên, một nhân viên học nhiều khóa.
CREATE TABLE IF NOT EXISTS course_employees (
    enrollment_id INTEGER PRIMARY KEY AUTOINCREMENT,
    course_id     INT,                   -- → courses.course_id
    employee_id   INT,                   -- → employees.employee_id
    status        VARCHAR,               -- xem COURSE_STATUS_CHOICES
    note          TEXT,
    created_at    DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at    DATETIME DEFAULT (datetime('now', 'localtime'))
);


-- ══════════════════════════════ CHỈ MỤC ═════════════════════════════════
CREATE INDEX IF NOT EXISTS idx_cand_name   ON candidates(full_name);
CREATE INDEX IF NOT EXISTS idx_cand_email  ON candidates(email);
CREATE INDEX IF NOT EXISTS idx_cand_phone  ON candidates(phone);
CREATE INDEX IF NOT EXISTS idx_cand_pool   ON candidates(pool_status);
CREATE INDEX IF NOT EXISTS idx_cand_years  ON candidates(years_experience);
CREATE INDEX IF NOT EXISTS idx_cand_title  ON candidates(current_title);

CREATE INDEX IF NOT EXISTS idx_cv_candidate     ON candidate_cvs(candidate_id, received_at);
CREATE UNIQUE INDEX IF NOT EXISTS idx_cv_hash   ON candidate_cvs(file_hash);
CREATE INDEX IF NOT EXISTS idx_exp_candidate    ON candidate_experiences(candidate_id, start_date);
CREATE INDEX IF NOT EXISTS idx_exp_cv           ON candidate_experiences(cv_id);
CREATE UNIQUE INDEX IF NOT EXISTS idx_cskill_unique ON candidate_skills(candidate_id, skill_id);
CREATE INDEX IF NOT EXISTS idx_cskill_skill     ON candidate_skills(skill_id);

CREATE INDEX IF NOT EXISTS idx_app_candidate ON applications(candidate_id, applied_at);
CREATE INDEX IF NOT EXISTS idx_app_position  ON applications(position_id, status);
CREATE INDEX IF NOT EXISTS idx_app_status    ON applications(status);

CREATE UNIQUE INDEX IF NOT EXISTS idx_intv_round ON interviews(application_id, round);
CREATE INDEX IF NOT EXISTS idx_intv_candidate    ON interviews(candidate_id, interview_date);
CREATE INDEX IF NOT EXISTS idx_intv_date         ON interviews(interview_date);
CREATE INDEX IF NOT EXISTS idx_fb_interview      ON interview_feedbacks(interview_id);
CREATE INDEX IF NOT EXISTS idx_fb_employee       ON interview_feedbacks(employee_id);

CREATE INDEX IF NOT EXISTS idx_eval_candidate ON candidate_evaluations(candidate_id, evaluated_at);
CREATE INDEX IF NOT EXISTS idx_eval_position  ON candidate_evaluations(position_id, ai_score);
CREATE INDEX IF NOT EXISTS idx_eval_cv        ON candidate_evaluations(cv_id);

CREATE INDEX IF NOT EXISTS idx_act_candidate ON candidate_activities(candidate_id, occurred_at);
CREATE INDEX IF NOT EXISTS idx_act_app       ON candidate_activities(application_id);
CREATE INDEX IF NOT EXISTS idx_act_type      ON candidate_activities(type);

CREATE UNIQUE INDEX IF NOT EXISTS idx_posreq_unique ON position_requirements(position_id, jd_hash);
CREATE INDEX IF NOT EXISTS idx_pos_dept   ON positions(department_id);
CREATE INDEX IF NOT EXISTS idx_pos_status ON positions(status);

CREATE INDEX IF NOT EXISTS idx_emp_name        ON employees(full_name);
CREATE INDEX IF NOT EXISTS idx_emp_code        ON employees(code);
CREATE INDEX IF NOT EXISTS idx_emp_dept        ON employees(department_id);
CREATE INDEX IF NOT EXISTS idx_emp_level       ON employees(level_id);
CREATE INDEX IF NOT EXISTS idx_emp_costcenter  ON employees(cost_center_id);
CREATE INDEX IF NOT EXISTS idx_emp_emptype     ON employees(employee_type_id);
CREATE INDEX IF NOT EXISTS idx_emp_termination ON employees(termination_date);

CREATE UNIQUE INDEX IF NOT EXISTS idx_ce_unique ON course_employees(course_id, employee_id);
CREATE INDEX IF NOT EXISTS idx_ce_course        ON course_employees(course_id);
CREATE INDEX IF NOT EXISTS idx_ce_employee      ON course_employees(employee_id);
CREATE INDEX IF NOT EXISTS idx_emptype_code     ON employee_types(code);
CREATE INDEX IF NOT EXISTS idx_costcenter_code  ON cost_centers(code);
CREATE INDEX IF NOT EXISTS idx_levels_name      ON levels(level_name);
CREATE INDEX IF NOT EXISTS idx_skill_name       ON skills(name);
CREATE INDEX IF NOT EXISTS idx_mail_templates_type ON mail_templates(type);
"""

# Bảng ảo TÌM KIẾM TOÀN VĂN. Tách khỏi SCHEMA_SQL vì FTS5 có thể không có trong
# bản SQLite đi kèm Python — thiếu thì bỏ qua, ứng dụng vẫn chạy (tìm kiếm lùi
# về LIKE). Không dùng `content=` (external content) để khỏi phải nuôi trigger;
# repository tự đồng bộ trong lúc thêm/sửa ứng viên.
FTS_SQL = """
CREATE VIRTUAL TABLE IF NOT EXISTS candidates_fts USING fts5(
    candidate_id UNINDEXED,
    full_name, current_title, skills_text, profile_summary, cv_text
);
"""

# =============================================================================
#  MIGRATIONS — ĐƯỜNG DUY NHẤT để dựng và sửa cấu trúc DB.
#
#  Danh sách các lượt (tên, SQL) xếp theo THỨ TỰ CHẠY. `cv_repository.init_db()`
#  chạy lần lượt và ghi dấu vết vào bảng `app_meta` với khóa "migration:<tên>":
#     • Máy chưa có file .db → chạy TẤT CẢ, bắt đầu từ lượt "0001" tạo bảng.
#     • Máy đã có DB rồi     → chỉ chạy những lượt CHƯA có dấu vết.
#  Nhờ vậy tắt mở lại app chỉ tốn đúng lượt vừa thêm, dữ liệu giữ nguyên.
#
#  THÊM MỘT THAY ĐỔI:
#     1. Thêm một dòng vào CUỐI danh sách, tên đánh số tăng dần cho dễ theo dõi.
#     2. KHÔNG sửa/xóa lượt cũ — máy khác đã chạy qua rồi, sửa cũng không chạy
#        lại, chỉ làm hai máy lệch cấu trúc nhau.
#     3. Đổi tên một lượt = tạo ra lượt mới → nó sẽ chạy lại trên mọi máy.
#
#  SQL của mỗi lượt chạy bằng `executescript` nên viết được NHIỀU CÂU, mỗi câu
#  kết thúc bằng dấu ";". Lượt nào lỗi thì DỪNG HẲN và không được đánh dấu —
#  lần mở app sau sẽ thử lại chính lượt đó.
#
#  Chỉ mục cho cột mới đặt CHUNG lượt với câu ALTER thêm cột đó.
#
#  Ví dụ một lượt thêm cột:
#      ("0002_candidate_linkedin", '''
#          ALTER TABLE candidates ADD COLUMN linkedin VARCHAR;
#          CREATE INDEX IF NOT EXISTS idx_cand_linkedin ON candidates(linkedin);
#      '''),
#
#  Ví dụ một lượt sửa DỮ LIỆU (không phải cấu trúc) — cũng nằm ở đây:
#      ("0003_rename_status_label", '''
#          UPDATE applications SET status = 'Screening'
#          WHERE LOWER(TRIM(COALESCE(status, ''))) = 'phone screen';
#      '''),
# =============================================================================
MIGRATIONS: list[tuple[str, str]] = [
    ("0001_initial_schema", SCHEMA_SQL),
    # Bỏ nhận xét ở CẤP BUỔI phỏng vấn: chỉ giữ nhận xét của TỪNG NGƯỜI phỏng vấn
    # (interview_feedbacks) — đó mới là thứ thực sự nhận được sau mỗi vòng. Nội
    # dung đang có trong hai cột này MẤT theo. `ALTER TABLE … DROP COLUMN` cần
    # SQLite ≥ 3.35 (Python 3.11 trở lên đều kèm bản mới hơn).
    ("0002_drop_interview_notes", '''
        ALTER TABLE interviews DROP COLUMN note;
        ALTER TABLE interviews DROP COLUMN summary;
    '''),
    # Đổi cấu trúc `employees` sang khớp file Excel "Personnel Data" (SSOT).
    # Dữ liệu nhân viên hiện có do người dùng tự xóa tay trước khi mở app với
    # bản này — lượt này chỉ đổi CẤU TRÚC, không đụng dữ liệu.
    #
    # Cột bỏ so với "0001": children_birthdays, education_field,
    # qualification_code, facility_country, facility_town, work_time_type,
    # working_time_pct, changing_dates, local_function (không còn nguồn nào —
    # cả Excel HC lẫn "DLVN Application Form" — cung cấp).
    # Cột thêm: sub_function (Excel "Function (Common)" — team/nhóm con trong
    # phòng ban, KHÔNG dùng để tra departments nữa).
    # Vẫn giữ dù Excel HC không có: nationality, marriage_status,
    # children_count, children_names — "DLVN Application Form" (đơn dự tuyển,
    # AI đọc — app/core/application_form.py) đang ghi các cột này.
    ("0003_rebuild_employees", '''
        ALTER TABLE employees ADD COLUMN sub_function VARCHAR;
        ALTER TABLE employees DROP COLUMN children_birthdays;
        ALTER TABLE employees DROP COLUMN education_field;
        ALTER TABLE employees DROP COLUMN qualification_code;
        ALTER TABLE employees DROP COLUMN facility_country;
        ALTER TABLE employees DROP COLUMN facility_town;
        ALTER TABLE employees DROP COLUMN work_time_type;
        ALTER TABLE employees DROP COLUMN working_time_pct;
        ALTER TABLE employees DROP COLUMN changing_dates;
        ALTER TABLE employees DROP COLUMN local_function;
    '''),
    # Bỏ `address` (Excel "Street (address)"): cột này nằm lạc ở tận cùng file
    # Excel, sau mấy ô header trống, và HR xác nhận không dùng tới — cột hợp lệ
    # cuối cùng của file là "Birthday". Địa chỉ vẫn còn đủ ở `permanent_address`
    # / `temporary_address` (+ `city`, `country`). Dữ liệu trong cột này MẤT theo.
    ("0004_drop_employees_address", '''
        ALTER TABLE employees DROP COLUMN address;
    '''),
    # Bỏ 5 cột "số liệu file Excel tự tính": chúng là CÔNG THỨC bên file (Age =
    # YEAR(NOW())-YEAR(ngày sinh), Year of service = (TODAY()-ngày vào)/365, hai
    # cột phân khoảng suy ra từ hai số đó) nên lưu lại chỉ là ảnh chụp lúc
    # import, càng để lâu càng sai. Bốn cột được tính động ngay trong câu truy
    # vấn (`EMPLOYEE_COMPUTED_SQL`) — giống cách `work_status` suy từ
    # termination_date. Riêng `birth_year` (file: "Year of birthday (year)") BỎ
    # HẲN, không tính lại: công thức của nó trong file trùng hệt cột Age nên đó
    # là cột dư, tiêu đề chỉ gây hiểu nhầm.
    ("0005_drop_employees_excel_figures", '''
        ALTER TABLE employees DROP COLUMN years_of_service;
        ALTER TABLE employees DROP COLUMN length_of_service;
        ALTER TABLE employees DROP COLUMN birth_year;
        ALTER TABLE employees DROP COLUMN age;
        ALTER TABLE employees DROP COLUMN age_range;
    '''),
    # Ba danh mục cuối cùng của sheet "Code" trong file HC còn nằm ngoài app —
    # nhân viên nhập tay nên sai chính tả không ai chặn. Đưa hết vào DB:
    #
    #   • `functions` — Excel "Function (Common)": nhóm nhỏ BÊN TRONG phòng ban
    #     (Production có PD, R&D có HHS · MOB · SOL…), nên gắn thẳng vào
    #     departments thay vì làm một danh mục phẳng nữa. CRUD nằm ngay màn hình
    #     Departments (mỗi phòng ban một ô nhập nhiều dòng).
    #
    #   • `code_lists` — gom NHIỀU danh mục MỘT CỘT vào chung một bảng, phân
    #     biệt bằng `type` (xem CODE_LIST_TYPES). Ba thứ đầu tiên (Qualification ·
    #     Qualification (VN) · ID issued place) chỉ là một danh sách chuỗi; tách
    #     mỗi thứ một bảng + một màn hình là nhân bản y hệt nhau ba lần. Danh mục
    #     một cột về sau chỉ cần thêm một `type`, không phải thêm bảng.
    #
    #   • `employees.qualification_vn` — Excel "Qualification (Việt Nam)". Trước
    #     đây bỏ qua khi import vì tưởng là bản dịch của "Qualification", nhưng
    #     sheet "Code" cho thấy hai cột là hai danh mục RIÊNG (8 giá trị so với
    #     7, cách chia bậc học khác nhau) nên lưu lại đúng như file.
    ("0006_functions_and_code_lists", '''
        CREATE TABLE IF NOT EXISTS functions (
            function_id   INTEGER PRIMARY KEY AUTOINCREMENT,
            department_id INT,             -- → departments.department_id
            function_name VARCHAR,
            description   TEXT,
            created_at    DATETIME DEFAULT (datetime('now', 'localtime')),
            updated_at    DATETIME DEFAULT (datetime('now', 'localtime'))
        );
        CREATE INDEX IF NOT EXISTS idx_functions_dept ON functions(department_id);

        CREATE TABLE IF NOT EXISTS code_lists (
            code_list_id INTEGER PRIMARY KEY AUTOINCREMENT,
            type         VARCHAR,          -- nhóm giá trị — xem CODE_LIST_TYPES
            value        VARCHAR,          -- giá trị, ghi y hệt file Excel
            sort_order   INT,              -- nhỏ hơn hiện trước
            description  TEXT,
            created_at   DATETIME DEFAULT (datetime('now', 'localtime')),
            updated_at   DATETIME DEFAULT (datetime('now', 'localtime'))
        );
        CREATE INDEX IF NOT EXISTS idx_code_lists_type ON code_lists(type);

        ALTER TABLE employees ADD COLUMN qualification_vn VARCHAR;
    '''),
    # HÀNG CHỜ mail chúc mừng sinh nhật. Trước đây lịch gửi do Outlook giữ
    # (`DeferredDeliveryTime` — mail nằm Outbox tới đúng ngày); cách đó chỉ chạy
    # được khi tài khoản gửi ĐÃ ĐĂNG NHẬP trong Outlook. Gửi từ hộp thư dùng
    # chung (chỉ được share, không login) thì Outlook bỏ qua giờ hẹn và gửi
    # ngay, nên cả tháng nhận mail cùng lúc. Giờ app tự giữ lịch: xếp hàng vào
    # bảng này, mỗi lần mở app gửi những dòng đã tới ngày.
    #
    # Bảng chỉ giữ KHÓA + TRẠNG THÁI, không chụp lại email/tên/đường dẫn thiệp:
    # mọi thứ đó tra lại từ `employees` + thư mục thiệp lúc gửi, nên không bao
    # giờ có chuyện dữ liệu chụp lúc xếp hàng bị cũ (thiệp bị xóa, đổi mail công
    # ty, nhân viên nghỉ việc sau khi xếp hàng).
    #
    # `due_date` lưu ISO 'yyyy-mm-dd' — CỐ TÌNH khác `employees.date_of_birth`
    # ('dd/mm/yyyy'): cột này để so sánh với hôm nay ngay trong SQL, chuỗi ISO
    # so sánh là ra đúng thứ tự ngày còn 'dd/mm/yyyy' thì không. Hàng chờ nối
    # với `employees` qua `employee_id`, không bao giờ nối theo ngày.
    #
    # Unique (employee_id, due_date): mỗi người mỗi sinh nhật đúng MỘT mail.
    # `birthday_mail.enqueue()` vẫn tra trước rồi mới ghi nên bình thường không
    # đụng tới index; index là lưới an toàn cho trường hợp mở app hai lần cùng lúc.
    ("0007_birthday_email_queue", '''
        CREATE TABLE IF NOT EXISTS birthday_emails (
            birthday_email_id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INT,        -- → employees.employee_id
            due_date    DATE,       -- 'yyyy-mm-dd' (ISO — xem ghi chú ở migration)
            status      VARCHAR,    -- xem BIRTHDAY_EMAIL_STATUS_CHOICES
            attempts    INT DEFAULT 0,
            last_error  TEXT,
            sent_at     DATETIME,
            created_at  DATETIME DEFAULT (datetime('now', 'localtime')),
            updated_at  DATETIME DEFAULT (datetime('now', 'localtime'))
        );
        CREATE UNIQUE INDEX IF NOT EXISTS idx_bd_emails_slot
            ON birthday_emails(employee_id, due_date);
        CREATE INDEX IF NOT EXISTS idx_bd_emails_due
            ON birthday_emails(status, due_date);
    '''),
]

# =============================================================================
#  DANH MỤC MỘT CỘT (bảng `code_lists`)
#
#  Mỗi `type` là MỘT danh sách giá trị người dùng tự bảo trì ở Master Data →
#  "Code lists". Khác `CANDIDATE_STATUS_CHOICES` & các hằng ở cuối file: những
#  hằng đó là giá trị app dựa vào để chạy logic nên phải cố định trong code,
#  còn ba danh mục này thuần dữ liệu HR — sửa là chuyện thường ngày, để trong
#  DB thì thêm/bớt không cần dựng lại app.
#
#  Hằng dưới đây chỉ liệt kê CÁC LOẠI (giá trị nằm trong DB). Thêm một danh mục
#  một cột về sau = thêm một hằng `CODE_TYPE_*` + một dòng ở `CODE_LIST_TYPES`,
#  không phải thêm bảng.
#
#  Khai ở ĐÂY (trước SEED_DATA) vì khối seed `code_lists` dùng tới.
# =============================================================================
CODE_TYPE_QUALIFICATION = "Qualification"            # Excel: "Qualification"
CODE_TYPE_QUALIFICATION_VN = "Qualification (VN)"    # Excel: "Qualification (Việt Nam)"
CODE_TYPE_ID_ISSUED_PLACE = "ID issued place"        # Excel: "Issued Place" (CCCD nơi cấp)

CODE_LIST_TYPES = [
    CODE_TYPE_QUALIFICATION,
    CODE_TYPE_QUALIFICATION_VN,
    CODE_TYPE_ID_ISSUED_PLACE,
]


# =============================================================================
#  SEED — DỮ LIỆU KHỞI TẠO cho các bảng danh mục (nguồn: file Code.xlsx).
#
#  `cv_repository._seed_master_data()` gọi trong init_db():
#    • Mỗi khối chỉ nạp MỘT LẦN cho mỗi file .db — dấu vết ở app_meta với khóa
#      "seed:<bảng>:v<version>". Xóa dòng đó nếu muốn nạp lại.
#    • Trong một lần nạp, dòng đã có sẵn (trùng theo cột ở `match`) bị bỏ qua
#      → không tạo bản ghi trùng và KHÔNG đụng dữ liệu người dùng tự nhập.
#    • Bổ sung danh mục về sau: thêm dòng vào `rows` rồi TĂNG `version`.
#    • Khối có `lookup` = {cột: (bảng, cột khớp, cột id)} thì ô tương ứng trong
#      `rows` ghi bằng TÊN (dễ đọc, dễ sửa) và seeder tra sang id lúc nạp — id
#      tự tăng nên không thể viết thẳng vào đây. Bảng được tra phải nằm TRƯỚC
#      trong SEED_DATA (dict giữ nguyên thứ tự khai báo) để nó đã có dữ liệu.
# =============================================================================
SEED_DATA: dict[str, dict] = {
    "departments": {
        "version": 1,
        "columns": ("department_name", "short_name"),
        "match": ("department_name",),
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
    # Nhóm nhỏ bên trong phòng ban (Excel "Function (Common)"). Cột thứ hai ghi
    # TÊN phòng ban, seeder tra sang department_id (xem `lookup`).
    #
    # Phòng ban của từng function suy ra từ chính file HC: ba sheet dữ liệu cũ
    # ("Termination", "trước update", "Truoc upda Jobtitle 24.Nov.20", ~1.700
    # dòng) ghép sẵn cặp "Business Unit (Department)" × "Function (Common)" →
    # 25/33 function có bằng chứng thực tế, lấy phòng ban áp đảo. `GCOM` khớp
    # short_name của Global COM. Bảy function còn lại (Test Automation Framework
    # · DC Testing · IA Testing · SOL Test · Annotation · Hardware · Project
    # Coordinator) không xuất hiện ở dòng dữ liệu nào — xếp vào R&D cho cùng họ
    # với các nhóm test đã có bằng chứng thuộc R&D, sửa lại được ở màn hình
    # Departments nếu chưa đúng.
    "functions": {
        "version": 1,
        "columns": ("function_name", "department_id"),
        "match": ("function_name",),
        "lookup": {"department_id": ("departments", "department_name", "department_id")},
        "rows": [
            ("FC",    "Facilities Control"),
            ("FIN",   "Finance"),
            ("BOM",   "Global Operations"),
            ("PD",    "Production"),
            ("ME",    "Manufacturing Engineering"),
            ("PL",    "Production Planning"),
            ("LG",    "Logistics"),
            ("QA",    "Quality Assurance"),
            ("HRA",   "Human Resource & Administration"),
            ("GP",    "Global Planning"),
            ("GL",    "Global Logistics"),
            ("COM",   "Local COM"),
            ("GCOM",  "Global COM"),
            ("MM",    "Materials Management"),
            ("GMM",   "Materials Management"),
            ("IT",    "Information Technology"),
            ("GPr",   "Global Procurement"),
            ("TS",    "Sales SEA"),
            # R&D — các nhóm sản phẩm & nhóm kiểm thử.
            ("R&D",                       "Research & Development"),
            ("HHS",                       "Research & Development"),
            ("MOB",                       "Research & Development"),
            ("FRS",                       "Research & Development"),
            ("SIS",                       "Research & Development"),
            ("SOL",                       "Research & Development"),
            ("Global Testing",            "Research & Development"),
            ("Data capture Testing",      "Research & Development"),
            ("Test Automation Framework", "Research & Development"),
            ("DC Testing",                "Research & Development"),
            ("IA Testing",                "Research & Development"),
            ("SOL Test",                  "Research & Development"),
            ("Annotation",                "Research & Development"),
            ("Hardware",                  "Research & Development"),
            ("Project Coordinator",       "Research & Development"),
        ],
    },
    # WC/WCA = White Collar · IBC/IBCA (Indirect) & DBC/DBCA (Direct) = Blue Collar.
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
    "cost_centers": {
        "version": 1,
        "columns": ("code", "group_function"),
        "match": ("code",),
        "rows": [
            ("VN1001", "VNPlant"), ("VN1002", "VNPlant"), ("VN1003", "VNPlant"),
            ("VN1004", "VNPlant"), ("VN1005", "VNPlant"), ("VN1006", "VNPlant"),
            ("VN1007", "VNPlant"), ("VN1008", "VNPlant"), ("VN1011", "VNPlant"),
            ("VN1012", "VNPlant"), ("VN1021", "VNPlant"), ("VN1023", "VNPlant"),
            ("VN1024", "VNPlant"), ("VN1031", "Corporate"), ("VN1032", "Corporate"),
            ("VN1033", "VNPlant"), ("VN1035", "Corporate"), ("VN1041", "Corporate"),
            ("VN1042", "VNPlant"), ("VN1051", "Corporate"), ("VN1052", "Corporate"),
            ("VN1054", "VNPlant"), ("VN1071", "Corporate"), ("VN1072", "Corporate"),
            ("VN3000", "R&D"), ("VN3001", "R&D"), ("VN3007", "Corporate"),
            ("VN3017", "R&D"), ("VN3024", "R&D"), ("VN3034", "R&D"),
            ("VN3041", "R&D"), ("VN3051", "R&D"), ("VN3061", "R&D"),
            ("VN4112", "Corporate"), ("VN4120", "Corporate"), ("VN4211", "Corporate"),
            ("VN6021", "Corporate"), ("VN7010", "Corporate"), ("VN7031", "Corporate"),
            ("VN7032", "Corporate"), ("VN7040", "Corporate"), ("VN7043", "Corporate"),
        ],
    },
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
    # Các danh mục MỘT CỘT, phân biệt bằng `type` (xem CODE_LIST_TYPES).
    # `sort_order` giữ đúng thứ tự các giá trị trong sheet "Code" để dropdown
    # trên app đọc ra giống hệt dropdown HR đang dùng bên Excel.
    "code_lists": {
        "version": 1,
        "columns": ("type", "value", "sort_order"),
        "match": ("type", "value"),
        "rows": [
            (CODE_TYPE_QUALIFICATION, "University",           1),
            (CODE_TYPE_QUALIFICATION, "Master",               2),
            (CODE_TYPE_QUALIFICATION, "High School",          3),
            (CODE_TYPE_QUALIFICATION, "Vocational Training",  4),
            (CODE_TYPE_QUALIFICATION, "College",              5),
            (CODE_TYPE_QUALIFICATION, "Secondary school",     6),
            (CODE_TYPE_QUALIFICATION, "PhD",                  7),

            (CODE_TYPE_QUALIFICATION_VN, "Tiến sĩ",    1),
            (CODE_TYPE_QUALIFICATION_VN, "Thạc sĩ",    2),
            (CODE_TYPE_QUALIFICATION_VN, "Đại Học",    3),
            (CODE_TYPE_QUALIFICATION_VN, "Cao đẳng",   4),
            (CODE_TYPE_QUALIFICATION_VN, "Trung cấp",  5),
            (CODE_TYPE_QUALIFICATION_VN, "THPT",       6),
            (CODE_TYPE_QUALIFICATION_VN, "THCS",       7),
            (CODE_TYPE_QUALIFICATION_VN, "Khác",       8),

            (CODE_TYPE_ID_ISSUED_PLACE,
             "Cục Cảnh sát quản lý hành chính về trật tự xã hội.", 1),
            (CODE_TYPE_ID_ISSUED_PLACE,
             "Cục trưởng cục cảnh sát ĐKQL cư trú và DLQG về dân cư", 2),
            (CODE_TYPE_ID_ISSUED_PLACE, "Bộ công an",  3),
            (CODE_TYPE_ID_ISSUED_PLACE, "Cục cản sát", 4),
        ],
    },
}


# =============================================================================
#  GIÁ TRỊ CỐ ĐỊNH (các ô "chọn 1 trong nhiều" trên giao diện)
# =============================================================================

# ── Đơn ứng tuyển: ĐANG Ở ĐÂU (applications.status) ──────────────────────
# Xếp theo đúng luồng tuyển dụng: chỉ số trong list = thứ tự giai đoạn, dùng khi
# cần sắp xếp/so sánh tiến độ (xem candidate_status_order). Ba nhãn "Not Proceed",
# "Rejected Offer", "Fail Probation Period" là nhánh DỪNG.
CANDIDATE_STATUS_CHOICES = [
    "New Application",
    "Screening",
    "Short List",
    "Not Proceed",
    "Technical Test",
    "First Interview",
    "Second Interview",
    "Third Interview",
    "Offer Approval",
    "Ready To Hire",
    "Rejected Offer",
    "Fail Probation Period",
]

# Trạng thái mặc định của đơn vừa tạo (nhập tay hoặc quét CV bằng AI).
CANDIDATE_STATUS_DEFAULT = CANDIDATE_STATUS_CHOICES[0]

# Bước kế tiếp trong luồng — gợi ý sẵn khi đổi trạng thái hàng loạt. Nhãn KHÔNG
# có mặt ở đây là điểm dừng của luồng, không còn bước kế tiếp mặc định.
CANDIDATE_STATUS_NEXT = {
    "New Application":  "Screening",
    "Screening":        "Short List",
    "Short List":       "Technical Test",
    "Technical Test":   "First Interview",
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

    Đơn chưa có trạng thái → bắt đầu từ đầu luồng. Trả về chuỗi rỗng khi trạng
    thái là điểm dừng hoặc là nhãn lạ (không có bước kế tiếp).
    """
    target = (status or "").strip().lower()
    if not target:
        return CANDIDATE_STATUS_DEFAULT
    for label, nxt in CANDIDATE_STATUS_NEXT.items():
        if label.lower() == target:
            return nxt
    return ""


# ── Đơn ứng tuyển: KẾT CỤC RA SAO (applications.final_status) ────────────
# Khác `status` ở chỗ status là "đang ở đâu", final_status là "kết cục ra sao".
FINAL_STATUS_CHOICES = [
    "Ongoing",
    "Pass",
    "Fail",
    "Considering",
    "Withdraw",
    "Could not contact",
    "Not proceed",
]
FINAL_STATUS_DEFAULT = "Ongoing"

# Đơn này từ đâu ra.
APPLICATION_ORIGIN_CHOICES = ["Applied", "Pool search", "Referral"]

# NƠI CUNG CẤP CV: sàn tuyển dụng, headhunt, người giới thiệu… — nạp sẵn nhưng
# người dùng gõ thêm được. Thêm sàn mới thì sửa thẳng list này.
CANDIDATE_SOURCE_CHOICES = [
    "Itviec", "VietnamWorks", "LinkedIn", "TopCV", "Referral", "HH- PSK",
    "HH- Adecco", "University", "Internal Sourced",
]

# Tool "AI CV Scan" đóng dấu này vào cột `source` khi quét cả thư mục CV. Nó chỉ
# nói hồ sơ vào app bằng ĐƯỜNG NÀO, chứ chưa biết CV lấy từ sàn nào — nên KHÔNG
# phải nguồn thật và không nằm trong CANDIDATE_SOURCE_CHOICES.
CANDIDATE_SOURCE_AUTO = "AI CV Scan"

# ── Ứng viên trong pool (candidates.pool_status) ─────────────────────────
# Chỉ `Active` mới lọt vào kết quả tìm kiếm mặc định. `Do Not Contact` là hàng
# rào CỨNG: chặn ngay ở bước gửi mail, không chỉ ẩn khỏi danh sách.
POOL_STATUS_CHOICES = ["Active", "Hired", "Do Not Contact", "Inactive"]
POOL_STATUS_DEFAULT = "Active"

# Nguyện vọng: sẵn sàng chuyển vùng làm việc.
RELOCATE_CHOICES = ["Yes", "No", "Negotiable"]

# Hồ sơ cũ hơn ngần này tháng (tính từ experience_as_of) bị gắn nhãn
# "Stale profile" trên màn hình tìm kiếm để nhắc xin CV mới.
STALE_PROFILE_MONTHS = 18

# ── Vị trí tuyển dụng ────────────────────────────────────────────────────
POSITION_STATUS_CHOICES = ["Open", "Paused", "Closed"]

# ── Mẫu mail ─────────────────────────────────────────────────────────────
# Loại mẫu chỉ để phân nhóm/lọc cho dễ tìm, KHÔNG ràng buộc: vị trí vẫn chọn
# được bất kỳ mẫu nào cho mỗi vòng.
#
# Riêng "Application Thank You" có ý nghĩa với luồng gửi mail: mẫu loại này
# KHÔNG gắn vào vị trí mà chọn thẳng lúc gửi, và gửi dạng MAIL THƯỜNG (không
# giờ giấc, không phải thư mời họp).
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

# ── Phỏng vấn ────────────────────────────────────────────────────────────
# 3 VÒNG PHỎNG VẤN của một vị trí: (cột trong positions, nhãn hiển thị, trạng
# thái đơn gợi ý sau khi gửi thư mời vòng đó).
INTERVIEW_ROUNDS = [
    ("mail_template_r1_id", "Interview Round 1", "First Interview"),
    ("mail_template_r2_id", "Interview Round 2", "Second Interview"),
    ("mail_template_r3_id", "Interview Round 3", "Third Interview"),
]

# Trạng thái HIỆN TẠI của đơn → vòng phỏng vấn gợi ý sẵn khi bấm Send email.
# Đơn vừa lọt short list / vừa xong bài test thì mời vòng 1; mời xong vòng 1
# trạng thái thành "First Interview" nên lần sau gợi ý vòng 2…
INTERVIEW_ROUND_BY_STATUS = {
    "Short List":       0,
    "Technical Test":   0,
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


INTERVIEW_MODE_CHOICES = ["Onsite", "Online", "Phone"]
INTERVIEW_STATUS_CHOICES = ["Scheduled", "Completed", "Cancelled", "No show"]
# Điểm của một vòng (interviews.overall_score) và của từng người phỏng vấn
# (interview_feedbacks.score).
INTERVIEW_SCORE_CHOICES = ["Pass", "Fail", "Consideration"]
INTERVIEWER_ROLE_CHOICES = ["Hiring Manager", "Technical", "HR", "Observer"]

# ── Lịch sử liên hệ ──────────────────────────────────────────────────────
ACTIVITY_TYPE_CHOICES = ["Email", "Call", "Status change", "Note"]
ACTIVITY_RESULT_CHOICES = ["Passed", "Failed", "No show", "Pending"]

# ── AI chấm điểm ─────────────────────────────────────────────────────────
EVALUATION_SOURCE_CHOICES = ["CV scan", "Pool search", "Re-rank"]

# Trọng số công thức `rule_score` (tổng = 100) — lọc thô bằng SQL/Python để chọn
# nhóm nhỏ trước khi gọi AI xếp hạng lại.
RULE_SCORE_WEIGHTS = {
    "must_have":    45,   # tỉ lệ khớp kỹ năng bắt buộc
    "nice_to_have": 15,   # tỉ lệ khớp kỹ năng ưu tiên
    "experience":   15,   # mức đáp ứng số năm kinh nghiệm (tính đến hôm nay)
    "education":    10,   # mức đáp ứng học vấn / chuyên ngành
    "title":        10,   # tương đồng chức danh
    "freshness":     5,   # độ mới của hồ sơ
}

# ── Kỹ năng ──────────────────────────────────────────────────────────────
SKILL_CATEGORY_CHOICES = ["Technical", "Tool", "Soft", "Language", "Certificate"]
SKILL_LEVEL_CHOICES = ["Basic", "Intermediate", "Advanced", "Expert"]
SKILL_SOURCE_CHOICES = ["CV", "Manual"]

# ── Nhân viên ────────────────────────────────────────────────────────────
GENDER_CHOICES = ["Male", "Female", "Other"]

# Nhãn trạng thái làm việc để hiển thị/lọc. Đây là giá trị SUY RA từ
# `employees.termination_date` (rỗng = Working, có ngày = Resigned).
EMPLOYEE_STATUS_CHOICES = ["Working", "Resigned"]

CONTRACT_PERMANENCY_CHOICES = ["Permanent", "Temporary", "Fixed Term", "Expat"]
DIRECT_INDIRECT_CHOICES = ["Direct", "Indirect"]
YES_NO_CHOICES = ["Y", "N"]                     # Marriage status (Yes)
MARITAL_STATUS_CHOICES = ["Single", "Married", "Divorced", "Widowed"]

COLLAR_CHOICES = ["Blue Collar", "White Collar"]
GROUP_FUNCTION_CHOICES = ["VNPlant", "Corporate", "R&D"]

# ── Mail chúc mừng sinh nhật ─────────────────────────────────────────────
# Trạng thái một dòng trong hàng chờ `birthday_emails`. Nhãn hiện thẳng trên
# bảng cho người dùng xem nên viết tiếng Anh.
#
# Vòng đời bình thường:  Pending → Sending → Sent
# Các nhánh rẽ:
#   • Failed    — gọi Outlook lỗi, thiếu email, hoặc không tìm thấy thiệp. CÒN
#                 thử lại được: lần mở app sau (còn trong hạn gửi bù) hoặc bấm
#                 "Send now" sẽ chạy lại.
#   • Missed    — quá hạn gửi bù (xem setting `birthday_catchup_days`). Không tự
#                 gửi nữa, chỉ để người dùng nhìn thấy mà xử lý tay.
#   • Cancelled — nhân viên đã nghỉ việc trước khi tới ngày sinh nhật.
#   • Sending   — đã giành được dòng, đang gọi Outlook. Dòng kẹt ở đây (app tắt
#                 giữa lúc gửi) được `reset_stale_sending()` đưa về Pending.
BIRTHDAY_EMAIL_PENDING = "Pending"
BIRTHDAY_EMAIL_SENDING = "Sending"
BIRTHDAY_EMAIL_SENT = "Sent"
BIRTHDAY_EMAIL_FAILED = "Failed"
BIRTHDAY_EMAIL_MISSED = "Missed"
BIRTHDAY_EMAIL_CANCELLED = "Cancelled"

BIRTHDAY_EMAIL_STATUS_CHOICES = [
    BIRTHDAY_EMAIL_PENDING,
    BIRTHDAY_EMAIL_SENDING,
    BIRTHDAY_EMAIL_SENT,
    BIRTHDAY_EMAIL_FAILED,
    BIRTHDAY_EMAIL_MISSED,
    BIRTHDAY_EMAIL_CANCELLED,
]

# ── Đào tạo ──────────────────────────────────────────────────────────────
# Loại khóa học — lưu dạng INT trong cột courses.course_type (chỉ số = giá trị lưu).
COURSE_TYPE_CHOICES = ["inhouse", "external", "funded"]  # 0, 1, 2
COURSE_STATUS_CHOICES = ["Not started", "Completed"]
