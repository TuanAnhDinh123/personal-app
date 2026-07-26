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

  departments ──1:N──> positions (+ jd_title, jd_file_path)
                          │
                          │ 1:N
                          ▼
                      candidates (+ cv_file_path)

  • departments (phòng ban)  1—N  positions (vị trí)
  • positions   (vị trí)     1—N  candidates
  • MỖI VỊ TRÍ CHỈ CÓ 1 JD → thông tin JD (tiêu đề + đường dẫn file) nằm THẲNG
    trong bảng positions; KHÔNG còn bảng job_descriptions riêng.
  • Đường dẫn file lưu thẳng: candidates.cv_file_path, positions.jd_file_path
    (không còn bảng document_files — file thực tế đã nằm sẵn trên máy).

  departments ──1:N──> employees
  courses  ──N:M──  employees   (qua bảng trung gian course_employees)
  • Một khóa học (courses) có NHIỀU nhân viên; một nhân viên tham gia NHIỀU
    khóa học. Mỗi dòng trong course_employees = 1 lượt ghi danh (course_id +
    employee_id, unique để không ghi danh trùng).
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
    short_name      VARCHAR,                 -- mã viết tắt (vd FIN, IT, R&D) — dùng khi xuất Excel
    manager_name    VARCHAR,
    description     TEXT,
    created_at      DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at      DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- ────────── MASTER: VỊ TRÍ TUYỂN DỤNG (kèm luôn JD của vị trí đó) ──────
-- Mỗi vị trí chỉ có ĐÚNG 1 mô tả công việc (JD) → 2 cột jd_* nằm ngay đây.
CREATE TABLE IF NOT EXISTS positions (
    position_id    INTEGER PRIMARY KEY AUTOINCREMENT,
    department_id  INT,                 -- tham chiếu mềm → departments.department_id
    position_code  VARCHAR,
    position_title VARCHAR,
    level          VARCHAR,             -- cấp bậc (Junior/Senior/Lead…)
    headcount      INT,                 -- số lượng cần tuyển
    status         VARCHAR,             -- Đang tuyển / Tạm dừng / Đã đóng
    jd_title       VARCHAR,             -- tiêu đề JD (để trống = dùng tên vị trí)
    jd_file_path   VARCHAR,             -- đường dẫn file JD trên máy
    mail_cc        VARCHAR,             -- CC mặc định khi gửi mail mời PV
    mail_subject   VARCHAR,             -- tiêu đề mẫu mail (hỗ trợ {name}{possion}{date}{time})
    mail_body      TEXT,                -- nội dung mẫu mail (HTML, hỗ trợ placeholder trên)
    created_at     DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at     DATETIME DEFAULT (datetime('now', 'localtime'))
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
    status           VARCHAR,           -- Mới / Phỏng vấn / Đạt / Loại…
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
CREATE TABLE IF NOT EXISTS employees (
    employee_id   INTEGER PRIMARY KEY AUTOINCREMENT,
    code          VARCHAR,               -- mã nhân viên nội bộ
    global_code   VARCHAR,               -- mã toàn cầu (Global Code)
    full_name     VARCHAR,               -- họ tên đầy đủ
    surname       VARCHAR,               -- họ
    name          VARCHAR,               -- tên
    middle_name   VARCHAR,               -- tên đệm
    date_of_birth DATE,
    gender        VARCHAR,               -- giới tính
    education     VARCHAR,               -- trình độ học vấn
    phone         VARCHAR,
    email         VARCHAR,
    level         VARCHAR,               -- cấp bậc
    department_id INT,                   -- tham chiếu mềm → departments.department_id
    address       VARCHAR,
    created_at    DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at    DATETIME DEFAULT (datetime('now', 'localtime'))
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
CREATE INDEX IF NOT EXISTS idx_positions_dept    ON positions(department_id);
CREATE INDEX IF NOT EXISTS idx_employees_name    ON employees(full_name);
CREATE INDEX IF NOT EXISTS idx_employees_code     ON employees(code);
CREATE INDEX IF NOT EXISTS idx_employees_dept     ON employees(department_id);
CREATE INDEX IF NOT EXISTS idx_ce_course           ON course_employees(course_id);
CREATE INDEX IF NOT EXISTS idx_ce_employee         ON course_employees(employee_id);
CREATE UNIQUE INDEX IF NOT EXISTS idx_ce_unique    ON course_employees(course_id, employee_id);
"""

# =============================================================================
#  MIGRATIONS — thêm cột cho bảng ĐÃ tồn tại (chạy an toàn nhiều lần).
#  Ví dụ:  "ALTER TABLE candidates ADD COLUMN linkedin VARCHAR",
# =============================================================================
MIGRATIONS: list[str] = [
    # Bỏ bảng document_files → lưu đường dẫn file thẳng vào candidates & positions.
    "ALTER TABLE candidates ADD COLUMN cv_file_path VARCHAR",
    # Dấu thời gian tạo / cập nhật cho MỌI bảng. Lưu ý: SQLite không cho dùng
    # default động (datetime('now')) trong ALTER TABLE → cột thêm cho bảng CŨ sẽ
    # NULL; bản ghi TẠO MỚI sau đó được điền qua init_db()._backfill_timestamps
    # và logic ghi (INSERT dùng DEFAULT của schema, UPDATE tự set updated_at).
    "ALTER TABLE departments ADD COLUMN created_at DATETIME",
    "ALTER TABLE departments ADD COLUMN updated_at DATETIME",
    "ALTER TABLE positions ADD COLUMN created_at DATETIME",
    "ALTER TABLE positions ADD COLUMN updated_at DATETIME",
    "ALTER TABLE candidates ADD COLUMN created_at DATETIME",
    "ALTER TABLE candidates ADD COLUMN updated_at DATETIME",
    # Cột 'batch' (đợt/lô quét CV) — chỉ lưu SỐ; thêm cho DB đã tồn tại.
    # LƯU Ý: index cho 'batch' PHẢI tạo Ở ĐÂY (sau ALTER), KHÔNG đặt trong
    # SCHEMA_SQL — vì executescript(SCHEMA_SQL) chạy TRƯỚC migration, DB cũ chưa
    # có cột batch sẽ khiến CREATE INDEX ném "no such column: batch" và hỏng cả
    # init_db (trang không mở được).
    "ALTER TABLE candidates ADD COLUMN batch INT",
    "CREATE INDEX IF NOT EXISTS idx_candidates_batch ON candidates(batch)",
    # Mẫu mail mời phỏng vấn gắn theo từng VỊ TRÍ (thêm cho DB đã tồn tại).
    "ALTER TABLE positions ADD COLUMN mail_cc VARCHAR",
    "ALTER TABLE positions ADD COLUMN mail_subject VARCHAR",
    "ALTER TABLE positions ADD COLUMN mail_body TEXT",
    # MỖI VỊ TRÍ = 1 JD → thông tin JD nằm luôn trong bảng positions (thêm cột
    # cho DB đã tồn tại). Bảng job_descriptions cũ bị XÓA HẲN, dữ liệu trong đó
    # KHÔNG chuyển sang (nhập lại ở form vị trí) — xem
    # cv_repository._drop_job_descriptions().
    "ALTER TABLE positions ADD COLUMN jd_title VARCHAR",
    "ALTER TABLE positions ADD COLUMN jd_file_path VARCHAR",
    # Bỏ các cột không dùng nữa (SQLite ≥ 3.35 hỗ trợ DROP COLUMN; DB mới đã
    # không có sẵn các cột này nên câu lệnh sẽ bị bỏ qua an toàn).
    "ALTER TABLE departments DROP COLUMN department_code",
    # Đổi tên cột course_employees.result → status (DB cũ). Bảng tạo mới đã có sẵn
    # cột 'status' nên câu lệnh này ném OperationalError ("no such column: result")
    # và được init_db() bỏ qua an toàn.
    "ALTER TABLE course_employees RENAME COLUMN result TO status",
    # Mã viết tắt bộ phận (vd FIN, IT, R&D) — thêm cho DB đã tồn tại.
    "ALTER TABLE departments ADD COLUMN short_name VARCHAR",
]

# Gợi ý cho các ô chọn ở giao diện (sửa tùy ý).
STATUS_CHOICES = ["Mới", "Đã liên hệ", "Phỏng vấn", "Đạt", "Loại", "Chờ"]
POSITION_STATUS_CHOICES = ["Đang tuyển", "Tạm dừng", "Đã đóng"]

# Nhân viên: giới tính & cấp bậc (level) — dùng cho ô lọc + form nhập.
GENDER_CHOICES = ["Nam", "Nữ", "Khác"]
EMPLOYEE_LEVEL_CHOICES = [
    "Intern", "Fresher", "Junior", "Middle", "Senior", "Lead", "Manager",
]

# Loại khóa học — lưu dạng INT trong cột courses.course_type (chỉ số = giá trị lưu).
COURSE_TYPE_CHOICES = ["inhouse", "external", "funded"]  # 0, 1, 2

# Trạng thái học của nhân viên trong 1 khóa (cột course_employees.status).
# "Not started" = chưa học · "Completed" = đã học xong.
COURSE_STATUS_CHOICES = ["Not started", "Completed"]

# Danh sách bảng do init_db quản lý — dùng khi cần dựng lại bảng trống lệch schema.
_MANAGED_TABLES = [
    "departments", "positions", "candidates",
    "employees", "courses", "course_employees",
]
