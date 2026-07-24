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

  departments ──1:N──> positions ──1:N──> job_descriptions (+ jd_file_path)
                          │
                          │ 1:N
                          ▼
                      candidates (+ cv_file_path)

  • departments (phòng ban)  1—N  positions (vị trí)
  • positions   (vị trí)     1—N  job_descriptions (JD) và  1—N  candidates
  • Đường dẫn file lưu thẳng: candidates.cv_file_path, job_descriptions.jd_file_path
    (không còn bảng document_files — file thực tế đã nằm sẵn trên máy).
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
    manager_name    VARCHAR,
    description     TEXT,
    created_at      DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at      DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- ───────────────────────── MASTER: VỊ TRÍ TUYỂN DỤNG ─────────────────
CREATE TABLE IF NOT EXISTS positions (
    position_id    INTEGER PRIMARY KEY AUTOINCREMENT,
    department_id  INT,                 -- tham chiếu mềm → departments.department_id
    position_code  VARCHAR,
    position_title VARCHAR,
    level          VARCHAR,             -- cấp bậc (Junior/Senior/Lead…)
    headcount      INT,                 -- số lượng cần tuyển
    status         VARCHAR,             -- Đang tuyển / Tạm dừng / Đã đóng
    created_at     DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at     DATETIME DEFAULT (datetime('now', 'localtime'))
);

-- ───────────────────────── MASTER: MÔ TẢ CÔNG VIỆC (JD) ──────────────
CREATE TABLE IF NOT EXISTS job_descriptions (
    jd_id        INTEGER PRIMARY KEY AUTOINCREMENT,
    position_id  INT,                   -- tham chiếu mềm → positions.position_id
    jd_title     VARCHAR,
    jd_file_path VARCHAR,                 -- đường dẫn file JD trên máy
    created_at   DATETIME DEFAULT (datetime('now', 'localtime')),
    updated_at   DATETIME DEFAULT (datetime('now', 'localtime'))
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

-- ───────────────────────── CHỈ MỤC (tăng tốc tìm kiếm) ───────────────
CREATE INDEX IF NOT EXISTS idx_candidates_name   ON candidates(full_name);
CREATE INDEX IF NOT EXISTS idx_candidates_email  ON candidates(email);
CREATE INDEX IF NOT EXISTS idx_candidates_phone  ON candidates(phone);
CREATE INDEX IF NOT EXISTS idx_candidates_pos    ON candidates(position_id);
CREATE INDEX IF NOT EXISTS idx_candidates_status ON candidates(status);
CREATE INDEX IF NOT EXISTS idx_positions_dept    ON positions(department_id);
CREATE INDEX IF NOT EXISTS idx_jd_pos            ON job_descriptions(position_id);
"""

# =============================================================================
#  MIGRATIONS — thêm cột cho bảng ĐÃ tồn tại (chạy an toàn nhiều lần).
#  Ví dụ:  "ALTER TABLE candidates ADD COLUMN linkedin VARCHAR",
# =============================================================================
MIGRATIONS: list[str] = [
    # Bỏ bảng document_files → lưu đường dẫn file thẳng vào candidates & jd.
    "ALTER TABLE candidates ADD COLUMN cv_file_path VARCHAR",
    "ALTER TABLE job_descriptions ADD COLUMN jd_file_path VARCHAR",
    # Dấu thời gian tạo / cập nhật cho MỌI bảng. Lưu ý: SQLite không cho dùng
    # default động (datetime('now')) trong ALTER TABLE → cột thêm cho bảng CŨ sẽ
    # NULL; bản ghi TẠO MỚI sau đó được điền qua init_db()._backfill_timestamps
    # và logic ghi (INSERT dùng DEFAULT của schema, UPDATE tự set updated_at).
    "ALTER TABLE departments ADD COLUMN created_at DATETIME",
    "ALTER TABLE departments ADD COLUMN updated_at DATETIME",
    "ALTER TABLE positions ADD COLUMN created_at DATETIME",
    "ALTER TABLE positions ADD COLUMN updated_at DATETIME",
    "ALTER TABLE job_descriptions ADD COLUMN created_at DATETIME",
    "ALTER TABLE job_descriptions ADD COLUMN updated_at DATETIME",
    "ALTER TABLE candidates ADD COLUMN created_at DATETIME",
    "ALTER TABLE candidates ADD COLUMN updated_at DATETIME",
    # Cột 'batch' (đợt/lô quét CV) — chỉ lưu SỐ; thêm cho DB đã tồn tại.
    # LƯU Ý: index cho 'batch' PHẢI tạo Ở ĐÂY (sau ALTER), KHÔNG đặt trong
    # SCHEMA_SQL — vì executescript(SCHEMA_SQL) chạy TRƯỚC migration, DB cũ chưa
    # có cột batch sẽ khiến CREATE INDEX ném "no such column: batch" và hỏng cả
    # init_db (trang không mở được).
    "ALTER TABLE candidates ADD COLUMN batch INT",
    "CREATE INDEX IF NOT EXISTS idx_candidates_batch ON candidates(batch)",
    # Bỏ các cột không dùng nữa (SQLite ≥ 3.35 hỗ trợ DROP COLUMN; DB mới đã
    # không có sẵn các cột này nên câu lệnh sẽ bị bỏ qua an toàn).
    "ALTER TABLE departments DROP COLUMN department_code",
    "ALTER TABLE job_descriptions DROP COLUMN summary",
    "ALTER TABLE job_descriptions DROP COLUMN requirements",
    "ALTER TABLE job_descriptions DROP COLUMN salary_range",
]

# Gợi ý cho các ô chọn ở giao diện (sửa tùy ý).
STATUS_CHOICES = ["Mới", "Đã liên hệ", "Phỏng vấn", "Đạt", "Loại", "Chờ"]
POSITION_STATUS_CHOICES = ["Đang tuyển", "Tạm dừng", "Đã đóng"]

# Danh sách bảng do init_db quản lý — dùng khi cần dựng lại bảng trống lệch schema.
_MANAGED_TABLES = [
    "departments", "positions", "job_descriptions", "candidates",
]
