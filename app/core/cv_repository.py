"""Tầng truy cập dữ liệu (data-access layer) cho tool Quản lý CV ứng viên.

Gói toàn bộ thao tác SQLite ở một chỗ để giao diện chỉ việc gọi hàm. Cấu trúc
bảng định nghĩa trong `app/core/cv_schema.py` (file thiết kế DB).

File .db mặc định:
    %APPDATA%\\PersonalToolbox\\candidates.db   (Windows)
    ~/.config/PersonalToolbox/candidates.db      (Linux/macOS — lúc dev)
"""
import os
import sqlite3

from app.core import cv_schema

# Cột được phép ghi cho từng bảng (chặn khóa lạ lọt vào câu INSERT/UPDATE).
DEPARTMENT_FIELDS = ["department_name", "short_name", "manager_name", "description"]
POSITION_FIELDS = ["department_id", "position_code", "position_title", "level",
                   "headcount", "status", "jd_file_path",
                   "mail_template_r1_id", "mail_template_r2_id",
                   "mail_template_r3_id"]
MAIL_TEMPLATE_FIELDS = ["name", "type", "mail_cc", "mail_subject", "mail_body"]
CANDIDATE_FIELDS = [
    "full_name", "email", "phone", "date_of_birth", "address",
    "position_id", "years_experience", "education", "applied_at", "status",
    "source", "batch", "fit_score", "fit_summary", "strengths", "weaknesses",
    "cv_file_path", "note",
]
# `employees` bám sát "Master HC file.xlsx" → rất nhiều cột. Giữ thứ tự NHÓM
# giống cv_schema.py cho dễ đối chiếu. Trạng thái làm việc không nằm trong danh
# sách này vì được suy ra từ `termination_date`.
EMPLOYEE_FIELDS = [
    # định danh + họ tên
    "code", "global_code",
    "full_name", "surname", "name", "middle_name",
    # thông tin cá nhân
    "date_of_birth", "gender", "place_of_birth", "native_place", "nationality",
    "religion", "marriage_status", "marital_status", "spouse_name", "spouse_dob",
    "children_count", "children_names", "children_birthdays",
    # liên hệ
    "phone", "email", "company_email", "address", "city", "country",
    "permanent_address", "temporary_address",
    "emergency_contact_name", "emergency_contact_relationship",
    # học vấn
    "education", "education_field", "major", "graduation_year", "school_name",
    "qualification", "qualification_code",
    # giấy tờ · ngân hàng · thuế · bảo hiểm
    "id_no", "id_issued_date", "id_issued_place",
    "passport_no", "passport_issued_date",
    "bank_account_no", "bank_address", "tax_code", "dependants",
    "insurance_book_no",
    # tổ chức & công việc
    "department_id", "cost_center_id", "employee_type_id", "level_id",
    "manager_name", "job_title", "current_position", "time_in_position",
    "facility_country", "facility_town", "local_function", "by_group",
    "labor_type", "production_line", "operator_skill", "driving_forklift",
    "working_hours_per_week", "smart_working_eligible", "er_jrf",
    # hợp đồng & thời gian làm việc
    "date_of_employment", "seniority_date", "contract_permanency",
    "work_time_type", "working_time_pct", "direct_indirect", "contract_type",
    "contract_start_date", "contract_end_date", "changing_date",
    "termination_date", "leaving_reason",
    # số liệu file Excel tự tính
    "years_of_service", "length_of_service", "birth_year", "age", "age_range",
    # ghi chú
    "changing_notes", "changing_dates", "updated_changing_date", "note",
]
COURSE_FIELDS = ["title", "content", "date", "location", "course_type"]
COURSE_EMPLOYEE_FIELDS = ["course_id", "employee_id", "status", "note"]
EMPLOYEE_TYPE_FIELDS = ["code", "collar", "description"]
COST_CENTER_FIELDS = ["code", "group_function", "name", "description"]
LEVEL_FIELDS = ["level_name", "sort_order", "description"]

# PK của mỗi bảng (dùng cho update/delete generic).
_PK = {
    "departments": "department_id",
    "positions": "position_id",
    "candidates": "candidate_id",
    "employees": "employee_id",
    "courses": "course_id",
    "course_employees": "enrollment_id",
    "employee_types": "employee_type_id",
    "cost_centers": "cost_center_id",
    "levels": "level_id",
    "mail_templates": "mail_template_id",
}


def _candidate_rows(path) -> int:
    """Đếm số ứng viên trong 1 file DB (an toàn, -1 nếu thiếu/không đọc được).

    Lưu ý: KHÔNG gọi khi file chưa tồn tại một cách vô ý — sqlite3.connect sẽ
    tạo file rỗng. Đã chặn sẵn bằng kiểm tra os.path.exists ở đây.
    """
    if not os.path.exists(path):
        return -1
    try:
        c = sqlite3.connect(path)
        try:
            return c.execute("SELECT COUNT(*) FROM candidates").fetchone()[0]
        finally:
            c.close()
    except sqlite3.Error:
        return -1


def _db_path() -> str:
    base = os.environ.get("APPDATA") or os.path.join(
        os.path.expanduser("~"), ".config")
    folder = os.path.join(base, "PersonalToolbox")
    os.makedirs(folder, exist_ok=True)
    new = os.path.join(folder, "candidates.sqlite")   # tên file mới
    old = os.path.join(folder, "candidates.db")        # tên file cũ

    if os.path.exists(old):
        try:
            if not os.path.exists(new):
                # Chỉ có file cũ → đổi tên sang .sqlite (một lần, giữ dữ liệu).
                os.rename(old, new)
            elif _candidate_rows(new) <= 0 < _candidate_rows(old):
                # Cả hai cùng tồn tại nhưng .sqlite RỖNG còn .db CÓ dữ liệu
                # → thay .sqlite rỗng bằng .db (tránh mất dữ liệu do file shadow).
                os.replace(old, new)
        except OSError:
            # Đổi tên thất bại (thường do app đang mở & khóa file .db).
            # Nếu .db có dữ liệu mà .sqlite chưa có → tạm dùng .db để KHÔNG
            # đọc nhầm file rỗng; lần mở sau (app đã đóng) sẽ đổi tên trót lọt.
            if _candidate_rows(old) > 0 >= _candidate_rows(new):
                return old
    return new


def get_connection() -> sqlite3.Connection:
    """Mở kết nối SQLite (row_factory=Row để truy cập cột theo tên).

    KHÔNG bật PRAGMA foreign_keys — thiết kế cố tình không dùng khóa ngoại.
    """
    conn = sqlite3.connect(_db_path())
    conn.row_factory = sqlite3.Row
    return conn


# Hậu tố đặt cho bảng cũ trong lúc di trú (migrate).
_LEGACY_SUFFIX = "__legacy"

# Đổi tên cột từ schema CŨ (phiên bản đầu) sang schema MỚI khi copy dữ liệu.
_LEGACY_COL_RENAME = {
    "candidates": {
        "id": "candidate_id",
        "dob": "date_of_birth",
        "applied_date": "applied_at",
        "cv_file": "cv_file_path",
    },
    "departments": {
        "id": "department_id",
        "name": "department_name",
    },
}


def init_db() -> None:
    """Tạo bảng nếu chưa có + di trú dữ liệu cũ + chạy migration.

    Gọi mỗi lần mở tool (rẻ). Nếu một bảng đang ở cấu trúc CŨ (thiếu cột PK
    đúng tên):
        • Bảng TRỐNG  → drop, tạo lại theo schema mới.
        • Bảng CÓ dữ liệu → đổi tên sang *__legacy, tạo bảng mới, rồi copy các
          cột khớp (có map tên cột) sang. KHÔNG mất dữ liệu.
    """
    with get_connection() as conn:
        _stash_legacy_tables(conn)
        conn.executescript(cv_schema.SCHEMA_SQL)
        _copy_legacy_data(conn)
        for stmt in cv_schema.MIGRATIONS:
            try:
                conn.execute(stmt)
            except sqlite3.OperationalError:
                pass  # thường là "duplicate column" — đã thêm rồi, bỏ qua
        _run_data_migrations(conn)
        _migrate_document_files(conn)
        _drop_job_descriptions(conn)
        _backfill_timestamps(conn)
        _seed_master_data(conn)


def _run_data_migrations(conn: sqlite3.Connection) -> None:
    """Chạy các lượt sửa DỮ LIỆU trong cv_schema.DATA_MIGRATIONS.

    Mỗi lượt chỉ chạy MỘT LẦN cho mỗi file DB — đánh dấu bằng khóa
    "data:<tên lượt>" trong app_meta, nên lần mở tool sau bỏ qua luôn, không
    quét lại bảng. Xóa dòng đánh dấu trong app_meta nếu muốn chạy lại.

    Lượt nào lỗi giữa chừng thì KHÔNG ghi dấu (cả init_db nằm trong một
    transaction → rollback), lần mở sau sẽ thử lại.
    """
    for name, statements in cv_schema.DATA_MIGRATIONS.items():
        key = f"data:{name}"
        if conn.execute("SELECT 1 FROM app_meta WHERE key = ?", (key,)).fetchone():
            continue
        for stmt in statements:
            conn.execute(stmt)
        conn.execute(
            "INSERT OR REPLACE INTO app_meta (key, value) VALUES "
            "(?, datetime('now', 'localtime'))", (key,))


def _table_exists(conn, name) -> bool:
    return conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
        (name,)).fetchone() is not None


def _stash_legacy_tables(conn: sqlite3.Connection) -> None:
    """Xử lý bảng lệch schema: drop nếu trống, đổi tên sang *__legacy nếu có data."""
    for table in cv_schema._MANAGED_TABLES:
        info = conn.execute(f"PRAGMA table_info({table})").fetchall()
        if not info:
            continue  # chưa tồn tại → executescript sẽ tạo
        cols = {row[1] for row in info}
        if _PK[table] in cols:
            continue  # đã đúng cấu trúc mới
        legacy = table + _LEGACY_SUFFIX
        conn.execute(f"DROP TABLE IF EXISTS {legacy}")  # dọn tàn dư lần trước
        n = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
        if n == 0:
            conn.execute(f"DROP TABLE {table}")
        else:
            conn.execute(f"ALTER TABLE {table} RENAME TO {legacy}")


def _copy_legacy_data(conn: sqlite3.Connection) -> None:
    """Copy dữ liệu từ các bảng *__legacy sang bảng mới rồi xóa bảng legacy."""
    for table in cv_schema._MANAGED_TABLES:
        legacy = table + _LEGACY_SUFFIX
        if not _table_exists(conn, legacy):
            continue
        old_cols = [r[1] for r in conn.execute(f"PRAGMA table_info({legacy})")]
        new_cols = {r[1] for r in conn.execute(f"PRAGMA table_info({table})")}
        rename = _LEGACY_COL_RENAME.get(table, {})

        pairs = []  # (cột_mới, cột_cũ) — chỉ giữ cột có chỗ trong bảng mới
        for oc in old_cols:
            nc = rename.get(oc, oc)
            if nc in new_cols:
                pairs.append((nc, oc))
        if pairs:
            new_list = ", ".join(nc for nc, _ in pairs)
            old_list = ", ".join(oc for _, oc in pairs)
            conn.execute(
                f"INSERT INTO {table} ({new_list}) SELECT {old_list} FROM {legacy}")

        conn.execute(f"DROP TABLE {legacy}")


def _backfill_timestamps(conn: sqlite3.Connection) -> None:
    """Điền created_at/updated_at cho bản ghi CŨ đang NULL (sau khi ALTER TABLE).

    Cột thêm qua ALTER TABLE không có default động nên bản ghi cũ để NULL. Điền
    một lần bằng thời gian hiện tại để giao diện không hiển thị ô trống; bản ghi
    tạo mới về sau đã có sẵn giá trị từ DEFAULT của schema.
    """
    for table in cv_schema._MANAGED_TABLES:
        cols = {r[1] for r in conn.execute(f"PRAGMA table_info({table})")}
        for col in ("created_at", "updated_at"):
            if col in cols:
                conn.execute(
                    f"UPDATE {table} SET {col} = datetime('now', 'localtime') "
                    f"WHERE {col} IS NULL")


def _seed_master_data(conn: sqlite3.Connection) -> None:
    """Nạp dữ liệu khởi tạo cho các bảng danh mục (cv_schema.SEED_DATA).

    • Mỗi khối seed chỉ chạy MỘT LẦN cho mỗi file .db — đánh dấu bằng khóa
      "seed:<bảng>:v<version>" trong bảng app_meta. Nhờ vậy nếu người dùng xóa
      bớt danh mục thì lần mở sau KHÔNG bị nạp lại.
    • Ngay trong lần nạp, dòng nào đã tồn tại (so theo các cột ở `match`, bỏ
      hoa/thường & khoảng trắng thừa) sẽ được bỏ qua → không tạo bản ghi trùng
      với dữ liệu người dùng đã tự nhập trước đó.
    • Muốn nạp lại: xóa dòng tương ứng trong app_meta (hoặc tăng `version` ở
      SEED_DATA khi bổ sung danh mục mới).
    """
    for table, spec in cv_schema.SEED_DATA.items():
        if not _table_exists(conn, table):
            continue
        key = f"seed:{table}:v{spec.get('version', 1)}"
        if conn.execute("SELECT 1 FROM app_meta WHERE key = ?", (key,)).fetchone():
            continue
        cols = list(spec["columns"])
        match = list(spec.get("match") or cols)
        placeholders = ", ".join("?" for _ in cols)
        where = " AND ".join(
            f"LOWER(TRIM(COALESCE({c}, ''))) = LOWER(TRIM(?))" for c in match)
        sql = (f"INSERT INTO {table} ({', '.join(cols)}) SELECT {placeholders} "
               f"WHERE NOT EXISTS (SELECT 1 FROM {table} WHERE {where})")
        for row in spec["rows"]:
            values = dict(zip(cols, row))
            conn.execute(sql, list(row) + [values[c] for c in match])
        conn.execute(
            "INSERT OR REPLACE INTO app_meta (key, value) VALUES "
            "(?, datetime('now', 'localtime'))", (key,))


def _migrate_document_files(conn: sqlite3.Connection) -> None:
    """Bỏ bảng document_files (thiết kế cũ) → đưa đường dẫn CV về candidates.

    Với DB đã lỡ tạo bảng document_files: chép đường dẫn/tên file của mỗi bản
    ghi về cột cv_file_path (theo candidate_id) rồi xóa hẳn bảng. Chạy an toàn
    nhiều lần (không có bảng thì bỏ qua).
    """
    if not _table_exists(conn, "document_files"):
        return
    file_cols = {r[1] for r in conn.execute("PRAGMA table_info(document_files)")}
    # Ưu tiên file_path (đường dẫn đầy đủ), thiếu thì dùng file_name.
    src = "file_path" if "file_path" in file_cols else "file_name"
    if "candidate_id" in file_cols:
        conn.execute(f"""
            UPDATE candidates SET cv_file_path = COALESCE((
                SELECT COALESCE(NULLIF(TRIM(f.{src}), ''), f.file_name)
                FROM document_files f
                WHERE f.candidate_id = candidates.candidate_id
                ORDER BY f.file_id LIMIT 1
            ), cv_file_path)
            WHERE cv_file_path IS NULL OR TRIM(cv_file_path) = ''
        """)
    conn.execute("DROP TABLE document_files")


def _drop_job_descriptions(conn: sqlite3.Connection) -> None:
    """Xóa hẳn bảng job_descriptions (thiết kế cũ).

    Mỗi vị trí chỉ có ĐÚNG 1 JD nên JD nằm luôn trong cột positions.jd_file_path.
    Dữ liệu trong bảng cũ KHÔNG chuyển sang (theo yêu cầu — nhập lại ở form vị
    trí), cũng không giữ bản sao. Chạy an toàn nhiều lần.
    """
    conn.execute("DROP TABLE IF EXISTS job_descriptions")


# ───────────────────────── CRUD generic dùng chung ──────────────────────

def _insert(table: str, allowed: list[str], data: dict) -> int:
    d = {k: data[k] for k in allowed if k in data}
    with get_connection() as conn:
        if not d:
            cur = conn.execute(f"INSERT INTO {table} DEFAULT VALUES")
            return cur.lastrowid
        cols = list(d)
        ph = ", ".join("?" for _ in cols)
        cur = conn.execute(
            f"INSERT INTO {table} ({', '.join(cols)}) VALUES ({ph})",
            [d[c] for c in cols])
        return cur.lastrowid


def _update(table: str, allowed: list[str], row_id: int, data: dict) -> None:
    d = {k: data[k] for k in allowed if k in data}
    if not d:
        return
    # updated_at luôn được cập nhật ở đây (SQLite không tự động làm việc này).
    sets = ", ".join(f"{c} = ?" for c in d)
    sets += ", updated_at = datetime('now', 'localtime')"
    params = [d[c] for c in d] + [row_id]
    with get_connection() as conn:
        conn.execute(f"UPDATE {table} SET {sets} WHERE {_PK[table]} = ?", params)


def _delete(table: str, row_id: int) -> None:
    with get_connection() as conn:
        conn.execute(f"DELETE FROM {table} WHERE {_PK[table]} = ?", (row_id,))


def _get(table: str, row_id: int):
    with get_connection() as conn:
        return conn.execute(
            f"SELECT * FROM {table} WHERE {_PK[table]} = ?", (row_id,)).fetchone()


# ───────────────────────────── PHÒNG BAN ────────────────────────────────

def list_departments():
    with get_connection() as conn:
        return conn.execute(
            "SELECT * FROM departments ORDER BY department_name").fetchall()


def get_department(dept_id):
    return _get("departments", dept_id)


def insert_department(data: dict) -> int:
    return _insert("departments", DEPARTMENT_FIELDS, data)


def update_department(dept_id, data: dict) -> None:
    _update("departments", DEPARTMENT_FIELDS, dept_id, data)


def delete_department(dept_id) -> None:
    _delete("departments", dept_id)


# ──────────────────────── LOẠI NHÂN VIÊN (employee_types) ───────────────

def list_employee_types():
    with get_connection() as conn:
        return conn.execute(
            "SELECT * FROM employee_types ORDER BY code").fetchall()


def list_employee_type_codes() -> list[str]:
    """Chỉ lấy danh sách mã (WC, WCA…) — tiện đổ vào ô chọn ở giao diện."""
    return [r["code"] for r in list_employee_types() if r["code"]]


def get_employee_type(type_id):
    return _get("employee_types", type_id)


def insert_employee_type(data: dict) -> int:
    return _insert("employee_types", EMPLOYEE_TYPE_FIELDS, data)


def update_employee_type(type_id, data: dict) -> None:
    _update("employee_types", EMPLOYEE_TYPE_FIELDS, type_id, data)


def delete_employee_type(type_id) -> None:
    _delete("employee_types", type_id)


# ─────────────────────── TRUNG TÂM CHI PHÍ (cost_centers) ────────────────

def list_cost_centers(group_function: str = ""):
    """Danh sách cost center, lọc tùy chọn theo Group Function (VNPlant/Corporate/R&D)."""
    sql = "SELECT * FROM cost_centers"
    params: list = []
    if group_function:
        sql += " WHERE group_function = ?"
        params.append(group_function)
    sql += " ORDER BY code"
    with get_connection() as conn:
        return conn.execute(sql, params).fetchall()


def list_cost_center_codes(group_function: str = "") -> list[str]:
    return [r["code"] for r in list_cost_centers(group_function) if r["code"]]


def get_cost_center(cost_center_id):
    return _get("cost_centers", cost_center_id)


def insert_cost_center(data: dict) -> int:
    return _insert("cost_centers", COST_CENTER_FIELDS, data)


def update_cost_center(cost_center_id, data: dict) -> None:
    _update("cost_centers", COST_CENTER_FIELDS, cost_center_id, data)


def delete_cost_center(cost_center_id) -> None:
    _delete("cost_centers", cost_center_id)


# ───────────────────────────── CẤP BẬC (levels) ──────────────────────────

def list_levels():
    """Cấp bậc theo thứ tự sort_order (dòng chưa đặt thứ tự xếp xuống cuối)."""
    with get_connection() as conn:
        return conn.execute(
            "SELECT * FROM levels "
            "ORDER BY COALESCE(sort_order, 9999), level_name").fetchall()


def list_level_names() -> list[str]:
    """Tên cấp bậc để đổ vào ô chọn."""
    return [r["level_name"] for r in list_levels() if r["level_name"]]


def get_level(level_id):
    return _get("levels", level_id)


def insert_level(data: dict) -> int:
    return _insert("levels", LEVEL_FIELDS, data)


def update_level(level_id, data: dict) -> None:
    _update("levels", LEVEL_FIELDS, level_id, data)


def delete_level(level_id) -> None:
    _delete("levels", level_id)


# ─────────────────────────── MẪU MAIL (mail_templates) ───────────────────

# Hậu tố đặt cho bản sao khi bấm Duplicate ở màn hình Mail templates.
_COPY_SUFFIX = "_copy"


def list_mail_templates(template_type: str = ""):
    """Danh sách mẫu mail, lọc tùy chọn theo loại (MAIL_TEMPLATE_TYPE_CHOICES)."""
    sql = "SELECT * FROM mail_templates"
    params: list = []
    if template_type:
        sql += " WHERE type = ?"
        params.append(template_type)
    sql += " ORDER BY type, name"
    with get_connection() as conn:
        return conn.execute(sql, params).fetchall()


def get_mail_template(template_id):
    return _get("mail_templates", template_id)


def insert_mail_template(data: dict) -> int:
    return _insert("mail_templates", MAIL_TEMPLATE_FIELDS, data)


def update_mail_template(template_id, data: dict) -> None:
    _update("mail_templates", MAIL_TEMPLATE_FIELDS, template_id, data)


def delete_mail_template(template_id) -> None:
    _delete("mail_templates", template_id)


def _copy_template_name(name: str, existing: set[str]) -> str:
    """Tên cho bản sao: '<tên>_copy'; đã có rồi thì thêm số ('_copy2', '_copy3'…).

    `existing` là tập tên đang có (đã hạ hoa/thường + strip) để so trùng.
    """
    base = (name or "").strip() + _COPY_SUFFIX
    if base.lower() not in existing:
        return base
    i = 2
    while f"{base}{i}".lower() in existing:
        i += 1
    return f"{base}{i}"


def duplicate_mail_template(template_id) -> int:
    """Nhân bản 1 mẫu mail: chép nguyên nội dung, tên thêm hậu tố '_copy'.

    Trả về id bản mới, hoặc 0 nếu không tìm thấy mẫu gốc.
    """
    row = _get("mail_templates", template_id)
    if row is None:
        return 0
    with get_connection() as conn:
        existing = {(r["name"] or "").strip().lower()
                    for r in conn.execute("SELECT name FROM mail_templates")}
    data = {f: row[f] for f in MAIL_TEMPLATE_FIELDS}
    data["name"] = _copy_template_name(row["name"], existing)
    return _insert("mail_templates", MAIL_TEMPLATE_FIELDS, data)


# ───────────────────────────── VỊ TRÍ ───────────────────────────────────

# Mỗi vị trí trỏ tới 3 mẫu mail (3 vòng phỏng vấn) — danh sách vị trí trả kèm
# TÊN của từng mẫu (mail_template_r1_name…) để bảng hiển thị chữ thay vì id.
_POSITION_SELECT = (
    "SELECT p.*, d.department_name, "
    "       t1.name AS mail_template_r1_name, "
    "       t2.name AS mail_template_r2_name, "
    "       t3.name AS mail_template_r3_name "
    "FROM positions p "
    "LEFT JOIN departments d     ON d.department_id = p.department_id "
    "LEFT JOIN mail_templates t1 ON t1.mail_template_id = p.mail_template_r1_id "
    "LEFT JOIN mail_templates t2 ON t2.mail_template_id = p.mail_template_r2_id "
    "LEFT JOIN mail_templates t3 ON t3.mail_template_id = p.mail_template_r3_id "
)


def list_positions():
    with get_connection() as conn:
        return conn.execute(
            _POSITION_SELECT + "ORDER BY p.position_title").fetchall()


def get_position(pos_id):
    return _get("positions", pos_id)


def insert_position(data: dict) -> int:
    return _insert("positions", POSITION_FIELDS, data)


def update_position(pos_id, data: dict) -> None:
    _update("positions", POSITION_FIELDS, pos_id, data)


def delete_position(pos_id) -> None:
    _delete("positions", pos_id)


# ───────────────────────────── ỨNG VIÊN ─────────────────────────────────

# Các cột TEXT được ô tìm kiếm toàn văn quét qua. KHÔNG gồm:
#   • field đã có ô lọc riêng dạng select (status, position, department);
#   • các cột NHẬN XÉT của AI (fit_summary/strengths/weaknesses) — văn bản dài
#     nên gõ 1–2 từ gần như dòng nào cũng khớp, gây nhiễu kết quả.
CANDIDATE_SEARCH_FIELDS = [
    "c.full_name", "c.email", "c.phone", "c.address", "c.education",
    "c.source", "c.note", "c.cv_file_path",
]


def list_batches():
    """Danh sách các 'batch' (đợt quét) khác nhau đang có trong DB, mới → cũ."""
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT DISTINCT batch FROM candidates "
            "WHERE batch IS NOT NULL AND TRIM(batch) <> '' "
            "ORDER BY batch").fetchall()
    return [r["batch"] for r in rows]


def search_candidates(keyword: str = "", position_id=None, status: str = "",
                      department_id=None, batch: str = ""):
    """Tìm ứng viên: từ khóa quét MỌI cột text; lọc theo vị trí / bộ phận /
    trạng thái / batch (các ô select).

    Từ khóa tách theo khoảng trắng → mỗi từ (token) phải khớp ÍT NHẤT một cột
    text (LIKE, khớp chuỗi con); các token ghép AND với nhau. Trả kèm
    `position_title` và `department_name` để hiển thị bảng cho tiện.
    """
    sql = [
        "SELECT c.*, p.position_title, d.department_name",
        "FROM candidates c",
        "LEFT JOIN positions p ON p.position_id = c.position_id",
        "LEFT JOIN departments d ON d.department_id = p.department_id",
        "WHERE 1=1",
    ]
    params: list = []
    kw = (keyword or "").strip()
    if kw:
        ors = " OR ".join(f"{col} LIKE ?" for col in CANDIDATE_SEARCH_FIELDS)
        for token in kw.split():
            sql.append(f"AND ({ors})")
            params += [f"%{token}%"] * len(CANDIDATE_SEARCH_FIELDS)
    if position_id:
        sql.append("AND c.position_id = ?")
        params.append(position_id)
    if department_id:
        sql.append("AND p.department_id = ?")
        params.append(department_id)
    if status:
        sql.append("AND c.status = ?")
        params.append(status)
    if batch not in (None, ""):
        try:
            batch = int(batch)
        except (TypeError, ValueError):
            pass
        sql.append("AND c.batch = ?")
        params.append(batch)
    sql.append("ORDER BY c.candidate_id DESC")
    with get_connection() as conn:
        return conn.execute(" ".join(sql), params).fetchall()


def get_candidate(candidate_id):
    return _get("candidates", candidate_id)


def insert_candidate(data: dict) -> int:
    return _insert("candidates", CANDIDATE_FIELDS, data)


def update_candidate(candidate_id, data: dict) -> None:
    _update("candidates", CANDIDATE_FIELDS, candidate_id, data)


def delete_candidate(candidate_id) -> None:
    _delete("candidates", candidate_id)


def count_candidates() -> int:
    with get_connection() as conn:
        return conn.execute("SELECT COUNT(*) FROM candidates").fetchone()[0]


def set_cv_file_path(candidate_id, path) -> None:
    """Cập nhật lại đường dẫn file CV (dùng khi định vị lại file đã bị di chuyển)."""
    _update("candidates", CANDIDATE_FIELDS, candidate_id, {"cv_file_path": path})


# ───────────────────────────── NHÂN VIÊN ────────────────────────────────

# Các cột TEXT được ô tìm kiếm toàn văn quét qua (bỏ các field đã có ô lọc riêng
# dạng select: department_id, gender, level). Chỉ gồm các cột hay dùng để TRA
# CỨU người — không quét hết ~90 cột (chậm & dễ khớp nhiễu).
EMPLOYEE_SEARCH_FIELDS = [
    "e.code", "e.global_code", "e.full_name", "e.surname", "e.name",
    "e.middle_name", "e.education", "e.phone", "e.email", "e.company_email",
    "e.address", "e.city", "e.job_title", "e.id_no",
]

# Trạng thái làm việc là cột SUY RA, không lưu trong DB: `termination_date` có
# giá trị (khác NULL/rỗng) = đã nghỉ việc. Mọi truy vấn danh sách nhân viên trả
# thêm cột `work_status` ("Working"/"Resigned") cho giao diện hiển thị.
EMPLOYEE_WORK_STATUS_SQL = (
    "CASE WHEN COALESCE(TRIM(e.termination_date), '') = '' "
    "THEN 'Working' ELSE 'Resigned' END")

# GLOBAL SCOPE: mọi nghiệp vụ đọc danh sách nhân viên (list/search/count) MẶC
# ĐỊNH bỏ qua người đã nghỉ việc (termination_date có giá trị) — trừ khi gọi với
# include_resigned=True (vd màn hình Employees khi tick "Include resigned").
_EXCLUDE_RESIGNED_SQL = "COALESCE(TRIM(e.termination_date), '') = ''"

# Phần SELECT + JOIN dùng chung cho list/search nhân viên: kèm tên của 4 bảng
# danh mục (bộ phận · cấp bậc · cost center · loại nhân viên) để bảng hiển thị
# TEXT thay vì id, và cột suy ra `work_status`.
_EMPLOYEE_SELECT = [
    "SELECT e.*, d.department_name, l.level_name,",
    "       cc.code AS cost_center_code, cc.group_function AS cost_center_group,",
    "       et.code AS employee_type_code, et.collar,",
    f"      {EMPLOYEE_WORK_STATUS_SQL} AS work_status",
    "FROM employees e",
    "LEFT JOIN departments d     ON d.department_id = e.department_id",
    "LEFT JOIN levels l          ON l.level_id = e.level_id",
    "LEFT JOIN cost_centers cc   ON cc.cost_center_id = e.cost_center_id",
    "LEFT JOIN employee_types et ON et.employee_type_id = e.employee_type_id",
]


def list_employees(include_resigned=False):
    sql = list(_EMPLOYEE_SELECT)
    if not include_resigned:
        sql.append(f"WHERE {_EXCLUDE_RESIGNED_SQL}")
    sql.append("ORDER BY e.full_name")
    with get_connection() as conn:
        return conn.execute(" ".join(sql)).fetchall()


def search_employees(keyword: str = "", department_id=None, gender: str = "",
                     level_id=None, codes=None, include_resigned=False):
    """Tìm nhân viên: từ khóa quét MỌI cột text; lọc theo bộ phận / giới tính /
    level (các ô select).

    Từ khóa tách theo khoảng trắng → mỗi token phải khớp ÍT NHẤT một cột text
    (LIKE); các token ghép AND. Trả kèm `department_name`/`level_name` để hiển
    thị bảng.

    `codes`: danh sách mã NV (thường dán nguyên cột từ Excel). Nếu có, lọc CHÍNH
    XÁC những nhân viên có `code` nằm trong danh sách (khớp không phân biệt hoa
    thường + bỏ khoảng trắng thừa).

    `include_resigned`: mặc định False → áp GLOBAL SCOPE, bỏ người đã nghỉ việc
    (`termination_date` có giá trị — xem `_EXCLUDE_RESIGNED_SQL`). True → bỏ áp
    scope, lấy luôn cả người đã nghỉ.
    """
    sql = _EMPLOYEE_SELECT + ["WHERE 1=1"]
    params: list = []
    if not include_resigned:
        sql.append(f"AND {_EXCLUDE_RESIGNED_SQL}")
    kw = (keyword or "").strip()
    if kw:
        ors = " OR ".join(f"{col} LIKE ?" for col in EMPLOYEE_SEARCH_FIELDS)
        for token in kw.split():
            sql.append(f"AND ({ors})")
            params += [f"%{token}%"] * len(EMPLOYEE_SEARCH_FIELDS)
    norm_codes = [c.strip().upper() for c in (codes or []) if c and c.strip()]
    if norm_codes:
        ph = ", ".join("?" for _ in norm_codes)
        sql.append(f"AND UPPER(TRIM(e.code)) IN ({ph})")
        params += norm_codes
    if department_id:
        sql.append("AND e.department_id = ?")
        params.append(department_id)
    if gender:
        sql.append("AND e.gender = ?")
        params.append(gender)
    if level_id:
        sql.append("AND e.level_id = ?")
        params.append(level_id)
    sql.append("ORDER BY e.employee_id DESC")
    with get_connection() as conn:
        return conn.execute(" ".join(sql), params).fetchall()


def count_employees(include_resigned=False) -> int:
    sql = "SELECT COUNT(*) FROM employees e"
    if not include_resigned:
        sql += f" WHERE {_EXCLUDE_RESIGNED_SQL}"
    with get_connection() as conn:
        return conn.execute(sql).fetchone()[0]


def get_employee(employee_id):
    return _get("employees", employee_id)


def insert_employee(data: dict) -> int:
    return _insert("employees", EMPLOYEE_FIELDS, data)


def update_employee(employee_id, data: dict) -> None:
    _update("employees", EMPLOYEE_FIELDS, employee_id, data)


def delete_employee(employee_id) -> None:
    _delete("employees", employee_id)


def find_employees_by_codes(codes):
    """Tìm nhân viên ĐÃ CÓ trong DB theo mã NV (`code`) — dùng để phát hiện
    trùng khi import Excel. Khớp không phân biệt hoa/thường + bỏ khoảng trắng
    thừa. Trả về [] nếu `codes` rỗng.
    """
    norm_codes = {c.strip().upper() for c in (codes or []) if c and c.strip()}
    if not norm_codes:
        return []
    ph = ", ".join("?" for _ in norm_codes)
    with get_connection() as conn:
        return conn.execute(
            "SELECT employee_id, code, full_name FROM employees "
            f"WHERE UPPER(TRIM(code)) IN ({ph})",
            list(norm_codes)).fetchall()


def find_employees_by_identity(code=None, id_no=None, full_name=None):
    """Tìm nhân viên ĐÃ CÓ trùng với một hồ sơ sắp thêm — dùng khi import đơn dự
    tuyển (nhập từng người một, thường chưa có mã NV).

    Khớp theo BẤT KỲ dấu hiệu định danh nào: mã NV, số CMND/CCCD, hoặc trùng cả
    họ tên (không phân biệt hoa/thường, bỏ khoảng trắng thừa). Không truyền gì
    thì trả về [].
    """
    conds, params = [], []
    for col, value in (("code", code), ("id_no", id_no), ("full_name", full_name)):
        text = " ".join(str(value or "").split())
        if text:
            conds.append(f"UPPER(TRIM({col})) = UPPER(?)")
            params.append(text)
    if not conds:
        return []
    with get_connection() as conn:
        return conn.execute(
            "SELECT employee_id, code, full_name, date_of_birth, id_no "
            f"FROM employees WHERE {' OR '.join(conds)}", params).fetchall()


# ───────────────────────── KHÓA HỌC / ĐÀO TẠO ───────────────────────────

def list_courses():
    with get_connection() as conn:
        return conn.execute(
            "SELECT * FROM courses ORDER BY date DESC, course_id DESC").fetchall()


def get_course(course_id):
    return _get("courses", course_id)


def insert_course(data: dict) -> int:
    return _insert("courses", COURSE_FIELDS, data)


def update_course(course_id, data: dict) -> None:
    _update("courses", COURSE_FIELDS, course_id, data)


def delete_course(course_id) -> None:
    _delete("courses", course_id)


# ───────────────── LIÊN KẾT KHÓA HỌC ↔ NHÂN VIÊN (nhiều-nhiều) ───────────

def list_course_employees(course_id):
    """Danh sách nhân viên đã ghi danh vào 1 khóa học (kèm thông tin nhân viên)."""
    with get_connection() as conn:
        return conn.execute(
            "SELECT ce.*, e.code, e.full_name, e.email, e.department_id, "
            "       d.department_name, d.short_name AS department_short "
            "FROM course_employees ce "
            "LEFT JOIN employees e   ON e.employee_id = ce.employee_id "
            "LEFT JOIN departments d ON d.department_id = e.department_id "
            "WHERE ce.course_id = ? "
            "ORDER BY e.full_name", (course_id,)).fetchall()


def list_employee_courses(employee_id):
    """Danh sách khóa học mà 1 nhân viên đã tham gia (kèm thông tin khóa học)."""
    with get_connection() as conn:
        return conn.execute(
            "SELECT ce.*, c.title, c.date, c.location, c.course_type "
            "FROM course_employees ce "
            "LEFT JOIN courses c ON c.course_id = ce.course_id "
            "WHERE ce.employee_id = ? "
            "ORDER BY c.date DESC", (employee_id,)).fetchall()


def enroll_employee(course_id, employee_id, data: dict | None = None) -> int:
    """Ghi danh 1 nhân viên vào 1 khóa học. Bỏ qua nếu đã ghi danh (INSERT OR IGNORE
    nhờ unique index course_id+employee_id). Trả về enrollment_id (0 nếu đã có sẵn).
    """
    d = {k: (data or {}).get(k) for k in COURSE_EMPLOYEE_FIELDS if k in (data or {})}
    d["course_id"] = course_id
    d["employee_id"] = employee_id
    cols = list(d)
    ph = ", ".join("?" for _ in cols)
    with get_connection() as conn:
        cur = conn.execute(
            f"INSERT OR IGNORE INTO course_employees ({', '.join(cols)}) "
            f"VALUES ({ph})", [d[c] for c in cols])
        return cur.lastrowid if cur.rowcount else 0


def unenroll_employee(course_id, employee_id) -> None:
    """Gỡ 1 nhân viên khỏi 1 khóa học."""
    with get_connection() as conn:
        conn.execute(
            "DELETE FROM course_employees WHERE course_id = ? AND employee_id = ?",
            (course_id, employee_id))


def update_enrollment(enrollment_id, data: dict) -> None:
    _update("course_employees", COURSE_EMPLOYEE_FIELDS, enrollment_id, data)


def get_enrollment(enrollment_id):
    return _get("course_employees", enrollment_id)


def search_course_employees(course_id=None, status: str = ""):
    """Tra cứu lượt ghi danh (course_employees) theo khóa học / trạng thái học.

    Chỉ lọc bằng 2 ô select (không có tìm kiếm toàn văn). Trả về TẤT CẢ cột của
    course_employees kèm thông tin nhân viên & tên khóa học để hiển thị bảng.
    """
    sql = [
        "SELECT ce.*, c.title AS course_title,",
        "       e.code, e.full_name, e.email, e.department_id,",
        "       d.department_name, d.short_name AS department_short",
        "FROM course_employees ce",
        "LEFT JOIN courses c     ON c.course_id = ce.course_id",
        "LEFT JOIN employees e   ON e.employee_id = ce.employee_id",
        "LEFT JOIN departments d ON d.department_id = e.department_id",
        "WHERE 1=1",
    ]
    params: list = []
    if course_id:
        sql.append("AND ce.course_id = ?")
        params.append(course_id)
    if status:
        sql.append("AND ce.status = ?")
        params.append(status)
    sql.append("ORDER BY e.full_name")
    with get_connection() as conn:
        return conn.execute(" ".join(sql), params).fetchall()


def count_course_employees() -> int:
    with get_connection() as conn:
        return conn.execute("SELECT COUNT(*) FROM course_employees").fetchone()[0]


def find_duplicates(email=None, phone=None, exclude_id=None):
    """Tìm ứng viên trùng theo EMAIL hoặc SĐT (bỏ qua khoảng trắng, không phân
    biệt hoa/thường với email). Trả về list rrow rỗng nếu không nhập gì.

    `exclude_id` để loại chính ứng viên đang sửa ra khỏi kết quả.
    """
    email = (email or "").strip()
    phone = (phone or "").strip()
    conds, params = [], []
    if email:
        conds.append("LOWER(TRIM(email)) = LOWER(?)")
        params.append(email)
    if phone:
        conds.append("TRIM(phone) = ?")
        params.append(phone)
    if not conds:
        return []
    sql = ("SELECT candidate_id, full_name, email, phone FROM candidates "
           f"WHERE ({' OR '.join(conds)})")
    if exclude_id is not None:
        sql += " AND candidate_id <> ?"
        params.append(exclude_id)
    with get_connection() as conn:
        return conn.execute(sql, params).fetchall()
