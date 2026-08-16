"""Tầng truy cập dữ liệu (data-access layer) cho nghiệp vụ Tuyển dụng & Nhân sự.

Gói toàn bộ thao tác SQLite ở một chỗ để giao diện chỉ việc gọi hàm. Cấu trúc
bảng định nghĩa trong `app/core/cv_schema.py` (file thiết kế DB); mô tả đầy đủ
trong `docs/db_design.md`.

File .db mặc định:
    %APPDATA%\\PersonalToolbox\\candidates.sqlite   (Windows)
    ~/.config/PersonalToolbox/candidates.sqlite      (Linux/macOS — lúc dev)

Cấu trúc DB dựng và cập nhật hoàn toàn bằng `cv_schema.MIGRATIONS`: máy chưa có
DB thì chạy tất cả các lượt, máy đã có thì chỉ chạy phần còn thiếu (xem
`init_db`). Không có cơ chế dựng lại DB nào khác.
"""
import hashlib
import os
import re
import sqlite3
from datetime import date, datetime

from app.core import cv_schema

# ─────────────────────── Cột được phép ghi cho từng bảng ─────────────────
# (chặn khóa lạ lọt vào câu INSERT/UPDATE)

DEPARTMENT_FIELDS = ["department_name", "short_name", "manager_name", "description"]
EMPLOYEE_TYPE_FIELDS = ["code", "collar", "description"]
COST_CENTER_FIELDS = ["code", "group_function", "name", "description"]
LEVEL_FIELDS = ["level_name", "sort_order", "description"]
SKILL_FIELDS = ["name", "aliases", "category", "description"]
MAIL_TEMPLATE_FIELDS = ["name", "type", "mail_cc", "mail_subject", "mail_body"]

POSITION_FIELDS = [
    "department_id", "position_code", "jrf_code", "position_title", "description",
    "required_experience", "salary_level", "starting_date", "level", "headcount",
    "status", "jd_file_path",
    "mail_template_r1_id", "mail_template_r2_id", "mail_template_r3_id", "note",
]
POSITION_REQUIREMENT_FIELDS = [
    "position_id", "jd_hash", "must_have", "nice_to_have", "min_years",
    "education", "major", "language_req", "level", "summary", "model", "parsed_at",
]

# Nhóm "ảnh chụp hồ sơ" — dùng chung giữa `candidates` và `candidate_cvs`, nên
# chép từ bản CV mới nhất sang ứng viên chỉ là một vòng lặp qua danh sách này.
PROFILE_SNAPSHOT_FIELDS = [
    "current_title", "industry", "years_experience", "experience_as_of",
    "education", "major", "languages", "skills_text", "profile_summary",
]
CANDIDATE_FIELDS = [
    "full_name", "email", "phone", "date_of_birth", "gender", "address", "city",
    "pool_status", "first_seen_at", "last_contacted_at", "latest_cv_id",
    "source", "note",
    *PROFILE_SNAPSHOT_FIELDS,
    "expected_salary", "salary_note", "available_from", "willing_to_relocate",
    "preferred_location",
]
CANDIDATE_CV_FIELDS = [
    "candidate_id", "file_path", "file_hash", "received_at", "batch", "source",
    "cv_text", "scanned_at", "scan_model",
    *PROFILE_SNAPSHOT_FIELDS,
]
CANDIDATE_EXPERIENCE_FIELDS = [
    "candidate_id", "cv_id", "company", "job_title", "industry",
    "start_date", "end_date", "as_of_date", "is_current", "description", "sort_order",
]
CANDIDATE_SKILL_FIELDS = [
    "candidate_id", "skill_id", "cv_id", "raw_name", "years", "level", "source",
]
APPLICATION_FIELDS = [
    "candidate_id", "position_id", "cv_id", "origin", "source", "status",
    "final_status", "phone_screen_date", "applied_at", "status_changed_at",
    "closed_at", "note",
]
INTERVIEW_FIELDS = [
    "application_id", "candidate_id", "round", "interview_date", "duration_minutes",
    "mode", "location", "overall_score", "next_step", "status",
    "mail_activity_id",
]
INTERVIEW_FEEDBACK_FIELDS = [
    "interview_id", "employee_id", "interviewer_name", "role", "score", "rating",
    "feedback", "strengths", "weaknesses", "submitted_at",
]
EVALUATION_FIELDS = [
    "candidate_id", "position_id", "application_id", "cv_id", "source",
    "rule_score", "ai_score", "matched_skills", "missing_skills",
    "summary", "strengths", "weaknesses", "model", "jd_hash", "extra_prompt",
    "evaluated_at",
]
ACTIVITY_FIELDS = [
    "candidate_id", "application_id", "type", "round", "occurred_at",
    "scheduled_at", "subject", "content", "mail_template_id", "mail_to",
    "mail_cc", "result", "from_status", "to_status",
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
    "emergency_contact_name", "emergency_contact_phone",
    "emergency_contact_relationship",
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
    "working_hours_per_week", "smart_working_eligible", "er_jrf", "is_interviewer",
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

# PK của mỗi bảng (dùng cho update/delete generic).
_PK = {
    "departments": "department_id",
    "employee_types": "employee_type_id",
    "cost_centers": "cost_center_id",
    "levels": "level_id",
    "skills": "skill_id",
    "mail_templates": "mail_template_id",
    "positions": "position_id",
    "position_requirements": "requirement_id",
    "candidates": "candidate_id",
    "candidate_cvs": "cv_id",
    "candidate_experiences": "experience_id",
    "candidate_skills": "candidate_skill_id",
    "applications": "application_id",
    "interviews": "interview_id",
    "interview_feedbacks": "feedback_id",
    "candidate_evaluations": "evaluation_id",
    "candidate_activities": "activity_id",
    "employees": "employee_id",
    "courses": "course_id",
    "course_employees": "enrollment_id",
}


# ════════════════════════════ KẾT NỐI & KHỞI TẠO ════════════════════════

def _db_path() -> str:
    base = os.environ.get("APPDATA") or os.path.join(
        os.path.expanduser("~"), ".config")
    folder = os.path.join(base, "PersonalToolbox")
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "candidates.sqlite")


class _Connection(sqlite3.Connection):
    """Kết nối tự ĐÓNG khi thoát khối `with`.

    `sqlite3.Connection` gốc chỉ commit/rollback ở `__exit__` chứ không đóng
    file, nên file .db còn bị giữ cho tới lượt gom rác. Trên Windows điều đó
    chặn hẳn việc đổi tên DB cũ lúc dựng lại schema (WinError 32).
    """

    def __exit__(self, exc_type, exc, tb):
        try:
            return super().__exit__(exc_type, exc, tb)
        finally:
            self.close()


def get_connection() -> sqlite3.Connection:
    """Mở kết nối SQLite (row_factory=Row để truy cập cột theo tên).

    Luôn dùng trong khối `with` — thoát khối là commit và ĐÓNG luôn kết nối.
    KHÔNG bật PRAGMA foreign_keys — thiết kế cố tình không dùng khóa ngoại.
    """
    conn = sqlite3.connect(_db_path(), factory=_Connection)
    conn.row_factory = sqlite3.Row
    return conn


# Có dùng được bảng ảo FTS5 hay không (một số bản SQLite không biên dịch kèm).
FTS_AVAILABLE = False


# Bảng ghi dấu vết migration. Phải tồn tại TRƯỚC khi chạy lượt đầu tiên nên tạo
# riêng ở đây, không nằm trong lượt nào cả.
_META_TABLE_SQL = (
    "CREATE TABLE IF NOT EXISTS app_meta (key VARCHAR PRIMARY KEY, value VARCHAR)")


def applied_migrations() -> set[str]:
    """Tên các lượt migration mà file DB hiện tại ĐÃ chạy."""
    if not os.path.exists(_db_path()):
        return set()
    with get_connection() as conn:
        conn.execute(_META_TABLE_SQL)
        return {r[0][len("migration:"):] for r in conn.execute(
            "SELECT key FROM app_meta WHERE key LIKE 'migration:%'")}


def init_db() -> None:
    """Chạy các lượt migration CÒN THIẾU + nạp danh mục khởi tạo.

    Gọi mỗi lần mở tool (rẻ — máy đã cập nhật thì chỉ tốn một câu SELECT).

        • Máy chưa có file .db → chạy TẤT CẢ `cv_schema.MIGRATIONS`, bắt đầu từ
          lượt "0001" tạo toàn bộ bảng.
        • Máy đã có DB rồi     → chỉ chạy những lượt chưa có dấu vết trong
          `app_meta`, nên thêm một cột chỉ tốn đúng lượt vừa thêm.

    Mỗi lượt chạy xong mới được đánh dấu, và đánh dấu nằm cùng transaction với
    chính lượt đó → lượt nào lỗi thì DỪNG HẲN (ném lỗi ra ngoài), không đánh
    dấu, lần mở app sau thử lại đúng lượt đó.
    """
    global FTS_AVAILABLE
    with get_connection() as conn:
        conn.execute(_META_TABLE_SQL)
        done = {r[0] for r in conn.execute(
            "SELECT key FROM app_meta WHERE key LIKE 'migration:%'")}

        for name, sql in cv_schema.MIGRATIONS:
            key = f"migration:{name}"
            if key in done:
                continue
            conn.executescript(sql)
            conn.execute("INSERT OR REPLACE INTO app_meta (key, value) VALUES (?, ?)",
                         (key, _now()))
            conn.commit()

        # Bảng ảo tìm kiếm toàn văn để NGOÀI migration: bản SQLite thiếu FTS5 sẽ
        # làm hỏng cả lượt, trong khi thiếu nó app vẫn chạy (tìm kiếm lùi về LIKE).
        try:
            conn.executescript(cv_schema.FTS_SQL)
            FTS_AVAILABLE = True
        except sqlite3.Error:
            FTS_AVAILABLE = False

        _seed_master_data(conn)


def _table_exists(conn, name) -> bool:
    return conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
        (name,)).fetchone() is not None


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


# ═══════════════════════ CRUD generic dùng chung ════════════════════════

def _insert(table: str, allowed: list[str], data: dict) -> int:
    d = {k: data[k] for k in allowed if k in data}
    with get_connection() as conn:
        return _insert_conn(conn, table, allowed, d)


def _insert_conn(conn, table: str, allowed: list[str], data: dict) -> int:
    """Bản dùng lại kết nối đang mở — cho các thao tác nhiều bảng trong 1 lượt."""
    d = {k: data[k] for k in allowed if k in data}
    if not d:
        return conn.execute(f"INSERT INTO {table} DEFAULT VALUES").lastrowid
    cols = list(d)
    ph = ", ".join("?" for _ in cols)
    return conn.execute(
        f"INSERT INTO {table} ({', '.join(cols)}) VALUES ({ph})",
        [d[c] for c in cols]).lastrowid


def _update(table: str, allowed: list[str], row_id: int, data: dict) -> None:
    with get_connection() as conn:
        _update_conn(conn, table, allowed, row_id, data)


def _update_conn(conn, table: str, allowed: list[str], row_id: int, data: dict) -> None:
    d = {k: data[k] for k in allowed if k in data}
    if not d:
        return
    sets = ", ".join(f"{c} = ?" for c in d)
    # Bảng chỉ-ghi-thêm (evaluations, activities) không có cột updated_at.
    if _has_column(conn, table, "updated_at"):
        sets += ", updated_at = datetime('now', 'localtime')"
    conn.execute(f"UPDATE {table} SET {sets} WHERE {_PK[table]} = ?",
                 [d[c] for c in d] + [row_id])


def _has_column(conn, table: str, column: str) -> bool:
    return any(r[1] == column for r in conn.execute(f"PRAGMA table_info({table})"))


def _delete(table: str, row_id: int) -> None:
    with get_connection() as conn:
        conn.execute(f"DELETE FROM {table} WHERE {_PK[table]} = ?", (row_id,))


def _get(table: str, row_id: int):
    with get_connection() as conn:
        return conn.execute(
            f"SELECT * FROM {table} WHERE {_PK[table]} = ?", (row_id,)).fetchone()


def now_text() -> str:
    """Thời điểm hiện tại dạng chuỗi SQLite ('yyyy-mm-dd HH:MM:SS')."""
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


_now = now_text   # tên ngắn dùng nội bộ trong file này


# ═══════════════════════════ DANH MỤC DÙNG CHUNG ════════════════════════

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


def list_employee_types():
    with get_connection() as conn:
        return conn.execute("SELECT * FROM employee_types ORDER BY code").fetchall()


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


# ───────────────────────────── KỸ NĂNG (skills) ──────────────────────────

def list_skills():
    with get_connection() as conn:
        return conn.execute(
            "SELECT * FROM skills ORDER BY category, name").fetchall()


def get_skill(skill_id):
    return _get("skills", skill_id)


def insert_skill(data: dict) -> int:
    return _insert("skills", SKILL_FIELDS, data)


def update_skill(skill_id, data: dict) -> None:
    _update("skills", SKILL_FIELDS, skill_id, data)


def delete_skill(skill_id) -> None:
    _delete("skills", skill_id)


def _skill_key(text) -> str:
    """Chuẩn hóa tên kỹ năng để so khớp: bỏ hoa/thường, khoảng trắng, dấu câu.

    "Node.js" · "node js" · "NodeJS" → "nodejs".
    """
    return re.sub(r"[^a-z0-9+#]", "", str(text or "").lower())


def skill_lookup() -> dict[str, int]:
    """Bảng tra {tên đã chuẩn hóa → skill_id}, gồm cả tên chính lẫn `aliases`."""
    table: dict[str, int] = {}
    for row in list_skills():
        sid = row["skill_id"]
        names = [row["name"] or ""] + str(row["aliases"] or "").split(";")
        for name in names:
            key = _skill_key(name)
            if key:
                table.setdefault(key, sid)
    return table


def resolve_skill(raw_name, lookup=None, create=False):
    """Tra một tên kỹ năng thô về skill_id (None nếu chưa có trong danh mục).

    `create=True` thì tự thêm vào danh mục `skills` khi chưa có — dùng lúc quét
    CV để danh mục tự lớn dần theo dữ liệu thật.
    """
    key = _skill_key(raw_name)
    if not key:
        return None
    table = skill_lookup() if lookup is None else lookup
    if key in table:
        return table[key]
    if not create:
        return None
    sid = insert_skill({"name": " ".join(str(raw_name).split())})
    table[key] = sid
    return sid


# ───────────────────────── MẪU MAIL (mail_templates) ─────────────────────

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


# ═════════════════════════════ VỊ TRÍ TUYỂN DỤNG ════════════════════════

# Danh sách vị trí trả kèm TÊN của 3 mẫu mail để bảng hiển thị chữ thay vì id.
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


# ───────────────── YÊU CẦU BÓC TÁCH TỪ JD (position_requirements) ────────

def file_hash(path) -> str:
    """Băm MD5 nội dung một file → nhận diện file trùng dù đã đổi tên.

    Trả về "" nếu không đọc được file.
    """
    try:
        with open(path, "rb") as fh:
            h = hashlib.md5()
            for chunk in iter(lambda: fh.read(1 << 20), b""):
                h.update(chunk)
        return h.hexdigest()
    except OSError:
        return ""


def text_hash(text) -> str:
    """Băm MD5 một chuỗi (dùng cho nội dung JD đã trích xuất)."""
    return hashlib.md5(str(text or "").encode("utf-8", "ignore")).hexdigest()


def get_position_requirements(position_id, jd_hash=""):
    """Bản bóc tách JD của một vị trí.

    Truyền `jd_hash` thì chỉ trả về bản khớp đúng nội dung JD hiện tại (JD đã
    sửa → không khớp → bên gọi biết là phải bóc lại). Không truyền thì trả bản
    mới nhất.
    """
    sql = "SELECT * FROM position_requirements WHERE position_id = ?"
    params: list = [position_id]
    if jd_hash:
        sql += " AND jd_hash = ?"
        params.append(jd_hash)
    sql += " ORDER BY requirement_id DESC LIMIT 1"
    with get_connection() as conn:
        return conn.execute(sql, params).fetchone()


def save_position_requirements(data: dict) -> int:
    """Ghi bản bóc tách JD. Cùng (position_id, jd_hash) thì ghi đè bản cũ."""
    data = dict(data)
    data.setdefault("parsed_at", _now())
    existing = get_position_requirements(data.get("position_id"),
                                         data.get("jd_hash") or "")
    if existing is not None and data.get("jd_hash"):
        _update("position_requirements", POSITION_REQUIREMENT_FIELDS,
                existing["requirement_id"], data)
        return existing["requirement_id"]
    return _insert("position_requirements", POSITION_REQUIREMENT_FIELDS, data)


# ═══════════════════════════════ ỨNG VIÊN ═══════════════════════════════

# Các cột TEXT được ô tìm kiếm toàn văn quét qua. KHÔNG gồm các cột NHẬN XÉT dài
# (profile_summary, cv_text) — gõ 1–2 từ là dòng nào cũng khớp, gây nhiễu.
CANDIDATE_SEARCH_FIELDS = [
    "c.full_name", "c.email", "c.phone", "c.address", "c.city",
    "c.current_title", "c.industry", "c.education", "c.major",
    "c.skills_text", "c.source", "c.note",
]

# Ứng viên + ĐƠN ỨNG TUYỂN MỚI NHẤT + lượt AI CHẤM MỚI NHẤT của đơn đó.
# Trạng thái tuyển dụng nằm ở `applications` nên bảng danh sách phải lấy qua
# đây; ứng viên chưa có đơn nào vẫn hiện (các cột của đơn để trống).
_CANDIDATE_SELECT = """
SELECT c.*,
       a.application_id, a.position_id, a.status, a.final_status,
       a.applied_at, a.origin, a.phone_screen_date,
       a.source AS application_source, a.note AS application_note,
       p.position_title, p.department_id, d.department_name,
       cv.batch, cv.received_at, cv.file_path AS cv_file_path,
       e.ai_score, e.rule_score, e.evaluated_at,
       e.summary AS fit_summary, e.strengths, e.weaknesses,
       (SELECT COUNT(*) FROM applications x WHERE x.candidate_id = c.candidate_id)
           AS application_count
FROM candidates c
LEFT JOIN applications a ON a.application_id = (
        SELECT x.application_id FROM applications x
        WHERE x.candidate_id = c.candidate_id
        ORDER BY COALESCE(x.applied_at, '') DESC, x.application_id DESC LIMIT 1)
LEFT JOIN positions p     ON p.position_id = a.position_id
LEFT JOIN departments d   ON d.department_id = p.department_id
LEFT JOIN candidate_cvs cv ON cv.cv_id = COALESCE(c.latest_cv_id, a.cv_id)
LEFT JOIN candidate_evaluations e ON e.evaluation_id = (
        SELECT y.evaluation_id FROM candidate_evaluations y
        WHERE y.candidate_id = c.candidate_id
          AND (a.position_id IS NULL OR y.position_id = a.position_id)
        ORDER BY COALESCE(y.evaluated_at, '') DESC, y.evaluation_id DESC LIMIT 1)
"""


def list_batches():
    """Các 'batch' (đợt quét) khác nhau đang có trong DB."""
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT DISTINCT batch FROM candidate_cvs "
            "WHERE batch IS NOT NULL ORDER BY batch").fetchall()
    return [r["batch"] for r in rows]


def search_candidates(keyword: str = "", position_id=None, status: str = "",
                      department_id=None, batch: str = "", pool_status: str = ""):
    """Tìm ứng viên (một dòng = một người, kèm ĐƠN MỚI NHẤT của họ).

    Từ khóa tách theo khoảng trắng → mỗi từ (token) phải khớp ÍT NHẤT một cột
    text (LIKE, khớp chuỗi con); các token ghép AND với nhau.

    Lọc theo vị trí / bộ phận / trạng thái đều áp lên ĐƠN MỚI NHẤT; muốn duyệt
    theo từng đơn thì dùng `list_applications`.
    """
    sql = [_CANDIDATE_SELECT, "WHERE 1=1"]
    params: list = []
    kw = (keyword or "").strip()
    if kw:
        ors = " OR ".join(f"{col} LIKE ?" for col in CANDIDATE_SEARCH_FIELDS)
        for token in kw.split():
            sql.append(f"AND ({ors})")
            params += [f"%{token}%"] * len(CANDIDATE_SEARCH_FIELDS)
    if position_id:
        sql.append("AND a.position_id = ?")
        params.append(position_id)
    if department_id:
        sql.append("AND p.department_id = ?")
        params.append(department_id)
    if status:
        sql.append("AND a.status = ?")
        params.append(status)
    if pool_status:
        sql.append("AND c.pool_status = ?")
        params.append(pool_status)
    if batch not in (None, ""):
        try:
            batch = int(batch)
        except (TypeError, ValueError):
            pass
        sql.append("AND cv.batch = ?")
        params.append(batch)
    sql.append("ORDER BY c.candidate_id DESC")
    with get_connection() as conn:
        return conn.execute(" ".join(sql), params).fetchall()


def get_candidate(candidate_id):
    return _get("candidates", candidate_id)


def get_candidate_full(candidate_id):
    """Ứng viên kèm đơn mới nhất & điểm AI mới nhất (giống dòng trong bảng)."""
    with get_connection() as conn:
        return conn.execute(
            _CANDIDATE_SELECT + " WHERE c.candidate_id = ?",
            (candidate_id,)).fetchone()


def insert_candidate(data: dict) -> int:
    data = dict(data)
    data.setdefault("pool_status", cv_schema.POOL_STATUS_DEFAULT)
    data.setdefault("first_seen_at", _now())
    with get_connection() as conn:
        cid = _insert_conn(conn, "candidates", CANDIDATE_FIELDS, data)
        _sync_fts(conn, cid)
    return cid


def update_candidate(candidate_id, data: dict) -> None:
    with get_connection() as conn:
        _update_conn(conn, "candidates", CANDIDATE_FIELDS, candidate_id, data)
        _sync_fts(conn, candidate_id)


def set_candidate_source(candidate_id, source: str, application_id=None) -> None:
    """Ghi NƠI CUNG CẤP CV (Itviec, VietnamWorks…) cho một ứng viên.

    Cột `source` có ở ba bảng và mô tả cùng một sự việc dưới ba góc: ứng viên
    biết đến từ đâu · đơn này đến từ đâu · bản CV này lấy ở đâu. Chúng phải khớp
    nhau, nếu không mỗi màn hình lại đọc ra một giá trị khác — nên ghi cả ba
    trong một lượt:

      • `candidates.source`
      • `applications.source` của đơn `application_id` (đơn đang hiển thị trên
        bảng); bỏ trống thì không đụng tới đơn nào
      • `candidate_cvs.source` của bản CV mới nhất (`candidates.latest_cv_id`)

    AI quét CV không suy ra được thông tin này (quét cả thư mục thì không biết
    file lấy từ sàn nào) nên đây là chỗ duy nhất điền — bằng tay.
    """
    source = (source or "").strip()
    now = _now()
    with get_connection() as conn:
        conn.execute(
            "UPDATE candidates SET source = ?, updated_at = ? WHERE candidate_id = ?",
            (source, now, candidate_id))
        if application_id:
            conn.execute(
                "UPDATE applications SET source = ?, updated_at = ? "
                "WHERE application_id = ?", (source, now, application_id))
        conn.execute(
            "UPDATE candidate_cvs SET source = ?, updated_at = ? WHERE cv_id = "
            "(SELECT latest_cv_id FROM candidates WHERE candidate_id = ?)",
            (source, now, candidate_id))


def delete_candidate(candidate_id) -> None:
    """Xóa ứng viên và MỌI dữ liệu con (CV, kinh nghiệm, đơn, phỏng vấn, lịch sử).

    Không có khóa ngoại nên phải tự dọn, nếu không sẽ còn lại dữ liệu mồ côi.
    """
    with get_connection() as conn:
        app_ids = [r[0] for r in conn.execute(
            "SELECT application_id FROM applications WHERE candidate_id = ?",
            (candidate_id,))]
        if app_ids:
            ph = ", ".join("?" for _ in app_ids)
            conn.execute(
                "DELETE FROM interview_feedbacks WHERE interview_id IN "
                f"(SELECT interview_id FROM interviews WHERE application_id IN ({ph}))",
                app_ids)
        for table in ("interviews", "applications", "candidate_activities",
                      "candidate_evaluations", "candidate_skills",
                      "candidate_experiences", "candidate_cvs"):
            conn.execute(f"DELETE FROM {table} WHERE candidate_id = ?", (candidate_id,))
        conn.execute("DELETE FROM candidates WHERE candidate_id = ?", (candidate_id,))
        _delete_fts(conn, candidate_id)


def count_candidates() -> int:
    with get_connection() as conn:
        return conn.execute("SELECT COUNT(*) FROM candidates").fetchone()[0]


def find_duplicates(email=None, phone=None, exclude_id=None):
    """Ứng viên đã có trong DB trùng email hoặc số điện thoại.

    Khớp không phân biệt hoa/thường & bỏ khoảng trắng thừa. Không truyền gì
    (hoặc cả hai đều rỗng) thì trả về [].
    """
    conds, params = [], []
    for col, value in (("email", email), ("phone", phone)):
        text = " ".join(str(value or "").split())
        if text:
            conds.append(f"LOWER(TRIM({col})) = LOWER(?)")
            params.append(text)
    if not conds:
        return []
    sql = ("SELECT candidate_id, full_name, email, phone FROM candidates "
           f"WHERE ({' OR '.join(conds)})")
    if exclude_id:
        sql += " AND candidate_id <> ?"
        params.append(exclude_id)
    with get_connection() as conn:
        return conn.execute(sql, params).fetchall()


# ───────────────────── TÌM KIẾM TOÀN VĂN (candidates_fts) ────────────────

def _sync_fts(conn, candidate_id) -> None:
    """Ghi lại dòng FTS của một ứng viên (xóa cũ → chèn mới). Lỗi thì bỏ qua."""
    if not FTS_AVAILABLE:
        return
    try:
        conn.execute("DELETE FROM candidates_fts WHERE candidate_id = ?",
                     (candidate_id,))
        conn.execute("""
            INSERT INTO candidates_fts
                (candidate_id, full_name, current_title, skills_text,
                 profile_summary, cv_text)
            SELECT c.candidate_id, c.full_name, c.current_title, c.skills_text,
                   c.profile_summary, COALESCE(cv.cv_text, '')
            FROM candidates c
            LEFT JOIN candidate_cvs cv ON cv.cv_id = c.latest_cv_id
            WHERE c.candidate_id = ?""", (candidate_id,))
    except sqlite3.Error:
        pass


def _delete_fts(conn, candidate_id) -> None:
    if not FTS_AVAILABLE:
        return
    try:
        conn.execute("DELETE FROM candidates_fts WHERE candidate_id = ?",
                     (candidate_id,))
    except sqlite3.Error:
        pass


def search_fts(query: str, limit: int = 200):
    """Tìm toàn văn (BM25) trên hồ sơ + nội dung CV. Trả [] nếu không có FTS5."""
    q = " ".join(str(query or "").split())
    if not q or not FTS_AVAILABLE:
        return []
    try:
        with get_connection() as conn:
            return conn.execute(
                "SELECT candidate_id, bm25(candidates_fts) AS rank "
                "FROM candidates_fts WHERE candidates_fts MATCH ? "
                "ORDER BY rank LIMIT ?", (q, limit)).fetchall()
    except sqlite3.Error:
        return []


# ═══════════════════ CÁC BẢN CV THEO THỜI GIAN (candidate_cvs) ══════════

def list_candidate_cvs(candidate_id):
    """Các bản CV của một ứng viên, MỚI NHẤT trước."""
    with get_connection() as conn:
        return conn.execute(
            "SELECT * FROM candidate_cvs WHERE candidate_id = ? "
            "ORDER BY COALESCE(received_at, '') DESC, cv_id DESC",
            (candidate_id,)).fetchall()


def get_candidate_cv(cv_id):
    return _get("candidate_cvs", cv_id)


def find_cv_by_hash(hash_value):
    """Bản CV đã quét có cùng mã băm nội dung (None nếu chưa từng quét)."""
    if not hash_value:
        return None
    with get_connection() as conn:
        return conn.execute(
            "SELECT * FROM candidate_cvs WHERE file_hash = ?",
            (hash_value,)).fetchone()


def add_candidate_cv(candidate_id, data: dict) -> int:
    """Thêm một bản CV và trỏ `candidates.latest_cv_id` sang bản này.

    Ảnh chụp hồ sơ (chức danh, số năm, kỹ năng…) được CHÉP lên bảng `candidates`
    để lọc nhanh — bản gốc theo từng thời điểm vẫn nằm nguyên ở đây.
    """
    data = dict(data)
    data["candidate_id"] = candidate_id
    data.setdefault("received_at", date.today().isoformat())
    with get_connection() as conn:
        cv_id = _insert_conn(conn, "candidate_cvs", CANDIDATE_CV_FIELDS, data)
        snapshot = {k: data[k] for k in PROFILE_SNAPSHOT_FIELDS if k in data}
        snapshot["latest_cv_id"] = cv_id
        _update_conn(conn, "candidates", CANDIDATE_FIELDS, candidate_id, snapshot)
        _sync_fts(conn, candidate_id)
    return cv_id


def set_cv_file_path(candidate_id, path) -> None:
    """Cập nhật đường dẫn file của bản CV mới nhất (khi file bị di chuyển)."""
    with get_connection() as conn:
        row = conn.execute(
            "SELECT latest_cv_id FROM candidates WHERE candidate_id = ?",
            (candidate_id,)).fetchone()
        cv_id = row["latest_cv_id"] if row else None
        if cv_id:
            conn.execute(
                "UPDATE candidate_cvs SET file_path = ?, "
                "updated_at = datetime('now', 'localtime') WHERE cv_id = ?",
                (path, cv_id))
        else:
            _insert_conn(conn, "candidate_cvs", CANDIDATE_CV_FIELDS,
                         {"candidate_id": candidate_id, "file_path": path,
                          "received_at": date.today().isoformat()})
            conn.execute(
                "UPDATE candidates SET latest_cv_id = last_insert_rowid() "
                "WHERE candidate_id = ?", (candidate_id,))


def candidate_cv_path(candidate_id) -> str:
    """Đường dẫn file CV mới nhất của một ứng viên ("" nếu chưa có)."""
    with get_connection() as conn:
        row = conn.execute(
            "SELECT cv.file_path FROM candidates c "
            "LEFT JOIN candidate_cvs cv ON cv.cv_id = c.latest_cv_id "
            "WHERE c.candidate_id = ?", (candidate_id,)).fetchone()
    return (row["file_path"] or "").strip() if row and row["file_path"] else ""


# ════════════ DÒNG THỜI GIAN CÔNG VIỆC (candidate_experiences) ══════════

def list_candidate_experiences(candidate_id, cv_id=None):
    sql = "SELECT * FROM candidate_experiences WHERE candidate_id = ?"
    params: list = [candidate_id]
    if cv_id:
        sql += " AND cv_id = ?"
        params.append(cv_id)
    sql += " ORDER BY COALESCE(sort_order, 999), COALESCE(start_date, '') DESC"
    with get_connection() as conn:
        return conn.execute(sql, params).fetchall()


def replace_candidate_experiences(candidate_id, cv_id, items) -> int:
    """Ghi lại toàn bộ kinh nghiệm bóc từ MỘT bản CV (xóa bản cũ của cv đó).

    `items` là list dict theo CANDIDATE_EXPERIENCE_FIELDS. Trả về số dòng đã ghi.
    """
    with get_connection() as conn:
        conn.execute("DELETE FROM candidate_experiences WHERE cv_id = ?", (cv_id,))
        for i, item in enumerate(items or []):
            row = dict(item)
            row["candidate_id"] = candidate_id
            row["cv_id"] = cv_id
            row.setdefault("sort_order", i)
            _insert_conn(conn, "candidate_experiences",
                         CANDIDATE_EXPERIENCE_FIELDS, row)
    return len(items or [])


# Các dạng ngày hay gặp trong CV — thử lần lượt.
_DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%Y", "%Y/%m", "%Y-%m", "%Y"]


def parse_date(value):
    """Chuỗi ngày (nhiều định dạng) → datetime.date; None nếu không đọc được.

    CV hay chỉ ghi tháng/năm hoặc mỗi năm → quy về ngày 01 của kỳ đó.
    """
    text = str(value or "").strip()
    if not text:
        return None
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    # Chuỗi datetime đầy đủ ("2026-08-15 09:30:00") → cắt lấy phần ngày.
    try:
        return datetime.strptime(text[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def sync_experience_years(candidate_id, cv_id=None) -> dict:
    """Tính lại số năm kinh nghiệm TỪ DÒNG THỜI GIAN rồi ghi đè ảnh chụp.

    Con số AI tự ước lượng trong CV thường lệch với tổng các khoảng thời gian
    thật (bỏ quên thực tập, cộng nhầm giai đoạn chồng nhau). Dòng thời gian có
    ngày tháng cụ thể nên đáng tin hơn — lấy nó làm chuẩn để bảng danh sách và
    màn hình chi tiết luôn hiện cùng một con số.

    Không có dòng kinh nghiệm nào thì giữ nguyên con số AI đưa.
    """
    years = experience_years(candidate_id)
    if not years["at_cv"]:
        return years
    data = {"years_experience": years["at_cv"]}
    if years["as_of"]:
        data["experience_as_of"] = years["as_of"]
    update_candidate(candidate_id, data)
    if cv_id:
        with get_connection() as conn:
            _update_conn(conn, "candidate_cvs", CANDIDATE_CV_FIELDS, cv_id, data)
    return years


def _merge_spans(spans):
    """Gộp các khoảng thời gian chồng nhau → tổng số ngày làm việc thực tế.

    Làm hai việc cùng lúc ở hai công ty (part-time, freelance) thì KHÔNG cộng
    dồn hai lần.
    """
    spans = sorted((s, e) for s, e in spans if s and e and e > s)
    total, cur_s, cur_e = 0, None, None
    for s, e in spans:
        if cur_e is None:
            cur_s, cur_e = s, e
        elif s <= cur_e:
            cur_e = max(cur_e, e)
        else:
            total += (cur_e - cur_s).days
            cur_s, cur_e = s, e
    if cur_e is not None:
        total += (cur_e - cur_s).days
    return total / 365.25


def experience_years(candidate_id, row=None) -> dict:
    """Số năm kinh nghiệm của một ứng viên — HAI con số, không phải một.

    Trả về dict:
        at_cv  — số năm TẠI THỜI ĐIỂM CV. Chắc chắn đúng: CV nói vậy.
        today  — số năm ƯỚC TÍNH ĐẾN HÔM NAY. Chỉ cộng thêm khoảng từ ngày nhận
                 CV tới nay cho các việc lúc đó đang làm dở — không ai biết họ
                 có nhảy việc hay nghỉ hay không, nên đây là SUY ĐOÁN.
        as_of  — mốc thời gian của con số `at_cv` (chuỗi ngày, "" nếu không rõ).
        stale  — hồ sơ đã cũ quá STALE_PROFILE_MONTHS tháng.

    Ưu tiên tính từ `candidate_experiences` (dòng thời gian công việc); không có
    dòng nào thì lùi về cột `candidates.years_experience` + `experience_as_of`.
    """
    today = date.today()
    rows = list_candidate_experiences(candidate_id)
    snapshot = row if row is not None else get_candidate(candidate_id)

    as_of = None
    if snapshot is not None:
        as_of = parse_date(snapshot["experience_as_of"]
                           if "experience_as_of" in snapshot.keys() else None)

    if rows:
        spans_cv, spans_now = [], []
        for r in rows:
            start = parse_date(r["start_date"])
            if not start:
                continue
            ref = parse_date(r["as_of_date"]) or as_of or today
            end = parse_date(r["end_date"])
            spans_cv.append((start, end or ref))
            spans_now.append((start, end or today))
            if as_of is None or ref > as_of:
                as_of = ref
        at_cv, now = _merge_spans(spans_cv), _merge_spans(spans_now)
    else:
        try:
            at_cv = float(snapshot["years_experience"]) if snapshot is not None \
                and snapshot["years_experience"] not in (None, "") else 0.0
        except (TypeError, ValueError):
            at_cv = 0.0
        gap = ((today - as_of).days / 365.25) if as_of else 0.0
        now = at_cv + max(0.0, gap)

    months = ((today.year - as_of.year) * 12 + today.month - as_of.month) if as_of else 0
    return {
        "at_cv": round(at_cv, 1),
        "today": round(now, 1),
        "as_of": as_of.isoformat() if as_of else "",
        "stale": bool(as_of) and months > cv_schema.STALE_PROFILE_MONTHS,
    }


# ═══════════════ KỸ NĂNG CỦA ỨNG VIÊN (candidate_skills) ════════════════

def list_candidate_skills(candidate_id):
    with get_connection() as conn:
        return conn.execute(
            "SELECT cs.*, s.name AS skill_name, s.category "
            "FROM candidate_skills cs "
            "LEFT JOIN skills s ON s.skill_id = cs.skill_id "
            "WHERE cs.candidate_id = ? "
            "ORDER BY COALESCE(s.name, cs.raw_name)", (candidate_id,)).fetchall()


def replace_candidate_skills(candidate_id, cv_id, names, create_missing=True) -> int:
    """Ghi lại kỹ năng của ứng viên từ danh sách tên thô đọc trong CV.

    Mỗi tên được tra về `skills` (kể cả qua `aliases`); chưa có thì tự thêm vào
    danh mục khi `create_missing`. Chỉ mục duy nhất (candidate_id, skill_id)
    chặn trùng, nên gọi lại nhiều lần vẫn an toàn.
    """
    # Tra danh mục (và thêm kỹ năng mới) TRƯỚC khi mở kết nối ghi: `resolve_skill`
    # tự mở kết nối riêng, gọi lồng trong transaction sẽ bị "database is locked".
    lookup = skill_lookup()
    resolved, seen = [], set()
    for name in names or []:
        raw = " ".join(str(name).split())
        if not raw:
            continue
        sid = resolve_skill(raw, lookup, create=create_missing)
        key = sid if sid is not None else _skill_key(raw)
        if key in seen:
            continue
        seen.add(key)
        resolved.append((raw, sid))

    with get_connection() as conn:
        conn.execute("DELETE FROM candidate_skills WHERE candidate_id = ?",
                     (candidate_id,))
        for raw, sid in resolved:
            _insert_conn(conn, "candidate_skills", CANDIDATE_SKILL_FIELDS, {
                "candidate_id": candidate_id, "skill_id": sid, "cv_id": cv_id,
                "raw_name": raw, "source": "CV",
            })
    return len(resolved)


# ═════════════════════ ĐƠN ỨNG TUYỂN (applications) ═════════════════════

_APPLICATION_SELECT = """
SELECT a.*, c.full_name, c.email, c.phone, c.pool_status,
       p.position_title, p.department_id, d.department_name,
       cv.file_path AS cv_file_path, cv.batch
FROM applications a
LEFT JOIN candidates c    ON c.candidate_id = a.candidate_id
LEFT JOIN positions p     ON p.position_id = a.position_id
LEFT JOIN departments d   ON d.department_id = p.department_id
LEFT JOIN candidate_cvs cv ON cv.cv_id = COALESCE(a.cv_id, c.latest_cv_id)
"""


def list_applications(candidate_id=None, position_id=None, status: str = ""):
    """Danh sách đơn ứng tuyển (mới nhất trước), lọc theo ứng viên / vị trí."""
    sql = [_APPLICATION_SELECT, "WHERE 1=1"]
    params: list = []
    if candidate_id:
        sql.append("AND a.candidate_id = ?")
        params.append(candidate_id)
    if position_id:
        sql.append("AND a.position_id = ?")
        params.append(position_id)
    if status:
        sql.append("AND a.status = ?")
        params.append(status)
    sql.append("ORDER BY COALESCE(a.applied_at, '') DESC, a.application_id DESC")
    with get_connection() as conn:
        return conn.execute(" ".join(sql), params).fetchall()


def get_application(application_id):
    return _get("applications", application_id)


def latest_application(candidate_id, position_id=None):
    """Đơn mới nhất của một ứng viên (tùy chọn: cho đúng một vị trí)."""
    rows = list_applications(candidate_id=candidate_id, position_id=position_id)
    return rows[0] if rows else None


def insert_application(data: dict) -> int:
    data = dict(data)
    data.setdefault("status", cv_schema.CANDIDATE_STATUS_DEFAULT)
    data.setdefault("final_status", cv_schema.FINAL_STATUS_DEFAULT)
    data.setdefault("applied_at", _now())
    data.setdefault("status_changed_at", data["applied_at"])
    return _insert("applications", APPLICATION_FIELDS, data)


def update_application(application_id, data: dict) -> None:
    _update("applications", APPLICATION_FIELDS, application_id, data)


def delete_application(application_id) -> None:
    """Xóa đơn kèm các buổi phỏng vấn & nhận xét của đơn đó."""
    with get_connection() as conn:
        conn.execute(
            "DELETE FROM interview_feedbacks WHERE interview_id IN "
            "(SELECT interview_id FROM interviews WHERE application_id = ?)",
            (application_id,))
        conn.execute("DELETE FROM interviews WHERE application_id = ?",
                     (application_id,))
        conn.execute("DELETE FROM applications WHERE application_id = ?",
                     (application_id,))


def ensure_application(candidate_id, position_id, data: dict | None = None) -> int:
    """Lấy đơn đang mở của (ứng viên, vị trí) — chưa có thì tạo mới.

    "Đang mở" = chưa vào nhánh dừng (`closed_at` rỗng). Nhờ vậy quét lại CV của
    một người cho cùng vị trí không sinh thêm đơn, còn ứng tuyển lại sau khi đã
    đóng đơn cũ thì tạo đơn mới, giữ nguyên lịch sử.
    """
    with get_connection() as conn:
        row = conn.execute(
            "SELECT application_id FROM applications "
            "WHERE candidate_id = ? AND position_id = ? "
            "  AND COALESCE(closed_at, '') = '' "
            "ORDER BY application_id DESC LIMIT 1",
            (candidate_id, position_id)).fetchone()
    if row:
        return row["application_id"]
    payload = dict(data or {})
    payload.update({"candidate_id": candidate_id, "position_id": position_id})
    return insert_application(payload)


def set_application_status(application_id, status: str, note: str = "") -> None:
    """Đổi trạng thái đơn + GHI LẠI vào lịch sử (candidate_activities).

    Vào nhánh dừng (không còn bước kế tiếp) thì đóng đơn luôn — `closed_at`.
    """
    app = get_application(application_id)
    if app is None:
        return
    old = app["status"] or ""
    now = _now()
    data = {"status": status, "status_changed_at": now}
    if not cv_schema.candidate_next_status(status):
        data["closed_at"] = now
    update_application(application_id, data)
    log_activity({
        "candidate_id": app["candidate_id"],
        "application_id": application_id,
        "type": "Status change",
        "occurred_at": now,
        "from_status": old,
        "to_status": status,
        "content": note,
    })


# ═══════════════════════ PHỎNG VẤN (interviews) ═════════════════════════

def list_interviews(application_id=None, candidate_id=None):
    """Các buổi phỏng vấn, kèm SỐ người đã gửi nhận xét cho từng buổi."""
    sql = ["SELECT i.*, p.position_title,",
           "  (SELECT COUNT(*) FROM interview_feedbacks f",
           "   WHERE f.interview_id = i.interview_id) AS feedback_count",
           "FROM interviews i",
           "LEFT JOIN applications a ON a.application_id = i.application_id",
           "LEFT JOIN positions p    ON p.position_id = a.position_id",
           "WHERE 1=1"]
    params: list = []
    if application_id:
        sql.append("AND i.application_id = ?")
        params.append(application_id)
    if candidate_id:
        sql.append("AND i.candidate_id = ?")
        params.append(candidate_id)
    sql.append("ORDER BY COALESCE(i.interview_date, '') DESC, i.round DESC")
    with get_connection() as conn:
        return conn.execute(" ".join(sql), params).fetchall()


def get_interview(interview_id):
    return _get("interviews", interview_id)


def insert_interview(data: dict) -> int:
    return _insert("interviews", INTERVIEW_FIELDS, data)


def update_interview(interview_id, data: dict) -> None:
    _update("interviews", INTERVIEW_FIELDS, interview_id, data)


def delete_interview(interview_id) -> None:
    with get_connection() as conn:
        conn.execute("DELETE FROM interview_feedbacks WHERE interview_id = ?",
                     (interview_id,))
        conn.execute("DELETE FROM interviews WHERE interview_id = ?", (interview_id,))


def save_interview(application_id, round_no, data: dict) -> int:
    """Ghi buổi phỏng vấn của MỘT VÒNG (chưa có thì tạo, có rồi thì cập nhật).

    Chỉ mục duy nhất (application_id, round) bảo đảm mỗi vòng đúng một dòng.
    """
    app = get_application(application_id)
    with get_connection() as conn:
        row = conn.execute(
            "SELECT interview_id FROM interviews "
            "WHERE application_id = ? AND round = ?",
            (application_id, round_no)).fetchone()
    payload = dict(data)
    payload.update({"application_id": application_id, "round": round_no})
    if app is not None:
        payload.setdefault("candidate_id", app["candidate_id"])
    if row:
        update_interview(row["interview_id"], payload)
        return row["interview_id"]
    return insert_interview(payload)


def list_interview_feedbacks(interview_id):
    """Nhận xét của từng người phỏng vấn, kèm tên lấy từ `employees` nếu có."""
    with get_connection() as conn:
        return conn.execute(
            "SELECT f.*, "
            "  COALESCE(NULLIF(TRIM(e.full_name), ''), f.interviewer_name) AS display_name, "
            "  e.job_title, d.department_name "
            "FROM interview_feedbacks f "
            "LEFT JOIN employees e   ON e.employee_id = f.employee_id "
            "LEFT JOIN departments d ON d.department_id = e.department_id "
            "WHERE f.interview_id = ? "
            "ORDER BY f.feedback_id", (interview_id,)).fetchall()


# Hai hàm dưới đây là bản "cả nhóm" của list_interviews / list_interview_feedbacks,
# dùng khi xuất Excel hàng loạt: hỏi một lần cho mọi ứng viên đang chọn thay vì
# lặp N lần. Chia mẻ vì SQLite giới hạn số tham số của một câu lệnh.
_SQL_VAR_LIMIT = 500


def _chunks(values, size=_SQL_VAR_LIMIT):
    seq = list(values)
    for i in range(0, len(seq), size):
        yield seq[i:i + size]


def list_interviews_by_candidates(candidate_ids):
    """Mọi buổi phỏng vấn của một NHÓM ứng viên, mỗi vòng MỚI NHẤT trước."""
    rows = []
    with get_connection() as conn:
        for chunk in _chunks(candidate_ids):
            marks = ",".join("?" * len(chunk))
            rows += conn.execute(
                "SELECT i.* FROM interviews i "
                f"WHERE i.candidate_id IN ({marks}) "
                "ORDER BY i.candidate_id, i.round, "
                "         COALESCE(i.interview_date, '') DESC, i.interview_id DESC",
                list(chunk)).fetchall()
    return rows


def list_feedbacks_by_interviews(interview_ids):
    """Nhận xét của MỌI buổi phỏng vấn trong danh sách, kèm tên người phỏng vấn."""
    rows = []
    with get_connection() as conn:
        for chunk in _chunks(interview_ids):
            marks = ",".join("?" * len(chunk))
            rows += conn.execute(
                "SELECT f.*, "
                "  COALESCE(NULLIF(TRIM(e.full_name), ''), f.interviewer_name) AS display_name, "
                "  e.job_title "
                "FROM interview_feedbacks f "
                "LEFT JOIN employees e ON e.employee_id = f.employee_id "
                f"WHERE f.interview_id IN ({marks}) "
                "ORDER BY f.interview_id, f.feedback_id", list(chunk)).fetchall()
    return rows


def save_interview_feedbacks(interview_id, entries) -> None:
    """Đặt LẠI toàn bộ nhận xét của một buổi phỏng vấn cho khớp `entries`.

    `entries` là list dict theo INTERVIEW_FEEDBACK_FIELDS; dict nào có sẵn
    `feedback_id` thì được CẬP NHẬT (giữ nguyên `submitted_at` cũ), không có thì
    chèn mới. Dòng đang nằm trong DB mà không còn trong `entries` bị xóa — form
    nhập là nguồn sự thật cho buổi đó, nên xóa người khỏi form là xóa luôn nhận
    xét của họ.
    """
    keep = {e["feedback_id"] for e in entries if e.get("feedback_id")}
    with get_connection() as conn:
        old = {r["feedback_id"] for r in conn.execute(
            "SELECT feedback_id FROM interview_feedbacks WHERE interview_id = ?",
            (interview_id,))}
        for feedback_id in old - keep:
            conn.execute("DELETE FROM interview_feedbacks WHERE feedback_id = ?",
                         (feedback_id,))
        for entry in entries:
            data = dict(entry)
            data["interview_id"] = interview_id
            feedback_id = data.pop("feedback_id", None)
            if feedback_id in old:
                _update_conn(conn, "interview_feedbacks",
                             INTERVIEW_FEEDBACK_FIELDS, feedback_id, data)
            else:
                data.setdefault("submitted_at", _now())
                _insert_conn(conn, "interview_feedbacks",
                             INTERVIEW_FEEDBACK_FIELDS, data)


def get_interview_feedback(feedback_id):
    return _get("interview_feedbacks", feedback_id)


def insert_interview_feedback(data: dict) -> int:
    data = dict(data)
    data.setdefault("submitted_at", _now())
    return _insert("interview_feedbacks", INTERVIEW_FEEDBACK_FIELDS, data)


def update_interview_feedback(feedback_id, data: dict) -> None:
    _update("interview_feedbacks", INTERVIEW_FEEDBACK_FIELDS, feedback_id, data)


def delete_interview_feedback(feedback_id) -> None:
    _delete("interview_feedbacks", feedback_id)


def list_interviewers():
    """Nhân viên có thể mời phỏng vấn: đang làm việc, ưu tiên người đã đánh dấu.

    `is_interviewer = 1` là cờ đánh dấu tay ở màn hình Employees. Không có ai
    được đánh dấu thì trả về toàn bộ nhân viên đang làm việc.
    """
    with get_connection() as conn:
        marked = conn.execute(
            "SELECT e.employee_id, e.full_name, e.job_title, d.department_name "
            "FROM employees e LEFT JOIN departments d "
            "  ON d.department_id = e.department_id "
            "WHERE COALESCE(e.is_interviewer, 0) = 1 "
            f"  AND {_EXCLUDE_RESIGNED_SQL} "
            "ORDER BY e.full_name").fetchall()
        if marked:
            return marked
        return conn.execute(
            "SELECT e.employee_id, e.full_name, e.job_title, d.department_name "
            "FROM employees e LEFT JOIN departments d "
            "  ON d.department_id = e.department_id "
            f"WHERE {_EXCLUDE_RESIGNED_SQL} "
            "ORDER BY e.full_name").fetchall()


# ═════════════ LỊCH SỬ AI CHẤM ĐIỂM (candidate_evaluations) ═════════════

def list_evaluations(candidate_id=None, position_id=None, limit=0):
    """Lịch sử chấm điểm, MỚI NHẤT trước. Không bao giờ ghi đè nên có bao nhiêu
    lượt chấm thì trả về bấy nhiêu dòng."""
    sql = ["SELECT ev.*, p.position_title, cv.received_at AS cv_received_at,",
           "       cv.file_path AS cv_file_path",
           "FROM candidate_evaluations ev",
           "LEFT JOIN positions p      ON p.position_id = ev.position_id",
           "LEFT JOIN candidate_cvs cv ON cv.cv_id = ev.cv_id",
           "WHERE 1=1"]
    params: list = []
    if candidate_id:
        sql.append("AND ev.candidate_id = ?")
        params.append(candidate_id)
    if position_id:
        sql.append("AND ev.position_id = ?")
        params.append(position_id)
    sql.append("ORDER BY COALESCE(ev.evaluated_at, '') DESC, ev.evaluation_id DESC")
    if limit:
        sql.append(f"LIMIT {int(limit)}")
    with get_connection() as conn:
        return conn.execute(" ".join(sql), params).fetchall()


def latest_evaluation(candidate_id, position_id=None):
    rows = list_evaluations(candidate_id, position_id, limit=1)
    return rows[0] if rows else None


def insert_evaluation(data: dict) -> int:
    """Ghi THÊM một lượt chấm. Không bao giờ sửa/xóa dòng cũ — đó là lịch sử."""
    data = dict(data)
    data.setdefault("evaluated_at", _now())
    return _insert("candidate_evaluations", EVALUATION_FIELDS, data)


def count_evaluations(candidate_id) -> int:
    with get_connection() as conn:
        return conn.execute(
            "SELECT COUNT(*) FROM candidate_evaluations WHERE candidate_id = ?",
            (candidate_id,)).fetchone()[0]


# ═══════════ LỊCH SỬ LIÊN HỆ (candidate_activities) ═════════════════════

def list_activities(candidate_id=None, application_id=None, limit=0):
    sql = ["SELECT ac.*, p.position_title, t.name AS mail_template_name",
           "FROM candidate_activities ac",
           "LEFT JOIN applications a   ON a.application_id = ac.application_id",
           "LEFT JOIN positions p      ON p.position_id = a.position_id",
           "LEFT JOIN mail_templates t ON t.mail_template_id = ac.mail_template_id",
           "WHERE 1=1"]
    params: list = []
    if candidate_id:
        sql.append("AND ac.candidate_id = ?")
        params.append(candidate_id)
    if application_id:
        sql.append("AND ac.application_id = ?")
        params.append(application_id)
    sql.append("ORDER BY COALESCE(ac.occurred_at, '') DESC, ac.activity_id DESC")
    if limit:
        sql.append(f"LIMIT {int(limit)}")
    with get_connection() as conn:
        return conn.execute(" ".join(sql), params).fetchall()


def log_activity(data: dict) -> int:
    """Ghi một sự kiện liên hệ + cập nhật `candidates.last_contacted_at`.

    Chỉ các loại thực sự CHẠM tới ứng viên (Email/Call) mới dời mốc liên hệ
    gần nhất — đổi trạng thái hay ghi chú nội bộ thì không.
    """
    data = dict(data)
    data.setdefault("occurred_at", _now())
    with get_connection() as conn:
        act_id = _insert_conn(conn, "candidate_activities", ACTIVITY_FIELDS, data)
        if data.get("type") in ("Email", "Call") and data.get("candidate_id"):
            conn.execute(
                "UPDATE candidates SET last_contacted_at = ?, "
                "updated_at = datetime('now', 'localtime') WHERE candidate_id = ?",
                (data["occurred_at"], data["candidate_id"]))
    return act_id


def count_activities(candidate_id) -> int:
    with get_connection() as conn:
        return conn.execute(
            "SELECT COUNT(*) FROM candidate_activities WHERE candidate_id = ?",
            (candidate_id,)).fetchone()[0]


# ═════════════════════════════ NHÂN VIÊN ════════════════════════════════

# Ô "Emergency Contact Name" trong "Master HC file.xlsx" gộp cả tên lẫn số điện
# thoại. Dữ liệu thật có đủ kiểu ngăn cách nên KHÔNG dò theo dấu ngăn mà dò cụm
# CHỈ GỒM chữ số + ký hiệu điện thoại (khoảng trắng, ngoặc, +, /, gạch, chấm)
# nằm ở ĐẦU hoặc CUỐI chuỗi — bắt được cả 3 kiểu đang có trong file:
#     "Nguyễn Văn A ⏎ 0903 991 962"  ·  "Nguyễn Văn A (0913484647)"
#     "0367842223 - Nguyễn Văn A"
_CONTACT_PHONE_TAIL_RE = re.compile(r"[\d(+][\d\s().+/\-]*[\d)]\s*$")
_CONTACT_PHONE_HEAD_RE = re.compile(r"^[\d(+][\d\s().+/\-]*[\d)]")
# Ký tự thừa còn sót ở hai đầu phần TÊN sau khi cắt SĐT ("Nguyễn Văn A-" → "…A").
_CONTACT_SEP_CHARS = " \t\r\n-–—,;:/|"
# Số chữ số tối thiểu để coi cụm đó là số điện thoại (số bàn ngắn nhất ~8 chữ
# số; để 6 cho rộng). Ít hơn → coi như một phần của tên (vd ô chỉ ghi "0").
_CONTACT_PHONE_MIN_DIGITS = 6


def split_contact_name_phone(text) -> tuple[str, str]:
    """Tách ô liên hệ khẩn cấp gộp "tên + số ĐT" → (tên, số ĐT).

    Không thấy số điện thoại thì trả về (cả chuỗi, ""). Nhiều số ngăn bởi "/"
    hoặc "," được gộp lại bằng "; " cho giống quy ước cột `phone`.
    """
    s = " ".join(str(text or "").split(" ")).strip()
    if not s:
        return "", ""
    m = _CONTACT_PHONE_TAIL_RE.search(s)
    if m and _count_digits(m.group()) >= _CONTACT_PHONE_MIN_DIGITS:
        return _tidy_name(s[:m.start()]), _tidy_phone(m.group())
    m = _CONTACT_PHONE_HEAD_RE.match(s)
    if m and _count_digits(m.group()) >= _CONTACT_PHONE_MIN_DIGITS:
        return _tidy_name(s[m.end():]), _tidy_phone(m.group())
    return _tidy_name(s), ""


def _count_digits(text: str) -> int:
    return sum(c.isdigit() for c in text)


def _tidy_name(text: str) -> str:
    return " ".join(text.strip(_CONTACT_SEP_CHARS).split())


def _tidy_phone(text: str) -> str:
    """Gộp nhiều số về "a; b"; bỏ cặp ngoặc bọc cả số ("(0913484647)").

    Chỉ bỏ khi ngoặc bọc TRỌN số — giữ nguyên mã vùng kiểu "(028) 3736 2323".
    """
    parts = []
    for part in re.split(r"[/,;]+", text):
        p = " ".join(part.split())
        if p.startswith("(") and p.endswith(")") and "(" not in p[1:-1]:
            p = p[1:-1].strip()
        if p:
            parts.append(p)
    return "; ".join(parts)


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
    """Tìm nhân viên ĐÃ CÓ trùng với một hồ sơ sắp thêm — dùng khi nhập từng
    người một (thường chưa có mã NV).

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


# ═══════════════════════ KHÓA HỌC / ĐÀO TẠO ════════════════════════════

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
    """Ghi danh 1 nhân viên vào 1 khóa học. Bỏ qua nếu đã ghi danh (nhờ unique
    index course_id+employee_id). Trả về enrollment_id (0 nếu đã có sẵn).
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
        return cur.lastrowid or 0


def unenroll_employee(course_id, employee_id) -> None:
    with get_connection() as conn:
        conn.execute(
            "DELETE FROM course_employees WHERE course_id = ? AND employee_id = ?",
            (course_id, employee_id))


def update_enrollment(enrollment_id, data: dict) -> None:
    _update("course_employees", COURSE_EMPLOYEE_FIELDS, enrollment_id, data)


def get_enrollment(enrollment_id):
    return _get("course_employees", enrollment_id)


def search_course_employees(course_id=None, status: str = ""):
    """Lượt ghi danh, lọc theo khóa học / trạng thái học."""
    sql = [
        "SELECT ce.*, c.title AS course_title, c.date AS course_date,",
        "       c.course_type, e.code, e.full_name, e.email,",
        "       d.department_name",
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
    sql.append("ORDER BY c.date DESC, e.full_name")
    with get_connection() as conn:
        return conn.execute(" ".join(sql), params).fetchall()


def count_course_employees() -> int:
    with get_connection() as conn:
        return conn.execute("SELECT COUNT(*) FROM course_employees").fetchone()[0]
