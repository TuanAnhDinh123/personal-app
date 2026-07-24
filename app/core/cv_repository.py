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
DEPARTMENT_FIELDS = ["department_name", "manager_name", "description"]
POSITION_FIELDS = ["department_id", "position_code", "position_title", "level",
                   "headcount", "status", "mail_cc", "mail_subject", "mail_body"]
JD_FIELDS = ["position_id", "jd_title", "jd_file_path"]
CANDIDATE_FIELDS = [
    "full_name", "email", "phone", "date_of_birth", "address",
    "position_id", "years_experience", "education", "applied_at", "status",
    "source", "batch", "fit_score", "fit_summary", "strengths", "weaknesses",
    "cv_file_path", "note",
]

# PK của mỗi bảng (dùng cho update/delete generic).
_PK = {
    "departments": "department_id",
    "positions": "position_id",
    "job_descriptions": "jd_id",
    "candidates": "candidate_id",
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
        _migrate_document_files(conn)
        _backfill_timestamps(conn)


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


def _migrate_document_files(conn: sqlite3.Connection) -> None:
    """Bỏ bảng document_files (thiết kế cũ) → đưa đường dẫn về candidates & jd.

    Với DB đã lỡ tạo bảng document_files: chép đường dẫn/tên file của mỗi bản
    ghi về cột cv_file_path (theo candidate_id) hoặc jd_file_path (theo jd_id),
    rồi xóa hẳn bảng. Chạy an toàn nhiều lần (không có bảng thì bỏ qua).
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
    if "jd_id" in file_cols:
        conn.execute(f"""
            UPDATE job_descriptions SET jd_file_path = COALESCE((
                SELECT COALESCE(NULLIF(TRIM(f.{src}), ''), f.file_name)
                FROM document_files f
                WHERE f.jd_id = job_descriptions.jd_id
                ORDER BY f.file_id LIMIT 1
            ), jd_file_path)
            WHERE jd_file_path IS NULL OR TRIM(jd_file_path) = ''
        """)
    conn.execute("DROP TABLE document_files")


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


# ───────────────────────────── VỊ TRÍ ───────────────────────────────────

def list_positions():
    with get_connection() as conn:
        return conn.execute(
            "SELECT p.*, d.department_name "
            "FROM positions p "
            "LEFT JOIN departments d ON d.department_id = p.department_id "
            "ORDER BY p.position_title").fetchall()


def get_position(pos_id):
    return _get("positions", pos_id)


def insert_position(data: dict) -> int:
    return _insert("positions", POSITION_FIELDS, data)


def update_position(pos_id, data: dict) -> None:
    _update("positions", POSITION_FIELDS, pos_id, data)


def delete_position(pos_id) -> None:
    _delete("positions", pos_id)


# ───────────────────────── MÔ TẢ CÔNG VIỆC (JD) ─────────────────────────

def list_job_descriptions():
    with get_connection() as conn:
        return conn.execute(
            "SELECT j.*, p.position_title "
            "FROM job_descriptions j "
            "LEFT JOIN positions p ON p.position_id = j.position_id "
            "ORDER BY j.jd_title").fetchall()


def get_job_description(jd_id):
    return _get("job_descriptions", jd_id)


def insert_job_description(data: dict) -> int:
    return _insert("job_descriptions", JD_FIELDS, data)


def update_job_description(jd_id, data: dict) -> None:
    _update("job_descriptions", JD_FIELDS, jd_id, data)


def delete_job_description(jd_id) -> None:
    _delete("job_descriptions", jd_id)


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
