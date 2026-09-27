"""Đẩy bản sao DB lên thư mục dùng chung (ổ mạng) — chiều ĐẨY LÊN của đồng bộ.

Mỗi máy làm việc trên file .db ở ổ local (`cv_repository._db_path()`); thư mục
dùng chung chỉ giữ BẢN CHỦ + lịch sử, không máy nào mở SQLite trực tiếp trên ổ
mạng (SQLite qua SMB/VPN là hỏng file). Bố cục thư mục dùng chung:

    <shared_db_folder>\\
    ├── candidates.sqlite    ← bản chủ
    ├── state.json           ← version · updated_by · updated_at · schema_version
    └── history\\
        └── candidates-20260927-091530-<tài khoản windows>.sqlite

Một lượt `push()`:
    1. Chụp DB local ra file tạm ở ổ LOCAL bằng `sqlite3.Connection.backup()`
       (nhất quán kể cả khi đang có lượt ghi dở — copy file thô thì không), ghi
       `sync:version` mới vào `app_meta` của bản chụp rồi `quick_check`.
    2. Copy bản chụp lên thư mục dùng chung dưới tên tạm, `os.replace()` đè lên
       `candidates.sqlite` → rớt mạng giữa chừng thì bản chủ cũ còn nguyên.
    3. Ghi `state.json` (cũng qua tên tạm + replace), lưu một bản vào
       `history\\`, xóa bớt bản cũ.
    4. Ghi `sync:version` vào DB local — chỉ sau khi bản chủ đã thay xong, để
       version local không bao giờ vượt bản chủ.

`version` nằm ở CẢ `state.json` lẫn `app_meta` của file .db: file .db được copy
đi đâu thì version đi theo đó.

Đây là backup chạy ngầm: thư mục không với tới được (chưa bật VPN, ổ chưa
mount) hay bất cứ lỗi nào khác đều chỉ ghi log và trả về `PushResult`, KHÔNG
ném lỗi ra ngoài. Setting `shared_db_folder` để trống = tắt hẳn tính năng.
"""
import datetime
import glob
import json
import os
import pathlib
import shutil
import sqlite3
import tempfile
import threading
from dataclasses import dataclass

from app.core import cv_repository, debuglog, settings

DB_NAME = "candidates.sqlite"
STATE_NAME = "state.json"
HISTORY_DIR = "history"
# Số bản giữ lại trong `history\` (~6 MB mỗi bản).
HISTORY_KEEP = 30

VERSION_KEY = "sync:version"

# Hai lượt đẩy cùng lúc (nút Back up now + đóng app) sẽ tranh nhau cùng file
# tạm và cùng số version → chạy lần lượt.
_LOCK = threading.Lock()


@dataclass
class PushResult:
    ok: bool                 # đã thay bản chủ thành công
    skipped: bool = False    # không làm gì vì tính năng tắt / thư mục không với tới
    message: str = ""        # mô tả ngắn (tiếng Anh — hiện được lên giao diện)
    version: int | None = None


def shared_folder() -> str:
    """Thư mục dùng chung đã cấu hình ("" = tính năng đang tắt)."""
    return str(settings.get("shared_db_folder", "") or "").strip()


def is_enabled() -> bool:
    return bool(shared_folder())


def read_state(folder: str) -> dict:
    """Đọc `state.json` của thư mục dùng chung ({} nếu chưa có / hỏng)."""
    try:
        with open(os.path.join(folder, STATE_NAME), encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def push(reason: str, folder: str | None = None) -> PushResult:
    """Đẩy DB local lên thư mục dùng chung. Không bao giờ ném lỗi.

    `reason` chỉ để ghi log + `state.json` (vd "close", "manual"). `folder` bỏ
    trống thì lấy từ setting `shared_db_folder`.
    """
    folder = (shared_folder() if folder is None else str(folder)).strip()
    if not folder:
        return PushResult(ok=False, skipped=True,
                          message="No shared database folder is set.")
    with _LOCK:
        try:
            result = _push(reason, folder)
        except Exception as exc:
            debuglog.exception(f"shared_db.push({reason}) failed — {folder}", exc)
            result = PushResult(ok=False, message=f"Backup failed: {exc}")
    if result.ok or result.skipped:
        debuglog.write(f"shared_db.push({reason}): {result.message}")
    return result


def push_with_timeout(reason: str, timeout: float) -> PushResult | None:
    """`push()` ở luồng phụ, chờ tối đa `timeout` giây (None = hết giờ).

    Dùng lúc đóng app: ổ mạng treo (VPN chập chờn) thì thao tác file có thể
    đứng hàng chục giây, không được giữ app lại vô thời hạn. Luồng là daemon
    nên hết giờ thì app thoát luôn — bản chủ không hỏng vì chỉ được thay bằng
    `os.replace()` sau khi file tạm đã ghi đủ.
    """
    box = []
    worker = threading.Thread(target=lambda: box.append(push(reason)),
                              name="shared-db-push", daemon=True)
    worker.start()
    worker.join(timeout)
    if worker.is_alive():
        debuglog.write(f"shared_db.push({reason}): gave up after {timeout:.0f}s "
                       f"— the shared folder is not responding")
        return None
    return box[0] if box else None


def _push(reason: str, folder: str) -> PushResult:
    # Không tự tạo thư mục gốc: ổ mạng chưa mount mà tạo thì hoặc lỗi, hoặc
    # tạo nhầm một thư mục local cùng tên.
    if not os.path.isdir(folder):
        return PushResult(ok=False, skipped=True,
                          message=f"Shared folder not reachable: {folder}")
    src = cv_repository._db_path()
    if not os.path.exists(src):
        return PushResult(ok=False, skipped=True,
                          message="No local database yet — nothing to back up.")

    login = cv_repository.windows_login() or "unknown"
    local_version = _read_version(src)
    remote_version = _as_int(read_state(folder).get("version"))
    version = max(local_version, remote_version) + 1

    work_dir = tempfile.mkdtemp(prefix="ptb-push-")
    try:
        snapshot = os.path.join(work_dir, DB_NAME)
        schema_version = _snapshot(src, snapshot, version)

        # Bản chủ: copy lên tên tạm RIÊNG của máy này rồi mới thay nguyên tử.
        master = os.path.join(folder, DB_NAME)
        tmp_master = os.path.join(folder, f".{DB_NAME}.{login}.tmp")
        shutil.copyfile(snapshot, tmp_master)
        if os.path.getsize(tmp_master) != os.path.getsize(snapshot):
            _remove_quietly(tmp_master)
            raise OSError("uploaded copy has the wrong size")
        os.replace(tmp_master, master)

        now = datetime.datetime.now()
        _write_json_atomic(os.path.join(folder, STATE_NAME), {
            "version": version,
            "updated_by": login,
            "updated_at": now.isoformat(timespec="seconds"),
            "schema_version": schema_version,
            "reason": reason,
        }, login)

        history_note = _save_history(folder, snapshot, now, login)
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

    _write_version(src, version)
    return PushResult(ok=True, version=version,
                      message=f"Backed up as version {version} to {folder}"
                              f"{history_note}")


def _snapshot(src: str, dest: str, version: int) -> str:
    """Chụp `src` ra `dest` bằng backup API, đóng dấu version → schema_version.

    Nguồn mở chế độ chỉ đọc: lượt đẩy không được đụng vào DB đang dùng.
    """
    source = sqlite3.connect(pathlib.Path(src).resolve().as_uri() + "?mode=ro",
                             uri=True)
    try:
        target = sqlite3.connect(dest)
        try:
            source.backup(target)
            target.execute("CREATE TABLE IF NOT EXISTS app_meta "
                           "(key VARCHAR PRIMARY KEY, value VARCHAR)")
            target.execute("INSERT OR REPLACE INTO app_meta (key, value) "
                           "VALUES (?, ?)", (VERSION_KEY, str(version)))
            target.commit()
            check = target.execute("PRAGMA quick_check").fetchone()[0]
            if check != "ok":
                raise sqlite3.DatabaseError(f"snapshot failed quick_check: {check}")
            return _schema_version(target)
        finally:
            target.close()
    finally:
        source.close()


def _schema_version(conn) -> str:
    """Tên lượt migration mới nhất mà DB đã chạy (vd "0009_users_and_audit_columns").

    Tên có tiền tố số đệm 0 nên so sánh chuỗi cũng đúng thứ tự.
    """
    row = conn.execute("SELECT MAX(key) FROM app_meta "
                       "WHERE key LIKE 'migration:%'").fetchone()
    return (row[0] or "")[len("migration:"):] if row else ""


def _read_version(db_path: str) -> int:
    try:
        conn = sqlite3.connect(db_path)
        try:
            row = conn.execute("SELECT value FROM app_meta WHERE key = ?",
                               (VERSION_KEY,)).fetchone()
        finally:
            conn.close()
    except sqlite3.Error:
        return 0
    return _as_int(row[0]) if row else 0


def _write_version(db_path: str, version: int) -> None:
    """Ghi version vào DB local. Lỗi ở đây không làm hỏng lượt đẩy đã xong —
    lượt sau vẫn lấy max(local, state.json) + 1 nên version không lùi."""
    try:
        conn = sqlite3.connect(db_path, timeout=10)
        try:
            conn.execute("INSERT OR REPLACE INTO app_meta (key, value) "
                         "VALUES (?, ?)", (VERSION_KEY, str(version)))
            conn.commit()
        finally:
            conn.close()
    except sqlite3.Error as exc:
        debuglog.exception("shared_db: could not store sync:version locally", exc)


def _write_json_atomic(path: str, data: dict, login: str) -> None:
    tmp = f"{path}.{login}.tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def _save_history(folder: str, snapshot: str, now, login: str) -> str:
    """Lưu một bản vào `history\\` rồi xóa bớt, chỉ giữ HISTORY_KEEP bản mới nhất.

    Bản chủ đã thay xong trước bước này, nên lỗi ở đây chỉ ghi log.
    """
    try:
        hist = os.path.join(folder, HISTORY_DIR)
        os.makedirs(hist, exist_ok=True)
        safe_login = "".join(c if c.isalnum() or c in "-_." else "_" for c in login)
        name = f"candidates-{now:%Y%m%d-%H%M%S}-{safe_login}.sqlite"
        tmp = os.path.join(hist, f".{name}.tmp")
        shutil.copyfile(snapshot, tmp)
        os.replace(tmp, os.path.join(hist, name))

        # Tên bắt đầu bằng mốc thời gian → sắp theo tên là sắp theo thời gian.
        files = sorted(glob.glob(os.path.join(hist, "candidates-*.sqlite")))
        for old in files[:-HISTORY_KEEP]:
            _remove_quietly(old)
        return ""
    except OSError as exc:
        debuglog.exception("shared_db: could not write history copy", exc)
        return " (history copy failed — see debug.log)"


def _remove_quietly(path: str) -> None:
    try:
        os.remove(path)
    except OSError:
        pass


def _as_int(value) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0
