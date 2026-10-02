"""Đồng bộ DB local ↔ thư mục dùng chung (ổ mạng) bằng cách copy NGUYÊN FILE.

Mỗi máy làm việc trên file .db ở ổ local (`cv_repository._db_path()`); thư mục
dùng chung chỉ giữ BẢN CHỦ + lịch sử, không máy nào mở SQLite trực tiếp trên ổ
mạng (SQLite qua SMB/VPN là hỏng file). Bố cục thư mục dùng chung:

    <shared_db_folder>\\
    ├── candidates.sqlite    ← bản chủ
    ├── state.json           ← version · db_id · updated_by · updated_at · schema_version
    └── history\\
        └── candidates-20260927-091530-<tài khoản windows>.sqlite

Hai chiều:

  ĐẨY LÊN — `push()` (lúc đóng app, nút Back up now):
    1. Chụp DB local ra file tạm ở ổ LOCAL bằng `sqlite3.Connection.backup()`
       (nhất quán kể cả khi đang có lượt ghi dở — copy file thô thì không), ghi
       `sync:version` mới + `sync:db_id` vào `app_meta` của bản chụp rồi
       `quick_check`.
    2. Copy bản chụp lên thư mục dùng chung dưới tên tạm, `os.replace()` đè lên
       `candidates.sqlite` → rớt mạng giữa chừng thì bản chủ cũ còn nguyên.
    3. Ghi `state.json` (cũng qua tên tạm + replace), lưu một bản vào
       `history\\`, xóa bớt bản cũ.
    4. Ghi version + db_id vào DB local — chỉ sau khi bản chủ đã thay xong, để
       version local không bao giờ vượt bản chủ.

  KÉO VỀ — `ensure_fresh()` (lúc mở app, định kỳ, nút Refresh, và từ
  `cv_repository.get_connection()` qua `_on_connect`):
    1. Đọc `state.json` (vài trăm byte), so `version` với `sync:version` của DB
       local. Không mới hơn → xong, KHÔNG copy gì.
    2. Mới hơn → kiểm tra `schema_version` (bản chủ mới hơn app → không kéo,
       báo cập nhật app) và `db_id` (khác dòng dõi → dừng hẳn).
    3. Copy bản chủ về tên tạm CẠNH file DB local, kiểm lại chính file đó
       (`quick_check`, version, db_id, schema).
    4. `cv_repository.replace_db_file()` thay nguyên tử — CHỈ khi không còn
       kết nối nào mở. Đang có kết nối thì giữ file đã tải làm "bản chờ" và
       thay ở lượt sau, không tải lại.
    5. Chạy lại `init_db()` trên file mới (migration còn thiếu so với app đang
       chạy, dòng `users` của máy này).

`version` và `db_id` nằm ở CẢ `state.json` lẫn `app_meta` của file .db: file
.db được copy đi đâu thì hai giá trị đó đi theo.

**`db_id` là lưới chắn tai nạn nặng nhất**: một UUID sinh MỘT LẦN cho mỗi dòng
dõi DB (ở lượt đẩy đầu tiên). Local và bản chủ khác `db_id` nghĩa là hai cơ sở
dữ liệu khác nhau (vd máy B lỡ dùng app một mình trước khi cấu hình) → không
kéo, không đẩy; chọn giữ bản nào là việc của con người. Ngoại lệ duy nhất: DB
local CHƯA TỪNG đồng bộ và CHƯA CÓ dữ liệu người dùng (máy mới cài) thì được
nhận bản chủ luôn — không có gì để mất.

**Chế độ chỉ đọc** (`shared_db_read_only`): vẫn kéo về, nhưng mọi lượt ghi bị
`cv_repository` chặn và `push()` không bao giờ chạy. Đây là rào chắn khi chưa có
khóa ghi: hai máy cùng ghi thì bản đẩy sau đè mất dữ liệu của bản trước.

Mọi lỗi (thư mục không với tới, file hỏng…) chỉ ghi log + cập nhật `status()`,
KHÔNG ném ra ngoài. Setting `shared_db_folder` để trống = tắt hẳn tính năng.
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
import time
import uuid
from dataclasses import dataclass

from app.core import cv_repository, cv_schema, debuglog, settings

DB_NAME = "candidates.sqlite"
STATE_NAME = "state.json"
HISTORY_DIR = "history"
# Số bản giữ lại trong `history\` (~6 MB mỗi bản).
HISTORY_KEEP = 30

VERSION_KEY = "sync:version"
DB_ID_KEY = "sync:db_id"

# `get_connection()` hỏi thư mục dùng chung nhiều nhất một lần mỗi chừng này
# giây — không thì mỗi lượt đọc lại đi hỏi ổ mạng.
CONNECT_CHECK_INTERVAL = 10
# Thời gian tối đa một lượt kéo chờ lượt khác đang chạy (tải 6 MB qua VPN).
LOCK_WAIT = 120

# Mọi lượt đẩy / kéo chạy lần lượt: hai lượt cùng lúc sẽ tranh nhau file tạm,
# số version, và lượt kéo không được thay file khi lượt đẩy đang chụp nó.
_LOCK = threading.Lock()

# Bản chủ đã tải về + đã kiểm, chờ lúc không còn kết nối mở để thay vào.
_pending_path: str | None = None
_pending_version = 0

_last_check = 0.0            # time.monotonic() của lượt kiểm gần nhất
_checker: threading.Thread | None = None


# ═══════════════════════════════ TRẠNG THÁI ══════════════════════════════

@dataclass(frozen=True)
class SyncStatus:
    """Ảnh chụp trạng thái đồng bộ cho giao diện (message bằng tiếng Anh).

    state: "off"      — chưa cấu hình thư mục dùng chung
           "synced"   — DB local khớp bản chủ (tính tới `synced_at`)
           "pending"  — đã tải bản mới, chờ thay file
           "offline"  — thư mục dùng chung không với tới được
           "blocked"  — dừng đồng bộ, cần người quyết định (khác db_id,
                        app cũ hơn bản chủ, bản chủ mới hơn chưa kéo…)
           "error"    — lỗi bất ngờ, chi tiết ở debug.log
    """
    state: str = "off"
    message: str = ""
    synced_at: datetime.datetime | None = None
    version: int = 0
    read_only: bool = False


@dataclass
class SyncResult:
    status: SyncStatus
    pulled: bool = False     # lượt này đã THAY file .db local


_status = SyncStatus()
_listeners: list = []


def status() -> SyncStatus:
    return _status


def add_listener(fn) -> None:
    """Đăng ký `fn(status, replaced)` — gọi mỗi lần trạng thái đổi.

    CÓ THỂ được gọi từ luồng phụ: phía giao diện phải tự chuyển về luồng
    chính (vd phát một Qt Signal).
    """
    _listeners.append(fn)


def _set_status(new: SyncStatus, replaced: bool = False) -> SyncStatus:
    global _status
    old, _status = _status, new
    if (old.state, old.message) != (new.state, new.message):
        debuglog.write(f"shared_db: {new.state} — {new.message}")
    if replaced or old != new:
        for fn in list(_listeners):
            try:
                fn(new, replaced)
            except Exception as exc:
                debuglog.exception("shared_db: status listener failed", exc)
    return new


def _status_of(state: str, message: str, **kw) -> SyncStatus:
    """Trạng thái mới, giữ lại `synced_at`/`version` cũ nếu không nói khác."""
    kw.setdefault("synced_at", _status.synced_at)
    kw.setdefault("version", _status.version)
    return SyncStatus(state=state, message=message, read_only=is_read_only(), **kw)


# ═══════════════════════════════ CẤU HÌNH ════════════════════════════════

def shared_folder() -> str:
    """Thư mục dùng chung đã cấu hình ("" = tính năng đang tắt)."""
    return str(settings.get("shared_db_folder", "") or "").strip()


def is_enabled() -> bool:
    return bool(shared_folder())


def is_read_only() -> bool:
    return bool(settings.get("shared_db_read_only", False))


def read_state(folder: str) -> dict:
    """Đọc `state.json` của thư mục dùng chung ({} nếu chưa có / hỏng)."""
    try:
        with open(os.path.join(folder, STATE_NAME), encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def app_schema_version() -> str:
    """Lượt migration mới nhất mà CODE của app đang chạy biết tới."""
    return cv_schema.MIGRATIONS[-1][0]


# ═══════════════════════════ KIỂM TRA AN TOÀN ════════════════════════════

MSG_DIFFERENT_DB = (
    "This computer's database and the shared database are different databases "
    "(they don't share the same history), so sync is stopped — nothing was "
    "downloaded or uploaded. Someone needs to decide which one to keep. Your "
    "local data is untouched.")
MSG_UPDATE_APP = (
    "The shared database was saved by a newer version of Personal Toolbox — "
    "please update the app on this computer. Until then it keeps working on "
    "its local copy.")


def _schema_too_new(schema: str) -> bool:
    # Tên lượt có tiền tố số đệm 0 → so chuỗi là so thứ tự.
    return bool(schema) and schema > app_schema_version()


@dataclass
class _LocalInfo:
    exists: bool
    version: int = 0
    db_id: str = ""
    has_user_data: bool = False


# Bảng không tính là "dữ liệu người dùng" khi xét DB có phải máy mới cài không:
# danh mục app tự nạp sẵn + dòng `users` app tự tạo cho máy.
_NOT_USER_DATA = set(cv_schema.SEED_DATA) | {"users"}


def _local_info(db_path: str) -> _LocalInfo:
    if not os.path.exists(db_path):
        return _LocalInfo(exists=False)
    conn = _connect_ro(db_path)
    try:
        meta = dict(conn.execute(
            "SELECT key, value FROM app_meta WHERE key IN (?, ?)",
            (VERSION_KEY, DB_ID_KEY)).fetchall()) if _has_table(conn, "app_meta") else {}
        has_data = False
        for table in cv_repository._PK:
            if table in _NOT_USER_DATA or not _has_table(conn, table):
                continue
            if conn.execute(f"SELECT EXISTS (SELECT 1 FROM {table})").fetchone()[0]:
                has_data = True
                break
    finally:
        conn.close()
    return _LocalInfo(exists=True, version=_as_int(meta.get(VERSION_KEY)),
                      db_id=str(meta.get(DB_ID_KEY) or ""), has_user_data=has_data)


def _may_adopt(local: _LocalInfo, remote_id: str) -> bool:
    """Có được THAY DB local bằng bản chủ mang `remote_id` không (xét dòng dõi)."""
    if not local.exists:
        return True
    if local.db_id and remote_id:
        return local.db_id == remote_id
    # DB local chưa từng đồng bộ (không có db_id): chỉ nhận bản chủ khi nó chưa
    # có dữ liệu gì — máy mới cài. Đã có dữ liệu thì đó là một DB khác.
    return not local.db_id and local.version == 0 and not local.has_user_data


# ═══════════════════════════════ KÉO VỀ ══════════════════════════════════

def ensure_fresh(reason: str = "check") -> SyncResult:
    """Kéo bản chủ về nếu nó mới hơn DB local. Không bao giờ ném lỗi.

    Chạy ĐỒNG BỘ (có đụng ổ mạng) — đừng gọi thẳng ở luồng giao diện, trừ lúc
    mở app qua `ensure_fresh_with_timeout`. Lượt khác (kéo hay đẩy) đang chạy
    thì CHỜ nó xong rồi tự chạy lượt của mình: người bấm Refresh phải nhận kết
    quả thật chứ không phải trạng thái cũ. Chờ quá `LOCK_WAIT` giây thì trả về
    trạng thái hiện tại.
    """
    global _last_check
    if not _LOCK.acquire(timeout=LOCK_WAIT):
        return SyncResult(_status)
    try:
        _last_check = time.monotonic()
        try:
            return _ensure_fresh(reason)
        except Exception as exc:
            debuglog.exception(f"shared_db.ensure_fresh({reason}) failed", exc)
            return SyncResult(_set_status(_status_of(
                "error", f"Sync check failed: {exc}")))
    finally:
        _LOCK.release()


def ensure_fresh_with_timeout(reason: str, timeout: float) -> SyncResult | None:
    """`ensure_fresh()` ở luồng phụ, chờ tối đa `timeout` giây (None = hết giờ).

    Dùng lúc mở app: kéo về TRƯỚC khi màn hình đầu tiên đọc DB, nhưng ổ mạng
    treo thì không giữ app lại. Hết giờ thì luồng vẫn chạy tiếp; nếu nó kéo
    được bản mới thì giao diện được báo qua `add_listener`.
    """
    box = []
    worker = threading.Thread(target=lambda: box.append(ensure_fresh(reason)),
                              name="shared-db-pull", daemon=True)
    worker.start()
    worker.join(timeout)
    if worker.is_alive():
        debuglog.write(f"shared_db.ensure_fresh({reason}): still running after "
                       f"{timeout:.0f}s — continuing in the background")
        return None
    return box[0] if box else None


def check_in_background(reason: str = "background") -> None:
    """Chạy `ensure_fresh()` ở luồng phụ, không chờ (bỏ qua nếu đang có lượt)."""
    global _checker, _last_check
    if _checker is not None and _checker.is_alive():
        return
    _last_check = time.monotonic()
    _checker = threading.Thread(target=lambda: ensure_fresh(reason),
                                name="shared-db-check", daemon=True)
    _checker.start()


def _on_connect() -> None:
    """Hook ở đầu `cv_repository.get_connection()`. KHÔNG chờ ổ mạng.

    - Có bản chờ (đã tải + đã kiểm) → thay ngay tại đây nếu được: việc này chỉ
      là đổi tên file ở ổ local.
    - Đã quá `CONNECT_CHECK_INTERVAL` giây kể từ lượt kiểm trước → kiểm ở luồng
      phụ; kết nối đang mở vẫn đọc bản local, bản mới (nếu có) thay ở lượt sau.
    """
    global _last_check
    if _pending_path is not None:
        if _LOCK.acquire(blocking=False):
            try:
                _apply_pending()
            finally:
                _LOCK.release()
        return
    if time.monotonic() - _last_check < CONNECT_CHECK_INTERVAL:
        return
    # Đóng dấu cả khi tính năng tắt: không thì mỗi kết nối lại đọc config.json.
    _last_check = time.monotonic()
    if not is_enabled():
        return
    check_in_background("connect")


def _ensure_fresh(reason: str) -> SyncResult:
    folder = shared_folder()
    if not folder:
        _drop_pending()
        return SyncResult(_set_status(SyncStatus(state="off",
                                                 read_only=is_read_only())))
    db_path = cv_repository._db_path()

    if not os.path.isdir(folder):
        # Bản đã tải từ trước vẫn thay được — nó đã được kiểm đầy đủ.
        if _pending_path is not None:
            return _apply_pending()
        return SyncResult(_set_status(_status_of(
            "offline", f"Shared folder not reachable: {folder}")))

    state = read_state(folder)
    master = os.path.join(folder, DB_NAME)
    remote_version = _as_int(state.get("version"))
    remote_id = str(state.get("db_id") or "")
    local = _local_info(db_path)
    now = datetime.datetime.now()

    if not state or not os.path.exists(master):
        _drop_pending()
        return SyncResult(_set_status(_status_of(
            "synced", "The shared folder has no database yet.",
            synced_at=now, version=local.version)))

    # Cùng version: kiểm dòng dõi vẫn phải làm (để báo sớm), nhưng không copy.
    if local.db_id and remote_id and local.db_id != remote_id:
        _drop_pending()
        return SyncResult(_set_status(_status_of("blocked", MSG_DIFFERENT_DB)))
    if remote_version <= local.version:
        _drop_pending()
        return SyncResult(_set_status(_status_of(
            "synced", f"Up to date (version {local.version}).",
            synced_at=now, version=local.version)))

    # Bản chủ mới hơn → hai lớp kiểm tra trước khi tải.
    if _schema_too_new(str(state.get("schema_version") or "")):
        _drop_pending()
        return SyncResult(_set_status(_status_of("blocked", MSG_UPDATE_APP)))
    if not _may_adopt(local, remote_id):
        _drop_pending()
        return SyncResult(_set_status(_status_of("blocked", MSG_DIFFERENT_DB)))

    if _pending_path is not None and _pending_version >= remote_version:
        return _apply_pending()

    staged, info = _download(master, db_path)
    # Kiểm lại trên CHÍNH file đã tải: state.json chỉ là lời hứa.
    if _schema_too_new(info["schema"]):
        _remove_quietly(staged)
        return SyncResult(_set_status(_status_of("blocked", MSG_UPDATE_APP)))
    if not _may_adopt(local, info["db_id"]):
        _remove_quietly(staged)
        return SyncResult(_set_status(_status_of("blocked", MSG_DIFFERENT_DB)))
    if info["version"] <= local.version:
        _remove_quietly(staged)
        return SyncResult(_set_status(_status_of(
            "synced", f"Up to date (version {local.version}).",
            synced_at=now, version=local.version)))

    _set_pending(staged, info["version"])
    debuglog.write(f"shared_db.ensure_fresh({reason}): downloaded version "
                   f"{info['version']} (local was {local.version})")
    return _apply_pending()


def _download(master: str, db_path: str) -> tuple[str, dict]:
    """Copy bản chủ về tên tạm cạnh DB local rồi kiểm file đó.

    Trả về (đường dẫn file tạm, {version, db_id, schema}). Cùng thư mục với DB
    local nên `os.replace` sau đó là nguyên tử (cùng ổ đĩa).
    """
    staged = db_path + ".incoming"
    tmp = staged + ".tmp"
    shutil.copyfile(master, tmp)
    try:
        conn = _connect_ro(tmp)
        try:
            check = conn.execute("PRAGMA quick_check").fetchone()[0]
            if check != "ok":
                raise sqlite3.DatabaseError(f"downloaded copy failed quick_check: {check}")
            meta = dict(conn.execute(
                "SELECT key, value FROM app_meta WHERE key IN (?, ?)",
                (VERSION_KEY, DB_ID_KEY)).fetchall())
            schema = _schema_version(conn)
        finally:
            conn.close()
        os.replace(tmp, staged)
    except Exception:
        _remove_quietly(tmp)
        raise
    return staged, {"version": _as_int(meta.get(VERSION_KEY)),
                    "db_id": str(meta.get(DB_ID_KEY) or ""), "schema": schema}


def _set_pending(path: str, version: int) -> None:
    global _pending_path, _pending_version
    if _pending_path and _pending_path != path:
        _remove_quietly(_pending_path)
    _pending_path, _pending_version = path, version


def _drop_pending() -> None:
    """Bỏ bản chờ (bản chủ đã đổi hướng — vd bị chặn, hoặc local đã kịp mới)."""
    global _pending_path, _pending_version
    if _pending_path:
        _remove_quietly(_pending_path)
    _pending_path, _pending_version = None, 0


def _apply_pending() -> SyncResult:
    """Thay DB local bằng bản chờ nếu không còn kết nối mở. Gọi khi giữ `_LOCK`."""
    global _pending_path, _pending_version
    path, version = _pending_path, _pending_version
    if path is None:
        return SyncResult(_status)
    if not os.path.exists(path):
        _pending_path, _pending_version = None, 0
        return SyncResult(_status)
    db_path = cv_repository._db_path()
    # Máy được ghi: giữ file cũ một bản phòng hờ — nếu nó có thay đổi chưa đẩy
    # lên thì vẫn còn đường lấy lại. Máy chỉ đọc không có gì đáng giữ.
    keep_old = None if is_read_only() else db_path + ".before-pull"
    try:
        swapped = cv_repository.replace_db_file(path, keep_old_as=keep_old)
    except OSError as exc:
        debuglog.exception("shared_db: could not replace the local database", exc)
        swapped = False
    if not swapped:
        return SyncResult(_set_status(_status_of(
            "pending", f"Version {version} downloaded — it will be loaded as soon "
                       f"as the database is free.")))
    _pending_path, _pending_version = None, 0
    debuglog.write(f"shared_db: local database replaced with version {version}")
    # File mới có thể thiếu migration mà app này có, và chưa có dòng `users`
    # của máy này (id người dùng cũ đã được quên trong replace_db_file).
    try:
        cv_repository.init_db()
    except Exception as exc:
        debuglog.exception("shared_db: init_db after pull failed", exc)
    return SyncResult(_set_status(_status_of(
        "synced", f"Updated to version {version}.",
        synced_at=datetime.datetime.now(), version=version), replaced=True),
        pulled=True)


# ═══════════════════════════════ ĐẨY LÊN ═════════════════════════════════

@dataclass
class PushResult:
    ok: bool                 # đã thay bản chủ thành công
    skipped: bool = False    # không làm gì vì tính năng tắt / thư mục không với tới
    message: str = ""        # mô tả ngắn (tiếng Anh — hiện được lên giao diện)
    version: int | None = None


def push(reason: str, folder: str | None = None) -> PushResult:
    """Đẩy DB local lên thư mục dùng chung. Không bao giờ ném lỗi.

    `reason` chỉ để ghi log + `state.json` (vd "close", "manual"). `folder` bỏ
    trống thì lấy từ setting `shared_db_folder`. Máy chỉ đọc không bao giờ đẩy.
    """
    folder = (shared_folder() if folder is None else str(folder)).strip()
    if not folder:
        return PushResult(ok=False, skipped=True,
                          message="No shared database folder is set.")
    if is_read_only():
        return PushResult(ok=False, skipped=True,
                          message="This computer is in read-only mode, so it "
                                  "never uploads its database.")
    with _LOCK:
        try:
            result = _push(reason, folder)
        except Exception as exc:
            debuglog.exception(f"shared_db.push({reason}) failed — {folder}", exc)
            result = PushResult(ok=False, message=f"Backup failed: {exc}")
    debuglog.write(f"shared_db.push({reason}): {result.message}")
    if result.ok:
        _set_status(_status_of("synced", f"Uploaded version {result.version}.",
                               synced_at=datetime.datetime.now(),
                               version=result.version))
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

    state = read_state(folder)
    has_master = bool(state) and os.path.exists(os.path.join(folder, DB_NAME))
    local = _local_info(src)
    remote_version = _as_int(state.get("version")) if has_master else 0
    remote_id = str(state.get("db_id") or "") if has_master else ""

    # Dòng dõi: chỉ đè lên bản chủ CỦA CHÍNH DB này.
    if has_master:
        if remote_id:
            same_db = local.db_id == remote_id
        else:
            # Bản chủ có từ trước khi có db_id: chỉ máy đã đẩy nó (version local
            # bắt kịp bản chủ) mới được đẩy tiếp.
            same_db = local.version > 0 and local.version >= remote_version
        if not same_db:
            _set_status(_status_of("blocked", MSG_DIFFERENT_DB))
            return PushResult(ok=False, message=MSG_DIFFERENT_DB)
        if remote_version > local.version:
            msg = (f"The shared database (version {remote_version}) is newer than "
                   f"this computer's copy (version {local.version}). Nothing was "
                   f"uploaded, so the newer data is not overwritten — refresh "
                   f"first.")
            _set_status(_status_of("blocked", msg))
            return PushResult(ok=False, message=msg)

    login = cv_repository.windows_login() or "unknown"
    version = max(local.version, remote_version) + 1
    db_id = local.db_id or remote_id or uuid.uuid4().hex

    work_dir = tempfile.mkdtemp(prefix="ptb-push-")
    try:
        snapshot = os.path.join(work_dir, DB_NAME)
        schema_version = _snapshot(src, snapshot, version, db_id)

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
            "db_id": db_id,
            "updated_by": login,
            "updated_at": now.isoformat(timespec="seconds"),
            "schema_version": schema_version,
            "reason": reason,
        }, login)

        history_note = _save_history(folder, snapshot, now, login)
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

    _write_sync_meta(src, version, db_id)
    return PushResult(ok=True, version=version,
                      message=f"Backed up as version {version} to {folder}"
                              f"{history_note}")


def _snapshot(src: str, dest: str, version: int, db_id: str) -> str:
    """Chụp `src` ra `dest` bằng backup API, đóng dấu version + db_id →
    trả về schema_version.

    Nguồn mở chế độ chỉ đọc: lượt đẩy không được đụng vào DB đang dùng.
    """
    source = _connect_ro(src)
    try:
        target = sqlite3.connect(dest)
        try:
            source.backup(target)
            target.execute("CREATE TABLE IF NOT EXISTS app_meta "
                           "(key VARCHAR PRIMARY KEY, value VARCHAR)")
            target.executemany("INSERT OR REPLACE INTO app_meta (key, value) "
                               "VALUES (?, ?)",
                               [(VERSION_KEY, str(version)), (DB_ID_KEY, db_id)])
            target.commit()
            check = target.execute("PRAGMA quick_check").fetchone()[0]
            if check != "ok":
                raise sqlite3.DatabaseError(f"snapshot failed quick_check: {check}")
            return _schema_version(target)
        finally:
            target.close()
    finally:
        source.close()


def _write_sync_meta(db_path: str, version: int, db_id: str) -> None:
    """Ghi version + db_id vào DB local. Lỗi ở đây không làm hỏng lượt đẩy đã
    xong — lượt sau vẫn lấy max(local, state.json) + 1 nên version không lùi,
    và db_id lấy lại được từ state.json."""
    try:
        conn = sqlite3.connect(db_path, timeout=10)
        try:
            conn.executemany("INSERT OR REPLACE INTO app_meta (key, value) "
                             "VALUES (?, ?)",
                             [(VERSION_KEY, str(version)), (DB_ID_KEY, db_id)])
            conn.commit()
        finally:
            conn.close()
    except sqlite3.Error as exc:
        debuglog.exception("shared_db: could not store sync metadata locally", exc)


# ═══════════════════════════════ TIỆN ÍCH ════════════════════════════════

def _connect_ro(path: str) -> sqlite3.Connection:
    """Kết nối CHỈ ĐỌC, không qua `cv_repository` (không đếm, không hook) —
    chỉ dùng trong module này khi đang giữ `_LOCK`."""
    return sqlite3.connect(pathlib.Path(path).resolve().as_uri() + "?mode=ro",
                           uri=True)


def _has_table(conn, name: str) -> bool:
    return conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
                        (name,)).fetchone() is not None


def _schema_version(conn) -> str:
    """Tên lượt migration mới nhất mà DB đã chạy (vd "0009_users_and_audit_columns").

    Tên có tiền tố số đệm 0 nên so sánh chuỗi cũng đúng thứ tự.
    """
    row = conn.execute("SELECT MAX(key) FROM app_meta "
                       "WHERE key LIKE 'migration:%'").fetchone()
    return (row[0] or "")[len("migration:"):] if row else ""


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


cv_repository.set_connect_hook(_on_connect)
