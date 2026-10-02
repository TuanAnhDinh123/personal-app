"""Đồng bộ DB local ↔ thư mục dùng chung (ổ mạng) bằng cách copy NGUYÊN FILE.

Mỗi máy làm việc trên file .db ở ổ local (`cv_repository._db_path()`); thư mục
dùng chung chỉ giữ BẢN CHỦ + lịch sử, không máy nào mở SQLite trực tiếp trên ổ
mạng (SQLite qua SMB/VPN là hỏng file). Bố cục thư mục dùng chung:

    <shared_db_folder>\\
    ├── candidates.sqlite    ← bản chủ
    ├── state.json           ← version · db_id · updated_by · updated_at · schema_version
    ├── lock.json            ← KHÓA GHI: ai đang giữ · nhịp tim · đang làm gì
    ├── release-request.json ← máy đang chờ nhờ máy giữ khóa nhả sớm
    └── history\\
        └── candidates-20260927-091530-<tài khoản windows>.sqlite

KHÓA GHI — mỗi lúc chỉ MỘT máy được ghi, app tự giữ/tự nhả:

    ① câu GHI đầu tiên (cổng ghi của `cv_repository`) → giành `lock.json`
    ② kiểm bản local có theo kịp bản chủ không — chậm hơn thì KHÔNG ghi, nhả
       khóa, kéo bản mới ở nền, báo người dùng bấm lại
    ③ ghi vào DB LOCAL như bình thường
    ④ hết ÂN HẠN (`GRACE` giây không ghi thêm) / máy kia xin / đóng app
       → đẩy lên rồi mới nhả khóa

  Nguyên tắc giữ cả mô hình an toàn: CHỈ NHẢ KHÓA KHI BẢN LOCAL ĐÃ ĐẨY LÊN
  XONG. Nhờ vậy bản local không bao giờ đi trước bản chủ mà lại không cầm
  khóa, nên kéo về lúc nào cũng an toàn. Ân hạn gom một loạt lượt ghi liền
  nhau thành MỘT lần đẩy (DB ~6 MB, đẩy mỗi lượt ghi qua VPN là quá chậm).

  Thao tác dài (quét CV, Bulk Import, Sync with Excel) bọc trong `session()`:
  giành khóa NGAY ĐẦU (máy kia đang ghi thì báo luôn, không đợi tới lượt ghi
  đầu tiên sau mấy phút gọi AI) và không đẩy giữa chừng.

  Khóa có NHỊP TIM (`HEARTBEAT` giây). App chết khi đang giữ khóa → nhịp tim
  đứng yên → máy kia giành lại sau `STALE_AFTER` giây. Máy giữ khóa mà không
  ghi được nhịp tim (rớt VPN) quá nửa thời gian đó thì tự ngừng nhận lượt ghi
  mới, để không bao giờ ghi trong lúc máy kia đã có quyền coi khóa là chết.

  `sync:local_seq` / `sync:pushed_seq` trong `app_meta`: sổ đếm lượt ghi. Lớn
  hơn nhau = DB local có thay đổi CHƯA ĐẨY ("dirty"). Bản local dirty thì
  không bao giờ bị thay bằng bản kéo về; nếu bản chủ cũng đã có thay đổi của
  máy khác thì DỪNG đồng bộ (`MSG_CONFLICT`) — chọn giữ bên nào là việc của
  con người. Không ai ghi chen vào thì lần mở app sau tự đẩy nốt.

Hai chiều:

  ĐẨY LÊN — `push()` (nút Back up now) và lượt nhả khóa (hết ân hạn, đóng app):
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
`cv_repository` chặn và `push()` không bao giờ chạy — cho máy chỉ cần XEM dữ
liệu. Máy không bật cờ này đều ghi được, khóa ghi lo việc chia lượt.

Thư mục dùng chung không với tới được (chưa bật VPN): ĐỌC vẫn chạy trên bản
local, nhưng GHI bị chặn — không giành được khóa thì không biết máy kia có đang
ghi không.

Lỗi đồng bộ (thư mục không với tới, file hỏng…) chỉ ghi log + cập nhật
`status()`, KHÔNG ném ra ngoài. Lỗi duy nhất được ném là `DatabaseBusyError` ở
cổng ghi — để chặn đúng lượt ghi đó, trước khi nó ghi được gì. Setting
`shared_db_folder` để trống = tắt hẳn tính năng.
"""
import contextlib
import dataclasses
import datetime
import glob
import json
import os
import pathlib
import shutil
import socket
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

LOCK_NAME = "lock.json"
RELEASE_REQUEST_NAME = "release-request.json"

VERSION_KEY = "sync:version"
DB_ID_KEY = "sync:db_id"
# Sổ đếm lượt ghi: local_seq tăng mỗi kết nối có ghi; pushed_seq = local_seq
# của bản đã đẩy lên. local_seq > pushed_seq ⇔ có thay đổi chưa đẩy.
LOCAL_SEQ_KEY = "sync:local_seq"
PUSHED_SEQ_KEY = "sync:pushed_seq"

# `get_connection()` hỏi thư mục dùng chung nhiều nhất một lần mỗi chừng này
# giây — không thì mỗi lượt đọc lại đi hỏi ổ mạng.
CONNECT_CHECK_INTERVAL = 10
# Thời gian tối đa một lượt kéo chờ lượt khác đang chạy (tải 6 MB qua VPN).
LOCK_WAIT = 120

# ─── Khóa ghi ───
# Giữ khóa thêm chừng này giây sau lượt ghi cuối rồi mới đẩy + nhả: người dùng
# hay sửa một mạch nhiều hồ sơ, gom lại thì chỉ tốn một lần đẩy.
GRACE = 45
# Nhịp tim của khóa — cũng là độ trễ tối đa để thấy lời "xin nhả" của máy kia.
HEARTBEAT = 5
# Nhịp tim ĐỨNG YÊN chừng này giây (đo bằng đồng hồ của máy đang nhìn, nên lệch
# giờ giữa hai máy không ảnh hưởng) → coi máy giữ khóa đã chết.
STALE_AFTER = 90
# …hoặc mốc giờ ghi trong khóa đã cũ hơn chừng này (cho khóa chết từ lâu mà máy
# này mới nhìn thấy lần đầu — đủ rộng để chịu được đồng hồ hai máy lệch nhau).
STALE_AFTER_WALL = 600
# Thời gian tối đa một lượt giành khóa chờ ổ mạng (VPN treo) ở cổng ghi.
ACQUIRE_TIMEOUT = 8
# `session()` gặp bản chủ mới hơn thì kéo về rồi giành lại, chờ tối đa chừng này.
SESSION_PULL_TIMEOUT = 60

# Mọi lượt đẩy / kéo chạy lần lượt: hai lượt cùng lúc sẽ tranh nhau file tạm,
# số version, và lượt kéo không được thay file khi lượt đẩy đang chụp nó.
_LOCK = threading.Lock()

# Bản chủ đã tải về + đã kiểm, chờ lúc không còn kết nối mở để thay vào.
_pending_path: str | None = None
_pending_version = 0

_last_check = 0.0            # time.monotonic() của lượt kiểm gần nhất
_checker: threading.Thread | None = None
# Lượt kiểm thấy local có thay đổi treo mà đẩy được → đẩy sau khi nhả `_LOCK`.
_recover_wanted = False
_recover_thread: threading.Thread | None = None


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

    lease: ""      — không ai giữ khóa ghi (hoặc chưa biết)
           "mine"  — máy này đang giữ khóa (đang trong ân hạn)
           khác    — mô tả máy khác đang giữ, vd "Lan (VN-PC12)"
    dirty: máy này đang có thay đổi chưa đẩy lên bản chủ.
    uploading: ĐANG đẩy lên ngay lúc này (chỉ bật trong đúng lượt đẩy).
    """
    state: str = "off"
    message: str = ""
    synced_at: datetime.datetime | None = None
    version: int = 0
    read_only: bool = False
    lease: str = ""
    dirty: bool = False
    uploading: bool = False


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
    """Trạng thái mới, giữ lại `synced_at`/`version`/`lease`/`dirty`/
    `uploading` cũ nếu không nói khác."""
    kw.setdefault("synced_at", _status.synced_at)
    kw.setdefault("version", _status.version)
    kw.setdefault("lease", _status.lease)
    kw.setdefault("dirty", _status.dirty)
    kw.setdefault("uploading", _status.uploading)
    return SyncStatus(state=state, message=message, read_only=is_read_only(), **kw)


def _set_lease_status(lease: str | None = None, dirty: bool | None = None) -> None:
    """Chỉ đổi phần khóa ghi của trạng thái, giữ nguyên mọi thứ khác."""
    changes = {}
    if lease is not None:
        changes["lease"] = lease
    if dirty is not None:
        changes["dirty"] = dirty
    if changes:
        _set_status(dataclasses.replace(_status, **changes))


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
MSG_CONFLICT = (
    "This computer has changes that were never uploaded, and another computer "
    "has saved newer data to the shared database in the meantime. Sync is "
    "stopped so neither side is overwritten — someone needs to decide which "
    "changes to keep. Your local data is untouched.")
MSG_LEASE_LOST = (
    "Another computer took over the shared database while this one still had "
    "changes that were not uploaded (the network was probably down for a "
    "while). Sync is stopped so nothing is overwritten — someone needs to "
    "decide which changes to keep. Your local data is untouched.")
MSG_NEWER = (
    "Another computer has saved newer data to the shared database. It is being "
    "downloaded now — please try again in a few seconds. Nothing was changed.")
MSG_OFFLINE = (
    "Can't reach the shared database folder (is the VPN connected?), so changes "
    "can't be saved right now. You can still view the data. Nothing was changed.")
MSG_NOT_RESPONDING = (
    "The shared database folder is not responding, so changes can't be saved "
    "right now. Check the network or VPN and try again. Nothing was changed.")


def _schema_too_new(schema: str) -> bool:
    # Tên lượt có tiền tố số đệm 0 → so chuỗi là so thứ tự.
    return bool(schema) and schema > app_schema_version()


@dataclass
class _LocalInfo:
    exists: bool
    version: int = 0
    db_id: str = ""
    has_user_data: bool = False
    local_seq: int = 0
    pushed_seq: int = 0

    @property
    def dirty(self) -> bool:
        """Có thay đổi chưa đẩy lên bản chủ."""
        return self.local_seq > self.pushed_seq


# Bảng không tính là "dữ liệu người dùng" khi xét DB có phải máy mới cài không:
# danh mục app tự nạp sẵn + dòng `users` app tự tạo cho máy.
_NOT_USER_DATA = set(cv_schema.SEED_DATA) | {"users"}


def _local_info(db_path: str) -> _LocalInfo:
    if not os.path.exists(db_path):
        return _LocalInfo(exists=False)
    conn = _connect_ro(db_path)
    try:
        meta = dict(conn.execute(
            "SELECT key, value FROM app_meta WHERE key IN (?, ?, ?, ?)",
            (VERSION_KEY, DB_ID_KEY, LOCAL_SEQ_KEY, PUSHED_SEQ_KEY)).fetchall()
        ) if _has_table(conn, "app_meta") else {}
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
                      db_id=str(meta.get(DB_ID_KEY) or ""), has_user_data=has_data,
                      local_seq=_as_int(meta.get(LOCAL_SEQ_KEY)),
                      pushed_seq=_as_int(meta.get(PUSHED_SEQ_KEY)))


def _same_lineage(local: _LocalInfo, state: dict) -> bool:
    """DB local có phải CÙNG dòng dõi với bản chủ đang có không (được ghi đè
    lên bản chủ / được ghi khi đang cầm khóa)."""
    remote_id = str(state.get("db_id") or "")
    if remote_id:
        return local.db_id == remote_id
    # Bản chủ có từ trước khi có db_id: chỉ máy đã đẩy nó (version local bắt
    # kịp bản chủ) mới tính là cùng dòng dõi.
    return local.version > 0 and local.version >= _as_int(state.get("version"))


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
    global _last_check, _recover_wanted
    if not _LOCK.acquire(timeout=LOCK_WAIT):
        return SyncResult(_status)
    try:
        _last_check = time.monotonic()
        try:
            result = _ensure_fresh(reason)
        except Exception as exc:
            debuglog.exception(f"shared_db.ensure_fresh({reason}) failed", exc)
            result = SyncResult(_set_status(_status_of(
                "error", f"Sync check failed: {exc}")))
    finally:
        _LOCK.release()
    # Đẩy nốt thay đổi còn treo — NGOÀI `_LOCK` vì lượt đẩy tự giữ nó.
    if _recover_wanted:
        _recover_wanted = False
        _start_recover_push()
    return result


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
    global _recover_wanted
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

    # Máy này đang cầm khóa ghi → chính nó là bên mới nhất; bản chủ không thể
    # mới hơn, và bản local đang có thay đổi chờ đẩy thì càng không được thay.
    if holding_lease():
        _drop_pending()
        return SyncResult(_status)

    _observe_lock(folder)
    state = read_state(folder)
    master = os.path.join(folder, DB_NAME)
    remote_version = _as_int(state.get("version"))
    remote_id = str(state.get("db_id") or "")
    local = _local_info(db_path)
    now = datetime.datetime.now()
    _set_lease_status(dirty=local.dirty)

    if not state or not os.path.exists(master):
        _drop_pending()
        if local.dirty and not is_read_only():
            _recover_wanted = True
        return SyncResult(_set_status(_status_of(
            "synced", "The shared folder has no database yet.",
            synced_at=now, version=local.version)))

    # Cùng version: kiểm dòng dõi vẫn phải làm (để báo sớm), nhưng không copy.
    if local.db_id and remote_id and local.db_id != remote_id:
        _drop_pending()
        return SyncResult(_set_status(_status_of("blocked", MSG_DIFFERENT_DB)))

    # Local có thay đổi chưa đẩy (app tắt khi đang offline, mất khóa…): KHÔNG
    # BAO GIỜ thay nó bằng bản kéo về. Không ai ghi chen vào thì đẩy nốt; có
    # người ghi chen vào rồi thì hai bên đã rẽ nhánh — người phải quyết.
    if local.dirty and not is_read_only():
        _drop_pending()
        if remote_version > local.version:
            return SyncResult(_set_status(_status_of("blocked", MSG_CONFLICT)))
        _recover_wanted = True
        # Lượt đẩy nốt chạy ngay sau đây và tự bật "uploading" — không báo trước.
        return SyncResult(_set_status(_status_of(
            "synced", "Uploading changes that were not uploaded yet.")))

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
    # Từ lúc tải tới lúc thay, máy này có thể đã kịp giành khóa / có thay đổi
    # chưa đẩy — bản tải về khi đó đã lỗi thời, thay vào là mất thay đổi.
    if holding_lease() or (not is_read_only() and _local_info(db_path).dirty):
        _drop_pending()
        return SyncResult(_status)
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
    """Đẩy DB local lên thư mục dùng chung (nút Back up now). Không ném lỗi.

    Đẩy là ghi đè bản chủ, nên cũng phải CẦM KHÓA GHI: đang giữ thì đẩy luôn
    (ân hạn vẫn chạy tiếp), chưa giữ thì giành khóa → đẩy → nhả ngay. Máy kia
    đang ghi thì không đẩy, trả về thông báo vì sao.

    `reason` chỉ để ghi log + `state.json` (vd "manual"). `folder` bỏ trống
    thì lấy từ setting `shared_db_folder`. Máy chỉ đọc không bao giờ đẩy.
    """
    folder = (shared_folder() if folder is None else str(folder)).strip()
    if not folder:
        return PushResult(ok=False, skipped=True,
                          message="No shared database folder is set.")
    if is_read_only():
        return PushResult(ok=False, skipped=True,
                          message="This computer is in read-only mode, so it "
                                  "never uploads its database.")
    try:
        took = _enter_lease("backup", allow_pull=False)
    except cv_repository.DatabaseBusyError as exc:
        # "Chưa phải lúc" chứ không phải lỗi → skipped (giao diện hiện cảnh báo).
        return PushResult(ok=False, skipped=True, message=str(exc))
    try:
        result = _push_locked(reason, folder)
    finally:
        _leave_lease()
    if took:
        _release_if_clean(reason)
    return result


def _push_locked(reason: str, folder: str) -> PushResult:
    """`_push` trong `_LOCK`, bắt mọi lỗi, cập nhật trạng thái. Gọi khi đang
    cầm khóa ghi.

    `uploading` chỉ bật trong đúng lượt đẩy — giao diện hiện "Uploading…" lúc
    đó rồi quay về bình thường; ân hạn trước đó chạy ngầm, không hiện gì.
    """
    with _LOCK:
        _set_status(dataclasses.replace(_status, uploading=True))
        try:
            result = _push(reason, folder)
        except Exception as exc:
            debuglog.exception(f"shared_db.push({reason}) failed — {folder}", exc)
            result = PushResult(ok=False, message=f"Backup failed: {exc}")
        finally:
            _set_status(dataclasses.replace(_status, uploading=False))
    debuglog.write(f"shared_db.push({reason}): {result.message}")
    if result.ok:
        _set_status(_status_of("synced", f"Uploaded version {result.version}.",
                               synced_at=datetime.datetime.now(),
                               version=result.version,
                               dirty=_local_info(cv_repository._db_path()).dirty))
    return result


def _run_with_timeout(fn, timeout: float, name: str):
    """Chạy `fn()` ở luồng daemon, chờ tối đa `timeout` giây.

    Trả về (xong?, kết quả). Ổ mạng treo thì thao tác file có thể đứng hàng
    chục giây và không hủy được — luồng cứ chạy tiếp, bên gọi thôi chờ.
    """
    box = []
    worker = threading.Thread(target=lambda: box.append(fn()), name=name, daemon=True)
    worker.start()
    worker.join(timeout)
    if worker.is_alive() or not box:
        return False, None
    return True, box[0]


def shutdown_with_timeout(timeout: float) -> PushResult | None:
    """Lúc đóng app: đẩy thay đổi còn chờ rồi nhả khóa, chờ tối đa `timeout`
    giây (None = hết giờ / không có gì để làm).

    Hết giờ thì app thoát luôn: bản chủ không hỏng vì chỉ được thay bằng
    `os.replace()` sau khi file tạm đã ghi đủ; còn khóa thì máy kia giành lại
    khi thấy nhịp tim đứng yên, và thay đổi chưa đẩy được lần mở app sau tự đẩy
    nốt (xem `_ensure_fresh`).
    """
    done, result = _run_with_timeout(_shutdown, timeout, "shared-db-shutdown")
    if not done:
        debuglog.write(f"shared_db.shutdown: gave up after {timeout:.0f}s — the "
                       "shared folder is not responding")
        return None
    return result


def _shutdown() -> PushResult | None:
    folder = shared_folder()
    if not folder or is_read_only():
        return None
    _STOP.set()
    with _LEASE:
        token = _lease_token
    if token is not None:
        # Đang cầm khóa → đẩy (nếu có gì để đẩy) rồi nhả, BỎ QUA việc đợi lượt
        # ghi đang chạy: app đang thoát, bản chụp vẫn nhất quán theo transaction.
        return _flush_and_release(token, "close", force=True)
    # Không cầm khóa: chỉ đẩy khi local có thay đổi treo, hoặc thư mục dùng
    # chung chưa có bản chủ nào (lần đầu cấu hình).
    if not os.path.isdir(folder):
        return None
    has_master = os.path.exists(os.path.join(folder, DB_NAME))
    if has_master and not _local_info(cv_repository._db_path()).dirty:
        return None
    return push("close", folder)


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
        if not _same_lineage(local, state):
            _set_status(_status_of("blocked", MSG_DIFFERENT_DB))
            return PushResult(ok=False, message=MSG_DIFFERENT_DB)
        if remote_version > local.version:
            # Đang cầm khóa mà bản chủ vẫn mới hơn = khóa đã bị máy khác giành
            # trong lúc máy này mất mạng. Không đè — có thay đổi treo thì là
            # rẽ nhánh, người phải quyết.
            msg = MSG_CONFLICT if local.dirty else (
                f"The shared database (version {remote_version}) is newer than "
                f"this computer's copy (version {local.version}). Nothing was "
                f"uploaded, so the newer data is not overwritten — refresh first.")
            _set_status(_status_of("blocked", msg))
            return PushResult(ok=False, message=msg)

    login = cv_repository.windows_login() or "unknown"
    version = max(local.version, remote_version) + 1
    db_id = local.db_id or remote_id or uuid.uuid4().hex

    work_dir = tempfile.mkdtemp(prefix="ptb-push-")
    try:
        snapshot = os.path.join(work_dir, DB_NAME)
        schema_version, pushed_seq = _snapshot(src, snapshot, version, db_id)

        # Bản chủ: copy lên tên tạm RIÊNG của máy này rồi mới thay nguyên tử.
        master = os.path.join(folder, DB_NAME)
        tmp_master = os.path.join(folder, f".{DB_NAME}.{login}.tmp")
        shutil.copyfile(snapshot, tmp_master)
        if os.path.getsize(tmp_master) != os.path.getsize(snapshot):
            _remove_quietly(tmp_master)
            raise OSError("uploaded copy has the wrong size")
        _replace(tmp_master, master)

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

    _write_sync_meta(src, version, db_id, pushed_seq)
    return PushResult(ok=True, version=version,
                      message=f"Backed up as version {version} to {folder}"
                              f"{history_note}")


def _snapshot(src: str, dest: str, version: int, db_id: str) -> tuple[str, int]:
    """Chụp `src` ra `dest` bằng backup API, đóng dấu version + db_id →
    trả về (schema_version, local_seq của bản chụp).

    Bản chụp được ghi `pushed_seq = local_seq`: máy nào kéo nó về đều thấy
    "không có gì chưa đẩy". Nguồn mở chế độ chỉ đọc: lượt đẩy không được đụng
    vào DB đang dùng.
    """
    source = _connect_ro(src)
    try:
        target = sqlite3.connect(dest)
        try:
            source.backup(target)
            target.execute("CREATE TABLE IF NOT EXISTS app_meta "
                           "(key VARCHAR PRIMARY KEY, value VARCHAR)")
            row = target.execute("SELECT value FROM app_meta WHERE key = ?",
                                 (LOCAL_SEQ_KEY,)).fetchone()
            seq = _as_int(row[0]) if row else 0
            target.executemany("INSERT OR REPLACE INTO app_meta (key, value) "
                               "VALUES (?, ?)",
                               [(VERSION_KEY, str(version)), (DB_ID_KEY, db_id),
                                (PUSHED_SEQ_KEY, str(seq))])
            target.commit()
            check = target.execute("PRAGMA quick_check").fetchone()[0]
            if check != "ok":
                raise sqlite3.DatabaseError(f"snapshot failed quick_check: {check}")
            return _schema_version(target), seq
        finally:
            target.close()
    finally:
        source.close()


def _write_sync_meta(db_path: str, version: int, db_id: str, pushed_seq: int) -> None:
    """Ghi version + db_id + pushed_seq vào DB local. Lỗi ở đây không làm hỏng
    lượt đẩy đã xong — lượt sau vẫn lấy max(local, state.json) + 1 nên version
    không lùi, db_id lấy lại được từ state.json; pushed_seq không ghi được thì
    local chỉ bị coi là còn thay đổi treo và đẩy thêm một lần thừa.

    `pushed_seq` là local_seq LÚC CHỤP: lượt ghi chen vào trong lúc đang đẩy đã
    tăng local_seq vượt con số này, nên vẫn được tính là chưa đẩy.
    """
    try:
        conn = sqlite3.connect(db_path, timeout=10)
        try:
            conn.executemany("INSERT OR REPLACE INTO app_meta (key, value) "
                             "VALUES (?, ?)",
                             [(VERSION_KEY, str(version)), (DB_ID_KEY, db_id),
                              (PUSHED_SEQ_KEY, str(pushed_seq))])
            conn.commit()
        finally:
            conn.close()
    except sqlite3.Error as exc:
        debuglog.exception("shared_db: could not store sync metadata locally", exc)


# ═══════════════════════════════ KHÓA GHI ════════════════════════════════
#
# Thứ tự khóa trong tiến trình (tránh deadlock): KHÔNG BAO GIỜ giữ `_LEASE`
# trong lúc lấy `_LOCK` hay làm việc với ổ mạng — `_LEASE` chỉ bảo vệ vài biến
# trong bộ nhớ, giữ trong tích tắc.

_LEASE = threading.RLock()          # bảo vệ các biến khóa ghi bên dưới
_ACQUIRE_LOCK = threading.Lock()    # mỗi lúc chỉ MỘT lượt giành khóa (nhiều luồng)
_STOP = threading.Event()           # app đang thoát → luồng nhịp tim dừng

_lease_token: str | None = None     # token của khóa MÁY NÀY đang giữ
_lease_operation = "editing"        # việc đang làm — ghi vào khóa cho máy kia thấy
_session_depth = 0                  # số `session()` đang mở (lồng nhau được)
_active = 0                         # lượt ghi + session đang chạy → chưa được nhả
_last_write = 0.0                   # monotonic lúc lượt ghi / session cuối kết thúc
_last_beat_ok = 0.0                 # monotonic lúc nhịp tim ghi được gần nhất
_observed: dict = {}                # {token khóa máy khác: (beat, monotonic lúc thấy)}

_EDITING = "editing"


def holding_lease() -> bool:
    """Máy này đang cầm khóa ghi."""
    with _LEASE:
        return _lease_token is not None


def _holding_usable() -> bool:
    """Đang cầm khóa VÀ nhịp tim gần đây còn ghi được — đủ chắc để nhận lượt
    ghi mới mà không hỏi lại ổ mạng. Gọi khi giữ `_LEASE`.

    Nửa `STALE_AFTER`: máy kia chỉ coi khóa là chết sau trọn `STALE_AFTER`
    giây nhịp tim đứng yên, nên máy này ngừng nhận lượt ghi từ trước đó rất xa.
    """
    return (_lease_token is not None
            and time.monotonic() - _last_beat_ok < STALE_AFTER / 2)


# ─── Cổng ghi (gắn vào cv_repository) ───

def _on_write_begin(conn) -> None:
    """Câu ghi đầu tiên của một kết nối: giành khóa (nếu dùng DB chung) rồi
    tăng sổ đếm lượt ghi. Không ghi được thì ném `DatabaseBusyError`."""
    if is_enabled():
        # Kết nối đang mở → không thay được file .db, nên không kéo bản mới ở đây.
        _enter_lease(_EDITING, allow_pull=False)
        conn._shared_lease = True
    _bump_local_seq(conn)


def _on_write_end(conn, failed: bool) -> None:
    if getattr(conn, "_shared_lease", False):
        conn._shared_lease = False
        _leave_lease()
        if not failed:
            _set_lease_status(dirty=True)


def _bump_local_seq(conn) -> None:
    """Tăng `sync:local_seq` NGAY TRÊN kết nối đang ghi: chung transaction với
    lượt ghi, nên lượt ghi rollback thì sổ đếm cũng lùi theo."""
    try:
        conn.execute("INSERT INTO app_meta (key, value) VALUES (?, '1') "
                     "ON CONFLICT(key) DO UPDATE SET value = CAST(value AS INTEGER) + 1",
                     (LOCAL_SEQ_KEY,))
    except sqlite3.Error as exc:
        debuglog.exception("shared_db: could not bump the write counter", exc)


# ─── Thao tác dài ───

@contextlib.contextmanager
def session(operation: str):
    """Giữ khóa ghi SUỐT một thao tác dài, vd `with shared_db.session("a CV scan"):`.

    - Giành khóa NGAY ĐẦU: máy kia đang ghi / không với tới thư mục dùng chung
      thì ném `DatabaseBusyError` trước khi thao tác bắt đầu (không đợi tới
      lượt ghi đầu tiên sau mấy phút gọi AI). Máy chỉ đọc → `ReadOnlyDatabaseError`.
    - Bản chủ mới hơn thì kéo về rồi giành lại luôn (chưa có kết nối nào mở).
    - Trong session không bao giờ đẩy/nhả, dù khoảng cách giữa hai lượt ghi dài
      hơn ân hạn; kết thúc thì ân hạn tính từ lúc đó.
    - `operation` hiện cho máy kia: "<ai> is running <operation>…".
    Lồng nhau được. Tính năng DB chung tắt thì không làm gì.
    """
    global _session_depth, _lease_operation
    if is_read_only():
        raise cv_repository.ReadOnlyDatabaseError()
    if not is_enabled():
        yield
        return
    _enter_lease(operation, allow_pull=True)
    with _LEASE:
        _session_depth += 1
        if _session_depth == 1:
            _lease_operation = operation
    try:
        yield
    finally:
        with _LEASE:
            _session_depth -= 1
            if _session_depth == 0:
                _lease_operation = _EDITING
        _leave_lease()


# ─── Vào / ra ───

def _enter_lease(operation: str, allow_pull: bool) -> bool:
    """Tính thêm một lượt dùng khóa (giành khóa nếu chưa giữ). Không được thì
    ném `DatabaseBusyError`. Trả về True nếu lượt này VỪA giành khóa mới.

    `allow_pull`: bản chủ mới hơn thì kéo về rồi giành lại — chỉ khi chắc chắn
    không có kết nối nào đang mở (không thì file .db không thay được).
    """
    global _active
    with _LEASE:
        if _holding_usable():
            _active += 1
            return False
    ok, message, fresh = _acquire_with_timeout(operation)
    if (not ok and message == MSG_NEWER and allow_pull
            and cv_repository.open_connection_count() == 0):
        ensure_fresh_with_timeout("before-write", SESSION_PULL_TIMEOUT)
        ok, message, fresh = _acquire_with_timeout(operation)
    if not ok:
        raise cv_repository.DatabaseBusyError(message)
    return fresh


def _leave_lease() -> None:
    global _active, _last_write
    with _LEASE:
        _active = max(0, _active - 1)
        _last_write = time.monotonic()


def _acquire_with_timeout(operation: str) -> tuple[bool, str, bool]:
    """`_acquire` ở luồng phụ, chờ tối đa `ACQUIRE_TIMEOUT` giây.

    Hết giờ mà luồng đó lát sau mới giành được thì nó KHÔNG tính lượt dùng
    (`ticket["abandoned"]`) — khóa vẫn được nhả bình thường khi hết ân hạn,
    thay vì kẹt vĩnh viễn với một lượt dùng không ai trả.
    """
    ticket = {"abandoned": False, "counted": False, "fresh": False}

    def attempt():
        try:
            return _acquire(operation, ticket)
        except Exception as exc:
            debuglog.exception("shared_db: acquiring the write lock failed", exc)
            return False, (f"Couldn't lock the shared database for writing: {exc}. "
                           "Nothing was changed.")

    done, result = _run_with_timeout(attempt, ACQUIRE_TIMEOUT, "shared-db-acquire")
    if done:
        ok, message = result
        return ok, message, ticket["fresh"]
    with _LEASE:
        if ticket["counted"]:          # xong đúng lúc vừa hết giờ
            return True, "", ticket["fresh"]
        ticket["abandoned"] = True
    return False, MSG_NOT_RESPONDING, False


def _count_in(ticket: dict) -> bool:
    """Tính thêm một lượt dùng khóa cho bên gọi (nếu bên gọi còn chờ). Gọi khi
    giữ `_LEASE`."""
    global _active
    if ticket["abandoned"]:
        return False
    _active += 1
    ticket["counted"] = True
    return True


def _acquire(operation: str, ticket: dict) -> tuple[bool, str]:
    """Một lượt giành khóa (luồng phụ — có thể đứng lâu vì ổ mạng).

    Trả về (True, "") — đã cầm khóa và đã tính lượt dùng — hoặc (False, thông
    báo tiếng Anh vì sao không ghi được).
    """
    global _lease_token, _lease_operation, _last_beat_ok, _last_write
    with _ACQUIRE_LOCK:
        with _LEASE:
            if _holding_usable():      # luồng khác vừa giành xong trong lúc chờ
                return (True, "") if _count_in(ticket) else (False, MSG_NOT_RESPONDING)
            current = _lease_token
        folder = shared_folder()
        if not folder:
            with _LEASE:
                _count_in(ticket)
            return True, ""
        if not os.path.isdir(folder):
            return False, MSG_OFFLINE
        path = _lock_path(folder)

        # Đang cầm khóa mà nhịp tim lâu chưa ghi được (vừa chập mạng): khóa trên
        # đĩa vẫn là của mình thì nhận lại luôn, không thì coi như đã mất.
        if current is not None:
            beat = _beat(folder, current)
            if beat:
                with _LEASE:
                    _last_write = time.monotonic()
                    return (True, "") if _count_in(ticket) else (False, MSG_NOT_RESPONDING)
            if beat is None:           # vẫn chưa ghi được nhịp tim — chưa biết còn giữ không
                return False, MSG_NOT_RESPONDING
            _on_lease_lost(current)

        token = uuid.uuid4().hex
        data = _new_lock(token, operation)
        for _attempt in range(4):
            if _write_lock_file(path, data, create=True):
                break
            lock = _read_lock(path)
            if lock is None:           # vừa có người nhả — thử tạo lại
                continue
            if _own_process(lock) or _own_dead(lock) or _is_stale(lock, path):
                if _steal(path, lock):
                    continue
            _request_release(folder)
            _set_lease_status(lease=_who(lock))
            return False, _busy_message(lock)
        else:
            return False, MSG_NOT_RESPONDING

        check = _read_lock(path)
        if not check or check.get("token") != token:
            return False, _busy_message(check or {})

        # Khóa trên đĩa đã là của mình → nhận vào bộ nhớ TRƯỚC khi kiểm độ mới
        # của bản local (để lỡ có lỗi giữa chừng thì vẫn biết mà nhả).
        with _LEASE:
            _lease_token = token
            _lease_operation = operation if operation != "backup" else _EDITING
            _last_beat_ok = _last_write = time.monotonic()
        # Lời "xin nhả" còn sót từ trước không dành cho lượt giữ khóa này.
        _remove_quietly(os.path.join(folder, RELEASE_REQUEST_NAME))

        problem, want_pull = _freshness_problem(folder)
        if problem:
            _drop_lease(token, folder)
            if want_pull:
                check_in_background("newer")
            return False, problem

        with _LEASE:
            counted = _count_in(ticket)
            ticket["fresh"] = True
        _start_lease_thread(token)
        _set_lease_status(lease="mine")
        debuglog.write(f"shared_db: write lock acquired ({operation})"
                       + ("" if counted else " — caller had given up"))
        return (True, "") if counted else (False, MSG_NOT_RESPONDING)


def _freshness_problem(folder: str) -> tuple[str, bool]:
    """Vừa giành khóa: bản local có được phép ghi tiếp không?

    Trả về ("", False) nếu được, hoặc (thông báo, có nên kéo bản mới không).
    Bản chủ mới hơn mà local không có gì treo → kéo về rồi người dùng bấm lại;
    local có thay đổi treo → hai bên đã rẽ nhánh, dừng hẳn.
    """
    state = read_state(folder)
    if not state or not os.path.exists(os.path.join(folder, DB_NAME)):
        return "", False
    local = _local_info(cv_repository._db_path())
    remote_version = _as_int(state.get("version"))
    if not _same_lineage(local, state):
        if _may_adopt(local, str(state.get("db_id") or "")):
            return MSG_NEWER, True
        _set_status(_status_of("blocked", MSG_DIFFERENT_DB))
        return MSG_DIFFERENT_DB, False
    if remote_version > local.version:
        if local.dirty:
            _set_status(_status_of("blocked", MSG_CONFLICT))
            return MSG_CONFLICT, False
        if _schema_too_new(str(state.get("schema_version") or "")):
            _set_status(_status_of("blocked", MSG_UPDATE_APP))
            return MSG_UPDATE_APP, False
        return MSG_NEWER, True
    return "", False


# ─── Nhịp tim · ân hạn · nhả ───

def _start_lease_thread(token: str) -> None:
    threading.Thread(target=_lease_loop, args=(token,), name="shared-db-lease",
                     daemon=True).start()


def _lease_loop(token: str) -> None:
    """Mỗi `HEARTBEAT` giây: ghi nhịp tim; hết ân hạn hoặc máy kia xin thì
    đẩy + nhả. Một luồng cho mỗi lượt giữ khóa, tự dừng khi khóa đổi chủ."""
    while not _STOP.wait(HEARTBEAT):
        with _LEASE:
            if _lease_token != token:
                return
        folder = shared_folder()
        if not folder:                 # vừa tắt tính năng DB chung
            _drop_lease(token, None)
            return
        try:
            beat = _beat(folder, token)
        except Exception as exc:
            debuglog.exception("shared_db: heartbeat failed", exc)
            beat = None
        if beat is False:
            _on_lease_lost(token)
            return
        if beat is None:               # ổ mạng không với tới — giữ khóa, nhịp sau thử lại
            continue
        with _LEASE:
            if _lease_token != token:
                return
            busy = _active > 0
            idle = time.monotonic() - _last_write >= GRACE
        if busy:
            continue
        requested = os.path.exists(os.path.join(folder, RELEASE_REQUEST_NAME))
        if idle or requested:
            _flush_and_release(token, "requested" if requested else "idle")


def _beat(folder: str, token: str):
    """Ghi một nhịp tim. True = xong · False = khóa không còn là của mình ·
    None = chưa ghi được (ổ mạng không với tới, file đang được ghi dở)."""
    global _last_beat_ok
    if not os.path.isdir(folder):
        return None
    path = _lock_path(folder)
    lock = _read_lock(path)
    if lock is None or (lock and lock.get("token") != token):
        return False
    if not lock:
        return None
    with _LEASE:
        operation = _lease_operation
    now = datetime.datetime.now()
    lock.update(heartbeat=now.isoformat(timespec="seconds"),
                heartbeat_epoch=time.time(),
                beat=_as_int(lock.get("beat")) + 1, operation=operation)
    _write_lock_file(path, lock, create=False)
    with _LEASE:
        _last_beat_ok = time.monotonic()
    return True


def _flush_and_release(token: str, reason: str, force: bool = False) -> PushResult | None:
    """Đẩy thay đổi còn chờ rồi nhả khóa.

    CHƯA ĐẨY ĐƯỢC THÌ KHÔNG NHẢ — đó là nguyên tắc giữ cả mô hình an toàn.
    Có lượt ghi chen vào trong lúc đang đẩy thì cũng chưa nhả, để nhịp sau đẩy
    tiếp. `force` (lúc đóng app): không đợi lượt ghi đang chạy.
    """
    global _lease_token
    folder = shared_folder()
    if not folder:
        _drop_lease(token, None)
        return None
    db_path = cv_repository._db_path()
    with _LEASE:
        if _lease_token != token or (_active > 0 and not force):
            return None

    result = None
    local = _local_info(db_path)
    if local.dirty or not os.path.exists(os.path.join(folder, DB_NAME)):
        result = _push_locked(reason, folder)
        if not result.ok:
            return result

    if not force and _local_info(db_path).dirty:
        return result
    with _LEASE:
        if _lease_token != token or (_active > 0 and not force):
            return result
        _lease_token = None
    _release_lock_file(folder, token)
    _remove_quietly(os.path.join(folder, RELEASE_REQUEST_NAME))
    _set_lease_status(lease="", dirty=_local_info(db_path).dirty)
    debuglog.write(f"shared_db: write lock released ({reason})")
    return result


def _release_if_clean(reason: str) -> None:
    """Nhả khóa ngay nếu không còn gì cần nó (sau lượt Back up now đã tự giành)."""
    with _LEASE:
        token = _lease_token
    if token is not None:
        _flush_and_release(token, reason)


def _drop_lease(token: str, folder: str | None) -> None:
    """Bỏ khóa KHÔNG đẩy gì — chỉ dùng khi chưa kịp ghi gì dưới khóa này."""
    global _lease_token
    with _LEASE:
        if _lease_token != token:
            return
        _lease_token = None
    if folder:
        _release_lock_file(folder, token)
    _set_lease_status(lease="")


def _on_lease_lost(token: str) -> None:
    """Khóa trên đĩa không còn là của mình (máy kia giành lại khi máy này mất
    mạng lâu). Có thay đổi chưa đẩy thì hai bên đã rẽ nhánh — dừng đồng bộ."""
    global _lease_token
    with _LEASE:
        if _lease_token != token:
            return
        _lease_token = None
    dirty = _local_info(cv_repository._db_path()).dirty
    debuglog.write("shared_db: the write lock was taken over by another computer"
                   + (" — with changes not uploaded" if dirty else ""))
    if dirty:
        _set_status(_status_of("blocked", MSG_LEASE_LOST, lease="", dirty=True))
    else:
        _set_lease_status(lease="")


# ─── File khóa ───

def _lock_path(folder: str) -> str:
    return os.path.join(folder, LOCK_NAME)


def _identity() -> dict:
    return {
        "login": cv_repository.windows_login() or "unknown",
        "machine": (os.environ.get("COMPUTERNAME") or socket.gethostname()
                    or "unknown").strip(),
        "pid": os.getpid(),
    }


def _new_lock(token: str, operation: str) -> dict:
    now = datetime.datetime.now()
    return {**_identity(), "token": token,
            "since": now.isoformat(timespec="seconds"),
            "heartbeat": now.isoformat(timespec="seconds"),
            "heartbeat_epoch": time.time(), "beat": 0,
            "operation": operation if operation != "backup" else _EDITING}


def _read_lock(path: str) -> dict | None:
    """None = không có khóa · {} = có file mà chưa đọc được (đang ghi dở / hỏng)."""
    for attempt in range(2):
        try:
            with open(path, encoding="utf-8") as fh:
                data = json.load(fh)
            return data if isinstance(data, dict) else {}
        except FileNotFoundError:
            return None
        except (OSError, ValueError):
            if attempt == 0:
                time.sleep(0.2)    # đang bị os.replace đè — đọc lại một lần
    return {}


def _write_lock_file(path: str, data: dict, create: bool) -> bool:
    """Ghi khóa qua file tạm. `create`: chỉ thành công khi CHƯA có khóa (tạo-
    nếu-chưa-có nguyên tử) → False nếu đã có người giữ."""
    tmp = f"{path}.{data['token']}.tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
    try:
        if not create:
            _replace(tmp, path)
        elif os.name == "nt":
            os.rename(tmp, path)       # Windows: lỗi nếu đích đã có
        else:
            os.link(tmp, path)
            os.remove(tmp)
        return True
    except FileExistsError:
        _remove_quietly(tmp)
        return False
    except OSError:
        _remove_quietly(tmp)
        raise


def _release_lock_file(folder: str, token: str) -> None:
    path = _lock_path(folder)
    lock = _read_lock(path)
    if lock and lock.get("token") == token:
        try:
            os.remove(path)
        except OSError as exc:
            debuglog.exception("shared_db: could not remove the lock file", exc)


def _request_release(folder: str) -> None:
    """Nhờ máy đang giữ khóa nhả sớm (nó thấy ở nhịp tim kế tiếp), thay vì bắt
    người dùng chờ hết ân hạn."""
    try:
        ident = _identity()
        _write_json_atomic(os.path.join(folder, RELEASE_REQUEST_NAME), {
            **ident, "at": datetime.datetime.now().isoformat(timespec="seconds"),
        }, f"{ident['login']}.{os.getpid()}")
    except OSError as exc:
        debuglog.exception("shared_db: could not write the release request", exc)


def _mine(lock: dict) -> bool:
    ident = _identity()
    return (str(lock.get("machine") or "").lower() == ident["machine"].lower()
            and str(lock.get("login") or "").lower() == ident["login"].lower())


def _own_process(lock: dict) -> bool:
    """Khóa do chính tiến trình này để lại (đang nhả dở, lượt giành bị bỏ chờ)."""
    return _mine(lock) and _as_int(lock.get("pid")) == os.getpid()


def _own_dead(lock: dict) -> bool:
    """Khóa của chính máy + tài khoản này mà tiến trình giữ nó đã chết (app vừa
    crash rồi mở lại) — giành lại ngay, khỏi đợi `STALE_AFTER`."""
    pid = _as_int(lock.get("pid"))
    return _mine(lock) and pid != os.getpid() and not _pid_alive(pid)


def _is_stale(lock: dict, path: str) -> bool:
    """Khóa của máy khác đã chết chưa — xem docstring của STALE_AFTER(_WALL)."""
    global _observed
    key = str(lock.get("token") or "?")
    beat = lock.get("beat")
    now = time.monotonic()
    seen = _observed.get(key)
    if seen is None or seen[0] != beat:
        _observed = {key: (beat, now)}
        unchanged = 0.0
    else:
        unchanged = now - seen[1]
    try:
        stamp = float(lock.get("heartbeat_epoch"))
    except (TypeError, ValueError):
        try:
            stamp = os.path.getmtime(path)
        except OSError:
            stamp = time.time()
    return unchanged >= STALE_AFTER or time.time() - stamp >= STALE_AFTER_WALL


def _steal(path: str, lock: dict) -> bool:
    """Dời khóa chết sang tên khác để tạo khóa mới. True = đã dời (hoặc nó vừa
    tự biến mất); False = không dời được / dời nhầm khóa còn sống (đã trả lại)."""
    grave = f"{path}.stale-{uuid.uuid4().hex}"
    try:
        os.rename(path, grave)
    except FileNotFoundError:
        return True
    except OSError:
        return False
    moved = _read_lock(grave) or {}
    if moved.get("token") != lock.get("token"):
        # Máy khác vừa tạo khóa mới chen vào giữa lúc đọc và lúc dời → trả lại.
        try:
            os.rename(grave, path)
        except OSError as exc:
            debuglog.exception("shared_db: could not put back a live lock", exc)
        return False
    _remove_quietly(grave)
    debuglog.write(f"shared_db: took over a dead lock of {lock.get('login')} on "
                   f"{lock.get('machine')} (since {lock.get('since')})")
    return True


def _observe_lock(folder: str) -> None:
    """Đọc khóa để giao diện biết ai đang ghi (và để `_is_stale` có mốc so)."""
    path = _lock_path(folder)
    lock = _read_lock(path)
    if not lock:
        _set_lease_status(lease="")
        return
    if _own_process(lock):
        return
    _set_lease_status(lease="" if _is_stale(lock, path) else _who(lock))


def _who(lock: dict) -> str:
    """Tên người giữ khóa để hiện lên giao diện, vd "Lan (VN-PC12)"."""
    login = str(lock.get("login") or "")
    if not login:
        return "Another computer"
    name = login
    try:
        conn = _connect_ro(cv_repository._db_path())
        try:
            if _has_table(conn, "users"):
                row = conn.execute("SELECT display_name FROM users WHERE "
                                   "windows_login = ?", (login.lower(),)).fetchone()
                if row and row[0]:
                    name = str(row[0])
        finally:
            conn.close()
    except sqlite3.Error:
        pass
    machine = str(lock.get("machine") or "")
    return f"{name} ({machine})" if machine else name


def _busy_message(lock: dict) -> str:
    who = _who(lock)
    since = _hhmm(lock.get("since"))
    since = f" (since {since})" if since else ""
    operation = str(lock.get("operation") or _EDITING)
    if operation == _EDITING:
        return (f"{who} is saving to the shared database right now{since}. They've "
                "been asked to finish up — please try again in a few seconds. "
                "Nothing was changed.")
    return (f"{who} is running {operation} on the shared database{since}. Please "
            "try again when it finishes. Nothing was changed.")


def _hhmm(iso) -> str:
    try:
        return datetime.datetime.fromisoformat(str(iso)).strftime("%H:%M")
    except (TypeError, ValueError):
        return ""


def _pid_alive(pid: int) -> bool:
    """Tiến trình `pid` còn chạy không (không chắc thì coi như còn)."""
    if pid <= 0:
        return False
    try:
        if os.name != "nt":
            os.kill(pid, 0)
            return True
        import ctypes
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        handle = kernel32.OpenProcess(0x1000, False, pid)  # QUERY_LIMITED_INFORMATION
        if not handle:
            return ctypes.get_last_error() == 5             # bị từ chối = vẫn còn
        try:
            code = ctypes.c_ulong()
            if not kernel32.GetExitCodeProcess(handle, ctypes.byref(code)):
                return True
            return code.value == 259                        # STILL_ACTIVE
        finally:
            kernel32.CloseHandle(handle)
    except ProcessLookupError:
        return False
    except Exception:
        return True


def _start_recover_push() -> None:
    """Đẩy nốt thay đổi còn treo từ phiên trước (ở nền, qua khóa ghi)."""
    global _recover_thread
    if is_read_only() or (_recover_thread is not None and _recover_thread.is_alive()):
        return

    def run():
        result = push("recover")
        debuglog.write(f"shared_db: recover push — {result.message}")

    _recover_thread = threading.Thread(target=run, name="shared-db-recover", daemon=True)
    _recover_thread.start()


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


def _replace(src: str, dst: str, attempts: int = 10, delay: float = 0.1) -> None:
    """`os.replace` có thử lại.

    Trên Windows, đè lên một file đang bị tiến trình khác mở — máy kia đang tải
    bản chủ về, antivirus đang quét file vừa ghi — báo `Access is denied`, mà
    thường chỉ kéo dài tích tắc. Hết lượt thử mới ném lỗi ra.
    """
    for attempt in range(attempts):
        try:
            os.replace(src, dst)
            return
        except PermissionError:
            if attempt == attempts - 1:
                raise
            time.sleep(delay)


def _write_json_atomic(path: str, data: dict, login: str) -> None:
    tmp = f"{path}.{login}.tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
    _replace(tmp, path)


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
        _replace(tmp, os.path.join(hist, name))

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
cv_repository.set_write_hooks(_on_write_begin, _on_write_end)
