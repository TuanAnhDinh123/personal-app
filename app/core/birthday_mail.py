"""Mail chúc mừng sinh nhật — LỊCH GỬI DO APP GIỮ (không nhờ Outlook nữa).

Trước đây cả tháng mail được đẩy sang Outlook một lượt, mỗi mail gắn
`DeferredDeliveryTime` để Outlook giữ ở Outbox tới đúng ngày. Cách đó chỉ chạy
khi tài khoản gửi ĐÃ ĐĂNG NHẬP trong Outlook; gửi từ hộp thư dùng chung (chỉ
được IT share, không login) thì Outlook bỏ qua giờ hẹn và gửi ngay — cả tháng
nhận mail cùng lúc.

Giờ chia làm ba nhịp, và NGƯỜI DÙNG DUYỆT Ở CẢ HAI ĐẦU:

  • `enqueue()`     — duyệt danh sách sinh nhật trong tháng rồi XẾP HÀNG vào
                      bảng `birthday_emails`. Không gửi gì cả.
  • `prepare_due()` — mỗi lần MỞ APP: dọn hàng chờ và trả về những dòng đã tới
                      ngày (kể cả muộn vài ngày, trong hạn
                      `birthday_catchup_days`). Vẫn CHƯA gửi.
  • `send_rows()`   — gửi đúng những dòng người dùng đã xác nhận ở modal.

Tách `prepare_due` khỏi `send_rows` chính là để chèn được modal xác nhận vào
giữa: app không bao giờ tự gửi mail mà không hỏi. Bấm Cancel thì mọi dòng vẫn
nằm nguyên trong hàng chờ, gửi tay sau được từ màn hình Queue.

Module này thuần Python, KHÔNG dính giao diện: cả hai hàm trên phải chạy được
lúc mở app khi người dùng chưa từng mở trang tool.

Hàng chờ chỉ lưu khóa + trạng thái. Email, họ tên và đường dẫn thiệp đều được
TRA LẠI lúc gửi, không chụp lại lúc xếp hàng — nhờ vậy không tồn tại các lỗi dữ
liệu cũ (thiệp bị xóa, đổi mail công ty, nhân viên nghỉ việc sau khi xếp hàng).
"""
import datetime
import os
import re
import sqlite3

from app.core import cv_repository as repo
from app.core import cv_schema, debuglog, outlook, settings

_IMAGE_EXTS = {".jpg", ".jpeg", ".png"}

DEFAULT_SUBJECT = "Happy Birthday, {name}!"

# Số ngày còn gửi bù khi hôm sinh nhật không mở app (setting đọc không ra thì
# dùng mốc này).
DEFAULT_CATCHUP_DAYS = 3

# Dòng giành được rồi mà quá lâu chưa xong thì coi như lượt gửi đã chết (app bị
# tắt giữa lúc gọi Outlook).
_STALE_SENDING_HOURS = 1


def _fill(text, name):
    return (text or "").replace("{name}", name or "")


def day_month(dob):
    """'dd/mm/yyyy' -> (ngày, tháng) dạng int, hoặc None nếu không đọc được."""
    try:
        d, m, _y = (dob or "").strip().split("/")
        return int(d), int(m)
    except (ValueError, AttributeError):
        return None


def birth_month(dob):
    dm = day_month(dob)
    return dm[1] if dm else None


def birthday_date(dob, year):
    """Ngày sinh nhật của `year` (chỉ ngày, không giờ) — mốc `due_date`.

    Ngày sinh 29/2 rơi vào năm không nhuận (hoặc dữ liệu ngày lệch như 31/4)
    thì lùi dần tối đa 3 ngày cho ra ngày hợp lệ, thay vì bỏ qua người đó.
    """
    dm = day_month(dob)
    if not dm:
        return None
    day, month = dm
    for offset in range(4):
        try:
            return datetime.date(year, month, day - offset)
        except ValueError:
            continue
    return None


_CANVA_FNAME_RE = re.compile(r"^\d+-(.+)$")


def _card_code_from_stem(stem):
    """Canva Bulk Create xuất file dạng '<stt>-<mã NV>' (vd '1-20170456',
    '2-20184578') — trả về phần mã NV. File không theo mẫu này (đặt tên trực
    tiếp bằng mã NV, kiểu cũ) thì trả nguyên tên."""
    m = _CANVA_FNAME_RE.match(stem)
    return m.group(1) if m else stem


def scan_cards(folder):
    """Map mã NV (chuẩn hóa hoa) -> tên file ảnh, trong `folder`."""
    images = {}
    if folder and os.path.isdir(folder):
        for fname in os.listdir(folder):
            stem, ext = os.path.splitext(fname)
            stem = stem.strip()
            if ext.lower() in _IMAGE_EXTS and stem:
                code = _card_code_from_stem(stem).strip()
                if code:
                    images[code.upper()] = fname
    return images


def cards_folder():
    return (settings.get("birthday_images_folder") or "").strip()


def catchup_days():
    """Số ngày còn gửi bù (setting nhập tay nên phải chịu được giá trị rác)."""
    try:
        return max(0, int(str(settings.get("birthday_catchup_days")).strip()))
    except (TypeError, ValueError):
        return DEFAULT_CATCHUP_DAYS


def subject_template():
    return (settings.get("birthday_subject") or "").strip() or DEFAULT_SUBJECT


def _card_path(folder, images, code):
    fname = images.get((code or "").strip().upper())
    return os.path.join(folder, fname) if fname else None


def _recipient(row):
    """Ưu tiên mail công ty, chỉ dùng mail cá nhân khi công ty chưa cấp (vd
    nhân viên mới)."""
    return ((row["company_email"] or "") or (row["email"] or "")).strip()


# ─────────────────────────────── XẾP HÀNG ────────────────────────────────

def month_candidates(today=None):
    """Người có sinh nhật TRONG THÁNG của `today`, kèm mọi thứ modal cần để
    người dùng duyệt.

    `repo.list_employees()` mặc định chỉ trả người đang làm việc nên người đã
    nghỉ không bao giờ lọt vào đây.
    """
    today = today or datetime.date.today()
    folder = cards_folder()
    images = scan_cards(folder)
    rows = []
    for emp in repo.list_employees():
        if birth_month(emp["date_of_birth"]) != today.month:
            continue
        due = birthday_date(emp["date_of_birth"], today.year)
        queued = (repo.find_birthday_email(emp["employee_id"], due.isoformat())
                  if due else None)
        rows.append({
            "employee_id": emp["employee_id"],
            "code": emp["code"],
            "name": emp["name"] or emp["full_name"] or emp["code"],
            "raw_full_name": emp["full_name"] or emp["code"],
            "email": _recipient(emp),
            "date_of_birth": emp["date_of_birth"],
            "card_path": _card_path(folder, images, emp["code"]),
            "due_date": due,
            # Đã qua sinh nhật trong tháng -> xếp hàng là gửi ngay lượt sau.
            "passed": bool(due and due < today),
            "queued_status": queued["status"] if queued else "",
        })
    return rows


def enqueue(rows):
    """Ghi các dòng đã duyệt vào hàng chờ. Trả về {"queued", "requeued", "skipped"}.

    Xếp hàng lại người đã có dòng thì KHÔNG tạo bản ghi trùng: đã `Sent` hoặc
    đang chờ thì bỏ qua, còn `Failed`/`Missed`/`Cancelled` thì đặt lại `Pending`
    (đây là đường để gửi lại người tháng trước bị lỗi).
    """
    result = {"queued": [], "requeued": [], "skipped": []}
    for row in rows:
        due = row.get("due_date")
        if not due:
            result["skipped"].append(row["raw_full_name"])
            continue
        due_iso = due.isoformat()
        existing = repo.find_birthday_email(row["employee_id"], due_iso)
        if existing:
            if existing["status"] in (cv_schema.BIRTHDAY_EMAIL_SENT,
                                      cv_schema.BIRTHDAY_EMAIL_PENDING,
                                      cv_schema.BIRTHDAY_EMAIL_SENDING):
                result["skipped"].append(row["raw_full_name"])
                continue
            repo.update_birthday_email(existing["birthday_email_id"], {
                "status": cv_schema.BIRTHDAY_EMAIL_PENDING,
                "last_error": None,
            })
            result["requeued"].append(row["raw_full_name"])
            continue
        try:
            repo.insert_birthday_email({
                "employee_id": row["employee_id"],
                "due_date": due_iso,
                "status": cv_schema.BIRTHDAY_EMAIL_PENDING,
                "attempts": 0,
            })
        except sqlite3.IntegrityError:
            # Lưới an toàn của unique (employee_id, due_date) — chỉ xảy ra khi
            # mở app hai lần và cùng xếp hàng. Coi như đã có, không phải lỗi.
            result["skipped"].append(row["raw_full_name"])
            continue
        result["queued"].append(row["raw_full_name"])
    return result


def month_is_queued(today=None) -> bool:
    """Tháng này đã xếp hàng chưa — dùng cho dòng nhắc trên trang tool.

    Không có ai sinh nhật trong tháng cũng trả True: không có gì phải nhắc.

    KHÔNG dùng `month_candidates()`: hàm này bị gọi mỗi lần nạp lại bảng, mà
    `month_candidates` quét cả thư mục thiệp — thư mục đó thường ở ổ mạng nên
    `os.listdir` có thể treo giao diện vài giây. Ở đây chỉ cần biết có dòng nào
    trong hàng chờ hay chưa, không cần biết ai có thiệp.
    """
    today = today or datetime.date.today()
    has_birthday = False
    for emp in repo.list_employees():
        if birth_month(emp["date_of_birth"]) != today.month:
            continue
        has_birthday = True
        due = birthday_date(emp["date_of_birth"], today.year)
        if due and repo.find_birthday_email(emp["employee_id"], due.isoformat()):
            return True
    return not has_birthday


# ──────────────────────────────── GỬI ────────────────────────────────────

def _send_row(row, folder, images, subject_tpl):
    """Gửi một dòng đã GIÀNH được. Trả về (thành công?, mô tả).

    Mọi dữ liệu người nhận tra lại tại đây, không dùng bản chụp lúc xếp hàng.
    """
    row_id = row["birthday_email_id"]
    display = row["full_name"] or row["employee_code"] or f"#{row['employee_id']}"

    if row["joined_employee_id"] is None:
        repo.mark_birthday_email_failed(
            row_id, "employee no longer in the database",
            status=cv_schema.BIRTHDAY_EMAIL_CANCELLED, count_attempt=False)
        return False, f"{display} — employee no longer in the database"

    if (row["termination_date"] or "").strip():
        repo.mark_birthday_email_failed(
            row_id, "employee resigned",
            status=cv_schema.BIRTHDAY_EMAIL_CANCELLED, count_attempt=False)
        return False, f"{display} — employee resigned"

    email = _recipient(row)
    if not email:
        repo.mark_birthday_email_failed(row_id, "missing email")
        return False, f"{display} — missing email"

    card = _card_path(folder, images, row["employee_code"])
    if not card or not os.path.isfile(card):
        repo.mark_birthday_email_failed(row_id, "no matching card")
        return False, f"{display} — no matching card"

    short_name = row["name"] or row["full_name"] or row["employee_code"]
    try:
        outlook.send_mail(
            email, _fill(subject_tpl, short_name), "",
            account_smtp=(settings.get("birthday_from_account") or "").strip() or None,
            attachments=[card], inline_attachment=True)
    except Exception as exc:
        repo.mark_birthday_email_failed(row_id, f"send failed: {exc}")
        return False, f"{display} — send failed: {exc}"

    repo.mark_birthday_email_sent(row_id)
    return True, display


def _blocked_reason(row, folder, images):
    """Lý do một dòng KHÔNG gửi được, "" nếu gửi được. Thứ tự kiểm tra = thứ tự
    ưu tiên hiển thị."""
    if row["joined_employee_id"] is None:
        return "employee no longer in the database"
    if (row["termination_date"] or "").strip():
        return "employee resigned"
    if not _recipient(row):
        return "missing email"
    card = _card_path(folder, images, row["employee_code"])
    if not card or not os.path.isfile(card):
        return "no matching card"
    return ""


# Hai lý do là VĨNH VIỄN (không bao giờ nên thử lại) -> Cancelled; còn lại là
# Failed, sửa được rồi lượt sau gửi tiếp.
_TERMINAL_REASONS = ("employee resigned", "employee no longer in the database")


def prepare_due(today=None, grace_days=None):
    """Dọn hàng chờ rồi trả về danh sách TỚI HẠN để người dùng DUYỆT. KHÔNG gửi.

    Trả về {"rows", "missed", "skipped"} — `skipped` là lý do cả lượt không chạy
    được (không có Outlook, thư mục thiệp chưa truy cập được…). Mỗi phần tử
    `rows`: row_id · display · email · due_date · card_path · blocked.

    Tên là `prepare_` chứ không phải `preview_` vì hàm này CÓ SỬA DB: ngoài việc
    dọn dòng kẹt và đánh `Missed`, nó ghi luôn lý do vào những dòng không gửi
    được (thiếu thiệp/mail → `Failed`, đã nghỉ việc → `Cancelled`). Làm vậy để
    bảng hàng chờ nhìn thấy được VÌ SAO một người không được gửi; nếu chỉ để
    nguyên `Pending` thì người dùng không có cách nào biết. `Failed` vẫn nằm
    trong diện tới hạn nên bổ sung thiệp là lượt sau gửi được ngay.
    """
    today = today or datetime.date.today()
    grace = catchup_days() if grace_days is None else max(0, int(grace_days))
    out = {"rows": [], "missed": 0, "skipped": ""}

    if not outlook.available():
        out["skipped"] = "Outlook is not available"
        return out

    repo.reset_stale_sending(
        (datetime.datetime.now()
         - datetime.timedelta(hours=_STALE_SENDING_HOURS)
         ).strftime("%Y-%m-%d %H:%M:%S"))

    earliest = today - datetime.timedelta(days=grace)
    rows = repo.due_birthday_emails(today.isoformat(), earliest.isoformat())

    # Thư mục thiệp thường nằm trên ổ mạng/OneDrive: lúc mới mở máy có thể chưa
    # mount. `scan_cards` trả {} cho thư mục thiếu, cứ chạy tiếp là đánh Failed
    # oan cho MỌI người tới hạn hôm đó -> thà để nguyên Pending, lần mở app sau
    # (hoặc trong hạn gửi bù) thử lại.
    #
    # Kiểm tra thư mục TRƯỚC bước đánh Missed: thư mục lỗi mấy ngày liền không
    # được phép ăn mất hạn gửi bù của người ta, vì đó không phải lỗi "không ai
    # mở app". Lượt nào thư mục lành thì mới tính lại hạn.
    folder = cards_folder()
    if not folder or not os.path.isdir(folder):
        if rows:
            out["skipped"] = "birthday cards folder is not reachable"
            debuglog.write(f"birthday_mail: {out['skipped']} ({folder!r}) — "
                           f"{len(rows)} email(s) left pending")
        return out

    out["missed"] = repo.expire_birthday_emails(earliest.isoformat())
    if not rows:
        return out

    images = scan_cards(folder)
    for row in rows:
        blocked = _blocked_reason(row, folder, images)
        if blocked:
            repo.mark_birthday_email_failed(
                row["birthday_email_id"], blocked,
                status=(cv_schema.BIRTHDAY_EMAIL_CANCELLED
                        if blocked in _TERMINAL_REASONS else ""),
                count_attempt=False)
        out["rows"].append({
            "row_id": row["birthday_email_id"],
            "display": (row["full_name"] or row["employee_code"]
                        or f"#{row['employee_id']}"),
            "code": row["employee_code"] or "—",
            "email": _recipient(row),
            "due_date": row["due_date"],
            "card_path": _card_path(folder, images, row["employee_code"]),
            "blocked": blocked,
        })
    return out


def send_rows(row_ids):
    """Gửi đúng những dòng được chỉ định (đã qua bước người dùng duyệt).

    Mọi thứ được TRA LẠI trong `_send_row` chứ không tin bản chụp của
    `prepare_due()`: giữa lúc hiện modal và lúc bấm Send, nhân viên có thể vừa
    nghỉ việc hoặc thiệp vừa bị xóa.
    """
    out = {"sent": [], "failed": [], "missed": 0, "skipped": ""}
    if not outlook.available():
        out["skipped"] = "Outlook is not available"
        return out

    folder = cards_folder()
    if not folder or not os.path.isdir(folder):
        out["skipped"] = "birthday cards folder is not reachable"
        return out

    images = scan_cards(folder)
    subject_tpl = subject_template()
    for row_id in row_ids:
        if not repo.claim_birthday_email(row_id):
            row = repo.get_birthday_email(row_id)
            name = (row["full_name"] if row else None) or f"#{row_id}"
            out["failed"].append(f"{name} — already sent or being sent")
            continue
        row = repo.get_birthday_email(row_id)
        ok, detail = _send_row(row, folder, images, subject_tpl)
        (out["sent"] if ok else out["failed"]).append(detail)

    debuglog.write(
        f"birthday_mail.send_rows: sent={len(out['sent'])} "
        f"failed={len(out['failed'])}"
        + (f" | failures: {'; '.join(out['failed'])}" if out["failed"] else ""))
    return out
