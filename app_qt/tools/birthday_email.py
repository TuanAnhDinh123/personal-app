"""Gửi mail chúc mừng sinh nhật — đính kèm ảnh thiệp xuất từ Canva Bulk Create.

Bấm "Send": tự quét bảng nhân viên tìm người có NGÀY SINH TRONG THÁNG HIỆN
TẠI, đối chiếu mã NV với thư mục ảnh đã cấu hình ở Cài đặt (tên file = mã NV)
→ hiện modal xác nhận (ai THIẾU ảnh bị đánh dấu, sẽ KHÔNG được gửi) → xác nhận
thì đẩy mail qua Outlook (có thể dùng tài khoản riêng, khác tài khoản mặc định
— cũng cấu hình ở Cài đặt) → báo lại số mail đã xếp hàng + tên những người
chưa được gửi.

Mail KHÔNG đi ngay: mỗi mail được HẸN GIỜ (DeferredDeliveryTime của Outlook)
đúng ngày sinh nhật của người đó, vào giờ chọn ở ô "Delivery time" — nên chạy
một lần đầu tháng là cả tháng tự gửi. Trong lúc chờ, mail nằm ở Outbox: tài
khoản Exchange thì server giữ hộ, tài khoản POP/IMAP thì Outlook phải đang mở
lúc tới hạn. Ai đã qua sinh nhật trong tháng thì gửi ngay (không hẹn được nữa).
"""
import csv
import datetime
import os
import re
import unicodedata

from PySide6.QtWidgets import (
    QCheckBox, QDialog, QFileDialog, QFrame, QHBoxLayout, QLabel, QVBoxLayout, QWidget,
)

from app.core import cv_repository as repo
from app.core import outlook
from app.core import settings
from app_qt import dialogs, theme, widgets
from app_qt.base_tool import BaseTool
from app_qt.components.modal import ModalDialog

_IMAGE_EXTS = {".jpg", ".jpeg", ".png"}

_DEFAULT_SEND_TIME = "08:00"

_DEFAULT_SUBJECT = "Happy Birthday, {name}!"


def _fill(text, name):
    return text.replace("{name}", name)


def _day_month(dob):
    """'dd/mm/yyyy' -> (ngày, tháng) dạng int, hoặc None nếu không đọc được."""
    try:
        d, m, _y = (dob or "").strip().split("/")
        return int(d), int(m)
    except (ValueError, AttributeError):
        return None


def _birth_month(dob):
    dm = _day_month(dob)
    return dm[1] if dm else None


def _time_slots():
    """Các mốc giờ cho ô "Delivery time" — mỗi 30 phút trong giờ hành chính."""
    return [f"{h:02d}:{m:02d}" for h in range(6, 21) for m in (0, 30)]


def _parse_time(text):
    """'HH:MM' -> datetime.time; đọc không được thì về mặc định 08:00."""
    try:
        h, m = (text or "").strip().split(":")
        return datetime.time(int(h), int(m))
    except (ValueError, AttributeError):
        return datetime.time(8, 0)


def _birthday_datetime(dob, at_time, year):
    """Thời điểm hẹn gửi = ngày/tháng sinh của `year`, vào giờ `at_time`.

    Ngày sinh 29/2 rơi vào năm không nhuận (hoặc dữ liệu ngày lệch như 31/4)
    thì lùi dần tối đa 3 ngày cho ra ngày hợp lệ, thay vì bỏ qua người đó.
    """
    dm = _day_month(dob)
    if not dm:
        return None
    day, month = dm
    for offset in range(4):
        try:
            d = datetime.date(year, month, day - offset)
        except ValueError:
            continue
        return datetime.datetime.combine(d, at_time)
    return None


def _fmt_when(dt):
    return dt.strftime("%d %b · %H:%M") if dt else "—"


def _title_case_name(name):
    """'NGUYỄN VĂN A' -> 'Nguyễn Văn A' — viết hoa chữ đầu mỗi từ, phần còn
    lại viết thường."""
    return " ".join(w.capitalize() for w in (name or "").split())


def _strip_vn_accents(text):
    """Bỏ dấu tiếng Việt (đ/Đ xử lý riêng vì NFD không tách được ký tự này)."""
    text = (text or "").replace("đ", "d").replace("Đ", "D")
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return unicodedata.normalize("NFC", text)


_CANVA_FNAME_RE = re.compile(r"^\d+-(.+)$")


def _card_code_from_stem(stem):
    """Canva Bulk Create xuất file dạng '<stt>-<mã NV>' (vd '1-20170456',
    '2-20184578') — trả về phần mã NV. File không theo mẫu này (đặt tên trực
    tiếp bằng mã NV, kiểu cũ) thì trả nguyên tên."""
    m = _CANVA_FNAME_RE.match(stem)
    return m.group(1) if m else stem


def _scan_images(folder):
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


def _chip(parent, text, color):
    lbl = QLabel(text, parent)
    r, g, b = widgets._hex_to_rgb(color)
    lbl.setStyleSheet(
        f"background: rgba({r},{g},{b},0.15); color:{color}; border-radius:10px;"
        " padding:3px 10px; font-size:12px; font-weight:600;")
    return lbl


class BirthdayEmailTool(BaseTool):
    name = "Birthday emails"
    description = "Send birthday cards (Canva-exported images) to employees by email."
    icon = "🎂"
    category = "Office"
    order = 7

    def build(self, parent=None):
        card = widgets.Card(parent)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(22, 20, 22, 18)
        lay.setSpacing(10)
        self._root = card

        widgets.section_label(card, "Email content")
        self.subject_field = widgets.text_row(card, "Subject")
        self.subject_field.set(_DEFAULT_SUBJECT)
        widgets.hint(card, "No email body — the card image is the message.")

        widgets.section_label(card, "Delivery")
        self.time_field = widgets.dropdown(card, "Delivery time", _time_slots())
        self.time_field.set(settings.get("birthday_send_time") or _DEFAULT_SEND_TIME)
        widgets.hint(
            card, "Emails are not sent right away: each one waits in Outlook's "
                  "Outbox and goes out on the employee's own birthday at this "
                  "time. Keep Outlook running and signed in so queued emails "
                  "can leave the Outbox.")

        send_bar = QHBoxLayout()
        send_bar.setContentsMargins(0, 8, 0, 0)
        send_bar.addWidget(widgets.button(card, "Review birthdays", variant="primary",
                                          icon="search", command=self._on_send))
        send_bar.addWidget(widgets.button(
            card, "Export CSV (missing cards)", variant="neutral",
            icon="file-text", command=self._export_missing_csv))
        send_bar.addStretch(1)
        lay.addLayout(send_bar)
        return card

    def build_body(self, parent):
        pass

    # -------------------------------------------------------------- quét + xác nhận
    def _on_send(self):
        if not outlook.available():
            dialogs.warning(self._root, "Outlook required",
                            "Sending email needs Outlook on Windows (pywin32).")
            return

        now = datetime.datetime.now()
        matches = [e for e in repo.list_employees()
                  if _birth_month(e["date_of_birth"]) == now.month]
        if not matches:
            dialogs.info(self._root, "No birthdays",
                        "No employee has a birthday this month.")
            return

        # Nhớ giờ vừa chọn để lần sau mở tool khỏi phải chỉnh lại.
        send_time = _parse_time(self.time_field.get())
        settings.update(birthday_send_time=send_time.strftime("%H:%M"))

        folder = settings.get("birthday_images_folder")
        images = _scan_images(folder)

        rows = []
        for emp in matches:
            code_norm = (emp["code"] or "").strip().upper()
            image_name = images.get(code_norm)
            send_at = _birthday_datetime(emp["date_of_birth"], send_time, now.year)
            rows.append({
                "code": emp["code"],
                "name": emp["name"] or emp["full_name"] or emp["code"],
                "full_name": _title_case_name(emp["full_name"] or emp["code"]),
                # Ưu tiên company email, chỉ dùng personal email khi công ty
                # chưa có (vd nhân viên mới chưa cấp mail công ty).
                "email": (emp["company_email"] or emp["email"] or "").strip(),
                "date_of_birth": emp["date_of_birth"],
                "image_path": os.path.join(folder, image_name) if image_name else None,
                "send_at": send_at,
                # Sinh nhật đã qua (hoặc đúng hôm nay nhưng lỡ giờ) thì không
                # hẹn được nữa -> gửi ngay.
                "scheduled": bool(send_at and send_at > now),
            })

        dlg = _BirthdayConfirmDialog(self._root, rows, folder)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        selected = dlg.selected_rows()
        if not selected:
            dialogs.info(self._root, "Nothing selected",
                        "No employee is ticked — nothing to send.")
            return
        self._send_all(selected)

    def _send_all(self, rows):
        account = settings.get("birthday_from_account").strip()
        subject_tpl = self.subject_field.get().strip() or _DEFAULT_SUBJECT

        queued, sent_now, not_sent = [], [], []
        for row in rows:
            display_name = row["full_name"]
            if not row["image_path"]:
                not_sent.append(f"{display_name} — no matching card")
                continue
            if not row["email"]:
                not_sent.append(f"{display_name} — missing email")
                continue
            try:
                outlook.send_mail(
                    row["email"], _fill(subject_tpl, row["name"]), "",
                    account_smtp=account or None, attachments=[row["image_path"]],
                    inline_attachment=True,
                    deferred_until=row["send_at"] if row["scheduled"] else None)
                if row["scheduled"]:
                    queued.append(f"{display_name} — {_fmt_when(row['send_at'])}")
                else:
                    sent_now.append(display_name)
            except Exception as exc:
                not_sent.append(f"{display_name} — send failed: {exc}")

        parts = [f"Handed to Outlook: {len(queued) + len(sent_now)}/{len(rows)}"]
        if queued:
            parts.append(f"Waiting in the Outbox until each birthday ({len(queued)}):\n"
                         + "\n".join(queued))
        if sent_now:
            parts.append("Sent immediately — birthday already passed this month "
                         f"({len(sent_now)}):\n" + "\n".join(sent_now))
        if not_sent:
            parts.append("Not sent:\n" + "\n".join(not_sent))
        msg = "\n\n".join(parts)

        if not_sent:
            dialogs.warning(self._root, "Done with skipped/failed", msg)
        else:
            dialogs.success(self._root, "Done",
                            msg + "\n\nKeep Outlook running so queued emails "
                                  "can leave the Outbox on the day.")

    # -------------------------------------------------------- xuất CSV cho Canva
    def _export_missing_csv(self):
        """Xuất CSV TOÀN BỘ nhân viên (không lọc theo tháng) CHƯA có card, để
        nạp vào Canva Bulk Create tạo 1 lần cho hết, khỏi phải làm lắt nhắt
        từng tháng. Cột `name` bỏ dấu tiếng Việt cho khớp mẫu thiệp. Cột
        date_of_birth ghi ngày/tháng SINH kèm NĂM HIỆN TẠI (không phải năm
        sinh thật) vì đây là ngày hiển thị trên thiệp, không phải để lộ tuổi.
        Cột name viết hoa chữ đầu (không phải IN HOA hết) và kèm dấu phẩy cuối
        để dán thẳng vào khung chữ chào trên thiệp, vd 'Anh,'."""
        today = datetime.date.today()
        all_employees = [e for e in repo.list_employees()
                         if _day_month(e["date_of_birth"])]
        if not all_employees:
            dialogs.info(self._root, "No employees",
                        "No employee has a usable date of birth.")
            return

        folder = settings.get("birthday_images_folder")
        images = _scan_images(folder)

        missing = []
        for emp in all_employees:
            code_norm = (emp["code"] or "").strip().upper()
            if code_norm in images:
                continue   # đã có card rồi, không cần xuất lại
            dm = _day_month(emp["date_of_birth"])
            raw_name = emp["name"] or emp["full_name"] or emp["code"] or ""
            missing.append({
                "code": emp["code"] or "",
                "name": _title_case_name(_strip_vn_accents(raw_name)) + ",",
                "date_of_birth": f"{dm[0]}/{dm[1]}/{today.year}",
            })

        if not missing:
            dialogs.info(self._root, "Nothing to export",
                        "Every employee already has a card.")
            return

        path, _ = QFileDialog.getSaveFileName(
            self._root, "Export employees without a card",
            "birthday_missing_cards.csv", "CSV (*.csv)", "",
            QFileDialog.Option.DontConfirmOverwrite)
        if not path:
            return
        if not path.lower().endswith(".csv"):
            path += ".csv"

        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=["code", "name", "date_of_birth"])
                writer.writeheader()
                writer.writerows(missing)
        except OSError as exc:
            dialogs.error(self._root, "Can't write file", str(exc))
            return

        dialogs.success(
            self._root, "Exported",
            f"Exported {len(missing)} employee(s) without a card to:\n{path}\n\n"
            "In Canva Bulk Create: connect this CSV, replace the name/date text "
            "boxes with the \"name\"/\"date_of_birth\" columns, then under Advanced "
            "settings set \"Name each page using\" → code, so the downloaded card "
            "filenames match employee codes.")


class _BirthdayConfirmDialog(ModalDialog):
    """Modal xác nhận trước khi gửi: liệt kê người có sinh nhật tháng này, đánh
    dấu ai THIẾU ảnh thiệp (người đó sẽ không được gửi mail)."""

    def __init__(self, parent, rows, folder):
        super().__init__(parent, "md")
        card, lay = self.build_shell(f"Birthdays this month · {len(rows)}")
        self._checks = []   # [(row, QCheckBox)] — bỏ tick để loại người đó (vd chỉ gửi test cho mình)

        if not folder:
            widgets.hint(
                card, "⚠ No cards folder configured — set it in Settings → "
                     "Birthday email. Nobody will be emailed.")

        has_any_card = any(r["image_path"] for r in rows)
        self.select_all_cb = QCheckBox("Select all", card)
        self.select_all_cb.setChecked(has_any_card)
        self.select_all_cb.setEnabled(has_any_card)
        self.select_all_cb.stateChanged.connect(self._toggle_all)
        lay.addWidget(self.select_all_cb)

        body = QWidget()
        col = QVBoxLayout(body)
        col.setContentsMargins(0, 0, 8, 0)
        col.setSpacing(8)
        if not rows:
            empty = QLabel("No employee has a birthday this month.")
            empty.setObjectName("DialogMsg")
            col.addWidget(empty)
        for row in rows:
            col.addWidget(self._row_card(body, row))
        col.addStretch(1)
        sa = widgets.scroll_area(body)
        lay.addWidget(sa, 1)
        self.set_grow_region(sa)

        widgets.hint(
            card, "Emails go to Outlook's Outbox and are delivered on each "
                  "employee's own birthday at the delivery time — Outlook must "
                  "stay running and signed in until then.")

        missing = sum(1 for r in rows if not r["image_path"])
        if missing:
            widgets.hint(
                card, f"⚠ {missing} employee(s) have no matching card and will "
                     "NOT be emailed.")
        instant = sum(1 for r in rows if r["image_path"] and not r["scheduled"])
        if instant:
            widgets.hint(
                card, f"⚠ {instant} birthday(s) already passed this month — those "
                     "emails cannot be scheduled and will be sent right away.")

        foot = QHBoxLayout()
        foot.addWidget(widgets.button(card, "Send", variant="primary", icon="mail",
                                      command=self.accept))
        foot.addWidget(widgets.button(card, "Cancel", variant="neutral", icon="x",
                                      command=self.reject))
        foot.addStretch(1)
        lay.addLayout(foot)

    def selected_rows(self):
        """Danh sách row còn được tick — bỏ tick thì loại khỏi đợt gửi."""
        return [row for row, cb in self._checks if cb.isChecked()]

    def _toggle_all(self, _state=None):
        """Tick/bỏ tick "Select all" -> áp cho mọi dòng CÓ card (danh sách dài
        thì tick/bỏ hết 1 phát nhanh hơn tự bỏ từng người)."""
        checked = self.select_all_cb.isChecked()
        for row, cb in self._checks:
            if cb.isEnabled():
                cb.setChecked(checked)

    def _row_card(self, parent, row):
        box = QFrame(parent)
        box.setObjectName("DetailCard")
        h = QHBoxLayout(box)
        h.setContentsMargins(14, 10, 14, 10)
        h.setSpacing(8)

        has_card = bool(row["image_path"])
        cb = QCheckBox(box)
        cb.setChecked(has_card)      # thiếu card thì không có gì để gửi -> mặc định bỏ tick
        cb.setEnabled(has_card)
        h.addWidget(cb)
        self._checks.append((row, cb))

        name_col = QVBoxLayout()
        name_col.setSpacing(2)
        name = QLabel(f"{row['full_name']}  ({row['code']})", box)
        name.setObjectName("DetailNamePlain")
        name_col.addWidget(name)
        email = QLabel(row["email"] or "No email on file", box)
        email.setObjectName("DetailMeta")
        name_col.addWidget(email)
        h.addLayout(name_col, 1)
        dob = QLabel(row["date_of_birth"] or "—", box)
        dob.setObjectName("Hint")
        h.addWidget(dob)
        if not has_card:
            h.addWidget(_chip(box, "No card", theme.PALETTE["--danger"]))
        elif row["scheduled"]:
            h.addWidget(_chip(box, _fmt_when(row["send_at"]),
                              theme.PALETTE["--success"]))
        else:
            h.addWidget(_chip(box, "Sends now", theme.PALETTE["--warning"]))
        return box
