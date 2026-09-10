"""Gửi mail chúc mừng sinh nhật — đính kèm ảnh thiệp xuất từ Canva Bulk Create.

Bấm "Review birthdays": tự quét bảng nhân viên tìm người có NGÀY SINH TRONG
THÁNG HIỆN TẠI, đối chiếu mã NV với thư mục ảnh đã cấu hình ở Cài đặt (tên file
= mã NV) → hiện modal xác nhận (ai THIẾU ảnh/mail bị đánh dấu, không xếp hàng
được) → bấm "Enqueue" thì danh sách được XẾP HÀNG vào bảng `birthday_emails`.

Mail KHÔNG đi lúc bấm Enqueue, và cũng không còn nhờ Outlook hẹn giờ nữa: MỖI
LẦN MỞ APP, `startup()` hiện **modal xác nhận** liệt kê những người tới hạn hôm
nay; bấm *Send now* mới gửi. Lý do bỏ cách cũ (`DeferredDeliveryTime`, mail nằm
Outbox tới đúng ngày): nó chỉ chạy khi tài khoản gửi ĐÃ ĐĂNG NHẬP trong Outlook,
còn hộp thư dùng chung (chỉ được IT share) thì Outlook bỏ qua giờ hẹn và gửi
ngay — cả tháng nhận mail cùng một lúc.

App KHÔNG BAO GIỜ tự gửi mail mà không hỏi. Modal đó chỉ hiện **một lần mỗi
ngày** (mở app 5 lần không bị hỏi 5 lần); bấm Cancel thì mọi dòng nằm nguyên
trong hàng chờ và gửi tay được bằng nút *Send due emails* ở màn hình Queue.

Toàn bộ logic gửi/xếp hàng nằm ở `app/core/birthday_mail.py` (thuần Python):
`startup()` phải chạy được khi người dùng chưa từng mở trang này.
"""
import csv
import datetime
import unicodedata

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QDialog, QFileDialog, QFrame, QHBoxLayout, QLabel,
    QVBoxLayout, QWidget,
)

from app.core import birthday_mail, config
from app.core import cv_repository as repo
from app.core import cv_schema, debuglog, outlook
from app_qt import dialogs, theme, widgets
from app_qt.base_tool import BaseTool
from app_qt.components.modal import ModalDialog
from app_qt.components.table import DataTable
from app_qt.components.task import Task

SECTION = "birthday_email"

# `last_prompt`: ngày gần nhất modal xác nhận lúc mở app đã hiện ra. Chỉ hỏi một
# lần mỗi ngày — mở app 5 lần không bị hỏi 5 lần. Đây là cấu hình RIÊNG của tool
# nên ở `config.py`, không phải `settings.py` (thiết lập chung có ô nhập ở
# màn hình Cài đặt).
_CONFIG_DEFAULTS = {"last_prompt": ""}

# Đợi hộp thoại của tool khác (Gate-Open Mail) đóng trước khi hỏi — 120 × 700ms
# ≈ 84 giây, quá đó thì bỏ lượt hỏi hôm nay.
_PROMPT_RETRY_MS = 700
_PROMPT_MAX_TRIES = 120


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


def _fmt_day(d):
    """datetime.date -> '23 Sep'."""
    return d.strftime("%d %b") if d else "—"


def _fmt_iso_day(value):
    """'yyyy-mm-dd' -> '23 Sep' cho cột Birthday của bảng hàng chờ."""
    try:
        return datetime.date.fromisoformat(str(value)).strftime("%d %b")
    except (TypeError, ValueError):
        return str(value or "")


def _chip(parent, text, color):
    lbl = QLabel(text, parent)
    r, g, b = widgets._hex_to_rgb(color)
    lbl.setStyleSheet(
        f"background: rgba({r},{g},{b},0.15); color:{color}; border-radius:10px;"
        " padding:3px 10px; font-size:12px; font-weight:600;")
    return lbl


def _person_row(parent, cb, name_text, meta_text, right_text, chip_label, chip_color):
    """Một dòng người trong modal: [tick] Tên (mã) / email … ngày [chip].

    Dùng chung cho cả modal xếp hàng và modal xác nhận gửi — hai modal khác nhau
    ở dữ liệu, còn khung dòng thì phải nhìn giống nhau.
    """
    box = QFrame(parent)
    box.setObjectName("DetailCard")
    h = QHBoxLayout(box)
    h.setContentsMargins(14, 10, 14, 10)
    h.setSpacing(8)
    h.addWidget(cb)

    col = QVBoxLayout()
    col.setSpacing(2)
    name = QLabel(name_text, box)
    name.setObjectName("DetailNamePlain")
    col.addWidget(name)
    meta = QLabel(meta_text, box)
    meta.setObjectName("DetailMeta")
    col.addWidget(meta)
    h.addLayout(col, 1)

    right = QLabel(right_text, box)
    right.setObjectName("Hint")
    h.addWidget(right)
    h.addWidget(_chip(box, chip_label, chip_color))
    return box


def _row_badge(row):
    """(nhãn chip, màu, chọn được?, tick sẵn?) cho một dòng trong modal.

    Thiếu thiệp/mail thì không có gì để gửi; đã gửi hoặc đang chờ thì xếp hàng
    lại là vô nghĩa → cả hai nhóm đều KHÔNG chọn được. Dòng từng lỗi/quá hạn thì
    chọn được và tick sẵn: đó chính là đường gửi lại.
    """
    status = row["queued_status"]
    if not row["card_path"]:
        return "No card", theme.PALETTE["--danger"], False, False
    if not row["email"]:
        return "No email", theme.PALETTE["--danger"], False, False
    if status == cv_schema.BIRTHDAY_EMAIL_SENT:
        return "Sent", theme.PALETTE["--success"], False, False
    if status in (cv_schema.BIRTHDAY_EMAIL_PENDING, cv_schema.BIRTHDAY_EMAIL_SENDING):
        return (f"Queued · {_fmt_day(row['due_date'])}",
                theme.TEXT_MUTED, False, False)
    if status:
        # Failed / Missed / Cancelled — xếp hàng lại được.
        return f"{status} · retry", theme.PALETTE["--warning"], True, True
    if row["passed"]:
        return "Passed · sends now", theme.PALETTE["--warning"], True, False
    return _fmt_day(row["due_date"]), theme.PALETTE["--success"], True, True


_QUEUE_COLUMNS = [
    ("employee", "Employee", 230, "left"),
    ("email", "Email", 230, "left"),
    ("due_date", "Birthday", 100, "center", _fmt_iso_day),
    ("status", "Status", 100, "center"),
    ("sent_at", "Sent at", 145, "center"),
    ("last_error", "Error", 260, "left"),
]


class BirthdayEmailTool(BaseTool):
    name = "Birthday emails"
    description = "Send birthday cards (Canva-exported images) to employees by email."
    icon = "🎂"
    category = "Office"
    order = 7
    fills_height = True
    # Mỗi lần mở app: gửi những mail đã tới ngày (xem startup()).
    auto_startup = True

    # Dựng thẳng thẻ full-height (giống EmployeeDbTool) thay cho khung mặc định.
    def build(self, parent=None):
        repo.init_db()
        card = widgets.Card(parent)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(22, 10, 22, 18)
        lay.setSpacing(10)
        self._root = card

        # Trang này KHÔNG có ô nhập nào: tiêu đề mail, thư mục thiệp, tài khoản
        # gửi và hạn gửi bù đều ở ⚙️ Settings → Birthday email. Để tiêu đề ở cả
        # hai chỗ thì vừa trùng, vừa ăn chiều cao của bảng hàng chờ — mà bảng đó
        # mới là thứ cần nhìn hằng ngày. Vì vậy cũng không còn section label nào
        # cho phần trên (không có field group để mà đặt tên).
        widgets.hint(
            card, "Queued emails live in this app, not in Outlook. Each one is "
                  "sent on the employee's own birthday, the first time you open "
                  "Personal Toolbox that day — after asking you to confirm. Miss "
                  f"that day and it still goes out within "
                  f"{birthday_mail.catchup_days()} day(s). Subject and cards "
                  "folder: Settings → Birthday email.")

        send_bar = QHBoxLayout()
        send_bar.setContentsMargins(0, 8, 0, 0)
        send_bar.addWidget(widgets.button(card, "Review birthdays", variant="primary",
                                          icon="search", command=self._on_review))
        send_bar.addWidget(widgets.button(
            card, "Export CSV (missing cards)", variant="neutral",
            icon="file-text", command=self._export_missing_csv))
        send_bar.addStretch(1)
        lay.addLayout(send_bar)

        # Nhắc khi tháng này chưa xếp hàng: hàng chờ chỉ có dòng cho tháng nào
        # người dùng đã bấm Enqueue, quên bấm là cả tháng không ai được gửi.
        self.notice_lbl = QLabel("", card)
        self.notice_lbl.setObjectName("Hint")
        self.notice_lbl.setWordWrap(True)
        lay.addWidget(self.notice_lbl)

        widgets.section_label(card, "Queue")
        self.table = DataTable(_QUEUE_COLUMNS, pk="birthday_email_id",
                               stretch_key="last_error", checkable=True,
                               menu_actions=[
                                   ("Send now", self._send_now),
                                   None,
                                   ("Remove from queue", self._remove_from_queue),
                               ])
        self._build_queue_bar(lay)
        lay.addWidget(self.table, 1)

        self.count_lbl = QLabel("", card)
        self.count_lbl.setObjectName("Hint")
        lay.addWidget(self.count_lbl)

        self._reload_queue()
        return card

    def build_body(self, parent):
        pass

    # ------------------------------------------------------------ hàng chờ
    def _build_queue_bar(self, lay):
        row = QHBoxLayout()
        row.setSpacing(10)
        self.sel_status = widgets.FilterSelect("Status")
        self.sel_status.set_options(cv_schema.BIRTHDAY_EMAIL_STATUS_CHOICES)
        self.sel_status.changed.connect(self._reload_queue)
        self.sel_status.setFixedWidth(180)
        row.addWidget(self.sel_status)
        row.addStretch(1)
        # Trigger TAY cho đúng lượt mà app hỏi lúc mở máy: bấm Cancel ở modal đó
        # thì cả ngày không bị hỏi lại, đây là đường vào lại.
        row.addWidget(widgets.button(
            self._root, "Send due emails", variant="primary", icon="mail",
            command=lambda: self._run_due_flow(self._root)))
        row.addWidget(widgets.button(self._root, "Reload", variant="neutral",
                                     icon="refresh", command=self._reload_queue))
        lay.addLayout(row)

    def _reload_queue(self):
        """Nạp lại bảng + hai dòng chữ phía trên/dưới. Gọi được cả sau lượt gửi
        tự động (chạy khi mở app), nên phải chịu được việc TRANG CHƯA DỰNG —
        trang tool chỉ dựng khi người dùng bấm vào sidebar. `count_lbl` là widget
        dựng SAU CÙNG trong build() nên có nó là chắc chắn có đủ phần còn lại.
        """
        if getattr(self, "count_lbl", None) is None:
            return
        year = datetime.date.today().year
        rows = repo.list_birthday_emails(status=self.sel_status.value(), year=year)
        self.table.set_rows([{
            "birthday_email_id": r["birthday_email_id"],
            "employee": f"{_title_case_name(r['full_name'] or '')} "
                        f"({r['employee_code'] or '—'})".strip(),
            "email": ((r["company_email"] or "") or (r["email"] or "")).strip(),
            "due_date": r["due_date"],
            "status": r["status"],
            "sent_at": r["sent_at"] or "",
            "last_error": r["last_error"] or "",
        } for r in rows])

        counts = {s: repo.count_birthday_emails(status=s, year=year)
                  for s in cv_schema.BIRTHDAY_EMAIL_STATUS_CHOICES}
        shown = " · ".join(f"{n} {s.lower()}" for s, n in counts.items() if n)
        self.count_lbl.setText(f"{year}: {shown}" if shown
                               else f"{year}: nothing queued yet")

        queued = birthday_mail.month_is_queued()
        self.notice_lbl.setText(
            "" if queued else
            "⚠ This month's birthdays are not queued yet — click "
            "\"Review birthdays\" so the emails can go out on the day.")
        self.notice_lbl.setVisible(not queued)

    def _send_now(self, rows):
        """Gửi ngay, bất kể còn bao lâu tới sinh nhật (menu chuột phải)."""
        if not outlook.available():
            dialogs.warning(self._root, "Outlook required",
                            "Sending email needs Outlook on Windows (pywin32).")
            return
        targets = [r for r in rows
                   if r["status"] not in (cv_schema.BIRTHDAY_EMAIL_SENT,
                                          cv_schema.BIRTHDAY_EMAIL_CANCELLED)]
        if not targets:
            dialogs.info(self._root, "Nothing to send",
                         "The selected row(s) are already sent or cancelled.")
            return
        if not dialogs.confirm(
                self._root, "Send now",
                f"Send {len(targets)} birthday email(s) right now, without waiting "
                "for the birthday?", ok_label="Send now"):
            return
        self._send_async(self._root, [r["birthday_email_id"] for r in targets])

    def _remove_from_queue(self, rows):
        if not dialogs.confirm(
                self._root, "Remove from queue",
                f"Remove {len(rows)} row(s) from the queue? The employees will "
                "not be emailed unless you queue them again.",
                ok_label="Remove"):
            return
        for row in rows:
            repo.delete_birthday_email(row["birthday_email_id"])
        self._reload_queue()

    # -------------------------------------------------- duyệt + xếp hàng
    def _on_review(self):
        rows = birthday_mail.month_candidates()
        if not rows:
            dialogs.info(self._root, "No birthdays",
                         "No employee has a birthday this month.")
            return

        dlg = _BirthdayConfirmDialog(self._root, rows,
                                     birthday_mail.cards_folder())
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        selected = dlg.selected_rows()
        if not selected:
            dialogs.info(self._root, "Nothing selected",
                         "No employee is ticked — nothing to queue.")
            return

        result = birthday_mail.enqueue(selected)
        self._reload_queue()

        parts = [f"Queued: {len(result['queued'])}"]
        if result["requeued"]:
            parts.append(f"Queued again after a previous failure "
                         f"({len(result['requeued'])}):\n"
                         + "\n".join(_title_case_name(n) for n in result["requeued"]))
        if result["skipped"]:
            parts.append(f"Skipped — already queued or sent ({len(result['skipped'])}):\n"
                         + "\n".join(_title_case_name(n) for n in result["skipped"]))
        dialogs.success(
            self._root, "Queued", "\n\n".join(parts)
            + "\n\nEach email goes out on the employee's own birthday, the first "
              "time you open Personal Toolbox that day.")

    # ------------------------------------------ hỏi & gửi khi mở app
    def startup(self, window):
        """Hỏi rồi gửi những mail đã tới ngày. Bốn cái bẫy ở đây, đừng gỡ:

        1. `startup()` chạy TRƯỚC `build()` — trang tool dựng lười lúc bấm
           sidebar, nên `self._root`/`self.table` CHƯA tồn tại. Chỉ được đọc
           DB/settings, lấy `window` làm cha hộp thoại.
        2. Không làm việc thẳng trong hàm này mà đẩy qua `QTimer.singleShot`:
           `MainWindow._run_startup_tasks` gọi lần lượt từng tool, và Gate-Open
           Mail (order 5, chạy TRƯỚC tool này) kết thúc bằng `dlg.exec()` — một
           event loop lồng. Làm việc thẳng ở đây thì hộp thoại đó còn mở là cả
           ngày không mail nào đi.
        3. `MainWindow._run_startup_tasks` nuốt mọi exception bằng `pass`, nên
           lỗi ở đây là mail âm thầm không gửi mà không có dấu vết → phải tự
           bọc try/except + debuglog.
        4. `repo.init_db()` gọi ở LUỒNG GIAO DIỆN (không chạy migration trong
           worker): tool này trước giờ chưa từng gọi, ăn theo tool khác.
        """
        if not outlook.available():
            return
        QTimer.singleShot(0, lambda: self._prompt_due_on_startup(window))

    def _prompt_due_on_startup(self, window, tries=0):
        """Lượt hỏi-rồi-gửi lúc mở app. CHỈ HỎI MỘT LẦN MỖI NGÀY.

        Bấm Cancel là cả ngày không hỏi lại (mở app 5 lần không bị hỏi 5 lần);
        đổi ý thì bấm "Send due emails" ở màn hình Queue. Dòng vẫn nằm nguyên
        trong hàng chờ nên hôm sau — còn trong hạn gửi bù — lại được hỏi.

        Khác Gate-Open Mail ở một điểm quan trọng: chỗ đó đóng dấu `last_scan`
        TRƯỚC khi làm việc nên Outlook lỗi là mất luôn ngày đó. Đây chỉ đóng dấu
        khi modal ĐÃ THỰC SỰ hiện ra.
        """
        # Gate-Open Mail (order 5, chạy trước tool này) có thể đang mở modal của
        # nó bằng `dlg.exec()`. Bật thêm modal đè lên thì rối, nên đợi trống đã.
        # Đợi mãi không được thì thôi, không đóng dấu ngày -> lần mở app sau vẫn
        # hỏi, và nút "Send due emails" ở màn hình Queue luôn là đường vào lại.
        if QApplication.activeModalWidget() is not None:
            if tries >= _PROMPT_MAX_TRIES:
                debuglog.write("birthday_email: another dialog stayed open, "
                               "skipped today's send prompt")
                return
            QTimer.singleShot(
                _PROMPT_RETRY_MS,
                lambda: self._prompt_due_on_startup(window, tries + 1))
            return
        try:
            repo.init_db()
        except Exception as exc:
            debuglog.write(f"birthday_email.startup: init_db failed: {exc}")
            return
        cfg = config.load(SECTION, _CONFIG_DEFAULTS)
        if cfg.get("last_prompt") == datetime.date.today().isoformat():
            return
        self._run_due_flow(window, silent_if_empty=True, stamp_prompt=True)

    def _run_due_flow(self, parent, silent_if_empty=False, stamp_prompt=False):
        """Chuẩn bị → hiện modal duyệt → gửi. Dùng chung cho lượt mở app và nút
        "Send due emails" trên màn hình Queue."""
        try:
            result = birthday_mail.prepare_due()
        except Exception as exc:
            debuglog.write(f"birthday_email.prepare_due failed: {exc}")
            if not silent_if_empty:
                dialogs.error(parent, "Birthday emails",
                              f"Couldn't read the queue:\n{exc}")
            return
        self._reload_queue()

        if result["skipped"]:
            if not silent_if_empty:
                dialogs.warning(parent, "Birthday emails",
                                f"Nothing to send — {result['skipped']}.")
            return
        if not result["rows"]:
            if not silent_if_empty:
                msg = "No birthday email is due today."
                if result["missed"]:
                    msg += (f"\n\n{result['missed']} email(s) passed the "
                            "catch-up window and are marked Missed.")
                dialogs.info(parent, "Nothing due", msg)
            return

        # Kéo cửa sổ lên trước rồi mới hỏi — lượt mở app có thể đang bị cửa sổ
        # khác che.
        if stamp_prompt:
            self._raise_window(parent)

        dlg = _SendDueDialog(parent, result["rows"], result["missed"])
        if stamp_prompt:
            cfg = config.load(SECTION, _CONFIG_DEFAULTS)
            cfg["last_prompt"] = datetime.date.today().isoformat()
            config.save(SECTION, cfg)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            debuglog.write("birthday_email: user cancelled the send prompt "
                           f"({len(result['rows'])} due)")
            return
        row_ids = dlg.selected_ids()
        if not row_ids:
            dialogs.info(parent, "Nothing selected",
                         "No employee is ticked — nothing was sent.")
            return
        self._send_async(parent, row_ids)

    @staticmethod
    def _raise_window(window):
        try:
            window.showNormal()
            window.raise_()
            window.activateWindow()
        except Exception:
            pass

    def _send_async(self, parent, row_ids):
        """Gửi ở luồng nền: Outlook khởi động lạnh có thể mất chục giây, không
        được để đơ cửa sổ."""
        def work(_emit):
            return birthday_mail.send_rows(row_ids)

        # `signals` được tạo ở luồng GUI nên Qt xếp hàng các slot dưới đây về
        # đúng luồng đó (đã kiểm chứng: slot chạy ở luồng chính, không phải
        # luồng nền) — nhờ vậy `_on_sent` chạm widget được. Đừng đổi sang gọi
        # thẳng hàm trong `work()`.
        self._task = Task(work, parent)
        self._task.signals.finished.connect(
            lambda result: self._on_sent(parent, result))
        self._task.signals.failed.connect(
            lambda msg: self._on_send_failed(parent, msg))
        # QThread trần: không chờ lúc thoát app thì đóng cửa sổ giữa lúc Outlook
        # đang khởi động lạnh sẽ ra "QThread: Destroyed while thread is running".
        app = QApplication.instance()
        if app is not None:
            app.aboutToQuit.connect(self._wait_for_task)
        self._task.start()

    def _wait_for_task(self):
        task = getattr(self, "_task", None)
        if task is not None and task.isRunning():
            task.wait(5000)

    def _on_sent(self, parent, result):
        self._reload_queue()
        self._report(parent, result, "Birthday emails")

    def _on_send_failed(self, parent, msg):
        debuglog.write(f"birthday_email.send_rows failed: {msg}")
        self._reload_queue()
        dialogs.error(parent, "Birthday emails", f"Sending failed:\n{msg}")

    @staticmethod
    def _report(parent, result, title):
        """Hộp thoại tổng kết một lượt gửi (dùng chung cho lượt tự động và
        "Send now")."""
        if result["skipped"]:
            dialogs.warning(parent, title,
                            f"Nothing was sent — {result['skipped']}.")
            return
        parts = [f"Sent: {len(result['sent'])}"]
        if result["sent"]:
            parts.append("\n".join(result["sent"]))
        if result["failed"]:
            parts.append(f"Not sent ({len(result['failed'])}):\n"
                         + "\n".join(result["failed"]))
        if result.get("missed"):
            parts.append(f"{result['missed']} email(s) passed the catch-up "
                         "window and are marked Missed — they will not be sent "
                         "automatically.")
        msg = "\n\n".join(parts)
        if result["failed"]:
            dialogs.warning(parent, title, msg)
        else:
            dialogs.success(parent, title, msg)

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
                         if birthday_mail.birth_month(e["date_of_birth"])]
        if not all_employees:
            dialogs.info(self._root, "No employees",
                         "No employee has a usable date of birth.")
            return

        images = birthday_mail.scan_cards(birthday_mail.cards_folder())

        missing = []
        for emp in all_employees:
            code_norm = (emp["code"] or "").strip().upper()
            if code_norm in images:
                continue   # đã có card rồi, không cần xuất lại
            dm = birthday_mail.day_month(emp["date_of_birth"])
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
    """Modal duyệt trước khi XẾP HÀNG: liệt kê người có sinh nhật tháng này, đánh
    dấu ai không xếp hàng được (thiếu thiệp/mail, hoặc đã xếp hàng/đã gửi rồi)."""

    def __init__(self, parent, rows, folder):
        super().__init__(parent, "md")
        card, lay = self.build_shell(f"Birthdays this month · {len(rows)}")
        self._checks = []   # [(row, QCheckBox)] — bỏ tick để loại người đó

        if not folder:
            widgets.hint(
                card, "⚠ No cards folder configured — set it in Settings → "
                      "Birthday email. Nobody can be queued.")

        # Tính nhãn/màu/chọn-được một lần rồi đi kèm từng row: cả phần đếm ở
        # cuối hộp thoại lẫn từng dòng đều đọc chung một kết quả.
        badged = [(row, _row_badge(row)) for row in rows]
        has_selectable = any(b[2] for _r, b in badged)
        self.select_all_cb = QCheckBox("Select all", card)
        self.select_all_cb.setChecked(
            has_selectable and all(b[3] for _r, b in badged if b[2]))
        self.select_all_cb.setEnabled(has_selectable)
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
        for row, badge in badged:
            col.addWidget(self._row_card(body, row, badge))
        col.addStretch(1)
        sa = widgets.scroll_area(body)
        lay.addWidget(sa, 1)
        self.set_grow_region(sa)

        widgets.hint(
            card, "Ticked employees are queued in this app. Each email is sent "
                  "on that employee's own birthday, the first time you open "
                  "Personal Toolbox that day — nothing is sent right now.")

        blocked = sum(1 for r, b in badged if not b[2] and not r["queued_status"])
        if blocked:
            widgets.hint(
                card, f"⚠ {blocked} employee(s) have no matching card or no email "
                      "address and cannot be queued.")
        passed = sum(1 for r, b in badged if r["passed"] and b[2])
        if passed:
            widgets.hint(
                card, f"⚠ {passed} birthday(s) already passed this month — tick "
                      "them only if you still want the email to go out on the "
                      "next app launch.")

        foot = QHBoxLayout()
        foot.addWidget(widgets.button(card, "Enqueue", variant="primary", icon="mail",
                                      command=self.accept))
        foot.addWidget(widgets.button(card, "Cancel", variant="neutral", icon="x",
                                      command=self.reject))
        foot.addStretch(1)
        lay.addLayout(foot)

    def selected_rows(self):
        """Danh sách row còn được tick — bỏ tick thì loại khỏi đợt xếp hàng."""
        return [row for row, cb in self._checks if cb.isChecked()]

    def _toggle_all(self, _state=None):
        """Tick/bỏ tick "Select all" -> áp cho mọi dòng XẾP HÀNG ĐƯỢC (danh sách
        dài thì tick/bỏ hết 1 phát nhanh hơn tự bỏ từng người)."""
        checked = self.select_all_cb.isChecked()
        for _row, cb in self._checks:
            if cb.isEnabled():
                cb.setChecked(checked)

    def _row_card(self, parent, row, badge):
        label, color, selectable, checked = badge
        cb = QCheckBox(parent)
        cb.setChecked(selectable and checked)
        cb.setEnabled(selectable)
        self._checks.append((row, cb))
        return _person_row(
            parent, cb,
            f"{_title_case_name(row['raw_full_name'])}  ({row['code']})",
            row["email"] or "No email on file",
            row["date_of_birth"] or "—", label, color)


class _SendDueDialog(ModalDialog):
    """Modal xác nhận NGAY TRƯỚC KHI GỬI — app không tự gửi mail mà không hỏi.

    Hiện lúc mở app (một lần mỗi ngày) và khi bấm "Send due emails" ở màn hình
    Queue. Bấm Cancel thì mọi dòng nằm nguyên trong hàng chờ, gửi tay sau được.
    """

    def __init__(self, parent, rows, missed=0):
        super().__init__(parent, "md")
        sendable = [r for r in rows if not r["blocked"]]
        card, lay = self.build_shell(f"Birthday emails to send · {len(sendable)}")
        self._checks = []   # [(row_id, QCheckBox)]

        self.select_all_cb = QCheckBox("Select all", card)
        self.select_all_cb.setChecked(bool(sendable))
        self.select_all_cb.setEnabled(bool(sendable))
        self.select_all_cb.stateChanged.connect(self._toggle_all)
        lay.addWidget(self.select_all_cb)

        body = QWidget()
        col = QVBoxLayout(body)
        col.setContentsMargins(0, 0, 8, 0)
        col.setSpacing(8)
        today_iso = datetime.date.today().isoformat()
        for row in rows:
            col.addWidget(self._row_card(body, row, today_iso))
        col.addStretch(1)
        sa = widgets.scroll_area(body)
        lay.addWidget(sa, 1)
        self.set_grow_region(sa)

        widgets.hint(
            card, "These are the queued birthday emails that are due. Sending "
                  "happens now, from this app — Cancel leaves everything in the "
                  "queue and you can send later with \"Send due emails\" on the "
                  "Queue list.")
        blocked = len(rows) - len(sendable)
        if blocked:
            widgets.hint(
                card, f"⚠ {blocked} queued email(s) cannot be sent — the reason "
                      "is shown on each row and saved in the queue.")
        if missed:
            widgets.hint(
                card, f"⚠ {missed} email(s) passed the catch-up window and are "
                      "marked Missed — they will not be sent automatically.")

        foot = QHBoxLayout()
        foot.addWidget(widgets.button(card, "Send now", variant="primary", icon="mail",
                                      command=self.accept))
        foot.addWidget(widgets.button(card, "Cancel", variant="neutral", icon="x",
                                      command=self.reject))
        foot.addStretch(1)
        lay.addLayout(foot)

    def selected_ids(self):
        return [row_id for row_id, cb in self._checks if cb.isChecked()]

    def _toggle_all(self, _state=None):
        checked = self.select_all_cb.isChecked()
        for _row_id, cb in self._checks:
            if cb.isEnabled():
                cb.setChecked(checked)

    def _row_card(self, parent, row, today_iso):
        blocked = row["blocked"]
        cb = QCheckBox(parent)
        cb.setChecked(not blocked)
        cb.setEnabled(not blocked)
        self._checks.append((row["row_id"], cb))

        when = _fmt_iso_day(row["due_date"])
        if blocked:
            label, color = blocked, theme.PALETTE["--danger"]
        elif row["due_date"] < today_iso:
            # Sinh nhật đã qua nhưng còn trong hạn gửi bù -> nói rõ là gửi muộn.
            label, color = f"Late · {when}", theme.PALETTE["--warning"]
        else:
            label, color = when, theme.PALETTE["--success"]
        return _person_row(
            parent, cb, f"{_title_case_name(row['display'])}  ({row['code']})",
            row["email"] or "No email on file", "", label, color)
