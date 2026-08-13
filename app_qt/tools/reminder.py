"""Nhắc phản hồi kết quả phỏng vấn — bản PySide6.

Quét lịch Outlook 1 tháng gần đây, liệt kê lịch phỏng vấn chưa phản hồi thành
bảng Yes/No: Yes = đã phản hồi (ẩn đi), No = soạn mail phản hồi hoặc đặt lịch
nhắc. Dùng lại app.core.outlook + helper của module Tk cũ; giao diện dựng Qt.
"""
import datetime

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame, QHBoxLayout, QLabel, QLineEdit, QVBoxLayout, QWidget,
)

from app.core import config, debuglog, outlook
from app.core.reminder_logic import (
    DEFAULTS, _extract_name, _fill_template, _month_ago,
)
from app_qt import dialogs, richtext, widgets
from app_qt.base_tool import BaseTool
from app_qt.components.dialog_base import build_dialog_shell

SECTION = "reminder"


def _clear_layout(lay):
    while lay.count():
        item = lay.takeAt(0)
        w = item.widget()
        if w is not None:
            w.deleteLater()
        elif item.layout() is not None:
            _clear_layout(item.layout())


class ReminderTool(BaseTool):
    name = "Interview Follow-up"
    description = "Scan the last month's interviews and remind you to send results to candidates."
    icon = "🔔"
    category = "Office"
    order = 6

    def build(self, parent=None):
        self._interviews = []
        cfg = config.load(SECTION, DEFAULTS)

        card = widgets.Card(parent)
        self._page = card
        lay = QVBoxLayout(card)
        lay.setContentsMargins(28, 24, 28, 24)
        lay.setSpacing(6)

        widgets.section_label(card, "Interview detection")
        self.var_keywords = widgets.text_row(card, "Title keywords (comma-separated)")
        self.var_keywords.set(cfg["keywords"])
        widgets.hint(card, "Events whose title contains any of these keywords "
                           "(case-insensitive) count as interviews.")
        self.var_exclude = widgets.text_row(card, "Skip emails on this domain (internal recipients)")
        self.var_exclude.set(cfg["exclude_domain"])

        widgets.section_label(card, "Candidate reply template")
        self.var_subject = widgets.text_row(card, "Subject")
        self.var_subject.set(cfg["subject"])
        lbl = QLabel("Body (supports {name}, {position}, {subject}, {date}, {time})")
        lbl.setObjectName("FieldLabel")
        lbl.setWordWrap(True)
        lay.addWidget(lbl)
        self.body_editor = richtext.RichText(card, height=11)
        self.body_editor.set_html(cfg["body"])
        lay.addWidget(self.body_editor)

        row = QHBoxLayout()
        row.addWidget(widgets.button(card, "Save config", variant="neutral",
                                     icon="save", command=self._save_config))
        row.addWidget(widgets.button(card, "Scan calendar (last month)", variant="primary",
                                     icon="refresh", command=self._scan_clicked))
        row.addStretch(1)
        lay.addLayout(row)

        if not outlook.available():
            widgets.hint(card, "⚠ Outlook not found (pywin32). Scanning/sending mail "
                               "only works on Windows with Outlook installed.")

        widgets.section_label(card, "Interviews found")
        self._table_holder = QWidget(card)
        self._table_layout = QVBoxLayout(self._table_holder)
        self._table_layout.setContentsMargins(0, 0, 0, 0)
        self._table_layout.setSpacing(6)
        lay.addWidget(self._table_holder)

        self._render_table()
        return card

    def build_body(self, parent):
        pass

    # -------------------------------------------------------------- config
    def _collect(self):
        cfg = config.load(SECTION, DEFAULTS)
        return {
            "keywords": self.var_keywords.get().strip(),
            "exclude_domain": self.var_exclude.get().strip(),
            "subject": self.var_subject.get().strip(),
            "body": self.body_editor.get_html(),
            "dismissed": cfg.get("dismissed", []),
        }

    def _save_config(self):
        config.save(SECTION, self._collect())
        self.info("Saved", "Config saved ✅")

    @staticmethod
    def _save_runtime(**changes):
        raw = dict(config._read_all().get(SECTION, {}))
        raw.update(changes)
        config.save(SECTION, raw)

    def _dismiss(self, appt):
        eid = appt.get("entry_id")
        dismissed = list(config.load(SECTION, DEFAULTS).get("dismissed", []))
        if eid and eid not in dismissed:
            dismissed.append(eid)
            self._save_runtime(dismissed=dismissed)
        self._interviews = [a for a in self._interviews if a is not appt]
        self._render_table()

    # ---------------------------------------------------------- quét lịch
    def _fetch_interviews(self):
        if not outlook.available():
            return []
        cfg = config.load(SECTION, DEFAULTS)
        kw_raw = self.var_keywords.get() if hasattr(self, "var_keywords") else cfg["keywords"]
        keywords = [k.strip().lower() for k in kw_raw.split(",") if k.strip()]
        today = datetime.date.today()
        appts = outlook.appointments_between(_month_ago(today), today)
        dismissed = set(cfg.get("dismissed", []))
        out = []
        for a in appts:
            if a.get("entry_id") in dismissed:
                continue
            subj = a["subject"].lower()
            if keywords and not any(kw in subj for kw in keywords):
                continue
            out.append(a)
        return out

    def _scan_clicked(self):
        if not outlook.available():
            self.error("Outlook required", "This feature needs Outlook on Windows (pywin32).")
            return
        # Hiện thẳng traceback lên dialog: quét lịch qua COM có thể lỗi vì rất
        # nhiều lý do bên ngoài (Outlook đang khởi động, profile khóa, quyền
        # truy cập...) mà app không có console để in ra.
        try:
            self._interviews = self._fetch_interviews()
        except Exception as exc:
            self.error("Calendar scan failed",
                       f"Could not read the Outlook calendar.\n\n{exc}\n\n———\n"
                       f"{debuglog.exception('reminder: scan calendar')}")
            return
        self._render_table()
        if not self._interviews:
            self.info("No events", "No interviews found in the last month.")

    # ----------------------------------------------------------- bảng
    def _render_table(self):
        _clear_layout(self._table_layout)
        if not self._interviews:
            empty = QLabel("No interviews yet (click Scan calendar).")
            empty.setObjectName("Hint")
            self._table_layout.addWidget(empty)
            return
        # header
        head = QFrame()
        head.setStyleSheet("background: #f0f2f8; border-radius: 8px;")
        hh = QHBoxLayout(head)
        hh.setContentsMargins(12, 8, 12, 8)
        h1 = QLabel("Interview (subject)")
        h1.setObjectName("SectionLabel")
        hh.addWidget(h1, 1)
        h2 = QLabel("Replied to candidate?")
        h2.setObjectName("SectionLabel")
        hh.addWidget(h2)
        self._table_layout.addWidget(head)
        for appt in self._interviews:
            self._table_layout.addWidget(self._table_row(appt))

    def _table_row(self, appt):
        row = QFrame()
        row.setStyleSheet("QFrame { border-bottom: 1px solid #e7ebf3; }")
        h = QHBoxLayout(row)
        h.setContentsMargins(12, 8, 12, 8)
        info = QVBoxLayout()
        info.setSpacing(2)
        subj = QLabel(appt["subject"] or "(no title)")
        subj.setWordWrap(True)
        info.addWidget(subj)
        start = appt.get("start")
        if start:
            when = QLabel(start.strftime("%d/%m/%Y %H:%M"))
            when.setObjectName("Hint")
            info.addWidget(when)
        h.addLayout(info, 1)
        h.addWidget(widgets.button(row, "Yes", variant="success", icon="check",
                                   command=lambda a=appt: self._dismiss(a)))
        h.addWidget(widgets.button(row, "No", variant="danger", icon="x",
                                   command=lambda a=appt: self._open_confirm(a)))
        return row

    # ----------------------------------------------------------- dialogs
    def _open_confirm(self, appt):
        from app_qt.dialogs import AppDialog
        r = AppDialog(
            self._page, "Confirm",
            f"Send the interview result now?\n\n{appt['subject']}",
            "question",
            buttons=[("No — Schedule a reminder", "neutral", 2),
                     ("Yes — Compose mail", "primary", 1)]).run()
        if r == 1:
            self._open_compose(appt)
        elif r == 2:
            self._open_schedule(appt)

    def _eligible_recipients(self, appt):
        exclude = self.var_exclude.get().strip().lower()
        out = []
        for e in appt.get("attendees", []):
            domain = e.split("@")[-1].lower() if "@" in e else ""
            if exclude and exclude in domain:
                continue
            out.append(e)
        return out

    def _open_compose(self, appt):
        dlg, card, lay = build_dialog_shell(self._page, "Compose candidate reply", size="md")

        lb1 = QLabel("To"); lb1.setObjectName("FieldLabel"); lay.addWidget(lb1)
        to_w = QLineEdit("; ".join(self._eligible_recipients(appt)))
        lay.addWidget(to_w)
        lb2 = QLabel("Subject"); lb2.setObjectName("FieldLabel"); lay.addWidget(lb2)
        subj_w = QLineEdit(_fill_template(self.var_subject.get().strip(), appt))
        lay.addWidget(subj_w)
        lb3 = QLabel("Body (select text, then B/I/U/color to format)")
        lb3.setObjectName("FieldLabel"); lb3.setWordWrap(True); lay.addWidget(lb3)
        body = richtext.RichText(card, height=16)
        body.set_html(_fill_template(self.body_editor.get_html(), appt))
        lay.addWidget(body)

        if not self._eligible_recipients(appt):
            widgets.hint(card, "⚠ No candidate email (outside the internal domain) found "
                               "in this event — please enter it manually in 'To'.")

        foot = QHBoxLayout()

        def do_send():
            to_value = to_w.text().strip()
            if not to_value:
                dialogs.warning(dlg, "Missing recipient", "Please enter a recipient email.")
                return
            try:
                outlook.send_mail(to_value, subj_w.text().strip(),
                                  body.get_text(), html=body.get_html())
            except Exception as exc:
                dialogs.error(dlg, "Send failed", f"Couldn't send:\n{exc}")
                return
            dlg.accept()
            self._dismiss(appt)
            dialogs.success(self._page, "Sent", "Reply sent ✅")

        foot.addWidget(widgets.button(card, "Send", variant="primary", icon="mail",
                                      command=do_send))
        foot.addWidget(widgets.button(card, "Cancel", variant="neutral", icon="x",
                                      command=dlg.reject))
        foot.addStretch(1)
        lay.addLayout(foot)
        dlg.exec()

    def _open_schedule(self, appt):
        dlg, card, lay = build_dialog_shell(self._page, "Schedule a follow-up reminder", size="sm")
        sub = QLabel(appt["subject"]); sub.setObjectName("DialogMsg"); sub.setWordWrap(True)
        lay.addWidget(sub)

        default_dt = (datetime.datetime.now() + datetime.timedelta(days=1)).replace(
            hour=9, minute=0, second=0, microsecond=0)

        def field(label, value):
            lb = QLabel(label); lb.setObjectName("FieldLabel"); lay.addWidget(lb)
            w = QLineEdit(value); lay.addWidget(w)
            return w

        date_w = field("Reminder date (dd/mm/yyyy)", default_dt.strftime("%d/%m/%Y"))
        time_w = field("Reminder time (HH:MM)", default_dt.strftime("%H:%M"))
        subj_w = field("Reminder title", f"Interview follow-up: {_extract_name(appt['subject'])}")

        foot = QHBoxLayout()

        def do_create():
            try:
                start = datetime.datetime.strptime(
                    f"{date_w.text().strip()} {time_w.text().strip()}", "%d/%m/%Y %H:%M")
            except ValueError:
                dialogs.warning(dlg, "Invalid format", "Date must be dd/mm/yyyy and time HH:MM.")
                return
            recipients = ", ".join(self._eligible_recipients(appt))
            body = (f"Reminder to send the interview result.\n\n"
                    f"Original event: {appt['subject']}\n"
                    f"Candidate: {recipients or '(email unknown)'}")
            try:
                outlook.create_appointment(subj_w.text().strip(), start,
                                           duration_minutes=30, body=body, reminder_minutes=15)
            except Exception as exc:
                dialogs.error(dlg, "Calendar error", f"Couldn't create the reminder:\n{exc}")
                return
            dlg.accept()
            dialogs.success(self._page, "Reminder set",
                            f"Reminder created for {start.strftime('%d/%m/%Y %H:%M')} ✅")

        foot.addWidget(widgets.button(card, "Create reminder", variant="primary",
                                      icon="calendar", command=do_create))
        foot.addWidget(widgets.button(card, "Cancel", variant="neutral", icon="x",
                                      command=dlg.reject))
        foot.addStretch(1)
        lay.addLayout(foot)
        dlg.exec()
