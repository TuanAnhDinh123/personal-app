"""Gửi mail theo lịch phỏng vấn (Outlook) — bản PySide6.

Quét lịch Outlook hôm nay, lọc theo từ khóa, soạn sẵn mail nhờ Security mở cổng
rồi hiện hộp thoại cho sửa/gửi. Tự chạy 1 lần/ngày khi mở app (auto_startup).
Dùng lại app.core.outlook (COM) — chỉ dựng lại giao diện bằng Qt.
"""
import datetime
import re
from html import escape

from PySide6.QtWidgets import QHBoxLayout, QLabel, QLineEdit

from app.core import config, outlook
from app_qt import dialogs, richtext, widgets
from app_qt.base_tool import BaseTool
from app_qt.components.dialog_base import build_dialog_shell

SECTION = "interview_gate"

DEFAULTS = {
    "keywords": "phỏng vấn, pv, interview",
    "to": "",
    "cc": "",
    "subject": "Đề nghị mở cổng đón ứng viên phỏng vấn ngày",
    "body": ("Kính gửi team Security,\n\n"
             "Hôm nay bộ phận tuyển dụng có lịch phỏng vấn sau, nhờ team hỗ trợ "
             "mở cổng đón ứng viên:\n\n"
             "Cảm ơn team!"),
    "auto": True,
    "last_scan": "",
}


class InterviewGateTool(BaseTool):
    name = "Gate-Open Mail"
    description = "Scan today's Outlook interviews and draft a mail asking Security to open the gate."
    icon = "📧"
    category = "Office"
    order = 5
    action_label = "Scan today's calendar"
    action_icon = "search"
    auto_startup = True

    def build_body(self, parent):
        cfg = config.load(SECTION, DEFAULTS)
        widgets.section_label(parent, "Interview detection")
        self.var_keywords = widgets.text_row(parent, "Title keywords (comma-separated)")
        self.var_keywords.set(cfg["keywords"])
        widgets.hint(parent, "Events whose title contains any of these keywords count "
                             "as interviews (case-insensitive).")
        widgets.section_label(parent, "Recipients")
        self.var_to = widgets.text_row(parent, "To (separate emails with ;)")
        self.var_to.set(cfg["to"])
        self.var_cc = widgets.text_row(parent, "CC (optional)")
        self.var_cc.set(cfg["cc"])
        widgets.section_label(parent, "Email template")
        self.var_subject = widgets.text_row(parent, "Subject")
        self.var_subject.set(cfg["subject"])
        self.body_box = widgets.richtext_area(
            parent, "Body (select text, then B/I/U/color to format)",
            value=cfg["body"], height=9)
        self.var_auto = widgets.checkbox(parent, "Auto-scan on app launch each morning",
                                         checked=cfg["auto"])
        row = QHBoxLayout()
        row.addWidget(widgets.button(parent, "Save config", variant="neutral",
                                     icon="save", command=self._save_config))
        row.addStretch(1)
        parent.layout().addLayout(row)
        if not outlook.available():
            widgets.hint(parent, "⚠ Outlook not found (pywin32). Scanning/sending mail "
                                 "only works on Windows with Outlook installed.")

    def _collect(self):
        return {
            "keywords": self.var_keywords.get().strip(),
            "to": self.var_to.get().strip(),
            "cc": self.var_cc.get().strip(),
            "subject": self.var_subject.get().strip(),
            "body": self.body_box.get_html(),
            "auto": bool(self.var_auto.get()),
            "last_scan": config.load(SECTION, DEFAULTS).get("last_scan", ""),
        }

    def _save_config(self):
        config.save(SECTION, self._collect())
        self.info("Saved", "Config saved ✅")

    # ---------------------------------------------------------- quét & gửi
    def run(self):
        config.save(SECTION, self._collect())
        self._scan_and_confirm(self._page, silent_if_empty=False)

    def startup(self, window):
        cfg = config.load(SECTION, DEFAULTS)
        if not cfg.get("auto"):
            return
        today = datetime.date.today().isoformat()
        if cfg.get("last_scan") == today:
            return
        cfg["last_scan"] = today
        config.save(SECTION, cfg)
        self._scan_and_confirm(window, window=window, silent_if_empty=True)

    def _scan_and_confirm(self, parent, window=None, silent_if_empty=False):
        if not outlook.available():
            if not silent_if_empty:
                dialogs.warning(parent, "Outlook required",
                                "This feature needs Outlook on Windows (pywin32).")
            return
        cfg = config.load(SECTION, DEFAULTS)
        try:
            appointments = outlook.today_appointments()
        except Exception as exc:
            if not silent_if_empty:
                dialogs.error(parent, "Outlook read error", f"Couldn't read the calendar:\n{exc}")
            return
        keywords = [k.strip().lower() for k in cfg["keywords"].split(",") if k.strip()]
        interviews = ([a for a in appointments
                       if any(kw in a["subject"].lower() for kw in keywords)]
                      if keywords else appointments)
        if not interviews:
            if not silent_if_empty:
                dialogs.info(parent, "No events",
                             "No interviews in Outlook today.")
            return
        subject, body = self._compose(interviews, cfg)
        if window is not None:
            try:
                window.showNormal()
                window.raise_()
                window.activateWindow()
                window._show_tool(self)
            except Exception:
                pass
        self._open_confirm(parent, cfg["to"], cfg["cc"], subject, body)

    @staticmethod
    def _extract_name(subject):
        m = re.search(r'\b(?:Mr|Ms)\.\s*(.+)', subject, re.IGNORECASE)
        return m.group(1).strip() if m else subject

    def _compose(self, interviews, cfg):
        """Trả về (tiêu đề, thân mail HTML) — danh sách lịch chèn trước 'Cảm ơn'."""
        lines = []
        for a in interviews:
            t = a["start"].strftime("%H:%M") if a["start"] else "??:??"
            name = self._extract_name(a["subject"])
            line = f"- {t} — {name}"
            if a["location"]:
                line += f" ({a['location']})"
            lines.append(escape(line))
        listing = "<br>".join(lines)
        body = richtext.insert_html(cfg["body"], listing, before_text="Cảm ơn")
        return cfg["subject"], body

    def _open_confirm(self, parent, to, cc, subject, body):
        dlg, card, lay = build_dialog_shell(parent, "Review & send email", size="md")

        def field(label, value, multiline=False):
            lb = QLabel(label); lb.setObjectName("FieldLabel")
            lb.setWordWrap(True)
            lay.addWidget(lb)
            if multiline:
                w = richtext.RichText(card, height=2)
                w.set_html(value)
            else:
                w = QLineEdit(value)
            lay.addWidget(w)
            return w

        to_w = field("To", to)
        cc_w = field("CC", cc)
        subj_w = field("Subject", subject)
        body_w = field("Body (select text, then B/I/U/color to format)",
                       body, multiline=True)

        foot = QHBoxLayout()

        def do_send():
            to_value = to_w.text().strip()
            if not to_value:
                dialogs.warning(dlg, "Missing recipient", "Please enter a recipient email.")
                return
            try:
                outlook.send_mail(to_value, subj_w.text().strip(),
                                  body_w.get_text(), cc=cc_w.text().strip(),
                                  html=body_w.get_html())
            except Exception as exc:
                dialogs.error(dlg, "Send failed", f"Couldn't send:\n{exc}")
                return
            dlg.accept()
            dialogs.success(parent, "Sent", "Email sent ✅")

        foot.addWidget(widgets.button(card, "Send", variant="primary", icon="mail",
                                      command=do_send))
        foot.addWidget(widgets.button(card, "Cancel", variant="neutral", icon="x",
                                      command=dlg.reject))
        foot.addStretch(1)
        lay.addLayout(foot)
        body_w.setMinimumHeight(round(dlg.modal_h * 0.5))   # vùng nội dung cao theo cỡ md
        dlg.exec()
