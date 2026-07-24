"""Gửi mail theo lịch phỏng vấn (Outlook) — bản PySide6.

Quét lịch Outlook hôm nay, lọc theo từ khóa, soạn sẵn mail nhờ Security mở cổng
rồi hiện hộp thoại cho sửa/gửi. Tự chạy 1 lần/ngày khi mở app (auto_startup).
Dùng lại app.core.outlook (COM) — chỉ dựng lại giao diện bằng Qt.
"""
import datetime
import re

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame, QHBoxLayout, QLabel, QLineEdit, QTextEdit, QVBoxLayout,
)

from app.core import config, outlook
from app_qt import dialogs, widgets
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
    name = "Gửi mail theo lịch"
    description = "Quét lịch phỏng vấn Outlook hôm nay và soạn mail nhờ Security mở cổng."
    icon = "📧"
    category = "Văn phòng"
    order = 5
    action_label = "Quét lịch hôm nay"
    action_icon = "search"
    auto_startup = True

    def build_body(self, parent):
        cfg = config.load(SECTION, DEFAULTS)
        widgets.section_label(parent, "Nhận diện lịch phỏng vấn")
        self.var_keywords = widgets.text_row(parent, "Từ khóa trong tiêu đề (cách nhau bởi dấu phẩy)")
        self.var_keywords.set(cfg["keywords"])
        widgets.hint(parent, "Sự kiện có tiêu đề chứa MỘT trong các từ khóa trên sẽ được "
                             "coi là lịch phỏng vấn (không phân biệt hoa thường).")
        widgets.section_label(parent, "Người nhận")
        self.var_to = widgets.text_row(parent, "Email nhận (nhiều email cách nhau bởi dấu ;)")
        self.var_to.set(cfg["to"])
        self.var_cc = widgets.text_row(parent, "CC (tùy chọn)")
        self.var_cc.set(cfg["cc"])
        widgets.section_label(parent, "Mẫu mail")
        self.var_subject = widgets.text_row(parent, "Tiêu đề")
        self.var_subject.set(cfg["subject"])
        self.body_box = widgets.text_area(parent, "Nội dung", value=cfg["body"], height=9)
        self.var_auto = widgets.checkbox(parent, "Tự động quét khi mở app mỗi sáng",
                                         checked=cfg["auto"])
        row = QHBoxLayout()
        row.addWidget(widgets.button(parent, "Lưu cấu hình", variant="neutral",
                                     icon="save", command=self._save_config))
        row.addStretch(1)
        parent.layout().addLayout(row)
        if not outlook.available():
            widgets.hint(parent, "⚠ Không tìm thấy Outlook (pywin32). Tính năng quét/gửi mail "
                                 "chỉ chạy trên Windows có cài Outlook.")

    def _collect(self):
        return {
            "keywords": self.var_keywords.get().strip(),
            "to": self.var_to.get().strip(),
            "cc": self.var_cc.get().strip(),
            "subject": self.var_subject.get().strip(),
            "body": self.body_box.get(),
            "auto": bool(self.var_auto.get()),
            "last_scan": config.load(SECTION, DEFAULTS).get("last_scan", ""),
        }

    def _save_config(self):
        config.save(SECTION, self._collect())
        self.info("Đã lưu", "Đã lưu cấu hình ✅")

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
                dialogs.warning(parent, "Cần Outlook",
                                "Tính năng này cần Outlook trên Windows (pywin32).")
            return
        cfg = config.load(SECTION, DEFAULTS)
        try:
            appointments = outlook.today_appointments()
        except Exception as exc:
            if not silent_if_empty:
                dialogs.error(parent, "Lỗi đọc Outlook", f"Không đọc được lịch:\n{exc}")
            return
        keywords = [k.strip().lower() for k in cfg["keywords"].split(",") if k.strip()]
        interviews = ([a for a in appointments
                       if any(kw in a["subject"].lower() for kw in keywords)]
                      if keywords else appointments)
        if not interviews:
            if not silent_if_empty:
                dialogs.info(parent, "Không có lịch",
                             "Hôm nay không có lịch phỏng vấn nào trong Outlook.")
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
        lines = []
        for a in interviews:
            t = a["start"].strftime("%H:%M") if a["start"] else "??:??"
            name = self._extract_name(a["subject"])
            line = f"- {t} — {name}"
            if a["location"]:
                line += f" ({a['location']})"
            lines.append(line)
        listing = "\n".join(lines)
        subject = cfg["subject"]
        body = cfg["body"]
        idx = body.rfind("\nCảm ơn")
        if idx != -1:
            body = body[:idx].rstrip() + f"\n\n{listing}" + body[idx:]
        else:
            body = body.rstrip() + f"\n\n{listing}"
        return subject, body

    def _open_confirm(self, parent, to, cc, subject, body):
        dlg, card, lay = build_dialog_shell(parent, "Kiểm tra & gửi mail", size="md")

        def field(label, value, multiline=False):
            lb = QLabel(label); lb.setObjectName("FieldLabel")
            lay.addWidget(lb)
            if multiline:
                w = widgets.TextEdit(); w.setAcceptRichText(False); w.setPlainText(value)
                w.setMinimumHeight(220)
            else:
                w = QLineEdit(value)
            lay.addWidget(w)
            return w

        to_w = field("Đến", to)
        cc_w = field("CC", cc)
        subj_w = field("Tiêu đề", subject)
        body_w = field("Nội dung", body, multiline=True)

        foot = QHBoxLayout()

        def do_send():
            to_value = to_w.text().strip()
            if not to_value:
                dialogs.warning(dlg, "Thiếu người nhận", "Vui lòng nhập email người nhận.")
                return
            try:
                outlook.send_mail(to_value, subj_w.text().strip(),
                                  body_w.toPlainText(), cc=cc_w.text().strip())
            except Exception as exc:
                dialogs.error(dlg, "Lỗi gửi mail", f"Không gửi được:\n{exc}")
                return
            dlg.accept()
            dialogs.success(parent, "Đã gửi", "Đã gửi mail ✅")

        foot.addWidget(widgets.button(card, "Gửi mail", variant="primary", icon="mail",
                                      command=do_send))
        foot.addWidget(widgets.button(card, "Hủy", variant="neutral", icon="x",
                                      command=dlg.reject))
        foot.addStretch(1)
        lay.addLayout(foot)
        body_w.setMinimumHeight(round(dlg.modal_h * 0.5))   # vùng nội dung cao theo cỡ md
        dlg.exec()
