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

from app.core import config, outlook
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
    name = "Nhắc phản hồi PV"
    description = "Quét lịch phỏng vấn 1 tháng gần đây và nhắc gửi kết quả cho ứng viên."
    icon = "🔔"
    category = "Văn phòng"
    order = 6

    def build(self, parent=None):
        self._interviews = []
        cfg = config.load(SECTION, DEFAULTS)

        card = widgets.Card(parent)
        self._page = card
        lay = QVBoxLayout(card)
        lay.setContentsMargins(28, 24, 28, 24)
        lay.setSpacing(6)

        widgets.section_label(card, "Nhận diện lịch phỏng vấn")
        self.var_keywords = widgets.text_row(card, "Từ khóa trong tiêu đề (cách nhau bởi dấu phẩy)")
        self.var_keywords.set(cfg["keywords"])
        widgets.hint(card, "Sự kiện có tiêu đề chứa MỘT trong các từ khóa trên (không phân biệt "
                           "hoa thường) sẽ được coi là lịch phỏng vấn.")
        self.var_exclude = widgets.text_row(card, "Bỏ qua email thuộc domain (người nhận nội bộ)")
        self.var_exclude.set(cfg["exclude_domain"])

        widgets.section_label(card, "Mẫu email phản hồi ứng viên")
        self.var_subject = widgets.text_row(card, "Tiêu đề")
        self.var_subject.set(cfg["subject"])
        lbl = QLabel("Nội dung (dùng được {name}, {position}, {subject}, {date}, {time} — "
                     "bôi đen chữ rồi bấm B/I/U/màu để định dạng)")
        lbl.setObjectName("FieldLabel")
        lbl.setWordWrap(True)
        lay.addWidget(lbl)
        self.body_editor = richtext.RichText(card, height=11)
        self.body_editor.set_html(cfg["body"])
        lay.addWidget(self.body_editor)

        row = QHBoxLayout()
        row.addWidget(widgets.button(card, "Lưu cấu hình", variant="neutral",
                                     icon="save", command=self._save_config))
        row.addWidget(widgets.button(card, "Quét lịch (1 tháng gần đây)", variant="primary",
                                     icon="refresh", command=self._scan_clicked))
        row.addStretch(1)
        lay.addLayout(row)

        if not outlook.available():
            widgets.hint(card, "⚠ Không tìm thấy Outlook (pywin32). Tính năng quét/gửi mail "
                               "chỉ chạy trên Windows có cài Outlook.")

        widgets.section_label(card, "Lịch phỏng vấn quét được")
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
        self.info("Đã lưu", "Đã lưu cấu hình ✅")

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
        try:
            appts = outlook.appointments_between(_month_ago(today), today)
        except Exception:
            return []
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
            self.error("Cần Outlook", "Tính năng này cần Outlook trên Windows (pywin32).")
            return
        self._interviews = self._fetch_interviews()
        self._render_table()
        if not self._interviews:
            self.info("Không có lịch", "Không tìm thấy lịch phỏng vấn nào trong 1 tháng gần đây.")

    # ----------------------------------------------------------- bảng
    def _render_table(self):
        _clear_layout(self._table_layout)
        if not self._interviews:
            empty = QLabel("Chưa có lịch phỏng vấn nào (bấm Quét lịch).")
            empty.setObjectName("Hint")
            self._table_layout.addWidget(empty)
            return
        # header
        head = QFrame()
        head.setStyleSheet("background: #f0f2f8; border-radius: 8px;")
        hh = QHBoxLayout(head)
        hh.setContentsMargins(12, 8, 12, 8)
        h1 = QLabel("Lịch phỏng vấn (subject)")
        h1.setObjectName("SectionLabel")
        hh.addWidget(h1, 1)
        h2 = QLabel("Đã phản hồi ứng viên?")
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
        subj = QLabel(appt["subject"] or "(không có tiêu đề)")
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
            self._page, "Xác nhận",
            f"Gửi phản hồi kết quả phỏng vấn ngay bây giờ?\n\n{appt['subject']}",
            "question",
            buttons=[("No — Đặt lịch nhắc khác", "neutral", 2),
                     ("Yes — Soạn mail", "primary", 1)]).run()
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
        dlg, card, lay = build_dialog_shell(self._page, "Soạn mail phản hồi ứng viên", min_width=700)

        lb1 = QLabel("Đến"); lb1.setObjectName("FieldLabel"); lay.addWidget(lb1)
        to_w = QLineEdit("; ".join(self._eligible_recipients(appt)))
        lay.addWidget(to_w)
        lb2 = QLabel("Tiêu đề"); lb2.setObjectName("FieldLabel"); lay.addWidget(lb2)
        subj_w = QLineEdit(_fill_template(self.var_subject.get().strip(), appt))
        lay.addWidget(subj_w)
        lb3 = QLabel("Nội dung (bôi đen chữ rồi bấm B/I/U/màu để định dạng)")
        lb3.setObjectName("FieldLabel"); lb3.setWordWrap(True); lay.addWidget(lb3)
        body = richtext.RichText(card, height=16)
        body.set_html(_fill_template(self.body_editor.get_html(), appt))
        lay.addWidget(body)

        if not self._eligible_recipients(appt):
            widgets.hint(card, "⚠ Không tìm thấy email ứng viên (ngoài domain nội bộ) trong "
                               "lịch này — vui lòng nhập tay ở ô 'Đến'.")

        foot = QHBoxLayout()

        def do_send():
            to_value = to_w.text().strip()
            if not to_value:
                dialogs.warning(dlg, "Thiếu người nhận", "Vui lòng nhập email người nhận.")
                return
            try:
                outlook.send_mail(to_value, subj_w.text().strip(),
                                  body.get_text(), html=body.get_html())
            except Exception as exc:
                dialogs.error(dlg, "Lỗi gửi mail", f"Không gửi được:\n{exc}")
                return
            dlg.accept()
            self._dismiss(appt)
            dialogs.success(self._page, "Đã gửi", "Đã gửi mail phản hồi ✅")

        foot.addWidget(widgets.button(card, "Gửi mail", variant="primary", icon="mail",
                                      command=do_send))
        foot.addWidget(widgets.button(card, "Hủy", variant="neutral", icon="x",
                                      command=dlg.reject))
        foot.addStretch(1)
        lay.addLayout(foot)
        dlg.resize(760, 760)
        dlg.exec()

    def _open_schedule(self, appt):
        dlg, card, lay = build_dialog_shell(self._page, "Đặt lịch nhắc phản hồi", min_width=460)
        sub = QLabel(appt["subject"]); sub.setObjectName("DialogMsg"); sub.setWordWrap(True)
        lay.addWidget(sub)

        default_dt = (datetime.datetime.now() + datetime.timedelta(days=1)).replace(
            hour=9, minute=0, second=0, microsecond=0)

        def field(label, value):
            lb = QLabel(label); lb.setObjectName("FieldLabel"); lay.addWidget(lb)
            w = QLineEdit(value); lay.addWidget(w)
            return w

        date_w = field("Ngày nhắc (dd/mm/yyyy)", default_dt.strftime("%d/%m/%Y"))
        time_w = field("Giờ nhắc (HH:MM)", default_dt.strftime("%H:%M"))
        subj_w = field("Tiêu đề lời nhắc", f"Phản hồi PV: {_extract_name(appt['subject'])}")

        foot = QHBoxLayout()

        def do_create():
            try:
                start = datetime.datetime.strptime(
                    f"{date_w.text().strip()} {time_w.text().strip()}", "%d/%m/%Y %H:%M")
            except ValueError:
                dialogs.warning(dlg, "Sai định dạng", "Ngày phải là dd/mm/yyyy và giờ là HH:MM.")
                return
            recipients = ", ".join(self._eligible_recipients(appt))
            body = (f"Nhắc gửi phản hồi kết quả phỏng vấn.\n\n"
                    f"Lịch gốc: {appt['subject']}\n"
                    f"Ứng viên: {recipients or '(chưa rõ email)'}")
            try:
                outlook.create_appointment(subj_w.text().strip(), start,
                                           duration_minutes=30, body=body, reminder_minutes=15)
            except Exception as exc:
                dialogs.error(dlg, "Lỗi tạo lịch", f"Không tạo được lịch nhắc:\n{exc}")
                return
            dlg.accept()
            dialogs.success(self._page, "Đã đặt lịch",
                            f"Đã tạo lời nhắc lúc {start.strftime('%d/%m/%Y %H:%M')} ✅")

        foot.addWidget(widgets.button(card, "Tạo lịch nhắc", variant="primary",
                                      icon="calendar", command=do_create))
        foot.addWidget(widgets.button(card, "Hủy", variant="neutral", icon="x",
                                      command=dlg.reject))
        foot.addStretch(1)
        lay.addLayout(foot)
        dlg.exec()
