"""Quét lịch phỏng vấn trong Outlook → soạn mail nhờ Security mở cổng.

Luồng hoạt động:
  • Mỗi sáng khi mở app (nếu bật "tự động"), tool quét calendar Outlook xem
    hôm nay có lịch nào tiêu đề chứa từ khóa phỏng vấn không — chỉ chạy 1
    lần/ngày (nhớ ngày đã quét trong config).
  • Nếu có, app soạn sẵn mail (người nhận, tiêu đề, nội dung kèm danh sách
    ứng viên + giờ) rồi HIỆN RA cho bạn xem/sửa, bấm "Gửi" mới thật sự gửi.
  • Ngoài ra luôn có nút bấm tay để quét bất cứ lúc nào.

Cấu hình (từ khóa, email Security, mẫu mail) lưu trong app/core/config.py
nên lần quét tự động vẫn dùng đúng thiết lập bạn đã lưu.
"""
import datetime
from tkinter import messagebox

import tkinter as tk
import ttkbootstrap as ttk

from app.core import config, outlook
from app.core.base_tool import BaseTool
from app.ui import theme, widgets

SECTION = "interview_gate"

DEFAULTS = {
    "keywords": "phỏng vấn, pv, interview",
    "to": "",
    "cc": "",
    "subject": "Đề nghị mở cổng đón ứng viên phỏng vấn ngày",
    "body": (
        "Kính gửi team Security,\n\n"
        "Hôm nay bộ phận tuyển dụng có lịch phỏng vấn sau, "
        "nhờ team hỗ trợ mở cổng đón ứng viên:\n\n"
        "Cảm ơn team!"
    ),
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
    auto_startup = True

    # ------------------------------------------------------------------ UI
    def build_body(self, parent):
        self._parent = parent
        cfg = config.load(SECTION, DEFAULTS)

        widgets.section_label(parent, "Nhận diện lịch phỏng vấn")
        self.var_keywords = widgets.text_row(
            parent, "Từ khóa trong tiêu đề (cách nhau bởi dấu phẩy)",
            placeholder=cfg["keywords"],
        )
        widgets.hint(
            parent,
            "Sự kiện nào có tiêu đề chứa MỘT trong các từ khóa trên sẽ được "
            "coi là lịch phỏng vấn (không phân biệt hoa thường).",
        )

        widgets.section_label(parent, "Người nhận")
        self.var_to = widgets.text_row(
            parent, "Email nhận (nhiều email cách nhau bởi dấu ;)",
            placeholder=cfg["to"],
        )
        self.var_cc = widgets.text_row(
            parent, "CC (tùy chọn)", placeholder=cfg["cc"],
        )

        widgets.section_label(parent, "Mẫu mail")
        self.var_subject = widgets.text_row(
            parent, "Tiêu đề", placeholder=cfg["subject"],
        )
        self.body_box = widgets.text_area(
            parent, "Nội dung", value=cfg["body"], height=9,
        )

        self.var_auto = widgets.checkbox(
            parent, "Tự động quét khi mở app mỗi sáng", checked=cfg["auto"],
        )

        row = tk.Frame(parent, bg=theme.CARD_BG)
        row.pack(fill="x", pady=(6, 0))
        ttk.Button(
            row, text="💾 Lưu cấu hình", bootstyle="secondary-outline",
            command=self._save_config,
        ).pack(side="left")

        if not outlook.available():
            widgets.hint(
                parent,
                "⚠ Không tìm thấy Outlook (pywin32). Tính năng quét/gửi mail "
                "chỉ chạy trên Windows có cài Outlook.",
            )

    # -------------------------------------------------------------- config
    def _collect(self):
        return {
            "keywords": self.var_keywords.get().strip(),
            "to": self.var_to.get().strip(),
            "cc": self.var_cc.get().strip(),
            "subject": self.var_subject.get().strip(),
            "body": self.body_box.get("1.0", "end-1c"),
            "auto": bool(self.var_auto.get()),
            # giữ nguyên ngày đã quét gần nhất
            "last_scan": config.load(SECTION, DEFAULTS).get("last_scan", ""),
        }

    def _save_config(self):
        config.save(SECTION, self._collect())
        messagebox.showinfo("Đã lưu", "Đã lưu cấu hình ✅")

    # ---------------------------------------------------------- quét & gửi
    def run(self):
        """Nút bấm tay: lưu cấu hình rồi quét + soạn mail ngay."""
        config.save(SECTION, self._collect())
        self._scan_and_confirm(parent=self._parent.winfo_toplevel(),
                               silent_if_empty=False)

    def startup(self, window):
        """Tự chạy khi mở app — mỗi ngày một lần."""
        cfg = config.load(SECTION, DEFAULTS)
        if not cfg.get("auto"):
            return
        today = datetime.date.today().isoformat()
        if cfg.get("last_scan") == today:
            return            # hôm nay đã quét rồi
        cfg["last_scan"] = today
        config.save(SECTION, cfg)   # đánh dấu trước, tránh quét lại trong ngày
        self._scan_and_confirm(parent=window, window=window,
                               silent_if_empty=True)

    def _scan_and_confirm(self, parent, window=None, silent_if_empty=False):
        if not outlook.available():
            if not silent_if_empty:
                messagebox.showwarning(
                    "Cần Outlook",
                    "Tính năng này cần Outlook trên Windows (pywin32).",
                )
            return

        cfg = config.load(SECTION, DEFAULTS)
        try:
            appointments = outlook.today_appointments()
        except Exception as exc:           # noqa: BLE001
            if not silent_if_empty:
                messagebox.showerror(
                    "Lỗi đọc Outlook", f"Không đọc được lịch:\n{exc}")
            return

        keywords = [k.strip().lower()
                    for k in cfg["keywords"].split(",") if k.strip()]
        if keywords:
            interviews = [
                a for a in appointments
                if any(kw in a["subject"].lower() for kw in keywords)
            ]
        else:
            interviews = appointments

        if not interviews:
            if not silent_if_empty:
                messagebox.showinfo(
                    "Không có lịch",
                    "Hôm nay không có lịch phỏng vấn nào trong Outlook.",
                )
            return

        subject, body = self._compose(interviews, cfg)

        if window is not None:             # quét tự động → đưa app ra trước
            try:
                window.deiconify()
                window.lift()
                window.show_tool(self)
            except Exception:
                pass

        self._open_confirm(parent, cfg["to"], cfg["cc"], subject, body)

    def _compose(self, interviews, cfg):
        date_str = datetime.date.today().strftime("%d/%m/%Y")
        lines = []
        for a in interviews:
            t = a["start"].strftime("%H:%M") if a["start"] else "??:??"
            line = f"- {t} — {a['subject']}"
            if a["location"]:
                line += f" ({a['location']})"
            lines.append(line)
        listing = "\n".join(lines)

        subject = cfg["subject"].rstrip() + f" {date_str}"

        body = cfg["body"]
        idx = body.rfind("\nCảm ơn")
        if idx != -1:
            body = body[:idx].rstrip() + f"\n\n{listing}" + body[idx:]
        else:
            body = body.rstrip() + f"\n\n{listing}"

        return subject, body

    # ---------------------------------------------------- hộp thoại xác nhận
    def _open_confirm(self, parent, to, cc, subject, body):
        dlg = tk.Toplevel(parent)
        dlg.title("Xác nhận gửi mail")
        dlg.configure(bg=theme.CONTENT_BG)
        dlg.geometry("640x560")
        dlg.transient(parent)
        dlg.grab_set()

        wrap = tk.Frame(dlg, bg=theme.CONTENT_BG)
        wrap.pack(fill="both", expand=True, padx=22, pady=20)

        tk.Label(
            wrap, text="Kiểm tra & gửi mail", bg=theme.CONTENT_BG,
            fg=theme.TEXT, font=(theme.FONT_FAMILY, 15, "bold"),
        ).pack(anchor="w", pady=(0, 12))

        def field(label, value, single=True, height=10):
            tk.Label(
                wrap, text=label, bg=theme.CONTENT_BG, fg=theme.TEXT,
                font=(theme.FONT_FAMILY, 9),
            ).pack(anchor="w", pady=(8, 3))
            if single:
                var = tk.StringVar(value=value)
                ttk.Entry(wrap, textvariable=var).pack(fill="x", ipady=3)
                return var
            box = tk.Text(
                wrap, height=height, wrap="word", relief="solid", bd=1,
                font=(theme.FONT_FAMILY, 10), bg="#ffffff", fg=theme.TEXT,
                highlightthickness=1, highlightbackground=theme.BORDER,
                padx=8, pady=6,
            )
            box.pack(fill="both", expand=True)
            box.insert("1.0", value)
            return box

        to_var = field("Đến", to)
        cc_var = field("CC", cc)
        subject_var = field("Tiêu đề", subject)
        body_box = field("Nội dung", body, single=False, height=11)

        actions = tk.Frame(wrap, bg=theme.CONTENT_BG)
        actions.pack(fill="x", pady=(16, 0))

        def do_send():
            to_value = to_var.get().strip()
            if not to_value:
                messagebox.showwarning(
                    "Thiếu người nhận", "Vui lòng nhập email người nhận.",
                    parent=dlg)
                return
            try:
                outlook.send_mail(
                    to_value, subject_var.get().strip(),
                    body_box.get("1.0", "end-1c"), cc=cc_var.get().strip(),
                )
            except Exception as exc:        # noqa: BLE001
                messagebox.showerror(
                    "Lỗi gửi mail", f"Không gửi được:\n{exc}", parent=dlg)
                return
            dlg.destroy()
            messagebox.showinfo("Đã gửi", "Đã gửi mail✅")

        ttk.Button(
            actions, text="Gửi mail", bootstyle="primary", command=do_send,
        ).pack(side="left", ipadx=10, ipady=3)
        ttk.Button(
            actions, text="Hủy", bootstyle="secondary-outline",
            command=dlg.destroy,
        ).pack(side="left", padx=(10, 0), ipady=3)
