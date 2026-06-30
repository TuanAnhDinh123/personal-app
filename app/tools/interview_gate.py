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
import os
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


def _log_path():
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, "PersonalToolbox", "interview_gate.log")


def _write_log(lines):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    path = _log_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "a", encoding="utf-8") as f:
        for line in lines:
            f.write(f"[{ts}] {line}\n")


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
        ttk.Button(
            row, text="📋 Xem log", bootstyle="secondary-outline",
            command=self._open_log,
        ).pack(side="left", padx=(8, 0))

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

    def _open_log(self):
        path = _log_path()
        if not os.path.exists(path):
            messagebox.showinfo("Chưa có log", f"Chưa có file log nào.\n({path})")
            return
        os.startfile(path)

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
        _write_log(["=== BẮT ĐẦU QUÉT ==="])
        try:
            appointments = outlook.today_appointments(log=lambda m: _write_log([m]))
        except Exception as exc:           # noqa: BLE001
            _write_log([f"[ERROR] Không đọc được lịch Outlook: {exc}"])
            if not silent_if_empty:
                messagebox.showerror(
                    "Lỗi đọc Outlook", f"Không đọc được lịch:\n{exc}")
            return

        keywords = [k.strip().lower()
                    for k in cfg["keywords"].split(",") if k.strip()]

        debug_lines = [f"[DEBUG] Tổng sự kiện hôm nay: {len(appointments)} | Keyword: {cfg['keywords']}"]
        for a in appointments:
            t = a["start"].strftime("%H:%M") if a["start"] else "??:??"
            matched = any(kw in a["subject"].lower() for kw in keywords) if keywords else True
            flag = "✓" if matched else "✗"
            entry = f"  [{flag}] {t} | {a['subject']}"
            if a["location"]:
                entry += f" | {a['location']}"
            debug_lines.append(entry)
        _write_log(debug_lines)

        if keywords:
            interviews = [
                a for a in appointments
                if any(kw in a["subject"].lower() for kw in keywords)
            ]
        else:
            interviews = appointments

        if not interviews:
            _write_log([f"Kết quả: không có lịch nào khớp keyword."])
            if not silent_if_empty:
                messagebox.showinfo(
                    "Không có lịch",
                    "Hôm nay không có lịch phỏng vấn nào trong Outlook.",
                )
            return

        log_lines = [f"Quét lịch: tìm thấy {len(interviews)}/{len(appointments)} lịch phỏng vấn:"]
        for a in interviews:
            t = a["start"].strftime("%H:%M") if a["start"] else "??:??"
            entry = f"  - {t} {a['subject']}"
            if a["location"]:
                entry += f" ({a['location']})"
            log_lines.append(entry)
        _write_log(log_lines)

        subject, body = self._compose(interviews, cfg)

        if window is not None:             # quét tự động → đưa app ra trước
            try:
                window.deiconify()
                window.lift()
                window.show_tool(self)
            except Exception:
                pass

        self._open_confirm(parent, cfg["to"], cfg["cc"], subject, body)

    @staticmethod
    def _extract_name(subject):
        """Lấy phần text sau 'Mr.' hoặc 'Ms.' trong subject mail."""
        import re
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

        # Tiêu đề giữ NGUYÊN như người dùng nhập ở UI (không tự thêm ngày).
        subject = cfg["subject"]

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
        dlg.geometry("960x840")          # to hơn 1.5× so với trước (640x560)
        dlg.transient(parent)
        dlg.grab_set()

        # --- vùng cuộn dọc: Canvas + Scrollbar, nội dung nằm trong `wrap` ---
        canvas = tk.Canvas(dlg, bg=theme.CONTENT_BG, highlightthickness=0)
        vbar = ttk.Scrollbar(dlg, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vbar.set)
        vbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        wrap = tk.Frame(canvas, bg=theme.CONTENT_BG)
        wrap_id = canvas.create_window((0, 0), window=wrap, anchor="nw")

        def _sync_scrollregion(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _stretch_wrap(event):
            # cho nội dung rộng bằng canvas (trừ padding hai bên)
            canvas.itemconfigure(wrap_id, width=event.width)

        wrap.bind("<Configure>", _sync_scrollregion)
        canvas.bind("<Configure>", _stretch_wrap)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-event.delta / 120), "units")

        dlg.bind("<MouseWheel>", _on_mousewheel)
        dlg.bind("<Destroy>", lambda e: dlg.unbind("<MouseWheel>"))

        # padding nội dung bên trong frame cuộn
        wrap.configure(padx=22, pady=20)

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
            _write_log([f"Đã gửi mail → {to_value}"])
            dlg.destroy()
            messagebox.showinfo("Đã gửi", "Đã gửi mail✅")

        ttk.Button(
            actions, text="Gửi mail", bootstyle="primary", command=do_send,
        ).pack(side="left", ipadx=10, ipady=3)
        ttk.Button(
            actions, text="Hủy", bootstyle="secondary-outline",
            command=dlg.destroy,
        ).pack(side="left", padx=(10, 0), ipady=3)
