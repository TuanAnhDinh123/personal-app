"""Nhắc phản hồi kết quả phỏng vấn cho ứng viên.

Luồng hoạt động:
  • Bấm "🔄 Quét lịch" để quét lịch Outlook từ HÔM NAY lùi về 1 tháng, lấy các
    sự kiện có tiêu đề chứa từ khóa phỏng vấn ("interview", "interview
    invitation"…). Cấu hình (từ khóa, mẫu mail) giống tính năng "Gửi mail theo
    lịch". Không quét tự động khi mở app (tránh chặn/đơ lúc khởi động).
  • Hiển thị các lịch đó thành 1 bảng: cột 1 là tiêu đề (subject) của lịch,
    cột 2 có 2 nút Yes / No cho câu hỏi "đã phản hồi ứng viên hay chưa?".
        - Bấm YES  → coi như đã phản hồi: xóa khỏi bảng, lần sau không hiện nữa.
        - Bấm NO   → hiện popup xác nhận "Gửi phản hồi ngay bây giờ?"
              · Yes → mở popup SOẠN MAIL (gửi xong thì xóa record như bấm Yes).
              · No  → mở popup ĐẶT LỊCH NHẮC khác.
  • Người nhận mail lấy từ chính lịch đó (loại bỏ email nội bộ, vd datalogic.com).

Các lịch đã xử lý được nhớ trong config (theo EntryID) nên không hiện lại.
"""
import calendar
import datetime
import re
from tkinter import messagebox

import tkinter as tk
import ttkbootstrap as ttk

from app.core import config, outlook
from app.core.base_tool import BaseTool
from app.ui import richtext, theme, widgets

SECTION = "reminder"

DEFAULTS = {
    "keywords": "interview, interview invitation, phỏng vấn, pv",
    "exclude_domain": "datalogic.com",
    "subject": "Interview Result — {position}",
    "body": (
        "Dear {name},\n\n"
        "Thank you for taking the time to interview for the {position} "
        "position with us. We truly appreciate your interest and the effort "
        "you put into the process.\n\n"
        "After careful consideration, we would like to inform you of the "
        "result of your interview:\n\n"
        "…\n\n"
        "Should you have any questions, please feel free to contact us.\n\n"
        "Best regards,\n"
    ),
    "dismissed": [],      # EntryID các lịch đã xử lý (đã phản hồi / đã gửi mail)
}


def _month_ago(d):
    """Trả về ngày cùng 'ngày' nhưng lùi lại 1 tháng (kẹp cuối tháng nếu cần)."""
    month = d.month - 1 or 12
    year = d.year - (1 if d.month == 1 else 0)
    day = min(d.day, calendar.monthrange(year, month)[1])
    return d.replace(year=year, month=month, day=day)


def _extract_name(subject):
    """Lấy tên ứng viên: phần sau 'Mr.'/'Ms.'/'Mrs.' nếu có, không thì cả subject."""
    m = re.search(r'\b(?:Mr|Ms|Mrs)\.\s*(.+)', subject, re.IGNORECASE)
    return m.group(1).strip() if m else subject.strip()


def _extract_position(subject):
    """Lấy vị trí ứng tuyển từ tiêu đề dạng 'Interview invitation for <vị trí> - Mr. ...'.

    Bỏ phần tên (từ 'Mr.'/'Ms.'/'Mrs.' trở đi) và các dấu ngăn cách ở cuối.
    """
    m = re.search(r'interview\s+(?:invitation\s+)?for\s+(.+)', subject,
                  re.IGNORECASE)
    if not m:
        return ""
    rest = re.split(r'\b(?:Mr|Ms|Mrs)\.', m.group(1), maxsplit=1,
                    flags=re.IGNORECASE)[0]
    return rest.strip().strip("-–—").strip()


def _fill_template(text, appt):
    """Thay các placeholder {name}/{position}/{subject}/{date}/{time} trong mẫu."""
    start = appt.get("start")
    return (
        text.replace("{name}", _extract_name(appt["subject"]))
            .replace("{position}", _extract_position(appt["subject"]))
            .replace("{subject}", appt["subject"])
            .replace("{date}", start.strftime("%d/%m/%Y") if start else "")
            .replace("{time}", start.strftime("%H:%M") if start else "")
    )


class ReminderTool(BaseTool):
    name = "Nhắc phản hồi PV"
    description = "Quét lịch phỏng vấn 1 tháng gần đây và nhắc gửi kết quả cho ứng viên."
    icon = "🔔"
    category = "Văn phòng"
    order = 6

    # ------------------------------------------------------------------ UI
    def build(self, parent):
        """Ghi đè: bố cục riêng (không dùng nút thao tác mặc định của BaseTool)."""
        outer = tk.Frame(parent, bg=theme.CONTENT_BG)
        card = tk.Frame(
            outer, bg=theme.CARD_BG,
            highlightbackground=theme.BORDER, highlightthickness=1,
        )
        card.pack(fill="x", anchor="n")
        inner = tk.Frame(card, bg=theme.CARD_BG)
        inner.pack(fill="both", expand=True, padx=28, pady=24)
        self.build_body(inner)
        return outer

    def build_body(self, parent):
        self._parent = parent
        self._interviews = []
        cfg = config.load(SECTION, DEFAULTS)

        # --- Khu vực cấu hình ---
        widgets.section_label(parent, "Nhận diện lịch phỏng vấn")
        self.var_keywords = widgets.text_row(
            parent, "Từ khóa trong tiêu đề (cách nhau bởi dấu phẩy)",
            placeholder=cfg["keywords"],
        )
        widgets.hint(
            parent,
            "Sự kiện có tiêu đề chứa MỘT trong các từ khóa trên (không phân biệt "
            "hoa thường) sẽ được coi là lịch phỏng vấn.",
        )
        self.var_exclude = widgets.text_row(
            parent, "Bỏ qua email thuộc domain (người nhận nội bộ)",
            placeholder=cfg["exclude_domain"],
        )

        widgets.section_label(parent, "Mẫu email phản hồi ứng viên")
        self.var_subject = widgets.text_row(
            parent, "Tiêu đề", placeholder=cfg["subject"],
        )
        tk.Label(
            parent,
            text="Nội dung  (dùng được {name}, {position}, {subject}, {date}, "
                 "{time} — bôi đen chữ rồi bấm B/I/U/màu để định dạng)",
            bg=theme.CARD_BG, fg=theme.TEXT, font=(theme.FONT_FAMILY, 9),
            justify="left", wraplength=620, anchor="w",
        ).pack(anchor="w", pady=(6, 4))
        self.body_editor = richtext.RichText(parent, height=11)
        self.body_editor.pack(fill="x", pady=(0, 10))
        self.body_editor.set_html(cfg["body"])

        row = tk.Frame(parent, bg=theme.CARD_BG)
        row.pack(fill="x", pady=(6, 0))
        ttk.Button(
            row, text="💾 Lưu cấu hình", bootstyle="secondary-outline",
            command=self._save_config,
        ).pack(side="left")
        ttk.Button(
            row, text="🔄 Quét lịch (1 tháng gần đây)", bootstyle="primary",
            command=self._scan_clicked,
        ).pack(side="left", padx=(10, 0))

        if not outlook.available():
            widgets.hint(
                parent,
                "⚠ Không tìm thấy Outlook (pywin32). Tính năng quét/gửi mail chỉ "
                "chạy trên Windows có cài Outlook.",
            )

        # --- Khu vực dữ liệu (bảng lịch phỏng vấn) ---
        widgets.section_label(parent, "Lịch phỏng vấn quét được")
        self._table_holder = tk.Frame(parent, bg=theme.CARD_BG)
        self._table_holder.pack(fill="x", pady=(2, 0))

        # KHÔNG quét ngay khi mở tool (quét calendar đồng bộ sẽ chặn UI, gây đơ).
        # Người dùng bấm "🔄 Quét lịch" — nút đã có hiệu ứng loading.
        self._render_table()

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
        messagebox.showinfo("Đã lưu", "Đã lưu cấu hình ✅")

    @staticmethod
    def _save_runtime(**changes):
        """Chỉ lưu các trường runtime (last_scan/dismissed) — KHÔNG đóng băng
        mẫu mail. Đọc thẳng phần đã lưu (bỏ qua DEFAULTS) để lần sau đổi
        DEFAULTS vẫn có tác dụng nếu người dùng chưa tự lưu cấu hình."""
        raw = dict(config._read_all().get(SECTION, {}))
        raw.update(changes)
        config.save(SECTION, raw)

    def _dismiss(self, appt):
        """Đánh dấu 1 lịch đã xử lý → nhớ EntryID & bỏ đúng dòng đó khỏi bảng.

        Không dựng lại cả bảng để KHÔNG làm khung cuộn nhảy về đầu."""
        eid = appt.get("entry_id")
        dismissed = list(config.load(SECTION, DEFAULTS).get("dismissed", []))
        if eid and eid not in dismissed:
            dismissed.append(eid)
            self._save_runtime(dismissed=dismissed)

        self._interviews = [a for a in self._interviews if a is not appt]
        rowf = self._row_widgets.pop(id(appt), None)
        if not self._interviews:
            self._render_table()          # hết dòng → hiện thông báo trống
        elif rowf is not None:
            rowf.destroy()                # xóa đúng dòng, giữ nguyên vị trí cuộn
        else:
            self._render_table()

    # ---------------------------------------------------------- quét lịch
    def _fetch_interviews(self):
        """Quét Outlook & trả về danh sách lịch phỏng vấn chưa xử lý (không dựng UI)."""
        if not outlook.available():
            return []
        cfg = config.load(SECTION, DEFAULTS)
        # ưu tiên giá trị đang gõ trên UI (nếu đã dựng), rồi tới config đã lưu
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
        """Nút '🔄 Quét lịch': quét lại theo giá trị đang nhập & cập nhật bảng.

        Không tự lưu mẫu mail (để không đóng băng default) — dùng nút 💾 để lưu."""
        if not outlook.available():
            messagebox.showwarning(
                "Cần Outlook",
                "Tính năng này cần Outlook trên Windows (pywin32).")
            return
        self._interviews = self._fetch_interviews()
        self._render_table()
        if not self._interviews:
            messagebox.showinfo(
                "Không có lịch",
                "Không tìm thấy lịch phỏng vấn nào trong 1 tháng gần đây.")

    # ----------------------------------------------------------- bảng dữ liệu
    def _render_table(self):
        for child in self._table_holder.winfo_children():
            child.destroy()
        self._row_widgets = {}   # id(appt) -> frame của dòng (để xóa lẻ từng dòng)

        if not self._interviews:
            tk.Label(
                self._table_holder,
                text="Chưa có lịch phỏng vấn nào (bấm 🔄 Quét lịch).",
                bg=theme.CARD_BG, fg=theme.MUTED,
                font=(theme.FONT_FAMILY, 9),
            ).pack(anchor="w", pady=6)
            return

        table = tk.Frame(
            self._table_holder, bg=theme.BORDER,
            highlightbackground=theme.BORDER, highlightthickness=1,
        )
        table.pack(fill="x")
        table.columnconfigure(0, weight=1)

        # Header
        head = tk.Frame(table, bg="#f0f2f8")
        head.grid(row=0, column=0, columnspan=2, sticky="ew")
        head.columnconfigure(0, weight=1)
        tk.Label(
            head, text="Lịch phỏng vấn (subject)", bg="#f0f2f8", fg=theme.TEXT,
            font=(theme.FONT_FAMILY, 9, "bold"), anchor="w", padx=12, pady=8,
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            head, text="Đã phản hồi ứng viên?", bg="#f0f2f8", fg=theme.TEXT,
            font=(theme.FONT_FAMILY, 9, "bold"), padx=12, pady=8,
        ).grid(row=0, column=1, sticky="e")

        for i, appt in enumerate(self._interviews, start=1):
            self._row_widgets[id(appt)] = self._table_row(table, i, appt)

    def _table_row(self, table, row_idx, appt):
        rowf = tk.Frame(table, bg=theme.CARD_BG)
        rowf.grid(row=row_idx, column=0, columnspan=2, sticky="ew", pady=(1, 0))
        rowf.columnconfigure(0, weight=1)

        start = appt.get("start")
        when = start.strftime("%d/%m/%Y %H:%M") if start else ""
        info = tk.Frame(rowf, bg=theme.CARD_BG)
        info.grid(row=0, column=0, sticky="w", padx=12, pady=9)
        tk.Label(
            info, text=appt["subject"] or "(không có tiêu đề)", bg=theme.CARD_BG,
            fg=theme.TEXT, font=(theme.FONT_FAMILY, 10), anchor="w",
            justify="left", wraplength=520,
        ).pack(anchor="w")
        if when:
            tk.Label(
                info, text=when, bg=theme.CARD_BG, fg=theme.MUTED,
                font=(theme.FONT_FAMILY, 8), anchor="w",
            ).pack(anchor="w")

        btns = tk.Frame(rowf, bg=theme.CARD_BG)
        btns.grid(row=0, column=1, sticky="e", padx=12, pady=6)
        ttk.Button(
            btns, text="Yes", bootstyle="success",
            command=lambda a=appt: self._on_yes(a),
        ).pack(side="left", ipadx=6)
        ttk.Button(
            btns, text="No", bootstyle="danger-outline",
            command=lambda a=appt: self._on_no(a),
        ).pack(side="left", padx=(8, 0), ipadx=6)
        return rowf

    # ------------------------------------------------------- xử lý nút Yes/No
    def _on_yes(self, appt):
        """Đã phản hồi → xóa khỏi bảng, lần sau không hiện."""
        self._dismiss(appt)

    def _on_no(self, appt):
        """Chưa phản hồi → hỏi có gửi phản hồi ngay không."""
        self._open_confirm(appt)

    # ------------------------------------------------------------- popup: confirm
    def _open_confirm(self, appt):
        parent = self._parent.winfo_toplevel()
        dlg = tk.Toplevel(parent)
        dlg.title("Xác nhận")
        dlg.configure(bg=theme.CONTENT_BG)
        dlg.geometry("460x400")
        dlg.transient(parent)
        dlg.grab_set()

        tk.Label(
            dlg, text="Gửi phản hồi kết quả phỏng vấn ngay bây giờ?",
            bg=theme.CONTENT_BG, fg=theme.TEXT,
            font=(theme.FONT_FAMILY, 12, "bold"), wraplength=400, justify="left",
        ).pack(anchor="w", padx=24, pady=(28, 6))
        tk.Label(
            dlg, text=appt["subject"], bg=theme.CONTENT_BG, fg=theme.MUTED,
            font=(theme.FONT_FAMILY, 9), wraplength=400, justify="left",
        ).pack(anchor="w", padx=24)

        actions = tk.Frame(dlg, bg=theme.CONTENT_BG)
        actions.pack(side="bottom", fill="x", padx=24, pady=20)

        def yes():
            dlg.destroy()
            self._open_compose(appt)

        def no():
            dlg.destroy()
            self._open_schedule(appt)

        ttk.Button(actions, text="Yes — Soạn mail", bootstyle="primary",
                   command=yes).pack(side="left", ipadx=8, ipady=2)
        ttk.Button(actions, text="No — Đặt lịch nhắc khác",
                   bootstyle="secondary-outline",
                   command=no).pack(side="left", padx=(10, 0), ipady=2)

    # ------------------------------------------------------------- popup: soạn mail
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
        parent = self._parent.winfo_toplevel()
        dlg = tk.Toplevel(parent)
        dlg.title("Soạn mail phản hồi")
        dlg.configure(bg=theme.CONTENT_BG)
        dlg.geometry("960x820")
        dlg.minsize(720, 520)
        dlg.transient(parent)
        dlg.grab_set()

        # --- Nút thao tác cố định ở ĐÁY (không cuộn theo) ---
        actions = tk.Frame(dlg, bg=theme.CONTENT_BG)
        actions.pack(side="bottom", fill="x", padx=24, pady=16)

        # --- Vùng cuộn dọc: Canvas + Scrollbar, nội dung nằm trong `wrap` ---
        canvas = tk.Canvas(dlg, bg=theme.CONTENT_BG, highlightthickness=0)
        vbar = ttk.Scrollbar(dlg, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vbar.set)
        vbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        wrap = tk.Frame(canvas, bg=theme.CONTENT_BG)
        wrap_id = canvas.create_window((0, 0), window=wrap, anchor="nw")
        wrap.configure(padx=24, pady=20)

        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfigure(wrap_id, width=e.width))
        wrap.bind("<Configure>",
                  lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        def _on_wheel(event):
            canvas.yview_scroll(int(-event.delta / 120), "units")
        dlg.bind("<MouseWheel>", _on_wheel)

        tk.Label(
            wrap, text="Soạn mail phản hồi ứng viên", bg=theme.CONTENT_BG,
            fg=theme.TEXT, font=(theme.FONT_FAMILY, 15, "bold"),
        ).pack(anchor="w", pady=(0, 10))

        def entry_field(label, value):
            tk.Label(
                wrap, text=label, bg=theme.CONTENT_BG, fg=theme.TEXT,
                font=(theme.FONT_FAMILY, 9),
            ).pack(anchor="w", pady=(8, 3))
            var = tk.StringVar(value=value)
            ttk.Entry(wrap, textvariable=var).pack(fill="x", ipady=3)
            return var

        to_var = entry_field("Đến", "; ".join(self._eligible_recipients(appt)))
        subject_var = entry_field(
            "Tiêu đề", _fill_template(self.var_subject.get().strip(), appt))

        tk.Label(
            wrap, text="Nội dung  (bôi đen chữ rồi bấm B/I/U/màu để định dạng)",
            bg=theme.CONTENT_BG, fg=theme.TEXT, font=(theme.FONT_FAMILY, 9),
        ).pack(anchor="w", pady=(8, 3))
        body_editor = richtext.RichText(wrap, height=18, bg=theme.CONTENT_BG)
        body_editor.pack(fill="x")
        body_editor.set_html(_fill_template(self.body_editor.get_html(), appt))

        if not self._eligible_recipients(appt):
            widgets.hint(
                wrap,
                "⚠ Không tìm thấy email ứng viên (ngoài domain nội bộ) trong lịch "
                "này — vui lòng nhập tay ở ô 'Đến'.", bg=theme.CONTENT_BG,
            )

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
                    body_editor.get_text(), html=body_editor.get_html())
            except Exception as exc:        # noqa: BLE001
                messagebox.showerror(
                    "Lỗi gửi mail", f"Không gửi được:\n{exc}", parent=dlg)
                return
            dlg.destroy()
            self._dismiss(appt)             # gửi xong -> bỏ khỏi bảng như bấm Yes
            messagebox.showinfo("Đã gửi", "Đã gửi mail phản hồi ✅")

        ttk.Button(actions, text="Gửi mail", bootstyle="primary",
                   command=do_send).pack(side="left", ipadx=10, ipady=3)
        ttk.Button(actions, text="Hủy", bootstyle="secondary-outline",
                   command=dlg.destroy).pack(side="left", padx=(10, 0), ipady=3)

    # ------------------------------------------------------- popup: đặt lịch nhắc
    def _open_schedule(self, appt):
        parent = self._parent.winfo_toplevel()
        dlg = tk.Toplevel(parent)
        dlg.title("Đặt lịch nhắc khác")
        dlg.configure(bg=theme.CONTENT_BG)
        dlg.geometry("480x420")
        dlg.transient(parent)
        dlg.grab_set()

        wrap = tk.Frame(dlg, bg=theme.CONTENT_BG)
        wrap.pack(fill="both", expand=True, padx=22, pady=20)

        tk.Label(
            wrap, text="Đặt lịch nhắc phản hồi", bg=theme.CONTENT_BG,
            fg=theme.TEXT, font=(theme.FONT_FAMILY, 14, "bold"),
        ).pack(anchor="w", pady=(0, 4))
        tk.Label(
            wrap, text=appt["subject"], bg=theme.CONTENT_BG, fg=theme.MUTED,
            font=(theme.FONT_FAMILY, 9), wraplength=420, justify="left",
        ).pack(anchor="w", pady=(0, 10))

        # mặc định: ngày mai 09:00
        default_dt = (datetime.datetime.now() + datetime.timedelta(days=1)).replace(
            hour=9, minute=0, second=0, microsecond=0)

        def entry(label, value):
            tk.Label(
                wrap, text=label, bg=theme.CONTENT_BG, fg=theme.TEXT,
                font=(theme.FONT_FAMILY, 9),
            ).pack(anchor="w", pady=(8, 3))
            var = tk.StringVar(value=value)
            ttk.Entry(wrap, textvariable=var).pack(fill="x", ipady=3)
            return var

        date_var = entry("Ngày nhắc (dd/mm/yyyy)", default_dt.strftime("%d/%m/%Y"))
        time_var = entry("Giờ nhắc (HH:MM)", default_dt.strftime("%H:%M"))
        subj_var = entry("Tiêu đề lời nhắc",
                         f"Phản hồi PV: {_extract_name(appt['subject'])}")

        actions = tk.Frame(wrap, bg=theme.CONTENT_BG)
        actions.pack(fill="x", pady=(20, 0))

        def do_create():
            try:
                start = datetime.datetime.strptime(
                    f"{date_var.get().strip()} {time_var.get().strip()}",
                    "%d/%m/%Y %H:%M")
            except ValueError:
                messagebox.showwarning(
                    "Sai định dạng",
                    "Ngày phải là dd/mm/yyyy và giờ là HH:MM.", parent=dlg)
                return
            recipients = ", ".join(self._eligible_recipients(appt))
            body = (
                f"Nhắc gửi phản hồi kết quả phỏng vấn.\n\n"
                f"Lịch gốc: {appt['subject']}\n"
                f"Ứng viên: {recipients or '(chưa rõ email)'}"
            )
            try:
                outlook.create_appointment(
                    subj_var.get().strip(), start, duration_minutes=30,
                    body=body, reminder_minutes=15)
            except Exception as exc:        # noqa: BLE001
                messagebox.showerror(
                    "Lỗi tạo lịch", f"Không tạo được lịch nhắc:\n{exc}",
                    parent=dlg)
                return
            dlg.destroy()
            messagebox.showinfo(
                "Đã đặt lịch",
                f"Đã tạo lời nhắc lúc {start.strftime('%d/%m/%Y %H:%M')} ✅")

        ttk.Button(actions, text="Tạo lịch nhắc", bootstyle="primary",
                   command=do_create).pack(side="left", ipadx=10, ipady=3)
        ttk.Button(actions, text="Hủy", bootstyle="secondary-outline",
                   command=dlg.destroy).pack(side="left", padx=(10, 0), ipady=3)
