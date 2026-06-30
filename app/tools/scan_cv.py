"""Quét CV bằng AI và đổi tên file CV hàng loạt.

Tính năng đổi tên:
  • Quét thư mục → tìm file PDF/Word
  • Tự động trích xuất tên ứng viên từ tên file
  • Hiển thị bảng xem trước, cho phép chỉnh tên trước khi đổi
  • Đổi tên theo format: {prefixcode}_{startcode}_{TenUngVien}.ext
    start code tăng dần cho từng file theo thứ tự sắp xếp
"""
import os
import re
import unicodedata
from pathlib import Path
from tkinter import messagebox

import tkinter as tk
import ttkbootstrap as ttk

from app.core import config
from app.core.base_tool import BaseTool
from app.ui import widgets, theme

_CV_EXTENSIONS = {".pdf", ".doc", ".docx"}

SECTION = "scan_cv"
DEFAULTS = {
    "folder":         "",
    "prefix":         "",
    "start":          "01",
    "noise_keywords": (
        "cv\n"
        "c.v\n"
        "resume\n"
        "hồ sơ\n"
        "ứng tuyển\n"
        "ứng viên\n"
        "phỏng vấn\n"
        "đơn xin việc"
    ),
}

_NOISE_NUMBER = re.compile(r"^(?:20\d{2}|\d+)$")


def _ascii_fold(s: str) -> str:
    s = s.replace("đ", "d").replace("Đ", "D")
    s = s.replace("ư", "u").replace("Ư", "U")
    s = s.replace("ơ", "o").replace("Ơ", "O")
    nfd = unicodedata.normalize("NFD", s)
    return "".join(c for c in nfd if ord(c) < 128)


def _parse_noise(raw: str) -> list[str]:
    result = []
    for line in raw.splitlines():
        for kw in line.split(","):
            kw = kw.strip()
            if kw:
                result.append(_ascii_fold(kw).lower())
    return result


def _title_vn(s: str) -> str:
    """Viết hoa chữ cái đầu mỗi từ, giữ nguyên dấu tiếng Việt."""
    return " ".join(w.capitalize() for w in s.split())


def _extract_name(stem: str, noise_list: list[str]) -> str:
    stem = unicodedata.normalize("NFC", stem)
    stem = re.sub(r"[\[\(][^\]\)]*[\]\)]", " ", stem)
    stem = re.sub(r"[_\-|/\\]+", " ", stem)

    words  = stem.split()
    folded = [_ascii_fold(w).lower() for w in words]

    max_win = max((len(kw.split()) for kw in noise_list), default=1)
    max_win = min(max_win, 4)

    keep = [True] * len(words)
    i = 0
    while i < len(words):
        removed = False
        for size in range(min(max_win, len(words) - i), 0, -1):
            chunk = " ".join(folded[i : i + size])
            if chunk in noise_list or _NOISE_NUMBER.match(chunk):
                for j in range(i, i + size):
                    keep[j] = False
                i += size
                removed = True
                break
        if not removed:
            i += 1

    raw = " ".join(w for w, k in zip(words, keep) if k)
    return _title_vn(raw)


def _build_filename(name: str, prefix: str, code: str, suffix: str) -> str:
    cv_code = (prefix + code).strip()   # prefix và startcode viết liền, không có _
    parts = [p for p in (cv_code, name) if p]
    return "_".join(parts) + suffix


def _seq_code(start_str: str, index: int) -> str:
    """Tính mã thứ tự: start_str='01', index=2 → '03'. Giữ nguyên số chữ số."""
    pad = len(start_str) if start_str.isdigit() else 2
    try:
        n = int(start_str) + index
    except ValueError:
        n = 1 + index
    return str(n).zfill(pad)


class ScanCvTool(BaseTool):
    name = "Quét CV (AI)"
    description = "Đổi tên hàng loạt file CV và đọc CV bằng AI để xuất bảng tổng hợp."
    icon = "🤖"
    category = "Trí tuệ nhân tạo"
    order = 10
    action_label = "Đổi tên file CV"
    action_style = "success"

    def build_body(self, parent):
        self._parent = parent
        cfg = config.load(SECTION, DEFAULTS)

        # ---- Thư mục ----
        widgets.section_label(parent, "Thư mục chứa CV")
        self.var_folder = widgets.file_row(parent, "Thư mục", mode="folder")
        if cfg["folder"]:
            self.var_folder.set(cfg["folder"])

        # ---- Mã CV — hai ô cạnh nhau ----
        widgets.section_label(parent, "Mã CV")

        code_row = tk.Frame(parent, bg=theme.CARD_BG)
        code_row.pack(fill="x", pady=(0, 6))

        col_l = tk.Frame(code_row, bg=theme.CARD_BG)
        col_l.pack(side="left", fill="x", expand=True, padx=(0, 12))
        tk.Label(
            col_l, text="Prefix code (4 số)",
            bg=theme.CARD_BG, fg=theme.TEXT,
            font=(theme.FONT_FAMILY, 9),
        ).pack(anchor="w", pady=(0, 4))
        self.var_prefix = tk.StringVar(value=cfg["prefix"])
        ttk.Entry(col_l, textvariable=self.var_prefix).pack(fill="x", ipady=4)

        col_r = tk.Frame(code_row, bg=theme.CARD_BG)
        col_r.pack(side="left", fill="x", expand=True)
        tk.Label(
            col_r, text="Start code (2 số, tăng dần)",
            bg=theme.CARD_BG, fg=theme.TEXT,
            font=(theme.FONT_FAMILY, 9),
        ).pack(anchor="w", pady=(0, 4))
        self.var_start = tk.StringVar(value=cfg["start"])
        ttk.Entry(col_r, textvariable=self.var_start).pack(fill="x", ipady=4)

        widgets.hint(
            parent,
            "Ví dụ: prefix=2506, start=01  →  250601_Nguyen Van A.pdf, "
            "250602_Tran Thi B.pdf, …",
        )

        # ---- Từ nhiễu ----
        widgets.section_label(parent, "Từ cần xóa khi trích tên ứng viên")
        self.noise_box = widgets.text_area(
            parent,
            "Mỗi từ / cụm từ một dòng (hoặc cách nhau bởi dấu phẩy):",
            value=cfg["noise_keywords"],
            height=6,
        )

        btn_row = tk.Frame(parent, bg=theme.CARD_BG)
        btn_row.pack(fill="x", pady=(4, 0))
        ttk.Button(
            btn_row, text="💾 Lưu cấu hình", bootstyle="secondary-outline",
            command=self._save_config,
        ).pack(side="left")

        # ---- AI (placeholder) ----
        widgets.section_label(parent, "Trích xuất AI (sắp ra mắt)")
        self.var_output = widgets.file_row(parent, "Xuất bảng tổng hợp ra", mode="save")
        widgets.checkbox(parent, "Họ tên, Email, Số điện thoại")
        widgets.checkbox(parent, "Số năm kinh nghiệm")
        widgets.checkbox(parent, "Kỹ năng & Học vấn")
        widgets.hint(
            parent,
            "⚠️ Tính năng trích xuất AI dùng Claude API — cần cấu hình API key trước khi chạy.",
        )

    # ----------------------------------------------------------------- config

    def _collect(self):
        return {
            "folder":         self.var_folder.get().strip(),
            "prefix":         self.var_prefix.get().strip(),
            "start":          self.var_start.get().strip(),
            "noise_keywords": self.noise_box.get("1.0", "end-1c"),
        }

    def _save_config(self):
        config.save(SECTION, self._collect())
        messagebox.showinfo("Đã lưu", "Đã lưu cấu hình ✅")

    # ----------------------------------------------------------------- hành động

    def run(self):
        folder = self.var_folder.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("Thiếu thư mục", "Vui lòng chọn thư mục chứa CV.")
            return

        prefix    = self.var_prefix.get().strip()
        start_str = self.var_start.get().strip() or "01"
        noise_raw = self.noise_box.get("1.0", "end-1c")
        noise     = _parse_noise(noise_raw)

        config.save(SECTION, self._collect())

        files = sorted(
            p for p in Path(folder).iterdir()
            if p.is_file() and p.suffix.lower() in _CV_EXTENSIONS
        )
        if not files:
            messagebox.showinfo(
                "Không có file",
                "Không tìm thấy file PDF/DOC/DOCX trong thư mục đã chọn.",
            )
            return

        self._open_preview(files, prefix, start_str, noise)

    # --------------------------------------------------------------- hộp thoại

    def _open_preview(
        self, files: list, prefix: str, start_str: str, noise: list[str]
    ):
        dlg = tk.Toplevel(self._parent.winfo_toplevel())
        dlg.title("Xem trước & xác nhận đổi tên")
        dlg.configure(bg=theme.CONTENT_BG)
        dlg.geometry("1020x660")
        dlg.transient(self._parent.winfo_toplevel())
        dlg.grab_set()

        tk.Label(
            dlg,
            text=f"Tìm thấy {len(files)} file  —  double-click dòng để chỉnh tên ứng viên",
            bg=theme.CONTENT_BG, fg=theme.TEXT,
            font=(theme.FONT_FAMILY, 13, "bold"),
        ).pack(anchor="w", padx=22, pady=(16, 2))
        tk.Label(
            dlg,
            text="Tên ứng viên tự động trích từ tên file gốc. "
                 "Double-click để chỉnh, cột 'Tên file mới' cập nhật ngay.",
            bg=theme.CONTENT_BG, fg=theme.MUTED,
            font=(theme.FONT_FAMILY, 9),
        ).pack(anchor="w", padx=22, pady=(0, 10))

        # ---- Treeview ----
        frm = tk.Frame(dlg, bg=theme.CONTENT_BG)
        frm.pack(fill="both", expand=True, padx=22)

        cols = ("original", "cname", "newname")
        tree = ttk.Treeview(frm, columns=cols, show="headings", height=22)
        tree.heading("original", text="Tên file gốc")
        tree.heading("cname",    text="Tên ứng viên  (double-click để sửa)")
        tree.heading("newname",  text="Tên file mới (xem trước)")
        tree.column("original", width=290, anchor="w", stretch=False)
        tree.column("cname",    width=230, anchor="w", stretch=False)
        tree.column("newname",  width=420, anchor="w")

        vsb = ttk.Scrollbar(frm, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        tree.pack(side="left", fill="both", expand=True)

        _path_map: dict[str, Path] = {}
        _code_map: dict[str, str]  = {}   # iid → mã thứ tự cố định của file đó

        for idx, p in enumerate(files):
            code = _seq_code(start_str, idx)
            name = _extract_name(p.stem, noise)
            new  = _build_filename(name, prefix, code, p.suffix)
            iid  = tree.insert("", "end", values=(p.name, name, new))
            _path_map[iid] = p
            _code_map[iid] = code

        # ---- Popup edit khi double-click ----
        def _on_double_click(event):
            region = tree.identify_region(event.x, event.y)
            iid    = tree.identify_row(event.y)
            if region != "cell" or not iid:
                return

            current = tree.set(iid, "cname")
            popup = tk.Toplevel(dlg)
            popup.title("Chỉnh tên ứng viên")
            popup.configure(bg=theme.CONTENT_BG)
            popup.resizable(False, False)
            popup.transient(dlg)
            popup.grab_set()
            popup.geometry(f"420x130+{event.x_root}+{event.y_root}")

            tk.Label(
                popup, text="Tên ứng viên:",
                bg=theme.CONTENT_BG, fg=theme.TEXT,
                font=(theme.FONT_FAMILY, 9),
            ).pack(anchor="w", padx=14, pady=(14, 4))

            var = tk.StringVar(value=current)
            entry = ttk.Entry(popup, textvariable=var, font=(theme.FONT_FAMILY, 10))
            entry.pack(fill="x", padx=14, ipady=3)
            entry.select_range(0, "end")
            entry.focus_set()

            def _save(_=None):
                new_name = var.get().strip()
                if new_name:
                    new_file = _build_filename(
                        new_name, prefix, _code_map[iid], _path_map[iid].suffix,
                    )
                    tree.set(iid, "cname",   new_name)
                    tree.set(iid, "newname", new_file)
                popup.destroy()

            entry.bind("<Return>", _save)
            entry.bind("<Escape>", lambda _: popup.destroy())

            btn_row = tk.Frame(popup, bg=theme.CONTENT_BG)
            btn_row.pack(fill="x", padx=14, pady=(8, 0))
            ttk.Button(
                btn_row, text="OK", bootstyle="primary", command=_save,
            ).pack(side="left", ipadx=10, ipady=3)
            ttk.Button(
                btn_row, text="Hủy", bootstyle="secondary-outline",
                command=popup.destroy,
            ).pack(side="left", padx=(8, 0), ipady=3)

        tree.bind("<Double-1>", _on_double_click)

        # ---- Nút hành động ----
        acts = tk.Frame(dlg, bg=theme.CONTENT_BG)
        acts.pack(fill="x", padx=22, pady=14)

        def do_rename():
            errors  = []
            renamed = 0
            skipped = 0
            for iid in tree.get_children():
                path     = _path_map[iid]
                cname    = tree.set(iid, "cname").strip()
                if not cname:
                    skipped += 1
                    continue
                new_name = _build_filename(cname, prefix, _code_map[iid], path.suffix)
                new_path = path.parent / new_name
                if new_path == path:
                    skipped += 1
                    continue
                try:
                    path.rename(new_path)
                    renamed += 1
                except Exception as exc:
                    errors.append(f"{path.name}: {exc}")

            dlg.destroy()
            msg = f"Đã đổi tên {renamed} file."
            if skipped:
                msg += f"\n{skipped} file bỏ qua (tên không thay đổi)."
            if errors:
                msg += "\n\nLỗi:\n" + "\n".join(errors)
                messagebox.showwarning("Hoàn thành (có lỗi)", msg)
            else:
                messagebox.showinfo("Hoàn thành ✅", msg)

        ttk.Button(
            acts, text="✅ Đổi tên tất cả",
            bootstyle="success", command=do_rename,
        ).pack(side="left", ipadx=14, ipady=5)
        ttk.Button(
            acts, text="Hủy",
            bootstyle="secondary-outline", command=dlg.destroy,
        ).pack(side="left", padx=(10, 0), ipady=5)
