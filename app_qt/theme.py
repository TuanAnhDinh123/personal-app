"""Tokens màu/khoảng cách + hàm nạp stylesheet QSS cho app PySide6.

Toàn bộ "CSS" của app nằm ở `theme.qss` (cùng thư mục). File này chỉ giữ:
  * các hằng số màu (để code Python cần màu — ví dụ vẽ shadow, chip icon — dùng
    chung một nguồn với QSS),
  * hàm build_stylesheet() nạp theme.qss và thay các biến {{--token}} bằng màu
    thật (QSS không có biến như CSS thật, nên ta tự thay bằng chuỗi).

Muốn đổi toàn bộ giao diện: sửa palette bên dưới HOẶC sửa thẳng theme.qss.
"""
import os

# ------------------------------------------------------------------ Palette
# Phong cách: dashboard SÁNG kiểu SaaS + sidebar TỐI làm điểm nhấn, accent indigo.
PALETTE = {
    # nền (vùng nội dung sáng)
    "--app-bg":        "#f4f6fb",   # nền toàn cửa sổ
    "--content-bg":    "#f4f6fb",   # vùng nội dung
    "--card-bg":       "#ffffff",   # thẻ trắng
    "--card-bg-hover": "#f8faff",
    "--input-bg":      "#ffffff",   # ô nhập trắng (phân biệt bằng viền)
    "--border":        "#e7ebf3",   # viền mảnh
    "--border-strong": "#d3dbe8",

    # chữ trên nền sáng
    "--text":          "#1f2735",   # chữ chính
    "--text-muted":    "#64748b",   # chữ phụ
    "--text-faint":    "#94a1b5",   # chữ rất nhạt

    # sidebar TỐI (điểm nhấn) — có bộ chữ riêng vì nền tối
    "--sidebar-bg":       "#0f1420",
    "--sidebar-text":     "#c4ccdb",
    "--sidebar-muted":    "#7b8798",
    "--sidebar-title":    "#5d6a7e",
    "--sidebar-hover":    "#252d3f",
    "--sidebar-border":   "#1d2534",

    # nhấn
    "--accent":        "#6366f1",   # indigo
    "--accent-hover":  "#5457e6",
    "--accent-press":  "#4a4dd0",
    "--accent-soft":   "#eef0fe",   # nền indigo nhạt (chip/hover trên nền sáng)
    "--row-hover":     "#f5f6fe",   # nền hover 1 dòng bảng

    # trạng thái nút
    "--success":       "#22c55e",
    "--success-hover": "#16a34a",
    "--danger":        "#ef4444",
    "--danger-hover":  "#dc2626",
    "--warning":       "#f59e0b",
    "--warning-hover": "#d97706",
    "--info":          "#3b82f6",
    "--info-hover":    "#2563eb",
    "--neutral":       "#ffffff",   # nút phụ: trắng viền (như "Xuất JSON" ở mẫu)
    "--neutral-hover": "#f1f4fa",

    # màu nhóm (chip icon)
    "--cat-files":     "#3b82f6",
    "--cat-data":      "#10b981",
    "--cat-office":    "#f59e0b",
    "--cat-hr":        "#8b5cf6",
    "--cat-default":   "#64748b",
}

# Truy cập nhanh trong code Python (không qua QSS).
ACCENT = PALETTE["--accent"]
CARD_BG = PALETTE["--card-bg"]
APP_BG = PALETTE["--app-bg"]
TEXT = PALETTE["--text"]
TEXT_MUTED = PALETTE["--text-muted"]
BORDER = PALETTE["--border"]
BORDER_STRONG = PALETTE["--border-strong"]

FONT_FAMILY = "Segoe UI"

CATEGORY_COLORS = {
    "Tệp & Tài liệu": PALETTE["--cat-files"],
    "Dữ liệu":        PALETTE["--cat-data"],
    "Văn phòng":      PALETTE["--cat-office"],
    "Tuyển dụng":     PALETTE["--cat-hr"],
    "Demo":           PALETTE["--accent"],
}


def category_color(category):
    return CATEGORY_COLORS.get(category, PALETTE["--cat-default"])


def asset(name):
    """Đường dẫn asset dạng url QSS (forward slash) — dùng cho image: url(...)."""
    p = os.path.join(os.path.dirname(__file__), "assets", name)
    return p.replace("\\", "/")


def build_stylesheet():
    """Đọc theme.qss, thay {{--token}} bằng màu thật, trả về chuỗi QSS."""
    path = os.path.join(os.path.dirname(__file__), "theme.qss")
    with open(path, encoding="utf-8") as f:
        qss = f.read()
    for token, value in PALETTE.items():
        qss = qss.replace("{{" + token + "}}", value)
    qss = qss.replace("{{CHECK_ICON}}", asset("check.svg"))
    qss = qss.replace("{{CHEVRON_ICON}}", asset("chevron-down.svg"))
    return qss
