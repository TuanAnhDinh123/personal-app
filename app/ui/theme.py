"""Bảng màu & font dùng chung cho toàn bộ giao diện.

Đổi giao diện toàn app chỉ cần sửa ở đây.
"""

# ----- Font -----
FONT_FAMILY = "Segoe UI"
FONT_EMOJI = "Segoe UI Emoji"

# ----- Thanh bên (sidebar sáng) -----
SIDEBAR_BG = "#ffffff"
SIDEBAR_BG_HOVER = "#eef0f7"
SIDEBAR_BG_ACTIVE = "#ebedfd"
SIDEBAR_FG = "#5c6573"
SIDEBAR_FG_ACTIVE = "#2a3354"
SIDEBAR_MUTED = "#9aa2b4"

# ----- Màu nhấn chính -----
ACCENT = "#6366f1"        # indigo
ACCENT_HOVER = "#7c83f6"

# ----- Màu riêng theo nhóm (dùng cho icon chip) -----
CATEGORY_COLORS = {
    "Tệp & Tài liệu": "#3b82f6",      # xanh dương
    "Dữ liệu": "#10b981",             # xanh lá
    "Văn phòng": "#f59e0b",           # cam
}
CATEGORY_DEFAULT_COLOR = "#64748b"    # xám xanh


def category_color(category):
    return CATEGORY_COLORS.get(category, CATEGORY_DEFAULT_COLOR)


# ----- Vùng nội dung (nền sáng) -----
CONTENT_BG = "#f4f6fb"
CARD_BG = "#ffffff"
BORDER = "#e6eaf2"
TEXT = "#1e2633"
MUTED = "#8a93a3"

# ----- Theme nền của ttkbootstrap -----
BOOTSTYLE_THEME = "flatly"
