"""Bản đồ emoji (icon cũ của tool) → tên icon line mới (bộ SVG trong assets/icons).

Nhờ vậy không phải sửa từng file tool: tool vẫn khai báo icon = "🤖"… còn giao
diện tra ra icon line tương ứng để vẽ cho đẹp & đồng bộ.
"""
EMOJI_ICON = {
    "🏠": "home",
    "🧰": "package",
    "⚙": "settings",
    "📊": "table",
    "💰": "scissors",
    "🏆": "award",
    "📇": "idcard",
    "🤖": "sparkles",
    "📝": "file-text",
    "📄": "files",
    "🙋": "users",
    "🏢": "building",
    "💼": "briefcase",
    "📋": "clipboard",
    "🧹": "eraser",
    "📧": "mail",
    "🔔": "bell",
}


def name_for(emoji, default="file"):
    return EMOJI_ICON.get(emoji, default)
