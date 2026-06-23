"""Lưu / đọc cấu hình của các tool ra một file JSON dùng chung.

Mỗi tool giữ cấu hình của mình trong một "section" riêng (thường là tên
module). Nhờ vậy phần chạy tự động lúc mở app vẫn đọc được cấu hình mà
người dùng đã lưu, kể cả khi chưa mở giao diện của tool đó.

File nằm ở:  %APPDATA%\\PersonalToolbox\\config.json  (Windows)
          ~/.config/PersonalToolbox/config.json     (Linux/macOS — lúc dev)
"""
import json
import os

_CACHE = None


def _config_path():
    base = os.environ.get("APPDATA") or os.path.join(
        os.path.expanduser("~"), ".config")
    folder = os.path.join(base, "PersonalToolbox")
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "config.json")


def _read_all():
    global _CACHE
    if _CACHE is None:
        try:
            with open(_config_path(), encoding="utf-8") as f:
                _CACHE = json.load(f)
        except (FileNotFoundError, ValueError):
            _CACHE = {}
    return _CACHE


def load(section, default=None):
    """Trả về dict cấu hình của một section (gộp lên trên `default` nếu có)."""
    data = dict(default or {})
    data.update(_read_all().get(section, {}))
    return data


def save(section, values):
    """Ghi đè cấu hình của một section rồi lưu xuống đĩa."""
    allcfg = _read_all()
    allcfg[section] = values
    with open(_config_path(), "w", encoding="utf-8") as f:
        json.dump(allcfg, f, ensure_ascii=False, indent=2)
