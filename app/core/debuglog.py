"""Ghi lỗi ra file log dùng chung để traceback không bị mất tăm.

Bản phát hành chạy bằng `pythonw.exe` (và cả file .exe sau khi build) đều không
có console, nên traceback in ra stderr không ai thấy. Mọi chỗ bắt lỗi rồi bỏ
qua nên gọi `exception()` để còn dấu vết đọc lại được, và `install()` gắn hook
cho những lỗi không ai bắt (kể cả trong thread) thay vì để app tắt im lặng.

File log:  %APPDATA%\\PersonalToolbox\\debug.log   (đầy 1 MB thì đổi tên .1)
"""
import datetime
import os
import sys
import threading
import traceback

_MAX_BYTES = 1_000_000


def log_path():
    base = os.environ.get("APPDATA") or os.path.join(
        os.path.expanduser("~"), ".config")
    folder = os.path.join(base, "PersonalToolbox")
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "debug.log")


def write(text):
    """Ghi một mục vào log kèm mốc thời gian; in ra console nếu đang có."""
    line = f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {text}"
    try:
        path = log_path()
        if os.path.exists(path) and os.path.getsize(path) > _MAX_BYTES:
            os.replace(path, path + ".1")
        with open(path, "a", encoding="utf-8") as fh:
            fh.write(line + "\n")
    except OSError:
        pass
    # pythonw.exe không có stderr (sys.stderr là None) -> phải kiểm tra trước.
    try:
        if sys.stderr is not None:
            print(line, file=sys.stderr, flush=True)
    except (OSError, ValueError):
        pass


def exception(context, exc=None):
    """Ghi traceback của lỗi đang xử lý; trả về chuỗi traceback để hiện lên UI."""
    if exc is not None:
        text = "".join(traceback.format_exception(
            type(exc), exc, exc.__traceback__))
    else:
        text = traceback.format_exc()
    write(f"{context}\n{text.rstrip()}")
    return text.rstrip()


def install():
    """Gắn hook ghi log cho lỗi không ai bắt, ở luồng chính và luồng con.

    PySide6 gọi sys.excepthook khi một slot (hàm xử lý click) ném lỗi, nên hook
    này bắt được cả lỗi từ các nút bấm trên giao diện.
    """
    write(f"--- app start (pid {os.getpid()}, {sys.executable}) ---")

    prev_hook = sys.excepthook

    def hook(exc_type, exc, tb):
        write("UNHANDLED EXCEPTION\n" + "".join(
            traceback.format_exception(exc_type, exc, tb)).rstrip())
        prev_hook(exc_type, exc, tb)

    sys.excepthook = hook

    def thread_hook(args):
        write(f"UNHANDLED EXCEPTION in thread {args.thread!r}\n" + "".join(
            traceback.format_exception(
                args.exc_type, args.exc_value, args.exc_traceback)).rstrip())

    threading.excepthook = thread_hook

    # faulthandler: crash tầng C (COM của Outlook, Qt) cũng để lại dấu vết.
    try:
        import faulthandler
        faulthandler.enable(file=open(log_path(), "a", encoding="utf-8"))
    except Exception:
        pass
