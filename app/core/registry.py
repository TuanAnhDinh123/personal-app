"""Tự động phát hiện và đăng ký tất cả tool trong gói app.tools.

Muốn thêm tác vụ mới: tạo 1 file trong app/tools/ với một class kế thừa
BaseTool — app sẽ tự nhận, KHÔNG cần sửa file nào khác.
"""
import importlib
import inspect
import pkgutil
import sys

from app import tools as tools_pkg
from app.core.base_tool import BaseTool


def _tool_module_names():
    """Tên các module trong app.tools — chạy được cả khi dev lẫn khi đóng gói.

    Khi chạy python main.py: module là file .py trên đĩa → iter_modules thấy.
    Khi chạy bản .exe (PyInstaller --onefile): module nằm trong archive, không
    trên đĩa, nên iter_modules theo đường dẫn trả rỗng. Lúc đó đọc thêm bảng
    module mà PyInstaller nhúng sẵn (thuộc tính `toc` của FrozenImporter).
    """
    names = {name for _, name, _ in pkgutil.iter_modules(tools_pkg.__path__)}
    if getattr(sys, "frozen", False):
        prefix = tools_pkg.__name__ + "."
        for importer in pkgutil.iter_importers(tools_pkg.__name__):
            toc = getattr(importer, "toc", None)
            if not toc:
                continue
            for mod in toc:
                if mod.startswith(prefix):
                    sub = mod[len(prefix):]
                    if sub and "." not in sub:   # chỉ lấy module con trực tiếp
                        names.add(sub)
    return sorted(names)


def discover_tools():
    """Quét gói app.tools, trả về danh sách tool đã sắp xếp."""
    found = []
    for module_name in _tool_module_names():
        module = importlib.import_module(f"app.tools.{module_name}")
        for _, obj in inspect.getmembers(module, inspect.isclass):
            is_tool = (
                issubclass(obj, BaseTool)
                and obj is not BaseTool
                and obj.__module__ == module.__name__
            )
            if is_tool:
                found.append(obj())
    found.sort(key=lambda t: (t.category, t.order, t.name))
    return found
