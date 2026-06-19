"""Tự động phát hiện và đăng ký tất cả tool trong gói app.tools.

Muốn thêm tác vụ mới: tạo 1 file trong app/tools/ với một class kế thừa
BaseTool — app sẽ tự nhận, KHÔNG cần sửa file nào khác.
"""
import importlib
import inspect
import pkgutil

from app import tools as tools_pkg
from app.core.base_tool import BaseTool


def discover_tools():
    """Quét gói app.tools, trả về danh sách tool đã sắp xếp."""
    found = []
    for _, module_name, _ in pkgutil.iter_modules(tools_pkg.__path__):
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
