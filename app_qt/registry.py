"""Tự động phát hiện tool trong gói app_qt.tools.

Giống app/core/registry.py: tạo 1 file trong app_qt/tools/ với class kế thừa
BaseTool là app tự nhận, không cần sửa file nào khác.
"""
import importlib
import inspect
import pkgutil

from app_qt import tools as tools_pkg
from app_qt.base_tool import BaseTool


def discover_tools():
    """Quét gói app_qt.tools, trả về danh sách tool đã sắp xếp."""
    found = []
    for _, module_name, _ in pkgutil.iter_modules(tools_pkg.__path__):
        module = importlib.import_module(f"app_qt.tools.{module_name}")
        for _, obj in inspect.getmembers(module, inspect.isclass):
            if (
                issubclass(obj, BaseTool)
                and obj is not BaseTool
                and obj.__module__ == module.__name__
                and not obj.__name__.startswith("_")
            ):
                found.append(obj())
    found.sort(key=lambda t: (t.category, t.order, t.name))
    return found
