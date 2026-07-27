"""Gộp / Tách PDF (stub — chưa gắn logic) — bản PySide6."""
from app_qt import widgets
from app_qt.base_tool import BaseTool


class PdfTool(BaseTool):
    name = "Merge / Split PDF"
    description = "Merge multiple PDFs into one, or split one PDF into pages."
    icon = "📄"
    category = "Files & Documents"
    order = 20
    action_label = "Process PDF"
    show_on_home = False

    def build_body(self, parent):
        widgets.dropdown(parent, "Mode", ["Merge PDFs", "Split by page"])
        widgets.file_row(parent, "PDF file / folder", mode="file")
        widgets.file_row(parent, "Save to", mode="folder")
