from .word_tools import read_docx, write_docx, edit_docx
from .excel_tools import read_xlsx, write_xlsx, edit_xlsx
from .ppt_tools import read_pptx, write_pptx, edit_pptx
from .filesystem_tools import (
    list_directory, read_file, write_file, edit_file,
    file_info, create_directory, move_file, delete_file,
)
try:
    from .pdf_processor import PDFProcessor
except ImportError:  # PyMuPDF 未安装时其余工具不受影响（server.py 会优雅降级）
    PDFProcessor = None

__all__ = [
    "read_docx", "write_docx", "edit_docx",
    "read_xlsx", "write_xlsx", "edit_xlsx",
    "read_pptx", "write_pptx", "edit_pptx",
    "list_directory", "read_file", "write_file", "edit_file",
    "file_info", "create_directory", "move_file", "delete_file",
    "PDFProcessor",
]
