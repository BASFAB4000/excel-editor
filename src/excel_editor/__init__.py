# This file initializes the excel_editor package and can be used to define package-level variables or imports.
"""
excel_editor – RISE Planungsexcel Editor

Öffentliche API:
    from excel_editor import ExcelEditor, ExcelReadConfig
"""

from .editor import ExcelEditor
from .models import ExcelReadConfig, CellInfo, RowData, SheetInfo

__all__ = [
    "ExcelEditor",
    "ExcelReadConfig",
    "CellInfo",
    "RowData",
    "SheetInfo",
]