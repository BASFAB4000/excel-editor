"""
Pydantic Modelle für den Excel Editor.
Hier werden alle Datenstrukturen definiert.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional, Union
from pydantic import BaseModel, field_validator


class ExcelReadConfig(BaseModel):
    """Konfiguration zum Einlesen der Excel-Datei."""

    file_path: Path
    sheet_name: Optional[str] = None  # None = aktives Sheet
    header_row: int = 1              # Zeile mit den Spaltenüberschriften (1-basiert)

    @field_validator("file_path")
    @classmethod
    def file_must_exist(cls, v: Path) -> Path:
        if not v.exists():
            raise ValueError(f"Datei nicht gefunden: {v}")
        if v.suffix not in (".xlsx", ".xlsm", ".xltx", ".xltm"):
            raise ValueError(f"Kein gültiges Excel-Format: {v.suffix}")
        return v


class CellInfo(BaseModel):
    """Repräsentiert eine einzelne Zelle mit Wert und Formatierung."""

    row: int
    column: int
    column_letter: str
    value: Optional[Any] = None
    bg_color: Optional[str] = None
    font_bold: bool = False
    font_color: Optional[str] = None


class RowData(BaseModel):
    """Repräsentiert eine Zeile mit allen Zellen."""

    row_index: int
    cells: List[CellInfo]

    def get_value(self, column_index: int) -> Optional[Any]:
        """Gibt den Wert einer Zelle anhand des Spaltenindex zurück."""
        for cell in self.cells:
            if cell.column == column_index:
                return cell.value
        return None


class SheetInfo(BaseModel):
    """Metadaten eines Sheets."""

    name: str
    max_row: int
    max_column: int
    headers: Dict[int, str]  # Spaltenindex -> Spaltenname
    header_row: int = 1       # Zeile in der die Header stehen
