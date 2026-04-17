"""
Kernlogik des Excel Editors – xlwings-Backend.

Verwendet xlwings (Windows COM) statt openpyxl, damit Excel (inkl. AutoSave /
AutoSync für OneDrive- und SharePoint-Dateien) aktiv bleibt.

Voraussetzung: Windows + Microsoft Excel installiert.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional

import xlwings as xw

from .models import CellInfo, ExcelReadConfig, RowData, SheetInfo


# ---------------------------------------------------------------------------
# Hilfsfunktionen (privat)
# ---------------------------------------------------------------------------

def _col_to_letter(col: int) -> str:
    """Wandelt 1-basierten Spaltenindex in Excel-Buchstabe um (A, B, …, AA, …)."""
    result = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _no_val_eq(cell_val: Any, search_val: Any) -> bool:
    """
    Vergleicht einen Zellwert mit einem Suchwert.
    xlwings gibt Integer-Zellen manchmal als float zurück (z.B. 1010.0).
    Wir versuchen numerischen Vergleich, Fallback ist String-Vergleich.
    """
    if cell_val is None:
        return False
    try:
        return int(float(cell_val)) == int(float(search_val))
    except (ValueError, TypeError):
        return str(cell_val) == str(search_val)


# ---------------------------------------------------------------------------
# Öffentliche Klasse
# ---------------------------------------------------------------------------

class ExcelEditor:
    """
    Haupt-Editor-Klasse (xlwings-Backend).

    Operiert via Excel COM – AutoSave/AutoSync auf OneDrive- und SharePoint-
    Dateien bleibt vollständig aktiv, da alle Änderungen durch Excel selbst
    vorgenommen werden.

    Wenn die Datei bereits in einer laufenden Excel-Instanz geöffnet ist,
    wird diese Instanz verwendet. Andernfalls wird eine versteckte Excel-
    Instanz gestartet.

    Verwendung:
        with ExcelEditor(config) as editor:
            editor.move_row_after("1010", "1026")
            editor.save()
    """

    def __init__(self, config: ExcelReadConfig) -> None:
        self.config = config
        self._app: Optional[xw.App] = None
        self._workbook: Optional[xw.Book] = None
        self._owns_app: bool = False

        resolved = config.file_path.resolve()

        # xw.Book(path) findet eine bereits geöffnete Mappe in jeder laufenden
        # Excel-Instanz automatisch, oder öffnet sie neu in einem sichtbaren
        # Excel-Fenster. Sichtbar (visible=True) ist notwendig damit Excel
        # OneDrive-/SharePoint-Dateien öffnen kann (Auth-Dialoge).
        try:
            self._workbook = xw.Book(str(resolved))
        except Exception as e:
            raise RuntimeError(
                f"Excel-Datei konnte nicht geöffnet werden: {resolved}\n"
                f"Stelle sicher dass:\n"
                f"  • Microsoft Excel installiert ist\n"
                f"  • die Datei existiert und nicht gesperrt ist\n"
                f"Details: {e}"
            )

        self._app = self._workbook.app
        self._owns_app = False  # wir schließen die Mappe nie automatisch
        self._worksheet = self._get_worksheet()

    # ------------------------------------------------------------------
    # Sheet-Zugriff
    # ------------------------------------------------------------------

    def _get_worksheet(self) -> xw.Sheet:
        """Gibt das konfigurierte Worksheet zurück."""
        if self.config.sheet_name:
            names = [s.name for s in self._workbook.sheets]
            if self.config.sheet_name not in names:
                raise ValueError(
                    f"Sheet '{self.config.sheet_name}' nicht gefunden. "
                    f"Verfügbar: {names}"
                )
            return self._workbook.sheets[self.config.sheet_name]
        return self._workbook.sheets.active

    def get_sheet_names(self) -> List[str]:
        """Gibt alle Sheet-Namen zurück."""
        return [s.name for s in self._workbook.sheets]

    def get_sheet_info(self) -> SheetInfo:
        """Gibt Metadaten des aktuellen Sheets zurück."""
        ws = self._worksheet
        header_row = self.config.header_row
        used = ws.used_range
        max_col = used.last_cell.column
        max_row = used.last_cell.row

        # Header-Zeile als Range einlesen (ein COM-Aufruf)
        header_vals = ws.range((header_row, 1), (header_row, max_col)).value
        if not isinstance(header_vals, list):
            header_vals = [header_vals]

        headers: Dict[int, str] = {}
        for i, val in enumerate(header_vals, start=1):
            headers[i] = str(val) if val is not None else f"Spalte_{i}"

        return SheetInfo(
            name=ws.name,
            max_row=max_row,
            max_column=max_col,
            headers=headers,
            header_row=header_row,
        )

    # ------------------------------------------------------------------
    # Daten lesen
    # ------------------------------------------------------------------

    def get_rows(
        self,
        min_row: Optional[int] = None,
        max_row: Optional[int] = None,
        skip_empty: bool = True,
    ) -> List[RowData]:
        """
        Liest alle Zeilen nach der Header-Zeile aus.
        Werte werden per Block-Read eingelesen (schnell, ein COM-Aufruf).

        Args:
            min_row: Erste Zeile (Standard: header_row + 1)
            max_row: Letzte Zeile (Standard: bis Ende)
            skip_empty: Leere Zeilen überspringen
        """
        if min_row is None:
            min_row = self.config.header_row + 1

        ws = self._worksheet
        used = ws.used_range
        last_row = max_row or used.last_cell.row
        last_col = used.last_cell.column

        if min_row > last_row:
            return []

        # Werte komplett auf einmal lesen
        data = ws.range((min_row, 1), (last_row, last_col)).value

        # xlwings gibt bei einer einzelnen Zeile eine flache Liste zurück
        if last_row == min_row or not isinstance(data[0], list):
            data = [data] if data else []

        rows: List[RowData] = []
        for row_offset, row_values in enumerate(data):
            actual_row = min_row + row_offset
            if row_values is None:
                row_values = [None] * last_col
            if not isinstance(row_values, list):
                row_values = [row_values]
            if skip_empty and all(v is None for v in row_values):
                continue

            cells = [
                CellInfo(
                    row=actual_row,
                    column=col_idx,
                    column_letter=_col_to_letter(col_idx),
                    value=val,
                )
                for col_idx, val in enumerate(row_values, start=1)
            ]
            rows.append(RowData(row_index=actual_row, cells=cells))

        return rows

    def get_row(self, row_index: int) -> Optional[RowData]:
        """Gibt eine einzelne Zeile anhand des Index zurück."""
        ws = self._worksheet
        last_col = ws.used_range.last_cell.column
        values = ws.range((row_index, 1), (row_index, last_col)).value
        if not isinstance(values, list):
            values = [values]
        if all(v is None for v in values):
            return None
        cells = [
            CellInfo(
                row=row_index,
                column=col_idx,
                column_letter=_col_to_letter(col_idx),
                value=val,
            )
            for col_idx, val in enumerate(values, start=1)
        ]
        return RowData(row_index=row_index, cells=cells)

    # ------------------------------------------------------------------
    # Daten schreiben
    # ------------------------------------------------------------------

    def edit_cell(self, row: int, column: int, new_value: Any) -> None:
        """
        Bearbeitet eine einzelne Zelle.
        Da die Änderung über Excel COM erfolgt, bleibt die Formatierung
        automatisch erhalten.

        Args:
            row: Zeilenindex (1-basiert)
            column: Spaltenindex (1-basiert)
            new_value: Neuer Wert der Zelle
        """
        self._worksheet.range(row, column).value = new_value

    def edit_row(self, row_index: int, updates: Dict[int, Any]) -> None:
        """
        Bearbeitet mehrere Zellen einer Zeile auf einmal.

        Args:
            row_index: Zeilenindex (1-basiert)
            updates: Dict mit {Spaltenindex: neuer_Wert}
        """
        for col, val in updates.items():
            self.edit_cell(row=row_index, column=col, new_value=val)

    # ------------------------------------------------------------------
    # Zeile verschieben
    # ------------------------------------------------------------------

    def _find_no_column(self) -> int:
        """
        Findet den Spaltenindex der 'No'-Spalte anhand des Headers.
        Wirft ValueError wenn keine 'No'-Spalte gefunden wird.
        """
        info = self.get_sheet_info()
        for idx, name in info.headers.items():
            if name.strip().lower() == "no":
                return idx
        raise ValueError(
            "Spalte 'No' nicht im Sheet gefunden. "
            f"Verfügbare Spalten: {list(info.headers.values())}"
        )

    def _find_row_by_no(self, no_value: Any) -> int:
        """
        Sucht den Zeilenindex der Zeile mit dem angegebenen 'No'-Wert.
        Liest die gesamte No-Spalte auf einmal (ein COM-Aufruf).
        Wirft ValueError wenn nicht gefunden.
        """
        no_col = self._find_no_column()
        ws = self._worksheet
        min_row = self.config.header_row + 1
        last_row = ws.used_range.last_cell.row

        col_values = ws.range((min_row, no_col), (last_row, no_col)).value
        if not isinstance(col_values, list):
            col_values = [col_values]

        for row_offset, val in enumerate(col_values):
            if _no_val_eq(val, no_value):
                return min_row + row_offset

        raise ValueError(f"Zeile mit No='{no_value}' nicht gefunden.")

    def move_row_after(self, source_no: Any, after_no: Any) -> int:
        """
        Verschiebt die Zeile mit No=source_no direkt NACH die Zeile mit No=after_no.

        Verwendet Excel COM (Rows.Cut + Rows.Insert), damit:
          - Formatierung, Farben und Formeln vollständig erhalten bleiben
          - AutoSave / AutoSync in Excel-Online-Dateien aktiv bleibt

        new_no = int((after_no + next_below_no) / 2)

        Bricht mit ValueError ab wenn:
          - source_no oder after_no nicht existieren
          - Kein ganzzahliger Mittelwert möglich
          - Das berechnete new_no bereits als No existiert
          - after_no und source_no identisch sind

        Args:
            source_no: No-Wert der zu verschiebenden Zeile
            after_no:  No-Wert der Zeile, nach der eingefügt wird

        Returns:
            new_no: Der tatsächlich vergebene No-Wert
        """
        no_col = self._find_no_column()
        ws = self._worksheet
        last_row = ws.used_range.last_cell.row

        # --- 1. Source- und After-Zeile finden ---
        source_idx = self._find_row_by_no(source_no)
        after_idx = self._find_row_by_no(after_no)

        if source_idx == after_idx:
            raise ValueError("Quell- und Zielzeile sind identisch.")

        try:
            after_no_int = int(float(after_no))
        except (ValueError, TypeError):
            raise ValueError(f"No='{after_no}' ist kein ganzzahliger Wert.")

        # --- 2. Nächste Zeile unter after_idx finden (source überspringen) ---
        below_no_int: Optional[int] = None
        for row_idx in range(after_idx + 1, last_row + 1):
            if row_idx == source_idx:
                continue
            val = ws.range(row_idx, no_col).value
            if val is None:
                continue
            try:
                below_no_int = int(float(val))
                break
            except (ValueError, TypeError):
                continue

        # --- 3. Neues No berechnen ---
        if below_no_int is None:
            new_no = after_no_int + 10
        else:
            raw_mid = (after_no_int + below_no_int) / 2
            new_no = int(raw_mid)
            if new_no <= after_no_int:
                raise ValueError(
                    f"Kein ganzzahliger Mittelwert möglich zwischen "
                    f"No={after_no_int} und No={below_no_int} "
                    f"(Mittelwert wäre {raw_mid}). "
                    f"Bitte Nummernlücke vergrößern."
                )

        # --- 4. Prüfen ob new_no bereits existiert ---
        try:
            self._find_row_by_no(new_no)
            raise ValueError(
                f"Berechnetes No={new_no} ist bereits vergeben. "
                f"Bitte Nummernlücke zwischen {after_no_int} und "
                f"{below_no_int} vergrößern."
            )
        except ValueError as e:
            if "bereits vergeben" in str(e):
                raise

        # --- 5. No-Wert in der Quellzeile VOR dem Verschieben setzen ---
        ws.range(source_idx, no_col).value = new_no

        # --- 6. Zeile via Excel COM verschieben (Cut + Insert) ---
        # Rows(n).Cut() markiert die Zeile als "ausgeschnitten".
        # Rows(m).Insert(xlShiftDown) fügt sie vor Zeile m ein → landet nach after_idx.
        # Die Indizes beziehen sich auf den aktuellen Zustand – kein manuelles Anpassen nötig.
        ws.api.Rows(source_idx).Cut()
        ws.api.Rows(after_idx + 1).Insert(Shift=-4121)  # -4121 = xlShiftDown

        return new_no

    # Rückwärtskompatibilität
    def renumber_and_move_row(self, source_no: Any, after_no: Any) -> tuple:
        """Alias für move_row_after (Rückwärtskompatibilität)."""
        new_no = self.move_row_after(source_no, after_no)
        return new_no, False

    def move_row_by_no(self, source_no: Any, after_no: Any) -> None:
        """Alias für move_row_after (Rückwärtskompatibilität)."""
        self.move_row_after(source_no, after_no)

    # ------------------------------------------------------------------
    # Speichern
    # ------------------------------------------------------------------

    def save(self, output_path: Optional[Path] = None) -> Path:
        """
        Speichert die Datei über Excel (AutoSave bleibt aktiv).

        Args:
            output_path: Zielpfad für eine Kopie (Standard: Original speichern)

        Returns:
            Pfad zur gespeicherten Datei
        """
        if output_path is not None:
            # SaveCopyAs speichert eine Kopie ohne die aktive Mappe zu schließen
            self._workbook.api.SaveCopyAs(str(output_path.resolve()))
            return output_path
        else:
            self._workbook.save()
            return self.config.file_path

    def close(self) -> None:
        """Schließt nichts automatisch – die Mappe bleibt in Excel offen
        damit AutoSave weiterläuft."""
        pass

    # Context Manager Support
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()