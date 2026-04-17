"""
CLI-Einstiegspunkt für den Excel Editor.

Nicht-interaktiv (z.B. von einer agentic AI aufgerufen):
    excel-editor --file <pfad> --sheet <name> --info

Interaktiv (Nutzer wird Schritt für Schritt gefragt):
    excel-editor
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pydantic import ValidationError

from .models import ExcelReadConfig
from .editor import ExcelEditor


# ---------------------------------------------------------------------------
# Argument Parser
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-editor",
        description=(
            "RISE Planungsexcel Editor.\n"
            "Ohne Argumente: interaktiver Modus (Nutzer wird gefragt).\n"
            "Mit --file: nicht-interaktiv, z.B. für agentic AI."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--file", "-f",
        type=lambda p: Path(p).expanduser().resolve(),
        default=None,
        help="Pfad zur Excel-Datei (optional – wird sonst abgefragt)",
    )
    parser.add_argument(
        "--sheet", "-s",
        type=str,
        default=None,
        help="Name des Sheets (optional – wird sonst abgefragt)",
    )
    parser.add_argument(
        "--list-sheets",
        action="store_true",
        help="Zeigt alle verfügbaren Sheets an und beendet",
    )
    parser.add_argument(
        "--info",
        action="store_true",
        help="Zeigt Metadaten des Sheets (Spalten, Dimensionen)",
    )
    parser.add_argument(
        "--rows",
        type=int,
        default=5,
        metavar="N",
        help="Anzahl der anzuzeigenden Zeilen (Standard: 5)",
    )
    parser.add_argument(
        "--header-row",
        type=int,
        default=None,
        metavar="N",
        help="Zeile mit den Spaltenüberschriften, 1-basiert (Standard: automatisch erkannt)",
    )
    # --- Zeile verschieben ---
    parser.add_argument(
        "--move-from",
        type=str,
        default=None,
        metavar="NO",
        help="'No'-Wert der Zeile, die verschoben werden soll",
    )
    parser.add_argument(
        "--move-after",
        type=str,
        default=None,
        metavar="NO",
        help="'No'-Wert der Zeile, NACH der eingefügt werden soll",
    )
    parser.add_argument(
        "--save",
        action="store_true",
        help="Datei nach der Bearbeitung speichern (überschreibt Original)",
    )
    parser.add_argument(
        "--output", "-o",
        type=lambda p: Path(p).expanduser().resolve(),
        default=None,
        help="Speichern unter neuem Pfad statt das Original zu überschreiben",
    )
    return parser


# ---------------------------------------------------------------------------
# Interaktive Eingabe
# ---------------------------------------------------------------------------

def _ask_file_path() -> Path:
    """Fragt den Nutzer nach dem Pfad zur Excel-Datei."""
    while True:
        raw = input("\nPfad zur Excel-Datei: ").strip().strip('"').strip("'")
        if not raw:
            print("  [!] Kein Pfad eingegeben. Bitte erneut versuchen.")
            continue
        path = Path(raw).expanduser().resolve()
        if not path.exists():
            print(f"  [!] Datei nicht gefunden: {path}")
            retry = input("  Erneut versuchen? [J/n]: ").strip().lower()
            if retry in ("n", "nein"):
                sys.exit(0)
            continue
        return path


def _ask_sheet(editor: ExcelEditor) -> str | None:
    """Zeigt verfügbare Sheets und fragt den Nutzer welches bearbeitet werden soll."""
    sheets = editor.get_sheet_names()
    print("\nVerfügbare Sheets:")
    for i, name in enumerate(sheets, start=1):
        print(f"  [{i}] {name}")

    if len(sheets) == 1:
        print(f"  → Nur ein Sheet vorhanden, verwende: '{sheets[0]}'")
        return sheets[0]

    while True:
        raw = input("\nSheet-Name oder Nummer eingeben (Enter = aktives Sheet): ").strip()
        if not raw:
            return None  # aktives Sheet
        # Nummer eingegeben?
        if raw.isdigit():
            idx = int(raw) - 1
            if 0 <= idx < len(sheets):
                return sheets[idx]
            print(f"  [!] Ungültige Nummer. Bitte 1–{len(sheets)} eingeben.")
            continue
        # Name eingegeben?
        if raw in sheets:
            return raw
        print(f"  [!] Sheet '{raw}' nicht gefunden.")


# ---------------------------------------------------------------------------
# Ausgabe-Funktionen
# ---------------------------------------------------------------------------

def print_sheet_info(editor: ExcelEditor) -> None:
    """Gibt Metadaten des Sheets aus."""
    info = editor.get_sheet_info()
    print(f"\nSheet:       {info.name}")
    print(f"Max Zeilen:  {info.max_row}")
    print(f"Max Spalten: {info.max_column}")
    print(f"\nSpalten:")
    for idx, name in info.headers.items():
        print(f"  [{idx:>3}] {name}")


def print_rows(editor: ExcelEditor, n: int) -> None:
    """Gibt die ersten n Datenzeilen aus."""
    info = editor.get_sheet_info()
    rows = editor.get_rows()

    print(f"\nErste {n} Datenzeilen:\n")
    for row_data in rows[:n]:
        print(f"  Zeile {row_data.row_index}:")
        for cell in row_data.cells:
            if cell.value is None:
                continue
            col_name = info.headers.get(cell.column, "?")
            print(
                f"    [{cell.column_letter}] "
                f"{col_name:<25} | "
                f"Wert: {str(cell.value):<20} | "
                f"BG: {cell.bg_color or 'keine':<12} | "
                f"Fett: {cell.font_bold}"
            )


# ---------------------------------------------------------------------------
# Auto-Detect Header-Zeile
# ---------------------------------------------------------------------------

def _detect_header_row(editor: ExcelEditor, max_scan: int = 20) -> int:
    """
    Erkennt automatisch die Header-Zeile: erste Zeile mit >= 3 gefüllten Zellen
    die überwiegend Strings enthält.
    Gibt 1 zurück wenn nichts gefunden wird.
    """
    try:
        ws = editor._worksheet
        last_col = ws.used_range.last_cell.column
        actual_max = min(max_scan, ws.used_range.last_cell.row)
        data = ws.range((1, 1), (actual_max, last_col)).value
        if not data:
            return 1
        # xlwings gibt bei einer einzelnen Zeile eine flache Liste zurück
        if not isinstance(data[0], list):
            data = [data]
        for row_idx, row_values in enumerate(data, start=1):
            if not row_values:
                continue
            values = [v for v in row_values if v is not None]
            if len(values) >= 3:
                string_count = sum(1 for v in values if isinstance(v, str))
                if string_count >= len(values) * 0.6:
                    return row_idx
    except Exception:
        pass
    return 1


# ---------------------------------------------------------------------------
# Einstiegspunkt
# ---------------------------------------------------------------------------

def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    # --- Pfad bestimmen: Argument oder interaktive Abfrage ---
    if args.file is not None:
        file_path = args.file
    else:
        print("=" * 55)
        print("  RISE Planungsexcel Editor – Interaktiver Modus")
        print("=" * 55)
        file_path = _ask_file_path()

    # Pydantic validiert Pfad und Dateiendung
    try:
        config = ExcelReadConfig(file_path=file_path)
    except (ValueError, ValidationError) as e:
        print(f"[FEHLER] {e}", file=sys.stderr)
        sys.exit(1)

    with ExcelEditor(config) as editor:

        # --list-sheets: nur Sheets ausgeben
        if args.list_sheets:
            print("Verfügbare Sheets:")
            for name in editor.get_sheet_names():
                print(f"  - {name}")
            return

        # --- Sheet bestimmen: Argument oder interaktive Abfrage ---
        if args.sheet is not None:
            try:
                config.sheet_name = args.sheet
                editor._worksheet = editor._get_worksheet()
            except ValueError as e:
                print(f"[FEHLER] {e}", file=sys.stderr)
                sys.exit(1)
        elif args.file is None:
            sheet_name = _ask_sheet(editor)
            if sheet_name:
                config.sheet_name = sheet_name
                editor._worksheet = editor._get_worksheet()

        # --- Header-Zeile bestimmen ---
        if args.header_row is not None:
            config.header_row = args.header_row
        else:
            # Auto-Detect: erste Zeile mit >= 3 gefüllten Zellen
            detected = _detect_header_row(editor)
            config.header_row = detected
            if detected != 1:
                print(f"  [i] Header-Zeile automatisch erkannt: Zeile {detected}")

        # Ausgabe
        if args.info or args.file is None:
            print_sheet_info(editor)

        # --- Zeile verschieben ---
        if args.move_from or args.move_after:
            if not args.move_from or not args.move_after:
                print("[FEHLER] --move-from und --move-after müssen zusammen angegeben werden.", file=sys.stderr)
                sys.exit(1)
            _do_move_row(editor, args)
        else:
            # Nur anzeigen wenn keine Aktion angegeben
            print_rows(editor, args.rows)


def _do_move_row(editor: ExcelEditor, args) -> None:
    """Führt die Move-Row-Aktion aus und fragt ggf. nach dem Speichern."""
    source_no = args.move_from
    after_no  = args.move_after

    # Source-Zeile prüfen
    try:
        source_idx = editor._find_row_by_no(source_no)
        after_idx  = editor._find_row_by_no(after_no)
    except ValueError as e:
        print(f"[FEHLER] {e}", file=sys.stderr)
        sys.exit(1)

    print(f"\nVerschiebe No={source_no!r} (Excel-Zeile {source_idx})")
    print(f"  einfügen nach No={after_no!r} (Excel-Zeile {after_idx})")

    # Verschieben
    try:
        new_no = editor.move_row_after(source_no, after_no)
    except ValueError as e:
        print(f"[FEHLER] {e}", file=sys.stderr)
        sys.exit(1)

    print(f"  [OK] Verschoben. Neues No={new_no}")

    # Speichern
    should_save = args.save or args.output is not None
    if not should_save and args.file is not None:
        # Interaktiv nachfragen wenn keine --save Flag gesetzt
        ans = input("\nDatei speichern? [J/n]: ").strip().lower()
        should_save = ans not in ("n", "nein")

    if should_save:
        try:
            saved_path = editor.save(output_path=args.output)
            print(f"  [OK] Gespeichert: {saved_path}")
        except PermissionError:
            target = args.output or args.file
            import tempfile, os
            local_example = Path(tempfile.gettempdir()) / "COP_TEST_ergebnis.xlsx"
            print(
                f"\n[FEHLER] Keine Schreibrechte auf: {target}\n"
                f"  Mögliche Ursachen:\n"
                f"    • Datei ist in Excel geöffnet (Excel sperrt die Datei)\n"
                f"    • Netzwerkpfad / SharePoint ist schreibgeschützt\n"
                f"\n  Lösung: Datei in Excel schließen, oder mit --output lokal speichern:\n"
                f"    --output \"{local_example}\"",
                file=sys.stderr,
            )
            sys.exit(1)
    else:
        print("  [i] Nicht gespeichert (--save nicht angegeben).")


if __name__ == "__main__":
    main()