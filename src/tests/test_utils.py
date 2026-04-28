"""
Tests für _find_sid_files: Ordner- und Datei-Erkennung per SID.
"""
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from excel_editor.cli import _find_sid_files, _MAX_SIDS


def _make_sid_structure(base: Path, sids_with_excel: list, sids_without_excel: list) -> None:
    """Baut eine SID-Ordnerstruktur auf."""
    for sid in sids_with_excel:
        folder = base / f"118_{sid} - Test System"
        folder.mkdir()
        wb = Workbook()
        wb.active["A1"] = "No"
        wb.save(folder / f"COP_{sid}_Migration.xlsx")

    for sid in sids_without_excel:
        folder = base / f"119_{sid} - Missing Excel"
        folder.mkdir()


class TestFindSidFiles(unittest.TestCase):

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())

    def tearDown(self):
        shutil.rmtree(self.tmp)

    def test_sid_folder_and_excel_found(self):
        _make_sid_structure(self.tmp, ["ZPP"], [])
        results = _find_sid_files(self.tmp, ["ZPP"])
        sid, folder, excel = results[0]
        self.assertEqual(sid, "ZPP")
        self.assertIsNotNone(folder)
        self.assertIsNotNone(excel)
        self.assertEqual(excel.name, "COP_ZPP_Migration.xlsx")

    def test_sid_folder_found_but_excel_missing(self):
        _make_sid_structure(self.tmp, [], ["ZMR"])
        results = _find_sid_files(self.tmp, ["ZMR"])
        sid, folder, excel = results[0]
        self.assertEqual(sid, "ZMR")
        self.assertIsNotNone(folder)
        self.assertIsNone(excel)

    def test_sid_folder_not_found(self):
        results = _find_sid_files(self.tmp, ["XXX"])
        sid, folder, excel = results[0]
        self.assertEqual(sid, "XXX")
        self.assertIsNone(folder)
        self.assertIsNone(excel)

    def test_ambiguous_sid_folder(self):
        """Zwei Ordner enthalten dieselbe SID → mehrdeutig."""
        (self.tmp / "100_DSC - System A").mkdir()
        (self.tmp / "200_DSC - System B").mkdir()
        results = _find_sid_files(self.tmp, ["DSC"])
        sid, folder, excel = results[0]
        self.assertEqual(sid, "DSC")
        self.assertIsNone(folder)
        self.assertIsNone(excel)

    def test_multiple_sids_mixed(self):
        _make_sid_structure(self.tmp, ["ZPP", "VA1"], ["ZMR"])
        results = _find_sid_files(self.tmp, ["ZPP", "VA1", "ZMR", "XXX"])
        by_sid = {s: (f, x) for s, f, x in results}

        self.assertIsNotNone(by_sid["ZPP"][1])   # Excel vorhanden
        self.assertIsNotNone(by_sid["VA1"][1])   # Excel vorhanden
        self.assertIsNone(by_sid["ZMR"][1])       # Ordner da, Excel fehlt
        self.assertIsNone(by_sid["XXX"][0])       # Ordner fehlt

    def test_sid_not_matched_as_substring(self):
        """'VA' sollte nicht 'VA1' matchen."""
        _make_sid_structure(self.tmp, ["VA1"], [])
        results = _find_sid_files(self.tmp, ["VA"])
        sid, folder, excel = results[0]
        self.assertIsNone(folder)

    def test_directory_not_found_exits(self):
        missing = self.tmp / "nicht_vorhanden"
        with self.assertRaises(SystemExit):
            _find_sid_files(missing, ["ZPP"])

    def test_max_sids_limit(self):
        """Mehr als _MAX_SIDS SIDs → sys.exit."""
        from excel_editor.cli import _run_sid_mode
        too_many = [f"S{i:02d}" for i in range(_MAX_SIDS + 1)]

        class FakeArgs:
            SID = too_many
            path = self.tmp

        with self.assertRaises(SystemExit):
            _run_sid_mode(FakeArgs())


if __name__ == "__main__":
    unittest.main()