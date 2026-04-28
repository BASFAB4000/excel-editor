"""
Tests für CLI-Argument-Parsing: --path + --file Kombination, --SID.
"""
import unittest
from pathlib import Path

from excel_editor.cli import build_parser


class TestCliParsing(unittest.TestCase):

    def test_path_and_file_combined(self):
        parser = build_parser()
        args = parser.parse_args([
            "--file", "COP_Migration_Template.xlsx",
            "--path", "/tmp/testdir",
        ])
        combined = args.path / args.file
        self.assertEqual(combined, Path("/tmp/testdir/COP_Migration_Template.xlsx"))

    def test_sid_parsed_as_list(self):
        parser = build_parser()
        args = parser.parse_args([
            "--path", "/tmp",
            "--SID", "ZPP", "ZMR", "DSC", "VA1",
        ])
        self.assertEqual(args.SID, ["ZPP", "ZMR", "DSC", "VA1"])

    def test_sid_single(self):
        parser = build_parser()
        args = parser.parse_args(["--path", "/tmp", "--SID", "ZPP"])
        self.assertEqual(args.SID, ["ZPP"])

    def test_no_sid_is_none(self):
        parser = build_parser()
        args = parser.parse_args(["--file", "test.xlsx"])
        self.assertIsNone(args.SID)

    def test_move_from_and_after(self):
        parser = build_parser()
        args = parser.parse_args([
            "--file", "test.xlsx",
            "--move-from", "1005",
            "--move-after", "1020",
        ])
        self.assertEqual(args.move_from, "1005")
        self.assertEqual(args.move_after, "1020")

    def test_save_flag(self):
        parser = build_parser()
        args = parser.parse_args(["--file", "test.xlsx", "--save"])
        self.assertTrue(args.save)

    def test_header_row(self):
        parser = build_parser()
        args = parser.parse_args(["--file", "test.xlsx", "--header-row", "3"])
        self.assertEqual(args.header_row, 3)


if __name__ == "__main__":
    unittest.main()