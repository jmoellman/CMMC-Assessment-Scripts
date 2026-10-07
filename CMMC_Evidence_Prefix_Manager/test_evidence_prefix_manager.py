import csv
import tempfile
import unittest
from pathlib import Path

from evidence_prefix_manager import (
    MappingRow, canonical_objective, execute_plans, load_mapping, match_file,
    prefixed_name, scan_sources,
)


class PrefixManagerTests(unittest.TestCase):
    def test_objective_normalization(self):
        self.assertEqual(canonical_objective("3.1.1A"), "3.1.1[a]")
        self.assertEqual(canonical_objective("AC.L2-3.1.1[B]"), "AC.L2-3.1.1[b]")

    def test_prefix_and_no_duplicate(self):
        row = MappingRow("3.1.1[a]", "ERL-AC-001")
        expected = "3.1.1[a]_ERL-AC-001_Users.xlsx"
        self.assertEqual(prefixed_name("Users.xlsx", row), expected)
        self.assertEqual(prefixed_name(expected, row), expected)
        self.assertEqual(prefixed_name("ERL-AC-001_Users.xlsx", row), expected)
        self.assertEqual(
            prefixed_name("3.1.1[a] Authorized User List.xlsx", row),
            "3.1.1[a]_ERL-AC-001_Authorized User List.xlsx",
        )

    def test_unique_erl_match(self):
        rows = [
            MappingRow("3.1.1[a]", "ERL-AC-001"),
            MappingRow("3.1.1[b]", "ERL-AC-002"),
        ]
        path = Path("ERL-AC-002 Account Process.pdf")
        plan = match_file(path, path, rows)
        self.assertEqual(plan.status, "Ready")
        self.assertEqual(plan.mapping.objective, "3.1.1[b]")

    def test_long_name_is_bounded(self):
        row = MappingRow("3.1.1[a]", "ERL-AC-001")
        result = prefixed_name(("A" * 300) + ".pdf", row)
        self.assertLessEqual(len(result), 240)
        self.assertTrue(result.endswith(".pdf"))

    def test_ambiguous_repeated_erl(self):
        rows = [
            MappingRow("3.1.1[a]", "ERL-001"),
            MappingRow("3.1.1[b]", "ERL-001"),
        ]
        path = Path("ERL-001 Evidence.pdf")
        plan = match_file(path, path, rows)
        self.assertEqual(plan.status, "Ambiguous")

    def test_multiple_folders_are_kept_separate(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            first, second = root / "AC_Evidence", root / "AU_Evidence"
            first.mkdir(); second.mkdir()
            (first / "ERL-001 Users.txt").write_text("one", encoding="utf-8")
            (second / "ERL-002 Logs.txt").write_text("two", encoding="utf-8")
            rows = [
                MappingRow("3.1.1[a]", "ERL-001"),
                MappingRow("3.3.1[a]", "ERL-002"),
            ]
            plans = scan_sources([first, second], rows)
            self.assertEqual(plans[0].relative_path.parts[0], "AC_Evidence")
            self.assertEqual(plans[1].relative_path.parts[0], "AU_Evidence")

    def test_load_mapping_and_hash_verified_copy(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            mapping = root / "mapping.csv"
            mapping.write_text(
                "Domain,Control ID,Objective ID,ERL #,Description\n"
                "Access Control,AC.L2-3.1.1,3.1.1[A],ERL-001,Users\n",
                encoding="utf-8",
            )
            rows = load_mapping(mapping)
            self.assertEqual(rows[0].objective, "3.1.1[a]")
            source = root / "source"
            source.mkdir()
            evidence = source / "ERL-001 User List.txt"
            evidence.write_text("unchanged evidence bytes", encoding="utf-8")
            plan = match_file(evidence, Path(evidence.name), rows)
            output = root / "output"
            log, counts = execute_plans([plan], source, output)
            self.assertEqual(counts["Copied"], 1)
            copied = output / "3.1.1[a]_ERL-001_User List.txt"
            self.assertTrue(copied.exists())
            with log.open(encoding="utf-8-sig", newline="") as handle:
                record = next(csv.DictReader(handle))
            self.assertEqual(record["hash_verified"], "TRUE")

    def test_ignyte_descriptive_erl_header(self):
        with tempfile.TemporaryDirectory() as temp:
            mapping = Path(temp) / "ignyte.xlsx"
            from openpyxl import Workbook
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Evidence Plan"
            sheet.append([
                "OBJECTIVE", "DOMAIN", "LEVEL", "SECURITY REQUIREMENT",
                "EVIDENCE TYPE", "EVIDENCE EXAMPLES",
                "ERL # (Suggested identifier for the evidence)",
                "Document Mapping (please add the ERL# as a prefix to the file name)",
            ])
            sheet.append(["3.1.1[a]", "AC", "L1", "Authorized users are identified.",
                          "Document", "Account documentation", "E-AC-01", "Access Policy.docx"])
            workbook.save(mapping)
            rows = load_mapping(mapping)
            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0].objective, "3.1.1[a]")
            self.assertEqual(rows[0].erl, "E-AC-01")
            self.assertEqual(rows[0].domain, "AC")


if __name__ == "__main__":
    unittest.main()
