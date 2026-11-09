"""Recorded real-consumer cases; these tests do not start LibreOffice."""
from io import BytesIO
import importlib.util
from pathlib import Path
import unittest
from zipfile import ZipFile

from conferencia_xlsx.compare import compare_snapshots
from conferencia_xlsx.reader import read_bytes, read_path

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("xlsx_rehearsal", ROOT / "tools/libreoffice_rehearsal.py")
REHEARSAL = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(REHEARSAL)
SAMPLES = ROOT / "src/conferencia_xlsx/samples"


class LibreOfficeFixtureTests(unittest.TestCase):
    def test_escaped_plain_and_rich_strings_match_recorded_real_consumer(self):
        spec = importlib.util.spec_from_file_location("xstring_rehearsal", ROOT / "tools/xstring_rehearsal.py")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        result = module.check_fixture()
        self.assertEqual(result["cases"], 14)
        self.assertEqual(result["cells"], 28)
        self.assertIs(result["app_matches_recorded_consumer"], True)
        self.assertIs(result["recorded_hashes_verified"], True)
        self.assertIs(result["external_consumer_rerun"], False)

    def test_recorded_consumer_hashes_oracles_and_eleven_known_changes(self):
        result = REHEARSAL.check_fixture()
        self.assertEqual(result["expected_changes"], 11)
        self.assertEqual(result["missing_cache_changes"], 1)
        self.assertIs(result["recorded_hashes_verified"], True)
        self.assertIs(result["external_consumer_rerun"], False)
        self.assertEqual(result["cache_freshness"], "unknown")

    def test_actual_file_reader_preserves_both_recorded_inputs(self):
        for name in ("antes.xlsx", "depois.xlsx"):
            path = SAMPLES / name
            original = path.read_bytes()
            self.assertEqual(read_path(path), read_bytes(original))
            self.assertEqual(path.read_bytes(), original)

    def test_document_metadata_changes_do_not_become_whole_file_equality(self):
        original = (SAMPLES / "antes.xlsx").read_bytes()
        output = BytesIO()
        with ZipFile(BytesIO(original)) as source, ZipFile(output, "w") as target:
            for info in source.infolist():
                data = source.read(info)
                if info.filename == "docProps/core.xml":
                    self.assertIn(b"Equipe Exemplo", data)
                    data = data.replace(b"Equipe Exemplo", b"Outra Equipe Ficticia")
                target.writestr(info, data)
        report = compare_snapshots(read_bytes(original), read_bytes(output.getvalue()))
        self.assertEqual(report["status"], "equal_in_scope")
        self.assertIs(report["identical_bytes"], False)
        self.assertEqual(report["summary"]["changes"], 0)


if __name__ == "__main__":
    unittest.main()
