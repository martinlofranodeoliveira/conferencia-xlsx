import csv
import hashlib
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import zipfile

from conferencia_xlsx.errors import AuditError
from conferencia_xlsx.publication import publish_bytes
from conferencia_xlsx.reports import bundle_bytes, csv_bytes, json_bytes, markdown_bytes, markdown_text, protect_csv


def example_report():
    source = {"sha256": "a" * 64, "size_bytes": 123, "epoch": "1900", "sheet_names": ["=Planilha"], "ignored_parts": [], "warning_count": 0}
    return {"schema_version": 1, "status": "different_in_scope", "identical_bytes": False,
            "scope": {"included": ["células", "fórmulas"], "excluded": ["formatação"], "cache_freshness": "unknown"},
            "before": source, "after": {**source, "sha256": "b" * 64}, "summary": {"changes": 1, "by_kind": {"literal": 1}},
            "changes": [{"kind": "literal", "sheet": "=Planilha", "coordinate": "A1", "before": {"kind": "string", "value": '<img src=x>|[a](https://example.invalid)\r\n"nome"'}, "after": {"kind": "number", "value": "1"}}]}


class ReportsTests(unittest.TestCase):
    def test_json_roundtrip_preserves_typed_text_and_has_no_nan(self):
        report = example_report()
        self.assertEqual(json.loads(json_bytes(report)), report)
        with self.assertRaises(ValueError):
            json_bytes({"bad": float("nan")})

    def test_decoded_string_controls_are_preserved_as_json_escapes_in_reports(self):
        report = example_report()
        original = {'kind': 'string', 'value': 'a\x00b\x01\r\n'}
        report['changes'][0]['before'] = original
        self.assertEqual(json.loads(json_bytes(report))['changes'][0]['before'], original)
        for result in [json_bytes(report), markdown_bytes(report), csv_bytes(report)]:
            self.assertNotIn(b'\x00', result)
            self.assertNotIn(b'\x01', result)
        rows = list(csv.reader(io.StringIO(csv_bytes(report).decode('utf-8-sig'), newline='')))
        self.assertEqual(json.loads(rows[1][3]), original)

    def test_markdown_values_do_not_become_images_links_or_table_rows(self):
        report = example_report()
        result = markdown_bytes(report).decode()
        self.assertNotIn('<img src=x>', result)
        self.assertNotIn('[a](https://example.invalid)', result)
        self.assertIn('&lt;img src=x&gt;', result)
        self.assertIn('\\|', result)
        self.assertIn('desconhecida', result)
        self.assertEqual(markdown_text('a\r\nb\rc\nd'), 'a<br>b<br>c<br>d')

    def test_csv_roundtrip_and_formula_prefix_protection(self):
        data = csv_bytes(example_report()).decode('utf-8-sig')
        rows = list(csv.reader(io.StringIO(data, newline='')))
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[1][1], "'=Planilha")
        self.assertEqual(json.loads(rows[1][3]), example_report()['changes'][0]['before'])
        for text in ['=1', '+2', '-3', '@call', ' \t=1', '\ufeff\r+1', '\u2003=1']:
            self.assertEqual(protect_csv(text), "'" + text)
        self.assertEqual(protect_csv('Texto'), 'Texto')

    def test_bundle_manifest_covers_exact_members_without_inputs_by_default(self):
        report = example_report()
        data = bundle_bytes(report)
        self.assertEqual(data, bundle_bytes(report))
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            self.assertIsNone(archive.testzip())
            self.assertEqual(set(archive.namelist()), {'comparacao.json', 'comparacao.md', 'alteracoes.csv', 'manifesto.json'})
            manifest = json.loads(archive.read('manifesto.json'))
            self.assertEqual(set(manifest['members']), set(archive.namelist()) - {'manifesto.json'})
            for name, record in manifest['members'].items():
                raw = archive.read(name)
                self.assertEqual(len(raw), record['size_bytes'])
                self.assertEqual(hashlib.sha256(raw).hexdigest(), record['sha256'])
        with zipfile.ZipFile(io.BytesIO(bundle_bytes(report, samples={'antes.xlsx': b'A', 'depois.xlsx': b'B'}))) as archive:
            self.assertEqual(archive.read('antes.xlsx'), b'A')
        with self.assertRaises(AuditError):
            bundle_bytes(report, samples={'../bad': b'no'})

    def test_markdown_includes_observed_outside_scope_inventory_and_warning_counts(self):
        report = example_report()
        report['before'] = {**report['before'], 'ignored_parts': {'drawings': 2}, 'warning_count': 1}
        report['after'] = {**report['after'], 'ignored_parts': {'drawings': 3, 'tables': 1}, 'warning_count': 2}
        text = markdown_bytes(report).decode()
        self.assertIn('| drawings | 2 | 3 |', text)
        self.assertIn('| tables | 0 | 1 |', text)
        self.assertIn('| Avisos do leitor — somente contagem | 1 | 2 |', text)

    def test_report_limit_refuses_instead_of_publishing_truncated_output(self):
        with patch('conferencia_xlsx.reports.MAX_REPORT_BYTES', 10):
            with self.assertRaises(AuditError):
                bundle_bytes(example_report())


class PublicationTests(unittest.TestCase):
    def test_new_output_is_complete_and_existing_destination_is_preserved(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'report.json'
            publish_bytes(path, b'{"complete":true}\n')
            with self.assertRaises(AuditError) as error:
                publish_bytes(path, b'changed')
            self.assertEqual(error.exception.code, 'output-exists')
            self.assertEqual(path.read_bytes(), b'{"complete":true}\n')
            self.assertEqual([p.name for p in Path(folder).iterdir()], ['report.json'])

    def test_link_failure_has_no_destructive_fallback_and_cleans_own_temp(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'report.json'
            with patch('conferencia_xlsx.publication.os.link', side_effect=OSError('synthetic PRIVATE detail')):
                with self.assertRaises(AuditError) as error:
                    publish_bytes(path, b'data')
            self.assertNotIn('PRIVATE', str(error.exception))
            self.assertFalse(path.exists())
            self.assertEqual(list(Path(folder).iterdir()), [])

    def test_directory_destination_is_not_replaced(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'existing'
            path.mkdir()
            (path / 'owned.txt').write_text('preserve')
            with self.assertRaises(AuditError):
                publish_bytes(path, b'data')
            self.assertEqual((path / 'owned.txt').read_text(), 'preserve')


if __name__ == '__main__':
    unittest.main()
