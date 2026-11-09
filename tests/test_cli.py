from contextlib import redirect_stdout, redirect_stderr
from hashlib import sha256
from importlib.resources import files
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import zipfile

from conferencia_xlsx.cli import main
from conferencia_xlsx.errors import AuditError


class CliTests(unittest.TestCase):
    def call(self, args):
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = main(args)
        return code, out.getvalue(), err.getvalue()

    def sources(self, root):
        paths = []
        for name in ['antes.xlsx', 'depois.xlsx']:
            path = root / name
            path.write_bytes(files('conferencia_xlsx').joinpath('samples', name).read_bytes())
            paths.append(path)
        return paths

    def test_real_comparison_exit1_bundle_and_input_hashes_unchanged(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            a, b = self.sources(root)
            hashes = [sha256(p.read_bytes()).hexdigest() for p in [a, b]]
            output = root / 'review.zip'
            code, text, error = self.call(['comparar', str(a), str(b), '--saida', str(output)])
            self.assertEqual((code, error), (1, ''))
            self.assertIn('Comparação completa:', text)
            self.assertNotIn('SUM', text)
            self.assertEqual(hashes, [sha256(p.read_bytes()).hexdigest() for p in [a, b]])
            with zipfile.ZipFile(output) as archive:
                report = json.loads(archive.read('comparacao.json'))
                self.assertEqual(report['status'], 'different_in_scope')
                self.assertGreater(report['summary']['changes'], 0)
                self.assertNotIn('antes.xlsx', archive.namelist())
                self.assertNotIn('depois.xlsx', archive.namelist())

    def test_equal_input_exit0_and_individual_json(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            a, _ = self.sources(root)
            output = root / 'same.json'
            code, _, error = self.call(['comparar', str(a), str(a), '--saida', str(output), '--formato', 'json'])
            self.assertEqual((code, error), (0, ''))
            report = json.loads(output.read_bytes())
            self.assertTrue(report['identical_bytes'])
            self.assertEqual(report['summary']['changes'], 0)

    def test_existing_output_exit2_and_original_bytes_preserved(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            a, b = self.sources(root)
            output = root / 'review.zip'
            output.write_bytes(b'previous content')
            code, text, error = self.call(['comparar', str(a), str(b), '--saida', str(output)])
            self.assertEqual(code, 2)
            self.assertEqual(text, '')
            self.assertIn('output-exists', error)
            self.assertEqual(output.read_bytes(), b'previous content')

    def test_invalid_source_diagnostic_has_no_content_path_or_traceback(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            bad = root / 'PRIVATE-input.xlsx'
            bad.write_bytes(b'PRIVATE contents, not a workbook')
            output = root / 'no-result.zip'
            code, text, error = self.call(['comparar', str(bad), str(bad), '--saida', str(output)])
            self.assertEqual(code, 2)
            self.assertEqual(text, '')
            self.assertNotIn('PRIVATE', error)
            self.assertNotIn('Traceback', error)
            self.assertFalse(output.exists())

    def test_public_audit_error_is_handled_as_exit2_without_handler_failure(self):
        with patch('conferencia_xlsx.cli.read_path', side_effect=AuditError('synthetic', 'Falha sanitizada.')):
            code, text, error = self.call(['comparar', 'PRIVATE-a', 'PRIVATE-b', '--saida', 'unused.zip'])
        self.assertEqual(code, 2)
        self.assertEqual(text, '')
        self.assertEqual(error, 'Erro [synthetic]: Falha sanitizada.\n')

    def test_demo_includes_only_original_synthetic_sources_and_reports(self):
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / 'demo.zip'
            code, text, error = self.call(['exemplo', '--saida', str(output)])
            self.assertEqual((code, error), (0, ''))
            self.assertIn('fictício', text)
            with zipfile.ZipFile(output) as archive:
                self.assertEqual(set(archive.namelist()), {'antes.xlsx', 'depois.xlsx', 'comparacao.json', 'comparacao.md', 'alteracoes.csv', 'manifesto.json'})
                self.assertEqual(archive.read('antes.xlsx'), files('conferencia_xlsx').joinpath('samples', 'antes.xlsx').read_bytes())


if __name__ == '__main__':
    unittest.main()
