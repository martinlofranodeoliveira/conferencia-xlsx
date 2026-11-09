import hashlib
import json
import unittest
from dataclasses import replace
from decimal import localcontext

from conferencia_xlsx.compare import CHANGE_KINDS, compare_snapshots
from conferencia_xlsx.errors import AuditError
from conferencia_xlsx.models import (
    Cache, Cell, Limits, Sheet, Snapshot, Value, canonical_number,
    coordinate_key, merge_key,
)


def number(text):
    return Value('number', canonical_number(text))


def literal(coordinate, value):
    return Cell(coordinate, literal=value)


def formula(coordinate, text, state='absent', value=None):
    return Cell(coordinate, formula=text, cache=Cache(state, value))


def sheet(name='Dados', cells=(), merges=(), state='visible'):
    return Sheet(
        name, state,
        tuple(sorted(cells, key=lambda cell: coordinate_key(cell.coordinate))),
        tuple(sorted(merges, key=merge_key)),
    )


def snapshot(*sheets, epoch='1900', sha=None, ignored=(), warnings=0):
    if not sheets:
        sheets = (sheet(),)
    representation = repr((epoch, sheets, ignored)).encode('utf-8')
    return Snapshot(
        sha or hashlib.sha256(representation).hexdigest(), len(representation),
        epoch, tuple(sheets), tuple(ignored), warnings,
    )


class ComparisonTests(unittest.TestCase):
    def test_equal_scope_is_separate_from_equal_bytes_and_ignored_inventory(self):
        a = snapshot(sheet(cells=(literal('A1', Value('string', 'Caderno')),)), sha='a' * 64)
        b = replace(a, sha256='b' * 64, ignored_parts=(('drawings', 2),), warning_count=1)
        result = compare_snapshots(a, b)
        self.assertEqual(result['status'], 'equal_in_scope')
        self.assertFalse(result['identical_bytes'])
        self.assertEqual(result['summary']['changes'], 0)
        self.assertEqual(result['summary']['by_kind'], dict.fromkeys(CHANGE_KINDS, 0))
        self.assertEqual(result['after']['ignored_parts'], {'drawings': 2})
        self.assertEqual(result['scope']['cache_freshness'], 'unknown')
        self.assertIn('appearance_styles_and_number_formats', result['scope']['excluded'])

    def test_identical_snapshot_has_identical_bytes_flag(self):
        a = snapshot(sheet(cells=(literal('A1', number('1')),)))
        result = compare_snapshots(a, a)
        self.assertTrue(result['identical_bytes'])
        self.assertEqual(result['status'], 'equal_in_scope')

    def test_impossible_reused_identity_rejects_conflicting_observations(self):
        a = snapshot(sheet(cells=(literal('A1', number('1')),)))
        b = replace(snapshot(sheet(cells=(literal('A1', number('2')),))), sha256=a.sha256)
        with self.assertRaises(AuditError) as raised:
            compare_snapshots(a, b)
        self.assertEqual(raised.exception.code, 'inconsistent-snapshot')
        with self.assertRaises(AuditError):
            compare_snapshots(a, replace(a, size_bytes=a.size_bytes + 1))

    def test_literal_types_boolean_number_string_and_error_are_distinct(self):
        a = snapshot(sheet(cells=(
            literal('A1', Value('string', '1')),
            literal('A2', number('1')),
            literal('A3', Value('error', '#DIV/0!')),
        )))
        b = snapshot(sheet(cells=(
            literal('A1', number('1')),
            literal('A2', Value('boolean', True)),
            literal('A3', Value('string', '#DIV/0!')),
        )))
        result = compare_snapshots(a, b)
        self.assertEqual([c['kind'] for c in result['changes']], ['literal'] * 3)
        self.assertEqual(result['changes'][0]['before'], {'kind': 'string', 'value': '1'})
        self.assertEqual(result['changes'][1]['after'], {'kind': 'boolean', 'value': True})

    def test_decimal_equality_is_exact_and_independent_of_decimal_context(self):
        with localcontext() as ctx:
            ctx.prec = 2
            a = snapshot(sheet(cells=(literal('A1', number('1.000e0')), literal('A2', number('9007199254740993')), literal('A3', number('1.000000000000000000001')))))
            b = snapshot(sheet(cells=(literal('A1', number('1')), literal('A2', number('9007199254740992')), literal('A3', number('1')))))
            result = compare_snapshots(a, b)
        self.assertEqual([c['coordinate'] for c in result['changes']], ['A2', 'A3'])
        self.assertEqual(result['changes'][0]['before']['value'], '9007199254740993')
        self.assertEqual(result['changes'][1]['before']['value'], '1.000000000000000000001')

    def test_formula_text_is_exact_even_if_results_are_equal(self):
        a = snapshot(sheet(cells=(formula('C1', '=A1+B1', 'value', number('3')),)))
        b = snapshot(sheet(cells=(formula('C1', '=SUM(A1:B1)', 'value', number('3')),)))
        result = compare_snapshots(a, b)
        self.assertEqual(len(result['changes']), 1)
        self.assertEqual(result['changes'][0], {
            'kind': 'formula', 'sheet': 'Dados', 'coordinate': 'C1',
            'before': '=A1+B1', 'after': '=SUM(A1:B1)',
        })

    def test_unchanged_formula_can_have_changed_stored_result(self):
        a = snapshot(sheet(cells=(formula('C1', '=A1+B1', 'value', number('3')),)))
        b = snapshot(sheet(cells=(formula('C1', '=A1+B1', 'value', number('4')),)))
        result = compare_snapshots(a, b)
        self.assertEqual([c['kind'] for c in result['changes']], ['cache'])
        self.assertEqual(result['changes'][0]['before'], {
            'state': 'value', 'value': {'kind': 'number', 'value': '3'}, 'freshness': 'unknown',
        })

    def test_formula_and_cache_changes_are_separate_and_ordered(self):
        a = snapshot(sheet(cells=(formula('C1', '=1', 'value', number('1')),)))
        b = snapshot(sheet(cells=(formula('C1', '=2', 'value', number('2')),)))
        self.assertEqual([c['kind'] for c in compare_snapshots(a, b)['changes']], ['formula', 'cache'])

    def test_absent_empty_and_present_empty_string_cache_are_distinct(self):
        states = [Cache('absent'), Cache('empty'), Cache('value', Value('string', ''))]
        for i, first in enumerate(states):
            for j, second in enumerate(states):
                with self.subTest(i=i, j=j):
                    a = snapshot(sheet(cells=(Cell('A1', formula='=IF(1,"","")', cache=first),)))
                    b = snapshot(sheet(cells=(Cell('A1', formula='=IF(1,"","")', cache=second),)))
                    result = compare_snapshots(a, b)
                    self.assertEqual(result['summary']['changes'], int(i != j))
                    if i != j:
                        self.assertEqual(result['changes'][0]['kind'], 'cache')
                        self.assertEqual(result['changes'][0]['after']['freshness'], 'unknown')

    def test_literal_to_formula_and_back_each_produce_one_full_content_change(self):
        a = snapshot(sheet(cells=(literal('A1', Value('string', '=1')),)))
        b = snapshot(sheet(cells=(formula('A1', '=1', 'value', number('1')),)))
        for first, second in [(a, b), (b, a)]:
            result = compare_snapshots(first, second)
            self.assertEqual(result['summary']['changes'], 1)
            change = result['changes'][0]
            self.assertEqual(change['kind'], 'formula')
            self.assertIn('literal', change['before'])
            self.assertIn('cache', change['after'])

    def test_added_removed_cells_keep_complete_content_without_duplicate_events(self):
        a = snapshot(sheet(cells=(literal('A1', Value('string', '')),)))
        b = snapshot(sheet(cells=(formula('B1', '=7', 'value', number('7')),)))
        result = compare_snapshots(a, b)
        self.assertEqual([c['kind'] for c in result['changes']], ['cell_removed', 'cell_added'])
        self.assertEqual(result['changes'][0]['before']['literal']['value'], '')
        self.assertEqual(result['changes'][1]['after']['formula'], '=7')
        self.assertEqual(result['changes'][1]['after']['cache']['freshness'], 'unknown')

    def test_added_removed_sheets_are_summarized_without_expanding_cells(self):
        a = snapshot(sheet('Comum'), sheet('Antiga', cells=(literal('A1', number('1')),), merges=('B2:C3',)))
        b = snapshot(sheet('Comum'), sheet('Nova', cells=(formula('A1', '=1'),)))
        result = compare_snapshots(a, b)
        self.assertEqual([c['kind'] for c in result['changes']], ['sheet_removed', 'sheet_added'])
        self.assertEqual(result['changes'][0]['before'], {'state': 'visible', 'cells': 1, 'merges': 1})
        self.assertEqual(result['after']['sheet_names'], ['Comum', 'Nova'])

    def test_case_renaming_is_removal_and_addition_not_inferred_identity(self):
        result = compare_snapshots(snapshot(sheet('Dados')), snapshot(sheet('dados')))
        self.assertEqual([c['kind'] for c in result['changes']], ['sheet_removed', 'sheet_added'])

    def test_common_sheet_order_changes_without_redundancy_for_simple_addition(self):
        a = snapshot(sheet('B'), sheet('A'))
        b = snapshot(sheet('A'), sheet('Nova'), sheet('B'))
        result = compare_snapshots(a, b)
        self.assertEqual([c['kind'] for c in result['changes']], ['sheet_added', 'sheet_order'])
        self.assertEqual(result['changes'][1]['before'], ['B', 'A'])
        self.assertEqual(result['changes'][1]['after'], ['A', 'B'])
        only_added = compare_snapshots(a, snapshot(sheet('B'), sheet('Nova'), sheet('A')))
        self.assertEqual([c['kind'] for c in only_added['changes']], ['sheet_added'])

    def test_epoch_state_merges_and_coordinates_have_deterministic_order(self):
        a = snapshot(sheet('B', cells=(literal('A10', number('1')), literal('A2', number('1')), literal('B1', number('1'))), merges=('J10:K11', 'A1:B2')), sheet('A'))
        b = snapshot(sheet('A'), sheet('B', cells=(literal('A10', number('2')), literal('A2', number('2')), literal('B1', number('2'))), merges=('A3:B4',), state='hidden'), epoch='1904')
        result = compare_snapshots(a, b)
        self.assertEqual([c['kind'] for c in result['changes']], [
            'epoch', 'sheet_order', 'sheet_state', 'merge_removed', 'merge_removed',
            'merge_added', 'literal', 'literal', 'literal',
        ])
        self.assertEqual([c['coordinate'] for c in result['changes'][-3:]], ['B1', 'A2', 'A10'])
        self.assertEqual(result['changes'][3]['coordinate'], 'A1:B2')
        self.assertEqual(json.dumps(result, ensure_ascii=False, allow_nan=False), json.dumps(compare_snapshots(a, b), ensure_ascii=False, allow_nan=False))

    def test_iso_date_text_is_conservative_and_numeric_serial_is_not_formatted(self):
        a = snapshot(sheet(cells=(literal('A1', Value('date', '2024-01-01')), literal('A2', number('45292')))))
        b = snapshot(sheet(cells=(literal('A1', Value('date', '2024-01-01T00:00:00')), literal('A2', number('45292')))))
        result = compare_snapshots(a, b)
        self.assertEqual([c['coordinate'] for c in result['changes']], ['A1'])
        self.assertIn('serial_dates_remain_numbers', result['scope']['number_policy'])

    def test_result_does_not_share_mutable_lists_with_inputs_or_later_results(self):
        a = snapshot(sheet(cells=(literal('A1', Value('string', '<b>|[x]\r\nTexto')),)))
        b = snapshot(sheet(cells=(literal('A1', Value('string', 'Depois')),)))
        original = repr((a, b))
        result = compare_snapshots(a, b)
        result['changes'][0]['before']['value'] = 'alterado'
        result['before']['sheet_names'].append('injetado')
        result['scope']['included'].append('impossível')
        self.assertEqual(repr((a, b)), original)
        fresh = compare_snapshots(a, b)
        self.assertEqual(fresh['before']['sheet_names'], ['Dados'])
        self.assertNotIn('impossível', fresh['scope']['included'])
        self.assertEqual(fresh['changes'][0]['before']['value'], '<b>|[x]\r\nTexto')

    def test_change_budget_is_inclusive_and_never_returns_partial_result(self):
        a = snapshot(sheet(cells=(literal('A1', number('1')), literal('B1', number('1')))))
        b = snapshot(sheet(cells=(literal('A1', number('2')), literal('B1', number('2')))))
        one = snapshot(sheet(cells=(literal('A1', number('2')), literal('B1', number('1')))))
        self.assertEqual(compare_snapshots(a, one, Limits(max_changes=1))['summary']['changes'], 1)
        with self.assertRaises(AuditError) as raised:
            compare_snapshots(a, b, Limits(max_changes=1))
        self.assertEqual(raised.exception.code, 'change-limit')
        self.assertEqual(a.sheets[0].cells[0].literal.value, '1')

    def test_forged_snapshots_and_limits_are_rejected_without_echoing_contents(self):
        good = snapshot()
        malformed = [
            replace(good, sha256='PRIVATE_SHA'),
            replace(good, size_bytes=True),
            replace(good, sheets=[sheet()]),
            replace(good, sheets=(sheet('Dados'), sheet('dados'))),
            replace(good, sheets=(Sheet('Dados', 'visible', (literal('A1', number('1')), literal('A1', number('2'))), ()),)),
            replace(good, sheets=(Sheet('Dados', 'visible', (literal('A2', number('1')), literal('A1', number('2'))), ()),)),
            replace(good, sheets=(sheet(cells=(literal('A1', Value('boolean', 1)),)),)),
            replace(good, sheets=(sheet(cells=(literal('A1', Value('number', True)),)),)),
            replace(good, sheets=(sheet(cells=(literal('A1', Value('number', '1.0')),)),)),
            replace(good, sheets=(sheet(cells=(literal('A1', Value([], 'x')),)),)),
            replace(good, sheets=(sheet(cells=(Cell('A1', formula='=PRIVATE_FORMULA'),)),)),
            replace(good, sheets=(sheet(cells=(Cell('A1', literal=number('1'), formula='=PRIVATE_FORMULA', cache=Cache('empty')),)),)),
            replace(good, sheets=(sheet(cells=(Cell('A1', formula='=1', cache=Cache('empty', number('1'))),)),)),
            replace(good, ignored_parts=(('same', 1), ('same', 2))),
            replace(good, ignored_parts=(('\ud800', 1),)),
        ]
        for malformed_snapshot in malformed:
            with self.subTest(snapshot=repr(malformed_snapshot)[:80]):
                with self.assertRaises(AuditError) as raised:
                    compare_snapshots(good, malformed_snapshot)
                self.assertNotIn('PRIVATE', str(raised.exception))
        for limits in [Limits(max_changes=True), Limits(max_cells=0), {'max_changes': 1}]:
            with self.assertRaises(AuditError):
                compare_snapshots(good, good, limits)


if __name__ == '__main__':
    unittest.main()
