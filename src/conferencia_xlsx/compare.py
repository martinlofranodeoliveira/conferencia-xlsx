"""Comparação determinística do conteúdo armazenado no escopo declarado.

Não calcula fórmulas, não usa resultados armazenados para inferir sua atualidade
e não tenta reconhecer registros que mudaram de linha ou de aba.
"""
from collections import Counter

from .errors import AuditError
from .models import (
    Cache, Cell, Limits, Sheet, Snapshot, Value, coordinate_key, merge_key,
    validate_limits, validate_snapshot,
)


CHANGE_KINDS = (
    'epoch', 'sheet_removed', 'sheet_added', 'sheet_order', 'sheet_state',
    'merge_removed', 'merge_added', 'cell_removed', 'cell_added', 'literal',
    'formula', 'cache',
)

INCLUDED = (
    'date_system', 'worksheet_names_order_visibility', 'merged_ranges',
    'typed_stored_literals', 'expanded_formula_text', 'stored_formula_cache',
)
EXCLUDED = (
    'appearance_styles_and_number_formats', 'charts_and_images',
    'structured_tables', 'data_validations', 'comments', 'hyperlinks',
    'defined_names', 'digital_signatures', 'external_link_targets',
    'mathematical_formula_equivalence', 'formula_recalculation',
)


def _value(value: Value | None) -> dict | None:
    if value is None:
        return None
    return {'kind': value.kind, 'value': value.value}


def _cache(cache: Cache | None) -> dict | None:
    if cache is None:
        return None
    return {'state': cache.state, 'value': _value(cache.value), 'freshness': 'unknown'}


def _content(cell: Cell) -> dict:
    return {
        'literal': _value(cell.literal),
        'formula': cell.formula,
        'cache': _cache(cell.cache),
    }


def _sheet_summary(sheet: Sheet) -> dict:
    return {'state': sheet.state, 'cells': len(sheet.cells), 'merges': len(sheet.merges)}


def _source(snapshot: Snapshot) -> dict:
    return {
        'sha256': snapshot.sha256,
        'size_bytes': snapshot.size_bytes,
        'epoch': snapshot.epoch,
        'sheet_names': [sheet.name for sheet in snapshot.sheets],
        'ignored_parts': dict(sorted(snapshot.ignored_parts)),
        'warning_count': snapshot.warning_count,
    }


def _same_identity_consistent(before: Snapshot, after: Snapshot) -> bool:
    """A repeated hash cannot consistently describe conflicting snapshot data.

    This is only an invariant check on the supplied observations, not proof
    that a caller-supplied digest belongs to a particular on-disk file.
    Warning counts describe the reader invocation and are not byte identity.
    """
    if before.sha256 != after.sha256:
        return False
    if (
        before.size_bytes != after.size_bytes
        or before.epoch != after.epoch
        or before.sheets != after.sheets
        or dict(before.ignored_parts) != dict(after.ignored_parts)
    ):
        raise AuditError(
            'inconsistent-snapshot',
            'Snapshots com a mesma identidade apresentam conteúdo inconsistente.',
        )
    return True


def compare_snapshots(before: Snapshot, after: Snapshot, limits: Limits = Limits()) -> dict:
    """Return a complete comparison or raise a sanitized :class:`AuditError`.

    ``sheet_order`` records the relative order of sheets present in both
    snapshots. Added/removed sheets get one structural change each, rather
    than one change for every cell they contain. Source metadata retains the
    complete order, including newly inserted sheets.

    ``max_changes`` is inclusive: a result with exactly that many changes is
    complete; the next change raises instead of returning a truncated list.
    """
    validate_limits(limits)
    validate_snapshot(before, limits)
    validate_snapshot(after, limits)
    identical_bytes = _same_identity_consistent(before, after)
    changes = []

    def add(kind, sheet=None, coordinate=None, old=None, new=None):
        if len(changes) >= limits.max_changes:
            raise AuditError(
                'change-limit',
                'A comparação ultrapassa o limite de alterações; nenhum resultado parcial foi produzido.',
            )
        changes.append({
            'kind': kind, 'sheet': sheet, 'coordinate': coordinate,
            'before': old, 'after': new,
        })

    if before.epoch != after.epoch:
        add('epoch', old=before.epoch, new=after.epoch)

    old_sheets = {sheet.name: sheet for sheet in before.sheets}
    new_sheets = {sheet.name: sheet for sheet in after.sheets}
    for sheet in before.sheets:
        if sheet.name not in new_sheets:
            add('sheet_removed', sheet=sheet.name, old=_sheet_summary(sheet))
    for sheet in after.sheets:
        if sheet.name not in old_sheets:
            add('sheet_added', sheet=sheet.name, new=_sheet_summary(sheet))

    old_common = [sheet.name for sheet in before.sheets if sheet.name in new_sheets]
    new_common = [sheet.name for sheet in after.sheets if sheet.name in old_sheets]
    if old_common != new_common:
        add('sheet_order', old=old_common, new=new_common)

    for name in old_common:
        old_sheet, new_sheet = old_sheets[name], new_sheets[name]
        if old_sheet.state != new_sheet.state:
            add('sheet_state', sheet=name, old=old_sheet.state, new=new_sheet.state)
        old_merges, new_merges = set(old_sheet.merges), set(new_sheet.merges)
        for reference in sorted(old_merges - new_merges, key=merge_key):
            add('merge_removed', sheet=name, coordinate=reference, old=reference)
        for reference in sorted(new_merges - old_merges, key=merge_key):
            add('merge_added', sheet=name, coordinate=reference, new=reference)

        old_cells = {cell.coordinate: cell for cell in old_sheet.cells}
        new_cells = {cell.coordinate: cell for cell in new_sheet.cells}
        for coordinate in sorted(old_cells.keys() | new_cells.keys(), key=coordinate_key):
            old_cell, new_cell = old_cells.get(coordinate), new_cells.get(coordinate)
            if old_cell is None:
                add('cell_added', name, coordinate, new=_content(new_cell))
            elif new_cell is None:
                add('cell_removed', name, coordinate, old=_content(old_cell))
            elif old_cell.formula is None and new_cell.formula is None:
                if old_cell.literal != new_cell.literal:
                    add('literal', name, coordinate, _value(old_cell.literal), _value(new_cell.literal))
            elif (old_cell.formula is None) != (new_cell.formula is None):
                add('formula', name, coordinate, _content(old_cell), _content(new_cell))
            else:
                if old_cell.formula != new_cell.formula:
                    add('formula', name, coordinate, old_cell.formula, new_cell.formula)
                if old_cell.cache != new_cell.cache:
                    add('cache', name, coordinate, _cache(old_cell.cache), _cache(new_cell.cache))

    counts = Counter(change['kind'] for change in changes)
    return {
        'schema_version': 1,
        'status': 'different_in_scope' if changes else 'equal_in_scope',
        'identical_bytes': identical_bytes,
        'scope': {
            'included': list(INCLUDED),
            'excluded': list(EXCLUDED),
            'coordinate_policy': 'exact_sheet_name_and_coordinate; no_row_identity_tracking',
            'formula_policy': 'exact_text_after_shared_formula_expansion; no_evaluation',
            'number_policy': 'exact_canonical_decimal; serial_dates_remain_numbers',
            'date_policy': 'stored_ISO_text; date_system_compared_separately',
            'cache_freshness': 'unknown',
        },
        'before': _source(before),
        'after': _source(after),
        'summary': {
            'changes': len(changes),
            'by_kind': {kind: counts[kind] for kind in CHANGE_KINDS},
        },
        'changes': changes,
    }
