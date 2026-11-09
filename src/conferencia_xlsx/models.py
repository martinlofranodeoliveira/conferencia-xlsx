"""Snapshots imutáveis e invariantes compartilhados pelo leitor e comparador."""
from dataclasses import dataclass, fields
import datetime
from decimal import Decimal, InvalidOperation
import re

from .errors import AuditError


@dataclass(frozen=True)
class Limits:
    max_file_bytes: int = 16 * 1024 * 1024
    max_expanded_bytes: int = 64 * 1024 * 1024
    max_parts: int = 2048
    max_cells: int = 100000
    max_grid_slots: int = 2000000
    max_xml_nodes: int = 500000
    max_sheets: int = 128
    max_changes: int = 100000
    max_text_chars: int = 1000000


@dataclass(frozen=True)
class Value:
    kind: str
    value: str | bool


@dataclass(frozen=True)
class Cache:
    state: str
    value: Value | None = None


@dataclass(frozen=True)
class Cell:
    coordinate: str
    literal: Value | None = None
    formula: str | None = None
    cache: Cache | None = None


@dataclass(frozen=True)
class Sheet:
    name: str
    state: str
    cells: tuple[Cell, ...]
    merges: tuple[str, ...]


@dataclass(frozen=True)
class Snapshot:
    sha256: str
    size_bytes: int
    epoch: str
    sheets: tuple[Sheet, ...]
    ignored_parts: tuple[tuple[str, int], ...] = ()
    warning_count: int = 0


def fail(code: str, message: str):
    raise AuditError(code, message)


def validate_limits(limits: Limits) -> Limits:
    if type(limits) is not Limits:
        fail('invalid-limits', 'Informe limites de recurso válidos.')
    for field in fields(limits):
        value = getattr(limits, field.name)
        if type(value) is not int or value < 1:
            fail('invalid-limits', 'Os limites precisam ser inteiros positivos.')
    return limits


def canonical_number(text: str) -> str:
    """Canonicaliza sem arredondamento ou dependência do contexto Decimal."""
    if type(text) is not str or len(text) > 1024 or not re.fullmatch(
        r'[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?', text
    ):
        fail('invalid-number', 'Um número armazenado não é reconhecido.')
    exponent = re.search(r'[eE]([+-]?[0-9]+)$', text)
    if exponent and (len(exponent[1].lstrip('+-')) > 7 or abs(int(exponent[1])) > 1000000):
        fail('number-limit', 'Um expoente numérico ultrapassa o limite suportado.')
    try:
        value = Decimal(text)
        sign, digits, power = value.as_tuple()
    except (InvalidOperation, ValueError):
        fail('invalid-number', 'Um número armazenado não é reconhecido.')
    if not any(digits):
        return '0'
    while digits and digits[0] == 0:
        digits = digits[1:]
    while digits[-1] == 0:
        digits = digits[:-1]
        power += 1
    return str(Decimal((sign, digits, power)))


def coordinate_key(coordinate: str) -> tuple[int, int]:
    if type(coordinate) is not str:
        fail('invalid-coordinate', 'Uma coordenada celular é inválida.')
    match = re.fullmatch(r'([A-Z]{1,3})([1-9][0-9]{0,6})', coordinate)
    if not match:
        fail('invalid-coordinate', 'Uma coordenada celular é inválida.')
    column = 0
    for letter in match[1]:
        column = column * 26 + ord(letter) - 64
    row = int(match[2])
    if row > 1048576 or column > 16384:
        fail('invalid-coordinate', 'Uma coordenada ultrapassa a grade XLSX.')
    return row, column


def column_name(column: int) -> str:
    result = ''
    while column:
        column, digit = divmod(column - 1, 26)
        result = chr(65 + digit) + result
    return result


def merge_key(reference: str) -> tuple[int, int, int, int]:
    if type(reference) is not str:
        fail('invalid-merge', 'Uma mesclagem é inválida.')
    parts = reference.split(':')
    if len(parts) != 2:
        fail('invalid-merge', 'Uma mesclagem é inválida.')
    row1, column1 = coordinate_key(parts[0])
    row2, column2 = coordinate_key(parts[1])
    if row1 > row2 or column1 > column2:
        fail('invalid-merge', 'Uma mesclagem contém um intervalo invertido.')
    return row1, column1, row2, column2


def normalize_merge(reference: str) -> str:
    if type(reference) is not str:
        fail('invalid-merge', 'Uma mesclagem é inválida.')
    reference = reference.replace('$', '')
    if ':' not in reference:
        reference = f'{reference}:{reference}'
    merge_key(reference)
    return reference


def validate_text(value: str, limits: Limits, *, required=False, allow_controls=False):
    if type(value) is not str or len(value) > limits.max_text_chars or (required and not value):
        fail('invalid-value', 'Um texto armazenado é inválido ou ultrapassa o limite.')
    try:
        value.encode('utf-8', 'strict')
    except UnicodeEncodeError:
        fail('invalid-value', 'Um texto armazenado contém Unicode inválido.')
    if not allow_controls and re.search(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', value):
        fail('invalid-value', 'Um texto armazenado contém controles inválidos.')


def validate_value(value: Value, limits: Limits):
    if type(value) is not Value or type(value.kind) is not str:
        fail('invalid-snapshot', 'O snapshot contém um valor tipado inválido.')
    if value.kind == 'boolean':
        if type(value.value) is not bool:
            fail('invalid-snapshot', 'O snapshot contém um booleano inválido.')
    elif value.kind in {'string', 'number', 'date', 'error'}:
        validate_text(value.value, limits, allow_controls=value.kind == 'string')
        if value.kind == 'number' and canonical_number(value.value) != value.value:
            fail('invalid-snapshot', 'O snapshot contém um número não canônico.')
        if value.kind in {'date', 'error'} and not value.value:
            fail('invalid-snapshot', 'O snapshot contém um valor armazenado vazio.')
        if value.kind == 'date':
            validate_iso_date(value.value)
    else:
        fail('invalid-snapshot', 'O snapshot contém um tipo celular desconhecido.')


def validate_iso_date(text: str):
    """Valida a representação ISO aceita, conservando sua grafia e precisão."""
    date = r'[0-9]{4}-[0-9]{2}-[0-9]{2}'
    time = r'[0-9]{2}:[0-9]{2}(?::[0-9]{2}(?:\.[0-9]+)?)?'
    match = re.fullmatch(rf'(?:(?P<date>{date})(?:T(?P<dt>{time}))?|(?P<time>{time}))(?P<tz>Z|[+-][0-9]{{2}}:[0-9]{{2}})?', text)
    try:
        if match is None:
            raise ValueError()
        if match['date']:
            datetime.date.fromisoformat(match['date'])
        clock = match['dt'] or match['time']
        if clock:
            datetime.time.fromisoformat(clock)
        zone = match['tz']
        if zone and zone != 'Z':
            hours, minutes = map(int, zone[1:].split(':'))
            if hours > 14 or minutes > 59 or (hours == 14 and minutes):
                raise ValueError()
    except (ValueError, OverflowError):
        fail('invalid-date', 'Uma data ISO armazenada não é reconhecida.')


def validate_snapshot(snapshot: Snapshot, limits: Limits = Limits()) -> Snapshot:
    validate_limits(limits)
    if type(snapshot) is not Snapshot:
        fail('invalid-snapshot', 'Informe um snapshot XLSX válido.')
    if type(snapshot.sha256) is not str or not re.fullmatch(r'[0-9a-f]{64}', snapshot.sha256):
        fail('invalid-snapshot', 'O snapshot contém uma identidade inválida.')
    if type(snapshot.size_bytes) is not int or not 0 < snapshot.size_bytes <= limits.max_file_bytes:
        fail('invalid-snapshot', 'O tamanho do snapshot é inválido.')
    if type(snapshot.epoch) is not str or snapshot.epoch not in ('1900', '1904'):
        fail('invalid-snapshot', 'O sistema de datas do snapshot é inválido.')
    if type(snapshot.sheets) is not tuple or not snapshot.sheets or len(snapshot.sheets) > limits.max_sheets:
        fail('invalid-snapshot', 'A lista de abas do snapshot é inválida.')
    names = set()
    cell_count = grid_slots = 0
    for sheet in snapshot.sheets:
        if type(sheet) is not Sheet:
            fail('invalid-snapshot', 'O snapshot contém uma aba inválida.')
        validate_text(sheet.name, limits, required=True)
        if len(sheet.name) > 31 or re.search(r'[\\/*?:\[\]]', sheet.name) or sheet.name.casefold() in names:
            fail('invalid-snapshot', 'O snapshot contém nomes de abas inválidos ou repetidos.')
        names.add(sheet.name.casefold())
        if type(sheet.state) is not str or sheet.state not in ('visible', 'hidden', 'veryHidden'):
            fail('invalid-snapshot', 'A visibilidade de uma aba é inválida.')
        if type(sheet.cells) is not tuple or type(sheet.merges) is not tuple:
            fail('invalid-snapshot', 'O snapshot precisa de sequências imutáveis.')
        if cell_count + len(sheet.cells) > limits.max_cells or len(sheet.merges) > limits.max_xml_nodes:
            fail('resource-limit', 'O snapshot ultrapassa os limites de células ou mesclagens.')
        positions = []
        for cell in sheet.cells:
            if type(cell) is not Cell:
                fail('invalid-snapshot', 'O snapshot contém uma célula inválida.')
            positions.append(coordinate_key(cell.coordinate))
            if cell.formula is None:
                if cell.literal is None or cell.cache is not None:
                    fail('invalid-snapshot', 'O snapshot contém conteúdo celular inconsistente.')
                validate_value(cell.literal, limits)
            else:
                validate_text(cell.formula, limits, required=True)
                if not cell.formula.startswith('=') or cell.literal is not None or type(cell.cache) is not Cache:
                    fail('invalid-snapshot', 'O snapshot contém uma fórmula inconsistente.')
                if type(cell.cache.state) is not str or cell.cache.state not in ('absent', 'empty', 'value'):
                    fail('invalid-snapshot', 'O snapshot contém um estado de cache inválido.')
                if cell.cache.state == 'value':
                    validate_value(cell.cache.value, limits)
                elif cell.cache.value is not None:
                    fail('invalid-snapshot', 'O snapshot contém um cache inconsistente.')
        if positions != sorted(set(positions)):
            fail('invalid-snapshot', 'O snapshot contém coordenadas repetidas ou fora de ordem.')
        if positions:
            grid_slots += max(p[0] for p in positions) * max(p[1] for p in positions)
        cell_count += len(positions)
        keys = [merge_key(ref) for ref in sheet.merges]
        if keys != sorted(set(keys)):
            fail('invalid-snapshot', 'O snapshot contém mesclagens repetidas ou fora de ordem.')
    if cell_count > limits.max_cells or grid_slots > limits.max_grid_slots:
        fail('resource-limit', 'O snapshot ultrapassa os limites de células ou de grade.')
    if type(snapshot.ignored_parts) is not tuple or type(snapshot.warning_count) is not int or snapshot.warning_count < 0:
        fail('invalid-snapshot', 'O inventário de cobertura do snapshot é inválido.')
    categories = set()
    for item in snapshot.ignored_parts:
        if type(item) is not tuple or len(item) != 2 or type(item[0]) is not str or not item[0] or type(item[1]) is not int or item[1] < 0 or item[0] in categories:
            fail('invalid-snapshot', 'O inventário de cobertura do snapshot é inválido.')
        validate_text(item[0], limits, required=True)
        categories.add(item[0])
    return snapshot
