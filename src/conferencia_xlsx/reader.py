"""Leitura limitada de snapshots XLSX sem salvar ou extrair as entradas."""
from __future__ import annotations

from collections import Counter
from binascii import crc32
from hashlib import sha256
from io import BytesIO
import os
from pathlib import PurePosixPath
import posixpath
import re
import stat
import struct
from urllib.parse import unquote, urlsplit
import warnings
import zipfile
import zlib

from defusedxml import ElementTree as SafeET
from openpyxl import load_workbook
from openpyxl.utils.datetime import from_ISO8601

from .errors import AuditError
from .models import (
    Limits, Value, Cache, Cell, Sheet, Snapshot, canonical_number,
    coordinate_key, column_name, merge_key, normalize_merge,
    validate_limits, validate_snapshot, validate_text, fail,
    validate_iso_date,
)

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CONTENT = 'http://schemas.openxmlformats.org/package/2006/content-types'
XLSX_MAIN = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
WORKSHEET_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
M = '{' + MAIN + '}'


def _part_name(name: str, *, directory=False):
    original = name
    if directory:
        name = name[:-1]
    if (not name or '\\' in name or '\x00' in name or name.startswith('/') or
            any(part in ('', '.', '..') for part in name.split('/')) or
            re.search(r'[\x00-\x1f]', name) or original != original.strip()):
        fail('invalid-package', 'O ZIP contém um caminho interno inválido.')


def _resolve(owner: str, target: str) -> str:
    if not target:
        fail('invalid-relationship', 'Um relacionamento interno é inválido.')
    try:
        uri = urlsplit(target)
        path = unquote(uri.path, encoding='utf-8', errors='strict')
    except (ValueError, UnicodeError):
        fail('invalid-relationship', 'Um relacionamento interno é inválido.')
    if uri.scheme or uri.netloc or uri.query or '\\' in path or '\x00' in path:
        fail('invalid-relationship', 'Um relacionamento interno não é uma parte local.')
    if not path:
        result = owner
    elif path.startswith('/'):
        result = posixpath.normpath(path[1:])
    else:
        result = posixpath.normpath(posixpath.join(posixpath.dirname(owner), path))
    if result in ('', '.', '..') or result.startswith('../'):
        fail('invalid-relationship', 'Um relacionamento sai do pacote local.')
    _part_name(result)
    return result


def _rels_name(owner: str) -> str:
    if not owner:
        return '_rels/.rels'
    return posixpath.join(posixpath.dirname(owner), '_rels', posixpath.basename(owner) + '.rels')


def _rels_owner(name: str) -> str:
    if name == '_rels/.rels':
        return ''
    parent, base = posixpath.split(name)
    if posixpath.basename(parent) != '_rels' or not base.endswith('.rels'):
        fail('invalid-package', 'Uma parte de relacionamentos possui caminho inválido.')
    return posixpath.join(posixpath.dirname(parent), base[:-5])


def _member_bytes(data: bytes, archive, info, remaining: int) -> bytes:
    """Confira a expansão real, sem confiar no tamanho descompactado do diretório."""
    # ZipFile continua validando nomes dos cabeçalhos locais e sobreposição.
    with archive.open(info):
        pass
    header = struct.unpack_from('<4s5H3I2H', data, info.header_offset)
    signature, _, flags, method, _, _, local_crc, local_compressed, local_size, name_size, extra_size = header
    if signature != b'PK\x03\x04' or flags != info.flag_bits or method != info.compress_type:
        fail('invalid-package', 'Os cabeçalhos ZIP da parte não são consistentes.')
    if not flags & 8:
        if local_crc != info.CRC or (local_compressed not in (info.compress_size, 0xffffffff)) or (local_size not in (info.file_size, 0xffffffff)):
            fail('invalid-package', 'Os tamanhos ou CRC dos cabeçalhos ZIP não são consistentes.')
    start = info.header_offset + 30 + name_size + extra_size
    end = start + info.compress_size
    if start < 0 or end > len(data):
        fail('invalid-package', 'Os limites de uma parte ZIP são inválidos.')
    compressed = data[start:end]
    if info.compress_type == zipfile.ZIP_STORED:
        payload = compressed
    else:
        inflater = zlib.decompressobj(-zlib.MAX_WBITS)
        payload = inflater.decompress(compressed, remaining + 1)
        if len(payload) > remaining:
            fail('expanded-limit', 'O pacote ultrapassa o limite real descompactado.')
        if not inflater.eof or inflater.unused_data or inflater.unconsumed_tail:
            fail('invalid-package', 'O fluxo comprimido de uma parte é inválido ou incompleto.')
    if len(payload) > remaining:
        fail('expanded-limit', 'O pacote ultrapassa o limite real descompactado.')
    if len(payload) != info.file_size or crc32(payload) & 0xffffffff != info.CRC:
        fail('invalid-package', 'O tamanho real ou CRC de uma parte ZIP não confere.')
    return payload


def _package(data: bytes, limits: Limits):
    parts = {}
    expanded = actual_expanded = 0
    try:
        with zipfile.ZipFile(BytesIO(data)) as archive:
            infos = archive.infolist()
            if len(infos) > limits.max_parts:
                fail('part-limit', 'O pacote ultrapassa o limite de partes.')
            names = set()
            for info in infos:
                name = info.filename
                if name != info.orig_filename or name in names:
                    fail('invalid-package', 'O ZIP contém partes repetidas ou ambíguas.')
                names.add(name)
                _part_name(name, directory=info.is_dir())
                mode = (info.external_attr >> 16) & 0o170000
                if mode not in (0, stat.S_IFREG, stat.S_IFDIR):
                    fail('invalid-package', 'O ZIP contém uma entrada não regular.')
                if info.flag_bits & (1 | 64):
                    fail('encrypted-package', 'Use uma planilha XLSX sem criptografia.')
                if info.compress_type not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED):
                    fail('unsupported-compression', 'O ZIP usa uma compressão não suportada.')
                expanded += info.file_size
                if expanded > limits.max_expanded_bytes:
                    fail('expanded-limit', 'O pacote ultrapassa o limite descompactado.')
                payload = _member_bytes(data, archive, info, limits.max_expanded_bytes - actual_expanded)
                actual_expanded += len(payload)
                if not info.is_dir():
                    parts[name] = payload
    except AuditError:
        raise
    except (zipfile.BadZipFile, NotImplementedError, RuntimeError, OSError, ValueError, OverflowError, struct.error, zlib.error):
        fail('invalid-package', 'O arquivo não é um pacote XLSX ZIP válido e íntegro.')
    if '[Content_Types].xml' not in parts or '_rels/.rels' not in parts:
        fail('invalid-package', 'O pacote não contém a estrutura XLSX necessária.')

    trees = {}
    nodes = 0
    def parse_xml(name):
        nonlocal nodes
        try:
            parser = SafeET.iterparse(BytesIO(parts[name]), events=('start', 'end'), forbid_dtd=True, forbid_entities=True, forbid_external=True)
            for event, node in parser:
                if event == 'start':
                    nodes += 1
                    if nodes > limits.max_xml_nodes:
                        fail('xml-limit', 'O pacote ultrapassa o limite de elementos XML.')
                    texts = node.attrib.values()
                else:
                    texts = (node.text, node.tail)
                for text in texts:
                    if text is not None:
                        validate_text(text, limits)
            trees[name] = parser.root
        except AuditError:
            raise
        except Exception:
            fail('invalid-xml', 'O pacote contém XML inválido ou recursos XML proibidos.')

    # O tipo de conteúdo, e não apenas o sufixo, determina quais partes são XML.
    parse_xml('[Content_Types].xml')
    content = trees['[Content_Types].xml']
    if content.tag != '{' + CONTENT + '}Types':
        fail('unsupported-format', 'O formato do pacote não é XLSX Transitional.')
    overrides = {}
    defaults = {}
    for node in content:
        if node.tag == '{' + CONTENT + '}Override':
            name = node.get('PartName', '')
            if not name.startswith('/'):
                fail('invalid-package', 'Uma declaração de parte é inválida.')
            name = _resolve('', name)
            if name in overrides or name not in parts:
                fail('invalid-package', 'Uma declaração de parte é ausente ou repetida.')
            overrides[name] = node.get('ContentType', '')
        elif node.tag == '{' + CONTENT + '}Default':
            extension = node.get('Extension', '').casefold()
            if not extension or extension in defaults:
                fail('invalid-package', 'Uma declaração de formato é inválida ou repetida.')
            defaults[extension] = node.get('ContentType', '')
        else:
            fail('invalid-package', 'Uma declaração de formato não é reconhecida.')
        if 'macroenabled' in node.get('ContentType', '').lower() or 'vbaproject' in node.get('ContentType', '').lower():
            fail('unsupported-macros', 'Planilhas com macros não são aceitas.')
    if any('vbaproject' in name.lower() for name in parts):
        fail('unsupported-macros', 'Planilhas com macros não são aceitas.')

    for name in parts:
        if name == '[Content_Types].xml':
            continue
        extension = name.rsplit('.', 1)[-1].casefold()
        content_type = overrides.get(name, defaults.get(extension, '')).lower().split(';', 1)[0].strip()
        if name.endswith(('.xml', '.rels')) or content_type.endswith('+xml') or content_type in ('application/xml', 'text/xml'):
            parse_xml(name)

    relationships = {}
    external_count = 0
    for name, tree in trees.items():
        if not name.endswith('.rels'):
            continue
        owner = _rels_owner(name)
        if owner and owner not in parts:
            fail('invalid-relationship', 'Uma parte de relacionamentos não possui origem.')
        if tree.tag != '{' + PKG + '}Relationships':
            fail('unsupported-format', 'Relacionamentos de outro formato não são suportados.')
        links = {}
        for node in tree:
            if node.tag != '{' + PKG + '}Relationship':
                fail('invalid-relationship', 'Um relacionamento não é reconhecido.')
            identity, kind, target = node.get('Id'), node.get('Type'), node.get('Target')
            if not identity or identity in links or not kind or not target:
                fail('invalid-relationship', 'Um relacionamento é inválido ou repetido.')
            mode = node.get('TargetMode', 'Internal')
            if mode == 'External':
                external_count += 1
                links[identity] = (kind, None)
            elif mode == 'Internal':
                resolved = _resolve(owner, target)
                if resolved not in parts:
                    fail('missing-part', 'Uma referência interna aponta para parte ausente.')
                links[identity] = (kind, resolved)
            else:
                fail('invalid-relationship', 'O modo de um relacionamento não é reconhecido.')
        relationships[owner] = links
    root_links = relationships.get('', {})
    offices = [(kind, target) for kind, target in root_links.values() if kind.endswith('/officeDocument')]
    if len(offices) != 1 or offices[0][1] is None:
        fail('invalid-package', 'O pacote não identifica uma pasta de trabalho local única.')
    kind, workbook_part = offices[0]
    if kind != REL + '/officeDocument' or overrides.get(workbook_part) != XLSX_MAIN:
        fail('unsupported-format', 'Somente pastas XLSX Transitional comuns são suportadas.')
    if [name for name, content_type in overrides.items() if content_type == XLSX_MAIN] != [workbook_part]:
        fail('invalid-workbook', 'O pacote possui declarações ambíguas de pasta de trabalho.')
    workbook = trees.get(workbook_part)
    if workbook is None or workbook.tag != M + 'workbook':
        fail('unsupported-format', 'Somente pastas XLSX Transitional comuns são suportadas.')
    return parts, trees, relationships, overrides, workbook_part, workbook, external_count


def _raw_cells(tree, limits: Limits):
    if tree.tag != M + 'worksheet':
        fail('unsupported-format', 'Uma aba não é uma worksheet XLSX Transitional.')
    sections = tree.findall(M + 'sheetData')
    if len(sections) != 1:
        fail('invalid-worksheet', 'Uma aba não contém uma região celular única.')
    result = {}
    previous_row = 0
    max_row = max_column = 0
    for row in sections[0]:
        if row.tag != M + 'row':
            fail('invalid-worksheet', 'O conteúdo de uma região celular não é reconhecido.')
        raw_row = row.get('r')
        if raw_row is not None and not re.fullmatch('[1-9][0-9]{0,6}', raw_row):
            fail('invalid-coordinate', 'O índice de uma linha é inválido.')
        row_number = int(raw_row) if raw_row is not None else previous_row + 1
        if row_number <= previous_row or row_number > 1048576:
            fail('invalid-coordinate', 'As linhas possuem índices inválidos ou repetidos.')
        previous_row = row_number
        column = 0
        for node in row:
            if node.tag != M + 'c':
                fail('invalid-cell', 'O conteúdo de uma célula não é reconhecido.')
            coordinate = node.get('r')
            if coordinate is None:
                column += 1
                coordinate = f'{column_name(column)}{row_number}'
            else:
                actual_row, column = coordinate_key(coordinate)
                if actual_row != row_number:
                    fail('invalid-coordinate', 'Uma célula não corresponde à linha que a contém.')
            actual_row, actual_column = coordinate_key(coordinate)
            if coordinate in result:
                fail('duplicate-coordinate', 'A aba contém coordenadas celulares repetidas.')
            kind = node.get('t', 'n')
            if kind not in ('n', 's', 'str', 'inlineStr', 'b', 'e', 'd'):
                fail('unsupported-cell', 'Um tipo celular não é suportado.')
            allowed = {M + 'v', M + 'f', M + 'is', M + 'extLst'}
            if any(child.tag not in allowed for child in node):
                fail('unsupported-cell', 'Um conteúdo celular não é suportado.')
            formulas, values, inline = node.findall(M + 'f'), node.findall(M + 'v'), node.findall(M + 'is')
            if len(formulas) > 1 or len(values) > 1 or len(inline) > 1:
                fail('invalid-cell', 'A célula contém elementos de conteúdo repetidos.')
            formula = formulas[0] if formulas else None
            if any(len(child) for child in formulas + values):
                fail('invalid-cell', 'Um valor ou fórmula contém conteúdo interno não reconhecido.')
            if formula is not None and formula.get('t', 'normal') not in ('normal', 'shared'):
                fail('unsupported-formula', 'Fórmulas matriciais ou dataTable não são suportadas.')
            if (formula is not None and inline) or (inline and kind != 'inlineStr') or (kind == 'inlineStr' and values):
                fail('invalid-cell', 'A célula contém representações de valor incompatíveis.')
            if inline:
                _validate_rich_text(inline[0])
            if formula is not None and kind == 'inlineStr':
                fail('unsupported-cell', 'Um cache de fórmula inline não é suportado.')
            if formula is not None and formula.get('t', 'normal') == 'shared':
                if not re.fullmatch('[0-9]+', formula.get('si', '')):
                    fail('invalid-formula', 'Uma fórmula compartilhada não possui identidade válida.')
            result[coordinate] = {'type': kind, 'formula': formula, 'v': values[0] if values else None, 'inline_node': inline[0] if inline else None}
            max_row, max_column = max(max_row, actual_row), max(max_column, actual_column)
            if len(result) > limits.max_cells:
                fail('cell-limit', 'A aba ultrapassa o limite de células presentes.')
    shared_masters = {}
    for coordinate, raw in result.items():
        formula = raw['formula']
        if formula is not None and formula.get('t') == 'shared' and formula.text:
            identity = formula.get('si')
            if identity in shared_masters:
                fail('invalid-formula', 'Uma fórmula compartilhada possui definições repetidas.')
            reference = normalize_merge(formula.get('ref'))
            shared_masters[identity] = merge_key(reference)
    for coordinate, raw in result.items():
        formula = raw['formula']
        if formula is not None and formula.get('t') == 'shared':
            bounds = shared_masters.get(formula.get('si'))
            row, column = coordinate_key(coordinate)
            if bounds is None or not (bounds[0] <= row <= bounds[2] and bounds[1] <= column <= bounds[3]):
                fail('invalid-formula', 'Uma fórmula compartilhada não possui definição ou intervalo correspondente.')
    merges = []
    for section in tree.findall(M + 'mergeCells'):
        for node in section:
            if node.tag != M + 'mergeCell':
                fail('invalid-merge', 'Uma definição de mesclagem não é reconhecida.')
            merges.append(normalize_merge(node.get('ref')))
    if len(merges) != len(set(merges)):
        fail('invalid-merge', 'A aba contém mesclagens repetidas.')
    return result, tuple(sorted(merges, key=merge_key)), max_row, max_column


def _validate_rich_text(node):
    """Não deixar o leitor ignorar conteúdo textual desconhecido silenciosamente."""
    tags = [child.tag for child in node]
    if any(tag not in (M+'t', M+'r', M+'rPh', M+'phoneticPr') for tag in tags) or tags.count(M+'t') > 1 or tags.count(M+'phoneticPr') > 1:
        fail('unsupported-cell', 'Uma estrutura textual não é suportada.')
    if M+'t' in tags and tags.index(M+'t') != 0:
        fail('invalid-cell', 'A ordem de uma estrutura textual é inválida.')
    for child in node:
        if child.tag == M+'t' and len(child):
            fail('unsupported-cell', 'Um texto contém elementos internos não reconhecidos.')
        if child.tag == M+'r':
            run_tags=[part.tag for part in child]
            if any(tag not in (M+'rPr', M+'t') for tag in run_tags) or run_tags.count(M+'t') != 1 or run_tags.count(M+'rPr') > 1:
                fail('unsupported-cell', 'Uma sequência de texto formatado não é suportada.')
            if len(child.find(M+'t')):
                fail('unsupported-cell', 'Um texto contém elementos internos não reconhecidos.')
        if child.tag == M+'rPh':
            if len(child) != 1 or child[0].tag != M+'t' or len(child[0]):
                fail('unsupported-cell', 'Uma anotação fonética não é suportada.')
        if child.tag == M+'phoneticPr' and len(child):
            fail('unsupported-cell', 'Uma anotação fonética não é suportada.')


def _decode_xstring(text: str, limits: Limits) -> str:
    """Decodifique ST_Xstring uma única vez, dentro de cada nó textual.

    Um prefixo _x005F_ protege o sublinhado de um token literal subsequente.
    Sem o sublinhado inicial, x005F_ é texto comum. Pares UTF-16 adjacentes
    no mesmo nó são reunidos; metades isoladas não são Unicode exportável.
    """
    units = re.sub(r'_x([0-9a-fA-F]{4})_', lambda match: chr(int(match[1], 16)), text)
    try:
        value = units.encode('utf-16-le', 'surrogatepass').decode('utf-16-le', 'strict')
    except UnicodeError:
        fail('invalid-value', 'Um texto armazenado contém Unicode inválido.')
    validate_text(value, limits, allow_controls=True)
    return value


def _rich_text_value(node, limits: Limits) -> str:
    # Não juntar os nós antes de decodificar: tokens repartidos entre runs
    # permanecem literais no formato, inclusive quando o texto concatenado
    # parece uma sequência de escape completa.
    snippets = []
    for child in node:
        if child.tag == M+'t':
            snippets.append(_decode_xstring(child.text or '', limits))
        elif child.tag == M+'r':
            snippets.append(_decode_xstring(child.find(M+'t').text or '', limits))
    value = ''.join(snippets)
    validate_text(value, limits, allow_controls=True)
    return value


def _stored_value(kind: str, text: str, shared_strings, limits: Limits) -> Value:
    if kind == 'n':
        return Value('number', canonical_number(text))
    if kind == 'b':
        if text not in ('0', '1'):
            fail('invalid-boolean', 'Um booleano armazenado não é reconhecido.')
        return Value('boolean', text == '1')
    if kind == 's':
        if not re.fullmatch('[0-9]+', text):
            fail('invalid-string', 'Um índice de texto compartilhado é inválido.')
        try:
            text = shared_strings[int(text)]
        except (IndexError, ValueError, OverflowError):
            fail('invalid-string', 'Um índice de texto compartilhado não existe.')
        validate_text(text, limits, allow_controls=True)
        return Value('string', text)
    if kind == 'str':
        return Value('string', _decode_xstring(text, limits))
    if kind == 'e':
        return Value('error', _decode_xstring(text, limits))
    if kind == 'd':
        validate_iso_date(text)
        try:
            if from_ISO8601(text) is None:
                raise ValueError()
        except Exception:
            fail('invalid-date', 'Uma data ISO armazenada não é reconhecida.')
        return Value('date', text)
    fail('unsupported-cell', 'Um valor armazenado não é suportado.')


def _cache(raw, shared_strings, limits: Limits) -> Cache:
    node = raw['v']
    if node is None:
        return Cache('absent')
    if node.text is None or node.text == '':
        return Cache('empty')
    return Cache('value', _stored_value(raw['type'], node.text, shared_strings, limits))


def read_bytes(data: bytes, limits: Limits = Limits()) -> Snapshot:
    validate_limits(limits)
    if type(data) is not bytes:
        fail('invalid-input', 'Forneça bytes de uma planilha XLSX.')
    if not data or len(data) > limits.max_file_bytes:
        fail('file-limit', 'A entrada está vazia ou ultrapassa o limite de bytes.')
    try:
        package = _package(data, limits)
        parts, trees, relationships, overrides, workbook_part, workbook_tree, external_count = package
        links = relationships.get(workbook_part, {})
        sheet_sections = workbook_tree.findall(M + 'sheets')
        if len(sheet_sections) != 1:
            fail('invalid-workbook', 'A pasta não contém uma lista única de abas.')
        sheets_raw = list(sheet_sections[0])
        if not sheets_raw or len(sheets_raw) > limits.max_sheets:
            fail('sheet-limit', 'A pasta precisa de abas dentro do limite configurado.')
        prepared, names, targets, ids = [], set(), set(), set()
        total_cells = total_slots = 0
        for sheet in sheets_raw:
            if sheet.tag != M + 'sheet':
                fail('invalid-workbook', 'Uma declaração de aba não é reconhecida.')
            name, identity = sheet.get('name'), sheet.get('{' + REL + '}id')
            validate_text(name, limits, required=True)
            if len(name) > 31 or re.search(r'[\\/*?:\[\]]', name) or name.casefold() in names or not identity or identity in ids:
                fail('invalid-workbook', 'As abas possuem nomes ou identidades inválidos ou repetidos.')
            names.add(name.casefold()); ids.add(identity)
            kind, target = links.get(identity, (None, None))
            if kind != REL + '/worksheet' or target is None:
                fail('unsupported-sheet', 'Somente abas worksheet locais são suportadas.')
            if target in targets or target not in trees or overrides.get(target) != WORKSHEET_TYPE:
                fail('invalid-workbook', 'Uma aba possui parte ausente, repetida ou inválida.')
            targets.add(target)
            state = sheet.get('state', 'visible')
            if state not in ('visible', 'hidden', 'veryHidden'):
                fail('invalid-workbook', 'Uma aba possui visibilidade inválida.')
            raw, merges, max_row, max_column = _raw_cells(trees[target], limits)
            total_cells += len(raw); total_slots += max_row * max_column
            if total_cells > limits.max_cells or total_slots > limits.max_grid_slots:
                fail('grid-limit', 'A pasta ultrapassa o limite de células ou slots da grade.')
            prepared.append((name, state, target, raw, merges, max_row, max_column))

        strings_links = [target for kind, target in links.values() if kind == REL + '/sharedStrings']
        if len(strings_links) > 1 or (strings_links and strings_links[0] is None):
            fail('invalid-workbook', 'O catálogo de textos compartilhados é inválido.')
        if strings_links:
            strings_tree = trees.get(strings_links[0])
            if strings_tree is None or strings_tree.tag != M+'sst':
                fail('invalid-string', 'O catálogo de textos compartilhados é inválido.')
            for node in strings_tree:
                if node.tag == M+'si':
                    _validate_rich_text(node)
                elif node.tag != M+'extLst':
                    fail('unsupported-cell', 'O catálogo contém uma estrutura textual não suportada.')
        # openpyxl 3.1.5 remove o literal "x005F_" incondicionalmente no SST.
        # A recuperação semântica usa os nós originais validados, sem esse
        # apagamento, enquanto o leitor continua resolvendo a grade/fórmulas.
        shared_strings = [_rich_text_value(node, limits) for node in strings_tree if node.tag == M+'si'] if strings_links else []
        results = []
        with warnings.catch_warnings(record=True) as captured:
            warnings.simplefilter('always')
            book = load_workbook(BytesIO(data), read_only=True, data_only=False, keep_links=False, rich_text=False)
            try:
                if book.sheetnames != [item[0] for item in prepared]:
                    fail('reader-mismatch', 'O leitor não confirmou a identidade das abas.')
                epoch = '1904' if book.epoch.year == 1904 else '1900'
                for name, state, target, raw, merges, max_row, max_column in prepared:
                    worksheet = book[name]
                    worksheet.reset_dimensions()
                    cells, seen = [], set()
                    if raw:
                        for row in worksheet.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_column):
                            for interpreted in row:
                                coordinate = getattr(interpreted, 'coordinate', None)
                                if coordinate not in raw:
                                    continue
                                if coordinate in seen:
                                    fail('reader-mismatch', 'O leitor produziu coordenadas celulares repetidas.')
                                seen.add(coordinate)
                                source = raw[coordinate]
                                if source['formula'] is not None:
                                    formula = interpreted.value
                                    if interpreted.data_type != 'f' or type(formula) is not str or not formula.startswith('='):
                                        fail('unsupported-formula', 'O leitor não resolveu uma fórmula comum ou compartilhada.')
                                    validate_text(formula, limits)
                                    cells.append(Cell(coordinate, formula=formula, cache=_cache(source, shared_strings, limits)))
                                elif source['type'] == 'inlineStr':
                                    if interpreted.data_type not in ('s', 'inlineStr'):
                                        fail('reader-mismatch', 'O leitor não confirmou uma célula textual.')
                                    inline = source['inline_node']
                                    text = _rich_text_value(inline, limits) if inline is not None else ''
                                    cells.append(Cell(coordinate, literal=Value('string', text)))
                                elif source['type'] == 'str' or (source['v'] is not None and source['v'].text not in (None, '')):
                                    text = source['v'].text if source['v'] is not None else ''
                                    value = _stored_value(source['type'], text or '', shared_strings, limits)
                                    if source['type'] in ('s', 'str') and interpreted.data_type not in ('s', 'str'):
                                        fail('reader-mismatch', 'O leitor não confirmou uma célula textual.')
                                    cells.append(Cell(coordinate, literal=value))
                    if seen != set(raw):
                        fail('reader-mismatch', 'O leitor não cobriu todas as células presentes.')
                    results.append(Sheet(name, state, tuple(sorted(cells, key=lambda cell: coordinate_key(cell.coordinate))), merges))
                warning_count = len(captured)
            finally:
                book.close()
        ignored = Counter({'external_relationships': external_count})
        prefixes = [('xl/drawings/', 'drawing_parts'), ('xl/charts/', 'chart_parts'), ('xl/media/', 'image_parts'), ('xl/tables/', 'table_parts'), ('xl/comments', 'comment_parts'), ('_xmlsignatures/', 'signature_parts'), ('xl/externalLinks/', 'external_link_parts'), ('xl/theme/', 'formatting_parts')]
        consumed = {'[Content_Types].xml', workbook_part, *targets}
        consumed.update(name for name in parts if name.endswith('.rels'))
        consumed.update(strings_links)
        for name in parts:
            if name in consumed:
                continue
            category = next((category for prefix, category in prefixes if name.startswith(prefix)), 'other_parts')
            if name.endswith('/styles.xml'):
                category = 'formatting_parts'
            ignored[category] += 1
        ignored['defined_names'] = len(workbook_tree.findall(M + 'definedNames/' + M + 'definedName'))
        for _, _, target, _, _, _, _ in prepared:
            tree = trees[target]
            ignored['hyperlinks'] += len(tree.findall(M + 'hyperlinks/' + M + 'hyperlink'))
            ignored['data_validations'] += len(tree.findall(M + 'dataValidations/' + M + 'dataValidation'))
            ignored['conditional_formats'] += len(tree.findall(M + 'conditionalFormatting'))
        snapshot = Snapshot(sha256(data).hexdigest(), len(data), epoch, tuple(results), tuple(sorted(ignored.items())), warning_count)
        return validate_snapshot(snapshot, limits)
    except AuditError:
        raise
    except Exception:
        fail('invalid-workbook', 'A estrutura da planilha não pôde ser lida de forma completa.')


def read_path(path, limits: Limits = Limits()) -> Snapshot:
    validate_limits(limits)
    try:
        # O_NONBLOCK prevents a named pipe from blocking before fstat rejects it.
        descriptor = os.open(path, os.O_RDONLY | getattr(os, 'O_BINARY', 0) | getattr(os, 'O_NONBLOCK', 0))
        with os.fdopen(descriptor, 'rb') as stream:
            before = os.fstat(stream.fileno())
            if not stat.S_ISREG(before.st_mode):
                fail('invalid-file', 'A entrada precisa ser um arquivo regular.')
            if before.st_size > limits.max_file_bytes:
                fail('file-limit', 'A entrada ultrapassa o limite de bytes.')
            data = stream.read(limits.max_file_bytes + 1)
            after = os.fstat(stream.fileno())
            signature = lambda value: (value.st_dev, value.st_ino, value.st_size, value.st_mtime_ns, value.st_ctime_ns)
            if signature(before) != signature(after) or len(data) != before.st_size:
                fail('input-changed', 'A entrada mudou durante a captura; tente novamente.')
    except AuditError:
        raise
    except (OSError, TypeError, ValueError, OverflowError):
        fail('input-read', 'Não foi possível capturar o arquivo de entrada.')
    return read_bytes(data, limits)
