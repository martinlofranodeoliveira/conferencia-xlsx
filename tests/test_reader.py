import dataclasses
import datetime
from decimal import localcontext
from hashlib import sha256
from io import BytesIO
import os
from pathlib import Path
import stat
import struct
import tempfile
import unittest
from unittest.mock import patch
from types import SimpleNamespace
import warnings
import zipfile
from binascii import crc32
from xml.etree import ElementTree as ET

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.datetime import CALENDAR_MAC_1904

from conferencia_xlsx.errors import AuditError
from conferencia_xlsx.models import Limits, Value, Cache, canonical_number
from conferencia_xlsx.reader import read_bytes, read_path, MAIN, REL, PKG, M
import conferencia_xlsx.reader as reader_module


def workbook_bytes(*, iso=False, epoch=None):
    book = Workbook(iso_dates=iso)
    if epoch:
        book.epoch = epoch
    sheet = book.active
    sheet.title = 'Dados'
    sheet['A1'] = 1
    sheet['B1'] = '1'
    sheet['C1'] = True
    sheet['D1'] = '=SUM(A1,1)'
    sheet['E1'] = ''
    sheet['F1'] = datetime.datetime(2026, 11, 9, 12, 30, 15, 123000)
    sheet['G1'] = '=SUM(1,2)'
    sheet['G1'].data_type = 's'
    sheet['H1'].font = Font(bold=True)
    sheet['A3'] = 'mesclagem'
    sheet.merge_cells('A3:B3')
    output = BytesIO()
    book.save(output)
    book.close()
    return output.getvalue()


def rewrite(data, replacements=None, xml=None, extra=()):
    output = BytesIO()
    with zipfile.ZipFile(BytesIO(data)) as source, zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as dest:
        for name in source.namelist():
            payload = source.read(name)
            if replacements and name in replacements:
                payload = replacements[name]
            if xml and name in xml:
                tree = ET.fromstring(payload)
                xml[name](tree)
                payload = ET.tostring(tree, encoding='utf-8', xml_declaration=True)
            dest.writestr(name, payload)
        for info, payload in extra:
            dest.writestr(info, payload)
    return output.getvalue()


def cell(tree, coordinate):
    return tree.find(f'.//{M}c[@r="{coordinate}"]')


def map_cells(snapshot):
    return {item.coordinate: item for item in snapshot.sheets[0].cells}


def shared_string_fixture(text, *, part='xl/sharedStrings.xml', related=True):
    """Uma pequena string compartilhada original, sem dados de terceiros."""
    table=ET.Element(M+'sst');entry=ET.SubElement(table,M+'si')
    if isinstance(text,tuple):
        for piece in text:ET.SubElement(ET.SubElement(entry,M+'r'),M+'t').text=piece
    else:ET.SubElement(entry,M+'t').text=text
    def cells(tree):
        node=cell(tree,'B1');node.clear();node.set('r','B1');node.set('t','s');ET.SubElement(node,M+'v').text='0'
    def content(tree):ET.SubElement(tree,'{http://schemas.openxmlformats.org/package/2006/content-types}Override',PartName='/'+part,ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml')
    def relationships(tree):
        if related:ET.SubElement(tree,'{'+PKG+'}Relationship',Id='strings',Type=REL+'/sharedStrings',Target=part.removeprefix('xl/'))
    mutations={'[Content_Types].xml':content,'xl/_rels/workbook.xml.rels':relationships}
    if related:mutations['xl/worksheets/sheet1.xml']=cells
    return rewrite(workbook_bytes(),xml=mutations,extra=[(part,ET.tostring(table,encoding='utf-8'))])


class ReaderTests(unittest.TestCase):
    def test_typed_literals_formula_literal_and_empty_string_survive(self):
        snapshot = read_bytes(workbook_bytes())
        cells = map_cells(snapshot)
        self.assertEqual(cells['A1'].literal, Value('number', '1'))
        self.assertEqual(cells['B1'].literal, Value('string', '1'))
        self.assertEqual(cells['C1'].literal, Value('boolean', True))
        self.assertEqual(cells['D1'].formula, '=SUM(A1,1)')
        self.assertEqual(cells['D1'].cache, Cache('empty'))
        self.assertEqual(cells['E1'].literal, Value('string', ''))
        self.assertEqual(cells['G1'].literal, Value('string', '=SUM(1,2)'))
        self.assertNotIn('H1', cells)
        self.assertEqual(snapshot.sheets[0].merges, ('A3:B3',))

    def test_numeric_precision_is_exact_and_decimal_context_independent(self):
        number = '900719925474099312345.123456789012345678901'
        def mutate(tree):
            cell(tree, 'A1').find(M+'v').text = number
        data = rewrite(workbook_bytes(), xml={'xl/worksheets/sheet1.xml': mutate})
        with localcontext() as context:
            context.prec = 2
            value = map_cells(read_bytes(data))['A1'].literal
            self.assertEqual(value, Value('number', number))
            self.assertEqual(canonical_number('1.0'), '1')
            self.assertEqual(canonical_number('1e0'), '1')
            self.assertEqual(canonical_number('-0.000'), '0')
            self.assertEqual(canonical_number('2.370'), '2.37')

    def test_formula_cache_absent_empty_and_typed_value_are_distinct(self):
        original = workbook_bytes()
        absent = rewrite(original, xml={'xl/worksheets/sheet1.xml': lambda t: cell(t, 'D1').remove(cell(t, 'D1').find(M+'v'))})
        def valued(tree):
            cell(tree, 'D1').find(M+'v').text = '900719925474099312345'
        present = rewrite(original, xml={'xl/worksheets/sheet1.xml': valued})
        self.assertEqual(map_cells(read_bytes(absent))['D1'].cache, Cache('absent'))
        self.assertEqual(map_cells(read_bytes(original))['D1'].cache, Cache('empty'))
        self.assertEqual(map_cells(read_bytes(present))['D1'].cache, Cache('value', Value('number', '900719925474099312345')))

    def test_error_boolean_string_date_caches_are_not_numbers(self):
        cases = [('e', '#DIV/0!', Value('error', '#DIV/0!')), ('b', '0', Value('boolean', False)), ('str', 'texto', Value('string', 'texto')), ('d', '2026-11-09', Value('date', '2026-11-09'))]
        for kind, text, expected in cases:
            def mutate(tree):
                node = cell(tree, 'D1'); node.set('t', kind); node.find(M+'v').text = text
            with self.subTest(kind=kind):
                snapshot = read_bytes(rewrite(workbook_bytes(), xml={'xl/worksheets/sheet1.xml': mutate}))
                self.assertEqual(map_cells(snapshot)['D1'].cache, Cache('value', expected))

    def test_iso_dates_preserve_stored_representation_and_epoch_is_separate(self):
        snapshot = read_bytes(workbook_bytes(iso=True))
        self.assertEqual(map_cells(snapshot)['F1'].literal, Value('date', '2026-11-09T12:30:15.123'))
        self.assertEqual(snapshot.epoch, '1900')
        mac = read_bytes(workbook_bytes(epoch=CALENDAR_MAC_1904))
        self.assertEqual(mac.epoch, '1904')
        self.assertEqual(map_cells(mac)['F1'].literal.kind, 'number')

    def test_serial_59_and_60_are_not_collapsed_by_date_conversion(self):
        def serial(text):
            def mutate(tree): cell(tree, 'F1').find(M+'v').text = text
            return map_cells(read_bytes(rewrite(workbook_bytes(), xml={'xl/worksheets/sheet1.xml': mutate})))['F1'].literal
        self.assertEqual(serial('59'), Value('number','59'))
        self.assertEqual(serial('60'), Value('number','6E+1'))

    def test_shared_formulas_expand_through_openpyxl(self):
        def mutate(tree):
            first = cell(tree, 'D1').find(M+'f'); first.set('t','shared'); first.set('si','0'); first.set('ref','D1:D2'); first.text='A1+1'
            row=ET.SubElement(tree.find(M+'sheetData'),M+'row',r='2')
            # Restore row order before the existing row3.
            sheet_data=tree.find(M+'sheetData');sheet_data.remove(row);sheet_data.insert(1,row)
            second=ET.SubElement(row,M+'c',r='D2');ET.SubElement(second,M+'f',t='shared',si='0');ET.SubElement(second,M+'v').text='3'
        cells=map_cells(read_bytes(rewrite(workbook_bytes(),xml={'xl/worksheets/sheet1.xml':mutate})))
        self.assertEqual(cells['D1'].formula,'=A1+1');self.assertEqual(cells['D2'].formula,'=A2+1')
        self.assertEqual(cells['D2'].cache,Cache('value',Value('number','3')))

    def test_malformed_shared_formula_or_special_formula_is_explicit_error(self):
        for kind in ['array','dataTable','unknown','shared']:
            def mutate(tree): cell(tree,'D1').find(M+'f').set('t',kind)
            with self.subTest(kind=kind),self.assertRaises(AuditError):
                read_bytes(rewrite(workbook_bytes(),xml={'xl/worksheets/sheet1.xml':mutate}))

    def test_dimension_is_not_trusted_and_implicit_coordinates_are_supported(self):
        def mutate(tree):
            tree.find(M+'dimension').set('ref','A1:A1')
            first=tree.find(M+'sheetData/'+M+'row');first.attrib.pop('r')
            for node in first:node.attrib.pop('r')
        snapshot=read_bytes(rewrite(workbook_bytes(),xml={'xl/worksheets/sheet1.xml':mutate}))
        self.assertIn('G1',map_cells(snapshot));self.assertIn('A3',map_cells(snapshot))

    def test_sparse_grid_and_cell_limits_are_checked_before_library_load(self):
        def mutate(tree):
            row=ET.SubElement(tree.find(M+'sheetData'),M+'row',r='1000');node=ET.SubElement(row,M+'c',r='Z1000');ET.SubElement(node,M+'v').text='1'
        data=rewrite(workbook_bytes(),xml={'xl/worksheets/sheet1.xml':mutate})
        with patch.object(reader_module,'load_workbook',side_effect=AssertionError('must not load')) as load:
            with self.assertRaises(AuditError) as caught:read_bytes(data,Limits(max_grid_slots=100))
            self.assertEqual(caught.exception.code,'grid-limit');load.assert_not_called()
        with self.assertRaises(AuditError):read_bytes(workbook_bytes(),Limits(max_cells=2))

    def test_node_budget_aborts_during_iterparse_before_full_tree(self):
        original=reader_module.SafeET.iterparse;starts=[]
        def observe(*args,**kwargs):
            class Tracking:
                def __init__(self):self.parser=original(*args,**kwargs)
                def __iter__(self):return self
                def __next__(self):
                    event,node=next(self.parser)
                    if event=='start':starts.append(node.tag)
                    return event,node
                @property
                def root(self):return self.parser.root
            return Tracking()
        with patch.object(reader_module.SafeET,'iterparse',side_effect=observe):
            with self.assertRaises(AuditError) as caught:read_bytes(workbook_bytes(),Limits(max_xml_nodes=4))
        self.assertEqual(caught.exception.code,'xml-limit');self.assertEqual(len(starts),5)

    def test_dtd_entities_and_invalid_xml_do_not_echo_content(self):
        for payload in [b'<!DOCTYPE x [<!ENTITY bad "PRIVATE-MARKER">]><x>&bad;</x>',b'<x>PRIVATE-MARKER']:
            with self.assertRaises(AuditError) as caught:read_bytes(rewrite(workbook_bytes(),replacements={'xl/workbook.xml':payload}))
            self.assertEqual(caught.exception.code,'invalid-xml');self.assertNotIn('PRIVATE-MARKER',str(caught.exception))

    def test_duplicate_traversal_symlink_compression_and_crc_are_rejected(self):
        data=workbook_bytes()
        duplicate=zipfile.ZipInfo('xl/workbook.xml')
        symlink=zipfile.ZipInfo('link');symlink.create_system=3;symlink.external_attr=(stat.S_IFLNK|0o777)<<16
        compressed=zipfile.ZipInfo('unsupported');compressed.compress_type=zipfile.ZIP_BZIP2
        for extra in [[(duplicate,b'x')],[('../escape',b'x')],[(symlink,b'target')],[(compressed,b'x')]]:
            with warnings.catch_warnings():
                warnings.simplefilter('ignore')
                damaged=rewrite(data,extra=extra)
            with self.assertRaises(AuditError):read_bytes(damaged)
        output=BytesIO()
        with zipfile.ZipFile(BytesIO(data)) as source,zipfile.ZipFile(output,'w',zipfile.ZIP_STORED) as dest:
            for name in source.namelist():dest.writestr(name,source.read(name))
        corrupt=output.getvalue().replace(b'mesclagem',b'alteracao',1)
        with self.assertRaises(AuditError) as caught:read_bytes(corrupt)
        self.assertEqual(caught.exception.code,'invalid-package')

    def test_encrypted_flag_and_expanded_part_file_limits_are_explicit(self):
        data=bytearray(workbook_bytes());offset=data.index(b'PK\x01\x02');data[offset+8]|=1
        with self.assertRaises(AuditError) as caught:read_bytes(bytes(data))
        self.assertEqual(caught.exception.code,'encrypted-package')
        for limits in [Limits(max_file_bytes=10),Limits(max_expanded_bytes=10),Limits(max_parts=1)]:
            with self.subTest(limits=limits),self.assertRaises(AuditError):read_bytes(workbook_bytes(),limits)

    def test_missing_internal_relation_and_external_relationship_inventory(self):
        relpart='xl/_rels/workbook.xml.rels'
        def missing(tree):tree[0].set('Target','worksheets/missing.xml')
        with self.assertRaises(AuditError) as caught:read_bytes(rewrite(workbook_bytes(),xml={relpart:missing}))
        self.assertEqual(caught.exception.code,'missing-part')
        def external(tree):ET.SubElement(tree,'{'+PKG+'}Relationship',Id='external',Type=REL+'/hyperlink',Target='https://example.invalid/private',TargetMode='External')
        snapshot=read_bytes(rewrite(workbook_bytes(),xml={relpart:external}))
        self.assertEqual(dict(snapshot.ignored_parts)['external_relationships'],1)

    def test_strict_macro_and_chartsheet_are_not_silently_skipped(self):
        def strict(tree):tree.tag='{http://purl.oclc.org/ooxml/spreadsheetml/main}workbook'
        with self.assertRaises(AuditError) as caught:read_bytes(rewrite(workbook_bytes(),xml={'xl/workbook.xml':strict}))
        self.assertEqual(caught.exception.code,'unsupported-format')
        with self.assertRaises(AuditError) as caught:read_bytes(rewrite(workbook_bytes(),extra=[('xl/vbaProject.bin',b'fake')]))
        self.assertEqual(caught.exception.code,'unsupported-macros')
        def chart(tree):tree[0].set('Type',REL+'/chartsheet')
        with self.assertRaises(AuditError) as caught:read_bytes(rewrite(workbook_bytes(),xml={'xl/_rels/workbook.xml.rels':chart}))
        self.assertEqual(caught.exception.code,'unsupported-sheet')

    def test_duplicate_coords_sheet_names_and_unknown_content_reject(self):
        def duplicate(tree):
            first=tree.find(M+'sheetData/'+M+'row');first.append(ET.fromstring(ET.tostring(first[0])))
        def unknown(tree):cell(tree,'A1').set('t','unrecognized')
        def nested(tree):ET.SubElement(cell(tree,'A1').find(M+'v'),M+'x').text='ignored'
        for mutate in [duplicate,unknown,nested]:
            with self.subTest(mutate=mutate.__name__),self.assertRaises(AuditError):read_bytes(rewrite(workbook_bytes(),xml={'xl/worksheets/sheet1.xml':mutate}))

    def test_unknown_nested_text_is_rejected_instead_of_becoming_empty(self):
        def mutate(tree):
            node=cell(tree,'B1').find(M+'is');node.clear();ET.SubElement(node,M+'unknown').text='PRIVATE-NESTED'
        with self.assertRaises(AuditError) as caught:read_bytes(rewrite(workbook_bytes(),xml={'xl/worksheets/sheet1.xml':mutate}))
        self.assertEqual(caught.exception.code,'unsupported-cell');self.assertNotIn('PRIVATE-NESTED',str(caught.exception))

    def test_literal_x005f_is_preserved_and_not_equal_to_empty_shared_text(self):
        from conferencia_xlsx.compare import compare_snapshots
        literal=read_bytes(shared_string_fixture('x005F_'));empty=read_bytes(shared_string_fixture(''))
        self.assertEqual(map_cells(literal)['B1'].literal,Value('string','x005F_'))
        result=compare_snapshots(literal,empty)
        self.assertEqual(result['status'],'different_in_scope')
        self.assertEqual([(x['kind'],x['coordinate']) for x in result['changes']],[('literal','B1')])

    def test_xstring_decodes_once_per_text_node_and_preserves_unicode(self):
        cases=[('_x0041_','A'),('_x005F_x0041_','_x0041_'),('_xD83D__xDE00_','😀'),('_x000D_','\r'),('_x000A_','\n'),('_x0008_','\b'),(('_x004','1_'),'_x0041_')]
        for encoded,expected in cases:
            with self.subTest(encoded=encoded):
                self.assertEqual(map_cells(read_bytes(shared_string_fixture(encoded)))['B1'].literal,Value('string',expected))
        with self.assertRaises(AuditError) as caught:read_bytes(shared_string_fixture('_xD83D_'))
        self.assertEqual(caught.exception.code,'invalid-value')

    def test_xml_contenttype_parts_have_same_budget_and_entity_guard_with_other_suffix(self):
        # Nomes de partes alternativos são aceitos; suas declarações MIME
        # precisam receber as mesmas verificações de XML antes do openpyxl.
        part='xl/sharedStrings.bin';data=shared_string_fixture('x005F_',part=part)
        self.assertEqual(map_cells(read_bytes(data))['B1'].literal,Value('string','x005F_'))
        payload=b'<!DOCTYPE sst><sst xmlns="'+MAIN.encode()+b'"><si><t>texto</t></si></sst>'
        with patch.object(reader_module,'load_workbook',side_effect=AssertionError('must not load')) as load:
            with self.assertRaises(AuditError) as caught:read_bytes(rewrite(data,replacements={part:payload}))
            self.assertEqual(caught.exception.code,'invalid-xml');load.assert_not_called()
        payload=b'<sst xmlns="'+MAIN.encode()+b'">'+b'<si><t>texto</t></si>'*400+b'</sst>'
        with patch.object(reader_module,'load_workbook',side_effect=AssertionError('must not load')) as load:
            with self.assertRaises(AuditError) as caught:read_bytes(rewrite(data,replacements={part:payload}),Limits(max_xml_nodes=500))
            self.assertEqual(caught.exception.code,'xml-limit');load.assert_not_called()

    def test_actual_deflate_length_and_crc_are_checked_beyond_advertised_size(self):
        # Uma amostra pequena corrompida declara só o primeiro byte de uma
        # parte: o restante do fluxo precisa ser conferido, não truncado.
        data=bytearray(rewrite(workbook_bytes(),extra=[('extra.bin',b'abc'*1024)]))
        with zipfile.ZipFile(BytesIO(data)) as archive:
            info=archive.getinfo('extra.bin');local=info.header_offset
        crc=crc32(b'a')&0xffffffff
        struct.pack_into('<I',data,local+14,crc);struct.pack_into('<I',data,local+22,1)
        central=data.index(b'PK\x01\x02');found=False
        while data[central:central+4]==b'PK\x01\x02':
            name_size,extra_size,comment_size=struct.unpack_from('<HHH',data,central+28)
            name=bytes(data[central+46:central+46+name_size])
            if name==b'extra.bin':
                struct.pack_into('<I',data,central+16,crc);struct.pack_into('<I',data,central+24,1);found=True;break
            central+=46+name_size+extra_size+comment_size
        self.assertTrue(found)
        with zipfile.ZipFile(BytesIO(data)) as archive:
            declared=sum(info.file_size for info in archive.infolist())
            self.assertEqual(archive.read('extra.bin'),b'a')
        with self.assertRaises(AuditError) as caught:read_bytes(bytes(data),Limits(max_expanded_bytes=declared))
        self.assertIn(caught.exception.code,('expanded-limit','invalid-package'))

    def test_read_path_captures_once_and_leaves_input_bytes_unchanged(self):
        data=workbook_bytes()
        with tempfile.TemporaryDirectory() as temp:
            path=Path(temp)/'entrada.xlsx';path.write_bytes(data);snapshot=read_path(path)
            self.assertEqual(snapshot.sha256,sha256(data).hexdigest());self.assertEqual(path.read_bytes(),data)
            with self.assertRaises(AuditError):read_path(Path(temp))
            if hasattr(os,'mkfifo'):
                pipe=Path(temp)/'pipe';os.mkfifo(pipe)
                with self.assertRaises(AuditError) as caught:read_path(pipe)
                self.assertEqual(caught.exception.code,'invalid-file')

    def test_descriptor_change_is_detected_by_matching_fstat_api(self):
        with tempfile.TemporaryDirectory() as temp:
            path=Path(temp)/'entrada.xlsx';path.write_bytes(workbook_bytes());actual=os.fstat;calls=[]
            def changed(fd):
                result=actual(fd);calls.append(fd)
                if len(calls)==2:
                    attrs={name:getattr(result,name) for name in ['st_mode','st_dev','st_ino','st_size','st_mtime_ns','st_ctime_ns']};attrs['st_size']+=1;return SimpleNamespace(**attrs)
                return result
            with patch.object(reader_module.os,'fstat',side_effect=changed):
                with self.assertRaises(AuditError) as caught:read_path(path)
            self.assertEqual(caught.exception.code,'input-changed');self.assertEqual(len(calls),2)

    def test_bad_values_limits_and_input_types_have_sanitized_errors(self):
        for text in ['NaN','Infinity','1e999999999','1 2']:
            with self.subTest(text=text),self.assertRaises(AuditError):canonical_number(text)
        for value in [None,'not bytes',bytearray(b'PK')]:
            with self.subTest(value=value),self.assertRaises(AuditError):read_bytes(value)
        with self.assertRaises(AuditError):read_bytes(workbook_bytes(),Limits(max_cells=True))
        def badboolean(tree):cell(tree,'C1').find(M+'v').text='2'
        with self.assertRaises(AuditError) as caught:read_bytes(rewrite(workbook_bytes(),xml={'xl/worksheets/sheet1.xml':badboolean}))
        self.assertEqual(caught.exception.code,'invalid-boolean')


if __name__ == '__main__':
    unittest.main()
