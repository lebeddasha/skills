"""Synthetic OOXML fixtures only; no client observations."""
import copy
import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET

SCRIPT = Path(__file__).resolve().parents[1] / 'scripts' / 'validate_workbook.py'
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def book(path, values=None, formula='SUM(B1:B2)', merge='D1:E1', extra=None):
    values = values or {}
    cells = {'A1': 'Задание', 'B1': 3, 'B2': 10, 'B3': 3, 'B4': 1,
             'B5': 60, 'B6': 2, 'B7': 0, 'B8': 210,
             'C1': 'средняя', 'C2': 'больше', 'C3': 'средняя',
             'C4': 10, 'C5': 10, 'C6': 10, 'F1': 'План', 'G1': 30}
    cells.update(values)
    root = ET.Element('worksheet', xmlns=NS)
    data = ET.SubElement(root, 'sheetData')
    rows = {}
    for address, value in cells.items():
        rowno = ''.join(c for c in address if c.isdigit())
        row = rows.get(rowno)
        if row is None:
            row = ET.SubElement(data, 'row', r=rowno)
            rows[rowno] = row
        cell = ET.SubElement(row, 'c', r=address)
        if isinstance(value, str):
            cell.set('t', 'inlineStr')
            ET.SubElement(ET.SubElement(cell, 'is'), 't').text = value
        elif value is not None:
            ET.SubElement(cell, 'v').text = str(value)
    cell = ET.SubElement(rows['1'], 'c', r='H1')
    ET.SubElement(cell, 'f').text = formula
    ET.SubElement(cell, 'v').text = '13'
    ET.SubElement(ET.SubElement(root, 'mergeCells', count='1'), 'mergeCell', ref=merge)
    files = {
        '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': f'<workbook xmlns="{NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Лист" sheetId="1" r:id="rId1"/></sheets></workbook>',
        'xl/_rels/workbook.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        'xl/worksheets/sheet1.xml': ET.tostring(root, encoding='utf-8'),
    }
    files.update(extra or {})
    with zipfile.ZipFile(path, 'w') as z:
        for name, value in files.items():
            z.writestr(name, value)


def manifest(original):
    ref = lambda a: 'Лист!' + a
    return {
        'schema_version': 1,
        'original_sha256': hashlib.sha256(original.read_bytes()).hexdigest(),
        'allowed_cells': [ref('F1')],
        'required_cells': [ref('F1')],
        'arithmetic': [{
            'name': 'Длительность', 'operation': 'duration',
            'inputs': {k: ref(v) for k, v in {
                'sets': 'B1', 'reps': 'B2', 'seconds_per_rep': 'B3',
                'sides': 'B4', 'rest_seconds': 'B5', 'rest_count': 'B6',
                'transition_seconds': 'B7'}.items()},
            'result': ref('B8')}],
        'transitions': [{
            'name': 'Синтетическая лестница',
            'require_regression': True,
            'states': [
                {'id': '1', 'kind': 'base', 'values': {'amplitude': ref('C1'), 'reps': ref('C4')}},
                {'id': '2', 'kind': 'progression', 'changed_parameter': 'amplitude', 'reason': 'Учебная проверка', 'values': {'amplitude': ref('C2'), 'reps': ref('C5')}},
                {'id': '3', 'kind': 'regression', 'changed_parameter': 'amplitude', 'reason': 'Условие кейса изменилось', 'values': {'amplitude': ref('C3'), 'reps': ref('C6')}}
            ]}],
        'observations': [{'cell': ref('F1'), 'status': 'planned'}],
    }


class ValidatorTests(unittest.TestCase):
    def setUp(self):
        self.assertTrue(SCRIPT.exists(), 'Validator is not implemented yet')
        spec = importlib.util.spec_from_file_location('validator', SCRIPT)
        self.mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(self.mod)
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.original = self.root / 'original.xlsx'
        self.output = self.root / 'filled.xlsx'
        book(self.original)
        book(self.output, {'F1': 'План: сравнить контроль'})
        self.m = manifest(self.original)

    def run_check(self):
        return self.mod.validate(self.original, self.output, self.m)

    def has(self, code):
        self.assertIn(code, [e['code'] for e in self.run_check()['errors']])

    def test_valid_and_read_only(self):
        before = self.original.read_bytes(), self.output.read_bytes()
        r = self.run_check()
        self.assertEqual(r['errors'], [])
        self.assertEqual(r['status'], 'mechanical_pass')
        self.assertEqual(before, (self.original.read_bytes(), self.output.read_bytes()))

    def test_original_changed(self):
        book(self.original, {'A1': 'Изменен'})
        self.has('original_hash')

    def test_same_path(self):
        r = self.mod.validate(self.original, self.original, self.m)
        self.assertIn('same_file', [e['code'] for e in r['errors']])

    def test_formula_preserved_even_if_allowed(self):
        book(self.output, formula='SUM(B1:B3)')
        self.m['allowed_cells'].append('Лист!H1')
        self.has('formula_changed')

    def test_header_untouched(self):
        book(self.output, {'A1': 'Не тот заголовок'})
        self.has('unexpected_cell_change')

    def test_merge_preserved(self):
        book(self.output, merge='D1:F1')
        self.has('sheet_structure')

    def test_lost_package_part(self):
        book(self.original, extra={'xl/media/image1.png': b'fake'})
        self.m['original_sha256'] = hashlib.sha256(self.original.read_bytes()).hexdigest()
        self.has('package_part')

    def test_wrong_duration(self):
        book(self.output, {'B8': 270})
        self.has('arithmetic')

    def test_two_changes(self):
        book(self.output, {'C5': 8})
        self.has('multiple_changes')

    def test_compare_adjacent_not_baseline(self):
        book(self.output, {'C3': 'средняя', 'C6': 8})
        self.has('multiple_changes')

    def test_empty_required(self):
        book(self.output, {'F1': '   '})
        self.has('required_empty')

    def test_placeholder(self):
        book(self.output, {'F1': 'TODO'})
        self.has('required_empty')

    def test_observed_needs_source(self):
        self.m['observations'][0]['status'] = 'observed'
        self.has('evidence_missing')

    def test_no_manifest_is_incomplete(self):
        r = self.mod.validate(self.original, self.output, None)
        self.assertNotEqual(r['status'], 'mechanical_pass')

    def test_missing_reference(self):
        self.m['arithmetic'][0]['inputs']['sets'] = 'Нет!A1'
        self.has('check_input')

    def test_negative_dose(self):
        book(self.output, {'B1': -3})
        self.has('check_input')

    def test_decimal_comma(self):
        book(self.output, {'B3': '2,5', 'B8': 195})
        self.m['allowed_cells'] += ['Лист!B3', 'Лист!B8']
        self.assertEqual(self.run_check()['errors'], [])

    def test_range_not_silently_parsed(self):
        book(self.output, {'B2': '8-10'})
        self.has('check_input')

    def test_bound_text_range_endpoints(self):
        book(self.output, {'B1': '3 по 8–10', 'B8': '192–210 с'})
        self.m['allowed_cells'] += ['Лист!B1', 'Лист!B8']
        select = lambda cell, text, index: {'cell': cell, 'expected_text': text, 'number_index': index}
        base = self.m['arithmetic'][0]
        base['inputs']['sets'] = select('Лист!B1', '3 по 8–10', 0)
        base['inputs']['reps'] = select('Лист!B1', '3 по 8–10', 1)
        base['result'] = select('Лист!B8', '192–210 с', 0)
        high = copy.deepcopy(base)
        high['inputs']['reps']['number_index'] = 2
        high['result']['number_index'] = 1
        self.m['arithmetic'].append(high)
        self.assertEqual(self.run_check()['errors'], [])

    def test_text_selector_stale_value(self):
        self.m['arithmetic'][0]['inputs']['sets'] = {'cell': 'Лист!B1', 'expected_text': '3 по 8–10', 'number_index': 0}
        self.has('check_input')

    def test_malformed_nested_check(self):
        self.m['arithmetic'] = [None]
        self.has('check_input')

    def test_malformed_transition_values(self):
        self.m['transitions'][0]['states'][1]['values'] = []
        self.has('check_input')

    def test_interval_duration(self):
        c = self.m['arithmetic'][0]
        del c['inputs']['reps']
        del c['inputs']['seconds_per_rep']
        c['inputs']['work_seconds'] = 'Лист!G1'
        self.assertEqual(self.run_check()['errors'], [])

    def test_no_change_is_not_progression(self):
        book(self.output, {'C2': 'средняя'})
        self.has('multiple_changes')

    def test_nonfinite_dose(self):
        book(self.output, {'B1': 'NaN'})
        self.has('check_input')

    def test_negative_text_selector(self):
        book(self.output, {'B1': '-3 подхода'})
        self.m['arithmetic'][0]['inputs']['sets'] = {'cell': 'Лист!B1', 'expected_text': '-3 подхода', 'number_index': 0}
        self.has('check_input')

    def test_product(self):
        self.m['arithmetic'].append({'name': 'Объем', 'operation': 'product', 'inputs': ['Лист!B1', 'Лист!B2'], 'result': 'Лист!G1'})
        self.assertEqual(self.run_check()['errors'], [])

    def test_missing_regression(self):
        self.m['transitions'][0]['states'][2]['kind'] = 'progression'
        self.has('regression_missing')

    def test_missing_parameter(self):
        del self.m['transitions'][0]['states'][1]['values']['reps']
        self.has('parameter_set')

    def test_unknown_manifest_key(self):
        self.m['arithmetci'] = []
        self.has('manifest')

    def test_malformed_workbook(self):
        self.output.write_text('not a zip')
        self.has('workbook_read')

    def test_no_arithmetic_coverage(self):
        self.m['arithmetic'] = []
        self.assertEqual(self.run_check()['status'], 'incomplete')

    def test_cli_report_cannot_overwrite_workbook(self):
        path = self.root / 'checks.json'
        path.write_text(json.dumps(self.m))
        before = self.output.read_bytes()
        r = subprocess.run([sys.executable, str(SCRIPT), 'check', str(self.original), str(self.output), '--manifest', str(path), '--report', str(self.output)], capture_output=True)
        self.assertNotEqual(r.returncode, 0)
        self.assertEqual(self.output.read_bytes(), before)


if __name__ == '__main__':
    unittest.main()
