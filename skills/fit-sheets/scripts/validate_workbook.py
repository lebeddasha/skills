#!/usr/bin/env python3
"""Read-only XLSX/XLSM checks. Python 3.9+, standard library only.

No Excel formula evaluation, macros, network calls, or workbook writes.
Mechanical checks are not a clinical or semantic approval.
"""
import argparse
import copy
from decimal import Decimal, InvalidOperation
import hashlib
import json
import math
from pathlib import Path
import posixpath
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
Q = lambda tag: '{' + NS + '}' + tag
EMPTY = {'', '-', '—', '...', '…', 'todo', 'tbd', 'заполнить'}
MAX_SIZE = 200 * 1024 * 1024


def digest(path):
    h = hashlib.sha256()
    with open(path, 'rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def canonical(node):
    """Ignore namespace prefixes and indentation, retain semantic XML structure."""
    return (node.tag, tuple(sorted(node.attrib.items())),
            (node.text or '').strip() if list(node) else (node.text or ''),
            tuple(canonical(child) for child in node))


def xml(data):
    if b'<!DOCTYPE' in data.upper() or b'<!ENTITY' in data.upper():
        raise ValueError('DTD/ENTITY не поддерживаются')
    return ET.fromstring(data)


class Workbook:
    def __init__(self, path):
        if Path(path).suffix.lower() not in {'.xlsx', '.xlsm'}:
            raise ValueError('Поддерживаются только .xlsx и .xlsm')
        with zipfile.ZipFile(path) as archive:
            infos = archive.infolist()
            if len(infos) > 10000 or sum(i.file_size for i in infos) > MAX_SIZE:
                raise ValueError('Книга превышает лимит проверки 200 МБ / 10000 частей')
            if len({i.filename for i in infos}) != len(infos):
                raise ValueError('Повторяющиеся части ZIP')
            self.parts = {i.filename: archive.read(i) for i in infos if not i.is_dir()}
        self.strings = []
        if 'xl/sharedStrings.xml' in self.parts:
            self.strings = [''.join(n.itertext()) for n in xml(self.parts['xl/sharedStrings.xml'])]
        rels = {n.get('Id'): n for n in xml(self.parts['xl/_rels/workbook.xml.rels'])}
        self.sheets = {}
        root = xml(self.parts['xl/workbook.xml'])
        for sheet in root.find(Q('sheets')):
            relation = rels[sheet.get('{' + REL + '}id')]
            if relation.get('TargetMode') == 'External':
                raise ValueError('Внешний лист не поддерживается')
            target = relation.get('Target')
            part = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl', target))
            tree = xml(self.parts[part])
            if tree.tag != Q('worksheet'):
                raise ValueError('Неподдерживаемый тип листа: ' + sheet.get('name'))
            cells = {c.get('r'): c for c in tree.iter(Q('c'))}
            self.sheets[sheet.get('name')] = {'part': part, 'tree': tree, 'cells': cells}

    def cell(self, ref):
        if not isinstance(ref, str) or '!' not in ref:
            raise ValueError('Нужна ссылка Лист!A1: ' + str(ref))
        name, address = ref.rsplit('!', 1)
        if not re.fullmatch(r'[A-Z]{1,3}[1-9][0-9]*', address):
            raise ValueError('Некорректная ячейка: ' + ref)
        if name not in self.sheets:
            raise ValueError('Нет листа: ' + name)
        return self.sheets[name]['cells'].get(address)

    def value(self, ref, allow_formula=False):
        c = self.cell(ref)
        if c is None:
            return None
        if c.find(Q('f')) is not None and not allow_formula:
            raise ValueError('Формула требует пересчета в Excel, не читаем кеш как факт: ' + ref)
        kind = c.get('t')
        if kind == 'e':
            raise ValueError('Ошибка Excel: ' + ref)
        if kind == 'inlineStr':
            return ''.join(n.text or '' for n in c.iter(Q('t')))
        value = c.findtext(Q('v'))
        if kind == 's':
            return self.strings[int(value)]
        return value

    def signature(self, c):
        if c is None:
            return None
        clone = copy.deepcopy(c)
        # Shared string numbering may change; compare actual rich text instead.
        if clone.get('t') == 's':
            clone.set('t', 'inlineStr')
            v = clone.find(Q('v'))
            idx = int(v.text)
            clone.remove(v)
            node = copy.deepcopy(list(xml(self.parts['xl/sharedStrings.xml']))[idx])
            node.tag = Q('is')
            clone.append(node)
        return canonical(clone)


def blank(value):
    return value is None or str(value).strip().lower() in EMPTY


def number(book, ref):
    if isinstance(ref, dict):
        if set(ref) != {'cell', 'expected_text', 'number_index'}:
            raise ValueError('Селектор числа: нужны cell, expected_text, number_index')
        value = book.value(ref['cell'])
        if not isinstance(ref['expected_text'], str) or value != ref['expected_text']:
            raise ValueError('Текст ячейки отличается от привязанного к расчету: ' + str(ref['cell']))
        index = ref['number_index']
        if type(index) is not int or index < 0:
            raise ValueError('number_index должен быть целым числом от 0')
        # Range separators are not negative signs; actual negative values stay negative.
        normalized = re.sub(r'(?<=\d)\s*[-–—]\s*(?=\d)', ' ', value)
        tokens = re.findall(r'[-+]?\d+(?:[.,]\d+)?', normalized)
        if index >= len(tokens):
            raise ValueError('В ячейке нет числа с индексом ' + str(index))
        value = tokens[index]
    else:
        value = book.value(ref)
    if value is None:
        raise ValueError('Нет числа: ' + str(ref))
    try:
        n = Decimal(str(value).strip().replace(',', '.'))
    except InvalidOperation:
        raise ValueError('Ожидается отдельное число или явный селектор: ' + str(ref))
    if not n.is_finite() or n < 0:
        raise ValueError('Число должно быть конечным и неотрицательным: ' + str(ref))
    return n


def validate(original, filled, manifest=None):
    original, filled = Path(original), Path(filled)
    report = {'status': 'incomplete', 'errors': [], 'warnings': [], 'checks': [],
              'limits': ['Формулы не пересчитываются.', 'Истинность наблюдений и полнота карты проверяются человеком.',
                         'Успех означает только прохождение перечисленных механических проверок.']}
    def error(code, detail):
        report['errors'].append({'code': code, 'detail': str(detail)})
    def warn(code, detail):
        report['warnings'].append({'code': code, 'detail': str(detail)})
    try:
        if original.resolve() == filled.resolve() or original.samefile(filled):
            error('same_file', 'Оригинал и результат должны быть разными файлами')
        a, b = Workbook(original), Workbook(filled)
        report['original_sha256'] = digest(original)
        report['filled_sha256'] = digest(filled)
    except (OSError, ValueError, KeyError, ET.ParseError, zipfile.BadZipFile, TypeError, IndexError) as exc:
        error('workbook_read', exc)
        report['status'] = 'failed'
        return report
    m = manifest if isinstance(manifest, dict) else {}
    coverage = bool(m)
    keys = {'schema_version', 'original_sha256', 'allowed_cells', 'required_cells', 'arithmetic', 'transitions', 'observations', 'not_applicable', 'package_changes'}
    if m:
        if set(m) - keys or m.get('schema_version') != 1:
            error('manifest', 'Неизвестные поля или schema_version != 1')
        for k in ['allowed_cells', 'required_cells', 'arithmetic', 'transitions', 'observations']:
            if not isinstance(m.get(k), list):
                error('manifest', k + ' должен быть списком')
                coverage = False
                m = dict(m, **{k: []})
        if not re.fullmatch(r'[a-f0-9]{64}', str(m.get('original_sha256', ''))):
            error('manifest', 'Нет корректного SHA-256 исходника, снятого до редактирования')
        elif m['original_sha256'] != report['original_sha256']:
            error('original_hash', 'Исходник изменился после фиксации контрольной суммы')
    else:
        warn('coverage', 'Нет манифеста: проверка дозировок, заполнения и переходов не выполнена')
    allowed = set()
    for ref in m.get('allowed_cells', []):
        try:
            a.cell(ref)
            b.cell(ref)
            allowed.add(ref)
        except (ValueError, TypeError) as exc:
            error('manifest', exc)
    if list(a.sheets) != list(b.sheets):
        error('sheet_names', 'Изменены имена, порядок или состав листов; требуется отдельное согласованное расширение шаблона')
    sheet_parts = {s['part'] for s in a.sheets.values()} | {s['part'] for s in b.sheets.values()}
    changes = m.get('package_changes', {})
    if not isinstance(changes, dict):
        error('manifest', 'package_changes должен быть объектом часть: причина')
        changes = {}
    for part in sorted(set(a.parts) | set(b.parts)):
        if part in sheet_parts or part == 'xl/sharedStrings.xml':
            continue
        old, new = a.parts.get(part), b.parts.get(part)
        same = old == new
        if not same and old is not None and new is not None and part.endswith(('.xml', '.rels')):
            try:
                same = canonical(xml(old)) == canonical(xml(new))
            except (ET.ParseError, ValueError):
                same = False
        if not same:
            # Only document bookkeeping can be explicitly excepted. Never styles,
            # drawings, VBA, relationships, defined names, or calculation chains.
            if part in {'docProps/core.xml', 'docProps/app.xml'} and changes.get(part):
                warn('metadata_change', part + ': ' + str(changes[part]))
            else:
                error('package_part', 'Изменена/потеряна часть книги: ' + part)
    for name in a.sheets.keys() & b.sheets.keys():
        sa, sb = a.sheets[name], b.sheets[name]
        for address in sa['cells'].keys() | sb['cells'].keys():
            ca, cb = sa['cells'].get(address), sb['cells'].get(address)
            ref = name + '!' + address
            fa = ca.find(Q('f')) if ca is not None else None
            fb = cb.find(Q('f')) if cb is not None else None
            if fa is not None:
                if fb is None or canonical(fa) != canonical(fb):
                    error('formula_changed', ref)
            if a.signature(ca) == b.signature(cb):
                continue
            if ref not in allowed:
                # Formula cached values can legitimately differ after recalculation.
                xa, xb = copy.deepcopy(ca), copy.deepcopy(cb)
                if fa is not None and fb is not None:
                    for x in [xa, xb]:
                        for v in list(x):
                            if v.tag == Q('v'):
                                x.remove(v)
                if xa is None or xb is None or canonical(xa) != canonical(xb):
                    error('unexpected_cell_change', ref)
            else:
                # Allow value/type changes, retain style/protection and other metadata.
                def attrs(c):
                    return {} if c is None else {k: v for k, v in c.attrib.items() if k not in {'r', 't'}}
                if attrs(ca) != attrs(cb):
                    error('cell_style', ref)
        def structure(tree):
            clone = copy.deepcopy(tree)
            data = clone.find(Q('sheetData'))
            if data is not None:
                for row in list(data):
                    for c in list(row):
                        row.remove(c)
                    if set(row.attrib) <= {'r'}:
                        data.remove(row)
            return canonical(clone)
        if structure(sa['tree']) != structure(sb['tree']):
            error('sheet_structure', name + ': объединения, размеры, проверки данных или оформление листа изменены')
    for name, sheet in b.sheets.items():
        for address, c in sheet['cells'].items():
            if c.get('t') == 'e':
                error('excel_error', name + '!' + address + ': ' + str(c.findtext(Q('v'))))
    for ref in m.get('required_cells', []):
        try:
            c = b.cell(ref)
            if c is not None and c.find(Q('f')) is not None:
                warn('formula_manual', ref + ': заполнена формулой, требуется пересчет')
            elif blank(b.value(ref)):
                error('required_empty', ref)
        except (ValueError, TypeError, IndexError) as exc:
            error('check_input', exc)
    for check in m.get('arithmetic', []):
        try:
            op, inputs = check['operation'], check['inputs']
            if op in {'sum', 'product'}:
                if not isinstance(inputs, list) or not inputs:
                    raise ValueError('sum/product: нужен непустой список ячеек')
                nums = [number(b, r) for r in inputs]
                expected = sum(nums) if op == 'sum' else math.prod(nums)
            elif op == 'duration':
                fields = {'sets', 'sides', 'rest_seconds', 'rest_count', 'transition_seconds'}
                fields |= {'work_seconds'} if 'work_seconds' in inputs else {'reps', 'seconds_per_rep'}
                if set(inputs) != fields:
                    raise ValueError('duration: неполные или лишние параметры')
                n = {k: number(b, r) for k, r in inputs.items()}
                for k in {'sets', 'sides', 'reps', 'rest_count'} & n.keys():
                    if n[k] != int(n[k]) or (k != 'rest_count' and n[k] == 0):
                        raise ValueError(k + ': нужен положительный целый счетчик (rest_count может быть 0)')
                work = n['work_seconds'] if 'work_seconds' in n else n['reps'] * n['seconds_per_rep']
                expected = n['sets'] * n['sides'] * work + n['rest_seconds'] * n['rest_count'] + n['transition_seconds']
            else:
                raise ValueError('Неизвестная операция: ' + str(op))
            actual = number(b, check['result'])
            tolerance = Decimal(str(check.get('tolerance', '0.01')))
            if not tolerance.is_finite() or tolerance < 0 or tolerance > 1:
                raise ValueError('Допуск должен быть от 0 до 1 единицы результата')
            report['checks'].append({'name': check.get('name', op), 'expected': str(expected), 'actual': str(actual)})
            if abs(expected - actual) > tolerance:
                error('arithmetic', str(check.get('name', op)) + ': ожидается ' + str(expected) + ', указано ' + str(actual))
        except (ValueError, KeyError, TypeError, IndexError, InvalidOperation) as exc:
            error('check_input', exc)
    for chain in m.get('transitions', []):
        try:
            states = chain['states']
            if len(states) < 2 or states[0]['kind'] != 'base':
                raise ValueError('Лестница должна начинаться с base и содержать минимум два состояния')
            if chain.get('require_regression') and not any(s['kind'] == 'regression' for s in states[1:]):
                error('regression_missing', chain.get('name', 'Лестница'))
            ids = [s['id'] for s in states]
            if len(set(ids)) != len(ids):
                raise ValueError('id уровней должны быть уникальны')
            previous = None
            for state in states:
                if not isinstance(state['values'], dict):
                    raise ValueError('values состояния должен быть объектом параметр: ячейка')
                values = {k: b.value(ref) for k, ref in state['values'].items()}
                if not values or any(blank(v) for v in values.values()):
                    raise ValueError('Не все параметры состояния заполнены')
                if previous is not None:
                    if state['kind'] not in {'progression', 'regression', 'repeat'}:
                        raise ValueError('Неверный kind уровня')
                    if set(previous) != set(values):
                        error('parameter_set', state['id'])
                    else:
                        changed = [k for k in values if values[k] != previous[k]]
                        wanted = 0 if state['kind'] == 'repeat' else 1
                        if len(changed) != wanted:
                            error('multiple_changes', state['id'] + ': изменены ' + ', '.join(changed))
                        elif changed and changed[0] != state.get('changed_parameter'):
                            error('declared_parameter', state['id'])
                    if blank(state.get('reason')):
                        error('reason_missing', state['id'])
                previous = values
            report['checks'].append({'name': chain.get('name', 'Лестница'), 'states': len(states)})
        except (ValueError, KeyError, TypeError, IndexError) as exc:
            error('check_input', exc)
    for item in m.get('observations', []):
        try:
            b.cell(item['cell'])
            if item['status'] not in {'planned', 'scenario', 'observed', 'unknown'}:
                raise ValueError('Неверный статус наблюдения')
            if item['status'] == 'observed' and (blank(item.get('evidence')) or blank(item.get('date'))):
                error('evidence_missing', item['cell'] + ': нужны evidence и date')
        except (ValueError, KeyError, TypeError) as exc:
            error('check_input', exc)
    not_applicable = m.get('not_applicable', {})
    if not isinstance(not_applicable, dict):
        error('manifest', 'not_applicable должен быть объектом')
        not_applicable = {}
    for group in ['required_cells', 'arithmetic', 'transitions', 'observations']:
        if not m.get(group):
            if not_applicable.get(group) and isinstance(not_applicable[group], str):
                warn('not_applicable', group + ': ' + not_applicable[group])
            else:
                coverage = False
                warn('coverage', 'Не проверено: ' + group)
    report['status'] = 'failed' if report['errors'] else ('mechanical_pass' if coverage else 'incomplete')
    return report


def write_report(path, result, protected):
    if path is None:
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return
    path = Path(path)
    if path.resolve() in {Path(p).resolve() for p in protected}:
        raise ValueError('Отчет не может перезаписать входной файл')
    # Exclusive creation also protects existing reports and hardlinks.
    with path.open('x', encoding='utf-8') as stream:
        json.dump(result, stream, ensure_ascii=False, indent=2)
        stream.write('\n')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    snapshot = sub.add_parser('snapshot', help='Снять SHA-256 до редактирования')
    snapshot.add_argument('original', type=Path)
    snapshot.add_argument('--report', type=Path)
    check = sub.add_parser('check', help='Сравнить оригинал и заполненную копию')
    check.add_argument('original', type=Path)
    check.add_argument('filled', type=Path)
    check.add_argument('--manifest', type=Path)
    check.add_argument('--report', type=Path)
    args = parser.parse_args()
    try:
        if args.command == 'snapshot':
            book = Workbook(args.original)
            result = {'schema_version': 1, 'original_sha256': digest(args.original), 'sheets': list(book.sheets)}
            protected = [args.original]
            code = 0
        else:
            manifest = json.loads(args.manifest.read_text(encoding='utf-8')) if args.manifest else None
            if manifest is not None and not isinstance(manifest, dict):
                raise ValueError('Манифест должен быть JSON-объектом')
            result = validate(args.original, args.filled, manifest)
            protected = [args.original, args.filled] + ([args.manifest] if args.manifest else [])
            code = {'mechanical_pass': 0, 'failed': 1, 'incomplete': 2}[result['status']]
        write_report(args.report, result, protected)
        return code
    except (OSError, ValueError, KeyError, ET.ParseError, zipfile.BadZipFile) as exc:
        print('Ошибка: ' + str(exc), file=sys.stderr)
        return 2


if __name__ == '__main__':
    sys.exit(main())
