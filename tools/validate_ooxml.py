"""Validate SpreadsheetML parts against the official ECMA-376 XSD.

Usage: python validate_ooxml.py workbook.xlsm ecma376-part4.zip
Requires lxml. Read-only: never rewrites the workbook.
Schema download: https://ecma-international.org/publications-and-standards/standards/ecma-376/
Use Part 4, 5th edition (2016), containing OfficeOpenXML-XMLSchema-Transitional.zip.
"""
import argparse
import hashlib
import io
import json
from pathlib import Path
import tempfile
import zipfile

from lxml import etree


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook', type=Path)
    parser.add_argument('ecma_part4', type=Path)
    args = parser.parse_args()
    passed, failures = [], {}
    with tempfile.TemporaryDirectory(prefix='employee-master-schema-') as directory:
        root = Path(directory)
        with zipfile.ZipFile(args.ecma_part4) as outer:
            data = outer.read('OfficeOpenXML-XMLSchema-Transitional.zip')
        with zipfile.ZipFile(io.BytesIO(data)) as inner:
            for name in inner.namelist():
                if Path(name).name != name or not name.endswith('.xsd'):
                    raise ValueError('Unexpected schema archive entry: ' + name)
                (root / name).write_bytes(inner.read(name))
        xml_parser = etree.XMLParser(resolve_entities=False, no_network=True)
        schema = etree.XMLSchema(etree.parse(str(root / 'sml.xsd'), parser=xml_parser))
        with zipfile.ZipFile(args.workbook) as book:
            if book.testzip() is not None:
                raise ValueError('ZIP CRC error')
            for name in book.namelist():
                if not name.endswith('.xml'):
                    continue
                doc = etree.fromstring(book.read(name), parser=xml_parser)
                if etree.QName(doc).namespace != 'http://schemas.openxmlformats.org/spreadsheetml/2006/main':
                    continue
                if schema.validate(doc):
                    passed.append(name)
                else:
                    failures[name] = [e.message for e in schema.error_log]
    print(json.dumps({'workbook': args.workbook.name,
                      'sha256': hashlib.sha256(args.workbook.read_bytes()).hexdigest(),
                      'passed_parts': len(passed), 'failed_parts': len(failures),
                      'errors': failures,
                      'scope': 'ECMA-376 SpreadsheetML XSD; not native Excel execution'},
                     ensure_ascii=False, indent=2))
    raise SystemExit(1 if failures else 0)


if __name__ == '__main__':
    main()
