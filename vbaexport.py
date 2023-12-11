"""Export VBA code from Excel files

Exports all of the VBA code from Excel files in the current directory into
folders named for the Excel workbook. Each code module is stored as a separate
text file that is named with a .bas extension.

Requires: [oletools](https://github.com/decalage2/oletools)

Install with: pip install oletools

Inspired by: https://www.xltrail.com/blog/auto-export-vba-commit-hook

Completely rewritten to take better advantage of Python3


Can be used as a Git pre-commit hook by creating
a `pre-commit` script in the `.git/hooks` directory

    #!/bin/sh

    python .git/hooks/vbaexport.py
    git add -- ./*.vba
"""
from oletools.olevba import VBA_Parser
from pathlib import Path

import shutil


EXCEL_FILE_EXTENSIONS = (
    '.xlsb',
    '.xls', '.xlsm',
    '.xla', '.xlam',
    '.xlt', '.xltm'
)


def export_vba(workbook_filename):
    vba_folder = workbook_filename.parent / f'{workbook_filename.name}.vba'
    vba_parser = VBA_Parser(workbook_filename)
    if not vba_parser.detect_vba_macros():
        # no macros, nothing to export
        return
    for x, y, filename, content in vba_parser.extract_macros():
        lines = content.replace('\r\n', '\n').split('\n')
        if not lines:
            continue
        code = []
        for line in lines:
            if line.startswith('Attribute'):
                # strip out the attribute statements
                continue
            code.append(line)
        if code and code[-1] == '':
            code.pop(len(code) - 1)
            non_empty_lines_of_code = len([c for c in code if c])
            if non_empty_lines_of_code == 0:
                # avoid writing empty files
                continue
            vba_folder.mkdir(parents=True, exist_ok=True)
            vba_file = vba_folder / f'{filename}.bas'
            vba_file.write_text('\n'.join(code), encoding='utf8')


if __name__ == '__main__':
    path = Path.cwd()
    for item in path.rglob('*'):
        if item.is_dir() and item.suffix == '.vba':
            # delete existing folders that end .vba so they can be recreated
            # this ensures any code that is removed, will be properly handled
            #   rather than lingering
            shutil.rmtree(item)
    for item in path.rglob('*'):
        if item.is_file() and item.suffix in EXCEL_FILE_EXTENSIONS:
            export_vba(item)
