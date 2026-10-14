"""Диагностирует частный пакет XLSM; основная книга не открывается и не изменяется."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
import copy

directory=Path(__file__).resolve().parent
base=directory/'WorkbookDirected/GeometrySolidProbe.xlsm'
with ZipFile(directory/'Candidate.xlsm') as patched, ZipFile(base) as original:
    changed={name:patched.read(name) for name in original.namelist() if patched.read(name)!=original.read(name)}
    original_parts={name:original.read(name) for name in original.namelist()}
    print('PATCHED_PARTS',list(changed))
    for label,parts in (
        ('NativeCodeOnly',['xl/vbaProject.bin']),
        ('Styles',['xl/vbaProject.bin','xl/styles.xml']),
        ('Workbook',['xl/vbaProject.bin','xl/workbook.xml']),
        ('Config',['xl/vbaProject.bin','xl/worksheets/sheet1.xml','xl/styles.xml']),
        ('Help',['xl/vbaProject.bin','xl/worksheets/sheet2.xml']),
        ('Chain',['xl/vbaProject.bin','xl/styles.xml','xl/workbook.xml','xl/worksheets/sheet1.xml','xl/worksheets/sheet2.xml','xl/sharedStrings.xml']),
    ):
        with ZipFile(directory/(label+'.xlsm'),'w',ZIP_DEFLATED) as out:
            for entry in original.infolist():
                out.writestr(copy.copy(entry),changed[entry.filename] if entry.filename in parts and entry.filename in changed else original_parts[entry.filename])
