"""ДЛЯ ТЕСТОВ: сравнивает контракт листов тяжёлого архива и текущей книги без Excel."""
from pathlib import Path
from xml.etree import ElementTree as E
from zipfile import ZipFile
import hashlib
import json
import re

e = Path(__file__).resolve().parent
root = e.parents[2]
inspection = json.loads((e / 'Ready30SourceInspection01.json').read_text('utf-8-sig'))
n = {'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def cell_position(address):
    letters, row = re.fullmatch(r'\$?([A-Z]+)\$?(\d+)',address).groups()
    column = 0
    for letter in letters:column = column * 26 + ord(letter) - 64
    return column,int(row)

def contains(sqref,address):
    column,row = cell_position(address)
    for area in sqref.split():
        ends=area.split(':');first=cell_position(ends[0]);last=cell_position(ends[-1])
        if first[0]<=column<=last[0] and first[1]<=row<=last[1]:return True
    return False

def inspect(path):
    with ZipFile(path) as z:
        strings = [''.join(x.itertext()) for x in E.fromstring(z.read('xl/sharedStrings.xml'))]
        book = E.fromstring(z.read('xl/workbook.xml'))
        names = {x.get('name'):x.text for x in book.findall('m:definedNames/m:definedName',n)}
        rels = {r.get('Id'):r.get('Target').lstrip('/') for r in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        grids = {}; codenames = {}; sheets = []; validations = {}
        for sheet in book.findall('m:sheets/m:sheet',n):
            name = sheet.get('name'); sheets.append(name)
            part = rels[sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')]
            xml = E.fromstring(z.read(part if part.startswith('xl/') else 'xl/'+part))
            prop = xml.find('m:sheetPr',n); codenames[name] = prop.get('codeName') if prop is not None else None
            validations[name] = [{'Range':v.get('sqref'),'Type':v.get('type'),'Formula1':v.findtext('m:formula1',default='',namespaces=n)} for v in xml.findall('m:dataValidations/m:dataValidation',n)]
            grid = {}
            for cell in xml.findall('m:sheetData/m:row/m:c',n):
                value = cell.find('m:v',n)
                if cell.get('t') == 'inlineStr': text = ''.join(cell.find('m:is',n).itertext())
                elif value is None: text = ''
                elif cell.get('t') == 's': text = strings[int(value.text)]
                else: text = value.text
                grid[cell.get('r')] = text
            grids[name] = grid
        formula = names['rngSystemSettings']; sheet,area = formula.split('!',1)
        col,row = re.match(r'\$?([A-Z]+)\$?(\d+)',area).groups()
        keys = sorted(v for address,v in grids[sheet.strip("'")].items() if re.fullmatch(col+r'\d+',address) and int(re.search(r'\d+',address)[0]) > int(row) and v and not v.startswith('['))
        key_column,_=cell_position(col+'1')
        value_column=key_column+1;value_letters=''
        while value_column:value_column,rem=divmod(value_column-1,26);value_letters=chr(65+rem)+value_letters
        setting_addresses={v:value_letters+re.search(r'\d+',address)[0] for address,v in grids[sheet.strip("'")].items() if re.fullmatch(col+r'\d+',address) and v in keys}
        lists={}
        for key in ('Geometry.Type','Mesh.CellGeometry','Mesh.StepMode','Results.SnapshotStorage','Output.CombinationMode'):
            address=setting_addresses[key]
            lists[key]=[{k:v for k,v in validation.items() if k!='Range'} for validation in validations[sheet.strip("'")] if contains(validation['Range'],address)]
        workbook_props=book.find('m:workbookPr',n)
        critical = ('rngNDMSectionGeometry','rngNDMSectionProperties','rngNDMMaterialDiagrams','rngNDMSectionContours','rngNDMElementResults','rngNDMElementGeometry','rngNDMSectionAnnotations','rngLoadCombinations','rngSystemSettings')
        return {'Sheets':sheets,'CodeNames':codenames,'WorkbookCodeName':workbook_props.get('codeName') if workbook_props is not None else None,'SettingKeys':keys,'FixedSettingValidations':lists,'Anchors':{key:names[key] for key in critical}}

source = Path(inspection['Source'])
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == inspection['SourceSHA256']
old = inspect(source); current = inspect(root / 'workbook/output/RC_Section_NDM.xlsm')
missing = sorted(set(current['SettingKeys']) - set(old['SettingKeys']))
extra = sorted(set(old['SettingKeys']) - set(current['SettingKeys']))
differences = {key:{'ReadySource':old['Anchors'][key],'Current':current['Anchors'][key]} for key in current['Anchors'] if old['Anchors'][key] != current['Anchors'][key]}
validation_differences={key:{'ReadySource':old['FixedSettingValidations'][key],'Current':current['FixedSettingValidations'][key]} for key in current['FixedSettingValidations'] if old['FixedSettingValidations'][key]!=current['FixedSettingValidations'][key]}
passed = not missing and not extra and not differences and not validation_differences and old['Sheets'] == current['Sheets'] and old['CodeNames'] == current['CodeNames'] and old['WorkbookCodeName']==current['WorkbookCodeName']
out = {'Status':'PASS_PASSIVE_STRUCTURE_CONTRACT' if passed else 'STRUCTURE_REVIEW_REQUIRED','SourceSHA256':inspection['SourceSHA256'],'CurrentSHA256':hashlib.sha256((root / 'workbook/output/RC_Section_NDM.xlsm').read_bytes()).hexdigest().upper(),'MissingSettings':missing,'ExtraSettings':extra,'DifferentAnchors':differences,'DifferentFixedValidations':validation_differences,'Source':old,'Current':current,'ExcelProcessesStarted':0,'BooksModified':False,'FullAcceptanceClaim':False}
(e / 'Ready30StructureComparison01.json').write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({k:v for k,v in out.items() if k not in ('Source','Current')},ensure_ascii=False,indent=2))
if not passed:raise SystemExit(1)
