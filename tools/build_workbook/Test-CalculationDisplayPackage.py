"""ДЛЯ ТЕСТОВ: сохранность пользовательского ввода, серые границы и точные PNG в оформленной книге."""
import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
from zipfile import ZipFile
from openpyxl import load_workbook
from openpyxl.styles.colors import COLOR_INDEX
from lxml import etree as E

ROOT=Path(__file__).resolve().parents[2]
S='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
spec=importlib.util.spec_from_file_location('presentation',Path(__file__).with_name('CalculationPresentation.py'))
presentation=importlib.util.module_from_spec(spec);spec.loader.exec_module(presentation)

# -------------------- ДЛЯ ТЕСТОВ --------------------
def test_rgb(color,theme):
    """Разрешает явный или тематический цвет; для серого tint белой/чёрной темы совпадает с RGB Excel."""
    if color is None:return None
    if color.get('rgb'):return color.get('rgb')[-6:].upper()
    if color.get('indexed'):return COLOR_INDEX[int(color.get('indexed'))][-6:].upper()
    if color.get('theme') is None:return None
    order=('lt1','dk1','lt2','dk2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink')
    entry=theme.find(A+'themeElements/'+A+'clrScheme/'+A+order[int(color.get('theme'))])[0]
    value=entry.get('lastClr',entry.get('val'));tint=float(color.get('tint','0'))
    channels=[int(value[i:i+2],16) for i in (0,2,4)]
    return ''.join(f'{round(c*(1+tint) if tint<0 else c+(255-c)*tint):02X}' for c in channels)

# -------------------- ДЛЯ ТЕСТОВ --------------------
def test_package(path):
    """Проверяет все существующие печатные границы и каждую иллюстрацию по её имени объекта."""
    parts=presentation.read_package(path);locations=presentation.sheet_parts(parts)
    styles=E.fromstring(parts['xl/styles.xml']);theme=E.fromstring(parts['xl/theme/theme1.xml'])
    formats=styles.find(S+'cellXfs');borders=styles.find(S+'borders');sheet=E.fromstring(parts[locations['Расчет'][1]])
    checked=0;failures=[]
    for cell in sheet.iter(S+'c'):
        col,row=presentation.CELL_ADDRESS.fullmatch(cell.get('r')).groups()
        if presentation.column_number(col)>56 or int(row)>59:continue
        border=borders[int(formats[int(cell.get('s','0'))].get('borderId','0'))]
        for edge in border:
            if edge.get('style'):
                checked+=1
                if test_rgb(edge.find(S+'color'),theme)!='808080':failures.append((cell.get('r'),E.QName(edge).localname))
    assert not failures,failures[:12]
    catalog=json.loads((ROOT/'tools/build_workbook/SectionIllustrations/Illustrations.json').read_text('utf-8-sig'))
    assets={name:name+'.png' for name in catalog['Names']};assets.update({x['Name']:x['File'] for x in catalog['Additional']})
    found={}
    for part,raw in parts.items():
        if not part.startswith('xl/drawings/drawing') or not part.endswith('.xml'):continue
        drawing=E.fromstring(raw);relationship=presentation.relationship_path(part)
        if relationship not in parts:continue
        links={r.get('Id'):r.get('Target') for r in E.fromstring(parts[relationship])}
        for pic in drawing.iter():
            if E.QName(pic).localname!='pic':continue
            props=next(n for n in pic.iter() if E.QName(n).localname=='cNvPr')
            name=props.get('name')
            if name not in assets:continue
            blip=next(n for n in pic.iter() if E.QName(n).localname=='blip')
            media=presentation.target_part(part,links[blip.get(presentation.R+'embed')])
            expected=(ROOT/'tools/build_workbook/SectionIllustrations'/assets[name]).read_bytes()
            assert parts[media]==expected,'Image differs from asset: '+name
            found[name]=hashlib.sha256(expected).hexdigest().upper()
    assert found.keys()==assets.keys(),sorted(assets.keys()-found.keys())
    return {'BorderEdgesChecked':checked,'NonGrayBorders':failures,'Images':found,'SHA256':hashlib.sha256(path.read_bytes()).hexdigest().upper()}

# -------------------- ДЛЯ ТЕСТОВ --------------------
def test_inputs(source,candidate):
    """Читает XLSM без записи: сравнивает формулы/значения всех входов, а не кеш пересчёта Excel."""
    before=load_workbook(source,read_only=True,data_only=False);after=load_workbook(candidate,read_only=True,data_only=False)
    count=0;differences=[]
    try:
        for name in ('Config','Сечения','Расчет'):
            old,new=before[name],after[name]
            current={cell.coordinate:cell.value for row in new.iter_rows() for cell in row if cell.value is not None}
            for row in old.iter_rows():
                for cell in row:
                    if cell.value is None:continue
                    if name=='Расчет' and cell.data_type=='f':continue
                    count+=1
                    actual=current.get(cell.coordinate)
                    if cell.value!=actual:differences.append((name,cell.coordinate,str(cell.value),str(actual)))
    finally:before.close();after.close()
    assert not differences,differences[:12]
    return {'ComparedInputCells':count,'Differences':differences}

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--source',type=Path,required=True);parser.add_argument('--candidate',type=Path,required=True);parser.add_argument('--clean',type=Path);parser.add_argument('--report',type=Path,required=True)
    args=parser.parse_args();result={'Status':'PASS','UserInputs':test_inputs(args.source,args.candidate),'Candidate':test_package(args.candidate)}
    if args.clean:result['Clean']=test_package(args.clean)
    args.report.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8');print(json.dumps(result,ensure_ascii=False))
