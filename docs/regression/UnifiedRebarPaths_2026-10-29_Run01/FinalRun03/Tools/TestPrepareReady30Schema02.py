"""ДЛЯ ТЕСТОВ: обновляет VBA/Main/Справку только в private Ready30 после полного PASS.

Не запускает Excel, не пересчитывает нагрузки, не публикует результат.
Исходные Results, схема и пользовательские данные проверяются отдельно и сохраняются.
"""
from pathlib import Path
from zipfile import ZipFile
from copy import deepcopy
from lxml import etree as ET
import argparse,json,hashlib,posixpath,re
from TestPrepareReady30Current01 import contract
from TestReady30SchemaTools02 import TestMergeStyles,TestMapSheetStyles,TestInlineStrings,TestPatchMain,TestUpdateHelpLinks,TestCells,TestCanonical,TestPayload

ROOT=Path(r'C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM')
PRIVATE=Path(__file__).resolve().parent
READY_SHA='6A21BE6E842D88317683D0BEB6228587C387014D24A8D42CA721785F664009C0'
OTHER_SHA='31F188692A01CC421A37CF962598778A5AD1FE6D9AB5D96DD41C29389C0919F7'
NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
REL='http://schemas.openxmlformats.org/package/2006/relationships'
OFFICE='http://schemas.openxmlformats.org/officeDocument/2006/relationships'

def TestSHA(data):
    """Возвращает SHA256 без нормализации содержимого."""
    return hashlib.sha256(data).hexdigest().upper()

def TestJSON(path):
    """Читает доказательство, включая допустимый UTF-8 BOM."""
    return json.loads(path.read_text('utf-8-sig'))

def TestSerialize(xml):
    """Сохраняет исходные namespace-префиксы lxml, включая mc:Ignorable."""
    return ET.tostring(xml,encoding='UTF-8',xml_declaration=True,standalone=True)

def TestResolvePart(part,target):
    """Разрешает внутреннюю связь пакета без обращения к сети или файловой системе."""
    result=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(part),target))
    assert result.startswith('xl/') and '..' not in result.split('/'),result
    return result

def TestSheets(package):
    """Возвращает настоящие имена частей пяти листов по workbook relationships."""
    book=ET.fromstring(package.read('xl/workbook.xml'))
    links={x.get('Id'):x.get('Target') for x in ET.fromstring(package.read('xl/_rels/workbook.xml.rels'))}
    sheets={x.get('name'):TestResolvePart('xl/workbook.xml',links[x.get('{'+OFFICE+'}id')]) for x in book.find('s:sheets',NS)}
    assert list(sheets)==['Config','Сечения','Справка','Расчет','Results']
    return book,sheets

def TestGeometryBounds(book,definitions):
    """ДЛЯ ТЕСТОВ: ограничивает поиск Main его настоящей именованной таблицей, независимо от совпадающих ключей соседних типов."""
    names={}
    for node in book.findall('s:definedNames/s:definedName',NS):
        name=node.get('name')
        if name not in {definition['RangeName'] for definition in definitions}:continue
        assert name not in names,'Duplicate geometry name: '+name
        match=re.fullmatch(r"(?:'Сечения'|Сечения)!\$?A\$?(\d+):\$?K\$?(\d+)",node.text or '')
        assert match,('Expected exact geometry table on Сечения A:K',name,node.text)
        names[name]=[int(match[1]),int(match[2])]
        assert names[name][0]<names[name][1]
    assert set(names)=={definition['RangeName'] for definition in definitions}
    return names

def TestHelpGraph(old,source,help_part,replacements,additions):
    """Переносит только принадлежащие справке рисунки; новые media-имена исключают замену пользовательских изображений."""
    rel_part=posixpath.join(posixpath.dirname(help_part),'_rels',posixpath.basename(help_part)+'.rels')
    relationships=ET.fromstring(source.read(rel_part))
    assert len(relationships)==1 and relationships[0].get('Type').endswith('/drawing')
    drawing=TestResolvePart(help_part,relationships[0].get('Target'))
    original=ET.fromstring(old.read(rel_part));assert len(original)==1
    assert TestResolvePart(help_part,original[0].get('Target'))==drawing
    # Другая часть книги не должна владеть заменяемым рисунком справки.
    for name in old.namelist():
        if not name.endswith('.rels') or name==rel_part or '/_rels/' not in name:continue
        owner=name.replace('/_rels/','/');owner=owner[:-5]
        for link in ET.fromstring(old.read(name)):
            if link.get('TargetMode')=='External':continue
            target=link.get('Target','')
            if target and TestResolvePart(owner,target)==drawing:raise AssertionError('Help drawing is shared with another owner: '+owner)
    drawing_rels=posixpath.join(posixpath.dirname(drawing),'_rels',posixpath.basename(drawing)+'.rels')
    children=ET.fromstring(source.read(drawing_rels));defaults={e.get('Extension'):e.get('ContentType') for e in ET.fromstring(old.read('[Content_Types].xml')) if e.tag.endswith('Default')}
    images=[]
    for link in children:
        assert link.get('Type').endswith('/image') and link.get('TargetMode')!='External'
        target=TestResolvePart(drawing,link.get('Target'));suffix=posixpath.basename(target).split('.')[-1]
        assert target.startswith('xl/media/') and suffix in defaults
        new='xl/media/ndm_help_current_'+posixpath.basename(target)
        assert new not in old.namelist() and new not in additions
        data=source.read(target);additions[new]=data
        link.set('Target',posixpath.relpath(new,posixpath.dirname(drawing)))
        images.append({'SourcePart':target,'NewPart':new,'Bytes':len(data),'SHA256':TestSHA(data)})
    replacements[rel_part]=source.read(rel_part)
    replacements[drawing]=source.read(drawing)
    replacements[drawing_rels]=TestSerialize(children)
    return images

def TestPrepare(launch_directory):
    """Все реальные действия требуют точного 182/122 PASS принятой main SHA и all-own closures."""
    launch=(ROOT/launch_directory).resolve();assert launch.is_relative_to(ROOT/'docs/regression')
    request=TestJSON(launch/'LaunchRequest.json');accept=TestJSON(ROOT/request['EvidenceDirectory']/'CompleteCurrentAcceptance.json');own=TestJSON(launch/'AllOwnClosuresReviewed.json')
    assert accept['Status']=='PASS' and accept['SourceSHA256']==request['SourceSHA256']
    assert accept['CoreSuites']==182 and accept['FocusedSingleCases']==122
    assert own['AllOwnIdentitiesTerminal'] and own['PrimaryLogsReviewed'] and own['SourceSHA256']==request['SourceSHA256']
    assert own['CoreMacrosReviewed']==182 and own['FocusedSingleCasesReviewed']==122 and own['FocusedRecordCount']==26
    main=ROOT/'workbook/output/RC_Section_NDM.xlsm';ready=ROOT/'workbook/output/NDM_Test_1232_30/RC_Section_NDM.xlsm';other=ready.with_name('RC_Section_NDM_ok.xlsm')
    assert TestSHA(main.read_bytes())==request['SourceSHA256'],'Full accepted source differs from current main'
    assert TestSHA(ready.read_bytes())==READY_SHA,'Ready30 changed; preserve and review new user edits'
    assert TestSHA(other.read_bytes())==OTHER_SHA,'The unrelated small example preservation fence differs'
    assert contract(main)==contract(ready),'Code names, settings, anchors or fixed Config lists need explicit compatibility review'
    definitions=TestJSON(ROOT/'tools/build_workbook/GeometryTables.json')
    replacements={};additions={};preserved=[]
    with ZipFile(ready) as old,ZipFile(main) as source:
        assert old.testzip() is None and source.testzip() is None
        assert not any('signature' in n.lower() for n in old.namelist())
        assert old.read('xl/theme/theme1.xml')==source.read('xl/theme/theme1.xml'),'Different themes need explicit color resolution; no output created'
        old_book,old_sheets=TestSheets(old);new_book,new_sheets=TestSheets(source)
        assert old_sheets==new_sheets
        section_bounds=TestGeometryBounds(new_book,definitions)
        assert TestGeometryBounds(old_book,definitions)==section_bounds,'Geometry table bounds changed'
        old_strings=ET.fromstring(old.read('xl/sharedStrings.xml'));new_strings=ET.fromstring(source.read('xl/sharedStrings.xml'))
        styles=ET.fromstring(old.read('xl/styles.xml'));new_styles=ET.fromstring(source.read('xl/styles.xml'))
        style_map,style_proof=TestMergeStyles(styles,new_styles);replacements['xl/styles.xml']=TestSerialize(styles)
        geometry_part=old_sheets['Сечения'];geometry=ET.fromstring(old.read(geometry_part));new_geometry=ET.fromstring(source.read(geometry_part))
        main_proof=TestPatchMain(geometry,new_geometry,old_strings,new_strings,style_map,definitions,section_bounds)
        geometry_links=TestUpdateHelpLinks(geometry,new_geometry);replacements[geometry_part]=TestSerialize(geometry)
        config_part=old_sheets['Config'];config=ET.fromstring(old.read(config_part));new_config=ET.fromstring(source.read(config_part));before_config={a:TestPayload(c) for a,c in TestCells(config).items()}
        config_links=TestUpdateHelpLinks(config,new_config)
        assert before_config=={a:TestPayload(c) for a,c in TestCells(config).items()}
        if config_links:replacements[config_part]=TestSerialize(config)
        help_part=old_sheets['Справка'];help_sheet=ET.fromstring(source.read(help_part));TestInlineStrings(help_sheet,new_strings);TestMapSheetStyles(help_sheet,style_map);replacements[help_part]=TestSerialize(help_sheet)
        images=TestHelpGraph(old,source,help_part,replacements,additions)
        # Только справочная печатная область может измениться; все остальные имена и свойства workbook сохранены.
        changed_names=[];new_names={(n.get('name'),n.get('localSheetId')):n for n in new_book.findall('s:definedNames/s:definedName',NS)}
        for name in old_book.findall('s:definedNames/s:definedName',NS):
            if name.get('name')=='_xlnm.Print_Area' and name.get('localSheetId')=='2':
                current=new_names[(name.get('name'),name.get('localSheetId'))]
                if name.text!=current.text:changed_names.append({'Name':name.get('name'),'Before':name.text,'After':current.text});name.text=current.text
        if changed_names:replacements['xl/workbook.xml']=TestSerialize(old_book)
        replacements['xl/vbaProject.bin']=source.read('xl/vbaProject.bin')
        assert old_sheets['Results'] not in replacements and old_sheets['Расчет'] not in replacements
        for name in replacements:assert name in old.namelist(),name
        original_parts={n:old.read(n) for n in old.namelist()}
    directory=PRIVATE/'Ready30Schema02';directory.mkdir(exist_ok=False)
    backup=directory/'Before_Ready30.xlsm';backup.write_bytes(ready.read_bytes());output=directory/'RC_Section_NDM.xlsm'
    with ZipFile(ready) as old,ZipFile(output,'w') as result:
        for item in old.infolist():result.writestr(item,replacements.get(item.filename,old.read(item.filename)))
        for name,data in additions.items():result.writestr(name,data)
    with ZipFile(output) as result:
        assert result.testzip() is None
        assert len(result.namelist())==len(set(result.namelist()))==len(original_parts)+len(additions)
        for name,data in original_parts.items():
            if name not in replacements:
                assert result.read(name)==data,name
                preserved.append({'Part':name,'Bytes':len(data),'SHA256':TestSHA(data)})
        for name,data in replacements.items():assert result.read(name)==data
        assert result.read(old_sheets['Results'])==original_parts[old_sheets['Results']]
        assert result.read(old_sheets['Расчет'])==original_parts[old_sheets['Расчет']]
    changed_parts=sorted(name for name,data in replacements.items() if data!=original_parts[name])
    proof={'Status':'PREPARED_PRIVATE_NATIVE_REVIEW_REQUIRED','LaunchDirectory':launch_directory,'AcceptedMainSHA256':request['SourceSHA256'],'AcceptedCommit':request['SourceCommit'],'Source':str(ready),'SourceSHA256':READY_SHA,'SourceBackup':str(backup),'Ready':str(output),'ReadySHA256':TestSHA(output.read_bytes()),'Published':False,'FullSourcePassVerified':True,'CurrentFullOwnClosed':True,'DataAndLoadsRecalculated':False,'ResultsPartByteExact':True,'CalculationSheetPartByteExact':True,'OriginalSharedStringsByteExact':True,'UnchangedParts':preserved,'ChangedOriginalParts':changed_parts,'RewrittenOriginalParts':sorted(replacements),'AddedHelpImages':images,'Styles':style_proof,'MainSchema':main_proof,'UpdatedGeometryHelpLinks':geometry_links,'UpdatedConfigHelpLinks':config_links,'ChangedPrintNames':changed_names,'OtherCopyUnchanged':True,'OnlyVBAChanged':False}
    (PRIVATE/'Ready30SchemaPreparation02.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({k:proof[k] for k in ['Status','AcceptedMainSHA256','SourceSHA256','ReadySHA256','Published','ResultsPartByteExact','CalculationSheetPartByteExact','ChangedOriginalParts']},ensure_ascii=False,indent=2))

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--launch-directory',required=True);args=parser.parse_args();TestPrepare(args.launch_directory)
