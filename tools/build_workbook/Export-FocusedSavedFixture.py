"""Читает контрольное сечение из XLSM без запуска старого VBA."""
import argparse
import hashlib
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def column(text):
    result=0
    for char in text:result=result*26+ord(char)-64
    return result

def export(source):
    names=('rngNDMSectionGeometry','rngNDMSectionContours','rngNDMSectionProperties','rngNDMSectionAnnotations','rngNDMMaterialDiagrams')
    with zipfile.ZipFile(source) as archive:
        strings=[]
        if 'xl/sharedStrings.xml' in archive.namelist():
            strings=[''.join(x.itertext()) for x in ET.fromstring(archive.read('xl/sharedStrings.xml')).findall('m:si',NS)]
        book=ET.fromstring(archive.read('xl/workbook.xml'))
        refs={x.attrib['name']:x.text for x in book.findall('m:definedNames/m:definedName',NS)}
        rels={x.attrib['Id']:x.attrib['Target'] for x in ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))}
        sheets={x.attrib['name']:rels[x.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']] for x in book.findall('m:sheets/m:sheet',NS)}
        output=[]
        for name in names:
            sheet,cell=refs[name].split('!')
            sheet=sheet.strip("'").replace("''", "'")
            col,row=re.match(r'\$?([A-Z]+)\$?(\d+)',cell).groups()
            row=int(row);col=column(col)
            target=sheets[sheet].lstrip('/')
            if not target.startswith('xl/'):target='xl/'+target
            values={}
            for entry in ET.fromstring(archive.read(target)).findall('.//m:sheetData/m:row/m:c',NS):
                letters,number=re.match(r'([A-Z]+)(\d+)',entry.attrib['r']).groups()
                kind=entry.attrib.get('t')
                value=entry.find('m:v',NS)
                if kind=='inlineStr':value=''.join(entry.find('m:is',NS).itertext())
                elif value is None:continue
                elif kind=='s':value=strings[int(value.text)]
                elif kind=='b':value=value.text=='1'
                elif kind in ('str','e'):value=value.text
                else:value=float(value.text)
                values[int(number),column(letters)]=value
            width=0
            while values.get((row,col+width)) not in (None,''):width+=1
            height=0
            while values.get((row+height,col)) not in (None,''):height+=1
            output.append({'Name':name,'Rows':height,'Columns':width,
                           'Values':[[values.get((row+r,col+c)) for c in range(width)] for r in range(height)]})
    return {'SourceSHA256':hashlib.sha256(source.read_bytes()).hexdigest().upper(),'Tables':output}

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source',type=Path,required=True);parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args();result=export(args.source)
    args.output.parent.mkdir(parents=True,exist_ok=True)
    args.output.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print([(x['Name'],x['Rows'],x['Columns']) for x in result['Tables']])
