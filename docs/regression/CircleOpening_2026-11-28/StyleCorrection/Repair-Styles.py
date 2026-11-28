from pathlib import Path
import importlib.util,json,copy,hashlib
from lxml import etree as E
r=Path.cwd();d=r/'outputs/CircleOpeningStyle_2026-11-28';spec=importlib.util.spec_from_file_location('presentation',r/'tools/build_workbook/CalculationPresentation.py');p=importlib.util.module_from_spec(spec);spec.loader.exec_module(p)
before=p.read_package(d/'Before.xlsm');parts=dict(before);locations=p.sheet_parts(parts);styles=p.xml(parts,'xl/styles.xml');mask=json.loads((d/'InputFillMask.json').read_text(encoding='utf-8-sig'))
allowed_text={'B8':'Смещение X','B9':'Смещение Y','E16':'-','F16':'-','G16':'-','H16':'-'}
for name,addresses in mask.items():
 part=locations[name][1];sheet=p.xml(parts,part);p.apply_user_input_fills(sheet,styles,set(addresses))
 if name=='Сечения':
  cells={c.get('r'):c for c in sheet.iter(p.M+'c')}
  for address,value in allowed_text.items():p.set_cell(cells[address],'literal',value)
 parts[part]=p.serialize(sheet)
parts['xl/styles.xml']=p.serialize(styles)
# Проверка, что меняются только заливки выбранных полей и шесть запрошенных подписей.
def canonical(node):return (node.tag,tuple(sorted((k,v) for k,v in node.attrib.items() if k!='s')),node.text,tuple(canonical(n) for n in node))
original_styles=p.xml(before,'xl/styles.xml');old_xfs=original_styles.find(p.M+'cellXfs');new_xfs=styles.find(p.M+'cellXfs');old_fills=original_styles.find(p.M+'fills');new_fills=styles.find(p.M+'fills');changes={};errors=[]
for name in mask:
 path=locations[name][1];a=p.xml(before,path);b=p.xml(parts,path);ca={n.get('r'):n for n in a.iter(p.M+'c')};cb={n.get('r'):n for n in b.iter(p.M+'c')}; changed=[]
 assert ca.keys()==cb.keys()
 for addr in ca:
  if canonical(ca[addr])!=canonical(cb[addr]) and not(name=='Сечения' and addr in allowed_text):errors.append([name,addr,'data changed'])
  old=copy.deepcopy(old_xfs[int(ca[addr].get('s','0'))]);new=copy.deepcopy(new_xfs[int(cb[addr].get('s','0'))]);of=old.get('fillId');nf=new.get('fillId')
  for xf in [old,new]:xf.attrib.pop('fillId',None);xf.attrib.pop('applyFill',None)
  if E.tostring(old)!=E.tostring(new):errors.append([name,addr,'non-fill style changed'])
  if of!=nf:changed.append(addr)
  if addr in mask[name]:
   fg=new_fills[int(nf)].find(p.M+'patternFill/'+p.M+'fgColor');assert fg is not None and fg.get('rgb')=='FFFFF2E6',(name,addr)
 changes[name]={'InputFields':len(mask[name]),'FillChanges':len(changed),'ChangedAddresses':changed}
changed_parts=[n for n in parts if parts[n]!=before[n]];assert set(changed_parts)=={'xl/styles.xml',locations['Config'][1],locations['Сечения'][1]};assert not errors, errors
candidate=r/'workbook/candidates/CircleOpening/RC_Section_NDM.xlsm';p.write_verified_package(candidate,parts)
# Реальный путь наложения шаблона: по новой политике белый стиль шаблона не сбрасывает заливку полей ввода.
p.apply(d/'NativeFields.xlsm',r/'tools/build_workbook/CalculationPresentation.zip',d/'FormattedRoundtrip.xlsm')
roundtrip=p.read_package(d/'FormattedRoundtrip.xlsm');roundstyles=p.xml(roundtrip,'xl/styles.xml');roundloc=p.sheet_parts(roundtrip)
for name,addresses in mask.items():assert set(addresses)<=p.user_input_fill_cells(p.xml(roundtrip,roundloc[name][1]),roundstyles)
proof={'Status':'PASS','SourceSHA256':hashlib.sha256((d/'Before.xlsm').read_bytes()).hexdigest().upper(),'CandidateSHA256':hashlib.sha256(candidate.read_bytes()).hexdigest().upper(),'InputFields':sum(len(a) for a in mask.values()),'Changes':changes,'TextChanges':allowed_text,'ChangedParts':changed_parts,'OtherPartsByteExact':True,'VbaProjectByteExact':parts['xl/vbaProject.bin']==before['xl/vbaProject.bin'],'AllFormulasAndOtherValuesExact':True,'NonFillStylesExact':True,'BuilderPresentationRoundtrip':'PASS','Errors':errors}
(d/'StyleAcceptance.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps({k:v for k,v in proof.items() if k!='Changes'},ensure_ascii=False,indent=2));print({k:{a:b for a,b in v.items() if a!='ChangedAddresses'} for k,v in changes.items()})
