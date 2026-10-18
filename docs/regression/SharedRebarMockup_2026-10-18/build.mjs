import fs from 'node:fs/promises';
import path from 'node:path';
import {Workbook, SpreadsheetFile} from '@oai/artifact-tool';

// Отдельный макет. Рабочая расчётная книга не изменяется.
const root=path.resolve(process.cwd(),'../../..');
const out=path.join(root,'outputs/01a12ece-8192-7141-89b3-f83f1176184c');
const filename='Config_LSection_SharedRebar_Proposal.xlsx';
const wb=Workbook.create();
const sh=wb.worksheets.add('Config');
const lists=wb.worksheets.add('_Lists');
const rg=address=>sh.getRange(address);
const ink='#000000', muted='#666666', section='#F2F2F2', header='#D9D9D9';
const capacity=8;
const data=Array.from({length:67},()=>Array(18).fill(null));
const write=(row,column,values)=>values.forEach((v,i)=>data[row-1][column-1+i]=v);
const mainRows=Array.from({length:capacity},(_,i)=>33+i), extraRows=Array.from({length:capacity},(_,i)=>46+i);
const dimensions=[
  {key:'B1',name:'Ширина R1',cell:'J11',value:250,note:'B1 — ширина верхнего прямоугольника R1.'},
  {key:'H1',name:'Высота R1',cell:'J12',value:550,note:'H1 — высота верхнего прямоугольника R1.'},
  {key:'B2',name:'Ширина R2',cell:'J13',value:600,note:'B2 — ширина нижнего прямоугольника R2.'},
  {key:'H2',name:'Высота R2',cell:'J14',value:250,note:'H2 — высота нижнего прямоугольника R2.'},
];
const sides=[
  ['R1.Top','R1: сверху','J11','Верхняя сторона прямоугольника R1; длина B1.'],
  ['R1.Bottom','R1: снизу','J11','Нижняя сторона прямоугольника R1; длина B1.'],
  ['R1.Left','R1: слева','J12','Левая сторона прямоугольника R1; длина H1.'],
  ['R1.Right','R1: справа','J12','Правая сторона прямоугольника R1; длина H1.'],
  ['R2.Top','R2: сверху','J13','Верхняя сторона прямоугольника R2; длина B2.'],
  ['R2.Bottom','R2: снизу','J13','Нижняя сторона прямоугольника R2; длина B2.'],
  ['R2.Left','R2: слева','J14','Левая сторона прямоугольника R2; длина H2.'],
  ['R2.Right','R2: справа','J14','Правая сторона прямоугольника R2; длина H2.'],
];
write(4,1,['Параметр','INPUT','INTERNAL','OUTPUT']);
write(5,1,['Length','mm','mm','m']);
write(7,1,['Тип сечения','LSection']);
write(5,8,['General','1. Общие настройки геометрии']);
write(6,8,['Ключ','Параметр','Значение','Ед.','Комментарий']);
write(7,8,['key','name','val','unit','note']);
write(8,8,['Type','Тип сечения','LSection','—','Постоянный код типа для ключей армирования.']);
write(9,8,['OffsetX','Смещение X',0,null,'Смещение всего сечения по X; плюс — вправо.']);
write(10,8,['OffsetY','Смещение Y',0,null,'Смещение всего сечения по Y; плюс — вверх.']);
dimensions.forEach((d,i)=>write(11+i,8,[d.key,d.name,d.value,null,d.note]));
write(16,8,['Edges','2. Участки контура']);
write(17,8,['Ключ стороны','Сторона','Длина','Ед.','Комментарий']);
write(18,8,['key','name','val','unit','note']);
sides.forEach((s,i)=>write(19+i,8,[s[0],s[1],null,null,s[3]]));
write(29,8,['Main','1. Основное армирование']);
write(30,8,['Ключ','Сторона','До оси','Диаметр','Количество','Отступ начала','Отступ конца','Комментарий']);
write(31,8,['key','side','as','d','n','s0','s1','note']);
write(32,8,['Units','Единицы',null,null,'шт',null,null,null]);
const count=[2,2,5,5,5,5,2,2];
const notes=['as — от края бетона до осей стержней.','d — диаметр основного ряда.','n — число стержней на выбранной стороне.','s0 и s1 — отступы вдоль линии размещения.','0 или пустой диаметр отключает ряд.','Ключ вычисляется по выбранной стороне геометрии.','Каждая сторона имеет свои параметры армирования.','Дополнительные ряды заданы в блоке Extra.'];
sides.forEach((s,i)=>write(mainRows[i],8,[null,null,40,20,count[i],40,40,notes[i]]));
write(42,8,['Extra','2. Дополнительное армирование']);
write(43,8,['Ключ','Сторона','Диаметр 2','Смещение 2','Привязка 2','Положение 2','Диаметр 3','Смещение 3','Привязка 3','Положение 3','Комментарий']);
write(44,8,['key','side','d2','a2','bind2','pos2','d3','a3','bind3','pos3','note']);
write(45,8,['Units','Единицы',null,null,'—','—',null,null,'—','—',null]);
const extra=[
  [16,0,'EachBar','Stacked',12,0,'EverySecondBar','SideBySide'],
  [16,60,'EachBar','Stacked',null,0,'EachBar','Stacked'],
  [12,50,'EverySecondBar','Stacked',null,0,'EachBar','Stacked'],
  [null,0,'EachBar','Stacked',null,0,'EachBar','Stacked'],
  [16,0,'EachBar','SideBySide',12,0,'EverySecondBar','Stacked'],
  [16,70,'EachBar','Stacked',12,70,'EverySecondBar','SideBySide'],
  [null,0,'EachBar','Stacked',null,0,'EachBar','Stacked'],
  [16,60,'EachBar','Stacked',null,0,'EachBar','Stacked'],
];
const extraNotes=['Смещения 2 и 3 отсчитываются от основного ряда.','0 или пустой диаметр отключает дополнительный ряд.','EverySecondBar — у первого, третьего и следующих нечётных стержней.','Ключ и сторона связаны с основным армированием.','SideBySide — рядом вдоль линии размещения стержней.','Stacked — в глубину бетона от основного ряда.','Все поля дополнительного армирования доступны для ввода.','Параметры противоположных сторон задаются независимо.'];
sides.forEach((s,i)=>write(extraRows[i],8,[null,null,...extra[i],extraNotes[i]]));
write(56,8,['General','1. Общие настройки геометрии']);
write(57,8,['Ключ','Параметр','Значение','Ед.','Комментарий']);
write(58,8,['key','name','val','unit','note']);
write(59,8,['Type','Тип сечения','Circle','—','Круглое сечение.']);
write(60,8,['OffsetX','Смещение X',0,null,'Смещение всего сечения по X; плюс — вправо.']);
write(61,8,['OffsetY','Смещение Y',0,null,'Смещение всего сечения по Y; плюс — вверх.']);
write(62,8,['D','Диаметр',300,null,'Диаметр бетонного круга.']);
write(64,8,['Edges','2. Участки контура']);
write(65,8,['Ключ участка','Участок','Длина','Ед.','Комментарий']);
write(66,8,['key','name','val','unit','note']);
write(67,8,['Circle','Окружность',null,null,'Полная длина бетонной окружности πD.']);
rg('A1:R67').values=data;
rg('A1:U69').format.font={name:'Arial',size:9,color:ink};
rg('A1:U69').format.rowHeight=15;
rg('A1:G69').format.columnWidth=17;
rg('H1:H69').format.columnWidth=32;
rg('I1:U69').format.columnWidth=17;
rg('G1:G69').format.columnWidth=4;
rg('A1:U69').format.wrapText=false;
rg('A1:R67').format.horizontalAlignment='center';
rg('A1:R67').format.verticalAlignment='center';
rg('A1:R67').format.fill='#FFFFFF';
sh.showGridLines=false;
lists.showGridLines=false;
lists.getRange('A1:AJ9').format.font={name:'Arial',size:9,color:ink};
function title(address,text){rg(address).merge();rg(address.split(':')[0]).values=[[text]];rg(address).format.font={name:'Arial',size:9,bold:true,color:ink};rg(address).format.horizontalAlignment='left';}
title('A3:D3','Единицы измерения');
title('H4:K4','Г-образное сечение');
title('H28:Q28','Общее армирование выбранного сечения');
title('H55:K55','Круглое сечение');
// Нативные ссылки добавляются после экспорта: движок расчёта пакета не исполняет HYPERLINK.
rg('L4').values=[['Подробнее']];
rg('R28').values=[['Подробнее']];
rg('L55').values=[['Подробнее']];
for(const a of ['L4','R28','L55'])rg(a).format.font={name:'Arial',size:9,color:'#0000FF',underline:'single'};
for(const [top,last,note] of [[5,14,'L'],[16,26,'L'],[29,40,'O'],[42,53,'R'],[56,62,'L'],[64,67,'L']]){
  const band=`H${top}:${note}${top}`;
  rg(band).format.fill=section;rg(band).format.font={name:'Arial',size:9,bold:true,color:ink};
  rg(`I${top}:${note}${top}`).merge();rg(`I${top}`).format.horizontalAlignment='left';
  rg(`H${top+1}:${note}${top+1}`).format.fill=header;
  rg(`H${top+1}:${note}${top+1}`).format.font={name:'Arial',size:9,bold:true,color:ink};
  rg(`H${top+2}:${note}${top+2}`).format.fill=section;
  rg(`H${top+2}:${note}${top+2}`).format.font={name:'Arial',size:9,italic:true,color:muted};
  rg(`${note}${top+1}:${note}${last}`).format.horizontalAlignment='right';
}
for(const r of [32,45])rg(`H${r}:R${r}`).format.font={name:'Arial',size:9,color:muted};
for(const address of ['A4:D5','A7:B7','H5:L26','H29:R53','H56:L67'])rg(address).format.borders={preset:'all',style:'dashed',color:ink};
rg('A4:D4').format.fill=header;rg('A4:D4').format.font={name:'Arial',size:9,bold:true,color:ink};
const help=[
  [9,'Параметры Г-сечения'],
  [10,'Размеры B1, H1, B2 и H2 задаются только в General.'],
  [11,'Длины восьми сторон R1/R2 вычисляются по этим размерам.'],
  [12,'Как задаётся армирование'],
  [13,'Выберите тип сечения в B7 и задайте параметры армирования.'],
  [14,'Ключ формируется из типа сечения и ключа стороны геометрии.'],
  [15,'Переименуйте ключ стороны — связанные ключи обновятся формулами.'],
  [16,'Дополнительные ряды относятся к той же стороне основного ряда.'],
  [17,'Для Г-сечения используются 8 строк, для круга — 1.'],
  [18,'Все ключи вычисляются формулами Excel. VBA в этой книге нет.'],
  [19,'Стороны подставляются из Edges выбранного типа сечения.'],
  [20,'При смене типа введённые параметры армирования сохраняются.'],
  [21,'В общей таблице 8 строк, без дополнительного запаса.'],
];
for(const [r,text] of help){rg(`A${r}:F${r}`).merge();rg(`A${r}`).values=[[text]];rg(`A${r}:F${r}`).format.horizontalAlignment='left';}
for(const r of [9,12])rg(`A${r}`).format.font={name:'Arial',size:9,bold:true,color:ink};
wb.names.add('rngUnitSettings',"='Config'!$A$4:$D$5");
wb.names.add('InputLengthUnit',"='Config'!$B$5");
wb.names.add('rngLSectionGeometry',"='Config'!$H$5:$L$26");
wb.names.add('rngRebarInput',"='Config'!$H$29:$R$53");
wb.names.add('rngCircleGeometry',"='Config'!$H$56:$L$67");
wb.names.add('selGeometryType',"='Config'!$B$7");
wb.names.add('rngLSectionEdgeKeys',"='Config'!$H$19:$H$26");
wb.names.add('rngLSectionEdgeNames',"='Config'!$I$19:$I$26");
wb.names.add('rngCircleEdgeKeys',"='Config'!$H$67");
wb.names.add('rngCircleEdgeNames',"='Config'!$I$67");
wb.names.add('rngCurrentSideKeys',"='_Lists'!$A$2:$A$9");
wb.names.add('rngCurrentSideNames',"='_Lists'!$B$2:$B$9");
lists.getRange('A1:B1').values=[['CurrentSideKey','CurrentSideName']];
for(let i=0;i<capacity;i++){
  const select=kind=>`=IF(selGeometryType="LSection",IFERROR(INDEX(rngLSectionEdge${kind},${i+1})&"",""),IF(selGeometryType="Circle",IFERROR(INDEX(rngCircleEdge${kind},${i+1})&"",""),""))`;
  lists.getRange('A'+(i+2)).formulas=[[select('Keys')]];
  lists.getRange('B'+(i+2)).formulas=[[select('Names')]];
}
const geometryRows=sides.map((s,i)=>19+i);
geometryRows.forEach((r,i)=>rg('J'+r).formulas=[['='+sides[i][2]]]);
rg('J67').formulas=[['=PI()*J62']];
const unitFormula='=InputLengthUnit';
const unitCells=['K9','K10',...dimensions.map((d,i)=>'K'+(11+i)),...geometryRows.map(r=>'K'+r),'J32','K32','M32','N32','J45','K45','N45','O45','K60','K61','K62','K67'];
unitCells.forEach(a=>rg(a).formulas=[[unitFormula]]);
const keyFormula=i=>`=IF(OR(selGeometryType="",INDEX(rngCurrentSideKeys,${i+1})=""),"",selGeometryType&"."&INDEX(rngCurrentSideKeys,${i+1}))`;
for(let i=0;i<capacity;i++){
  rg('H'+mainRows[i]).formulas=[[keyFormula(i)]];
  rg('I'+mainRows[i]).formulas=[[`=IF(H${mainRows[i]}="","",INDEX(rngCurrentSideNames,${i+1})&"")`]];
  rg('H'+extraRows[i]).formulas=[['=H'+mainRows[i]+'&""']];
  rg('I'+extraRows[i]).formulas=[['=I'+mainRows[i]+'&""']];
}
const validations=[];
let nextColumn=2;
const col=n=>{let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
function choice(address,name,options,sourceFormulas){
  const c=col(nextColumn++);
  lists.getRange(c+'1').values=[[name]];
  const dest=lists.getRange(`${c}2:${c}${options.length+1}`);
  if(sourceFormulas)dest.formulas=sourceFormulas.map(f=>[f]);else dest.values=options.map(v=>[v]);
  wb.names.add(name,`='_Lists'!$${c}$2:$${c}$${options.length+1}`);
  rg(address).dataValidation={rule:{type:'list',formula1:'='+name}};
  validations.push({address,name});
}
choice('B5','lstInputLength',['mm','cm','m']);
choice('B7','lstGeometryType',['LSection','Circle']);
const inputCells=['B7','J9','J10',...dimensions.map(d=>d.cell),'B5','J60','J61','J62'];
for(let i=0;i<capacity;i++){
  const r=mainRows[i],e=extraRows[i];
  inputCells.push(...['J','K','L','M','N'].map(c=>c+r));
  for(const [c,kind,values] of [['L','Bind2',['EachBar','EverySecondBar']],['M','Pos2',['Stacked','SideBySide']],['P','Bind3',['EachBar','EverySecondBar']],['Q','Pos3',['Stacked','SideBySide']]]){
    choice(c+e,`lst${kind}_${i+1}`,values);inputCells.push(c+e);
  }
  inputCells.push(...['J','K','N','O'].map(c=>c+e));
  for(const c of ['J','K','M','N'])rg(c+r).dataValidation={rule:{type:'decimal',operator:'greaterThanOrEqual',formula1:0}};
  rg('L'+r).dataValidation={rule:{type:'whole',operator:'greaterThanOrEqual',formula1:0}};
  for(const c of ['J','K','N','O'])rg(c+e).dataValidation={rule:{type:'decimal',operator:'greaterThanOrEqual',formula1:0}};
}
inputCells.forEach(a=>{rg(a).format.fill='#FFFFFF';rg(a).format.horizontalAlignment='center';});
for(const d of dimensions)rg(d.cell).dataValidation={rule:{type:'decimal',operator:'greaterThan',formula1:0}};
rg('J62').dataValidation={rule:{type:'decimal',operator:'greaterThan',formula1:0}};
for(const a of ['J9:J14','J19:J26','J33:N40','J46:K53','N46:O53','J60:J62','J67'])rg(a).setNumberFormat('General');
for(const r of mainRows)rg('L'+r).setNumberFormat('0');
wb.recalculate();
const expected=sides.map(s=>'LSection.'+s[0]);
const actual=rg('H33:H40').values.flat();
if(JSON.stringify(actual)!==JSON.stringify(expected))throw Error('Dynamic key formulas did not calculate: '+JSON.stringify(actual));
const lengths=geometryRows.map(r=>rg('J'+r).values[0][0]);
if(JSON.stringify(lengths)!==JSON.stringify([250,250,550,550,600,600,250,250]))throw Error('L-section dimension formulas failed.');
console.log((await wb.inspect({kind:'region',sheetId:'Config',range:'H33:K40',maxChars:1600,tableMaxRows:8,tableMaxCols:4,include:'values,formulas'})).ndjson);
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!|not implemented',options:{useRegex:true,maxResults:20},summary:'Formula error scan'});
console.log(errors.ndjson);
if(errors.ndjson.includes('"kind":"match"'))throw Error('Unexpected formula errors.');
await fs.mkdir(out,{recursive:true});
const file=await SpreadsheetFile.exportXlsx(wb);await file.save(path.join(out,filename));
const captureRanges=[{name:'Geometry',range:'H4:L26'},{name:'Rebar',range:'H28:R53'},{name:'CircleGeometry',range:'H55:L67'},{name:'CircleRebar',range:'H28:R53',type:'Circle'}];
for(const {name,range,type} of captureRanges){
  if(type){rg('B7').values=[[type]];wb.recalculate();}
  const image=await wb.render({sheetName:'Config',range,scale:1.5,format:'png'});
  await fs.writeFile(name+'_Artifact.png',new Uint8Array(await image.arrayBuffer()));
  if(type){rg('B7').values=[['LSection']];wb.recalculate();}
}
await fs.writeFile('Manifest.json',JSON.stringify({file:path.join(out,filename),capacity,geometry:'H5:L26',circleGeometry:'H56:L67',rebar:'H29:R53',geometryRows,mainRows,extraRows,dimensions,sides:sides.map((s,i)=>({key:s[0],name:s[1],lengthCell:'J'+geometryRows[i],sourceCell:s[2]})),captureRanges,commentRanges:['L8:L14','L19:L26','O33:O40','R46:R53','L59:L62','L67'],expectedKeys:expected,validations,inputCells,unitCells,mainWorkbookSHA256:'ABAA60FCB940F5C850571F515BF4C7C82DA474E7002B7DD122C54BB4E147C45C'},null,2));
console.log('SHARED_REBAR_MOCKUP: '+path.join(out,filename));
