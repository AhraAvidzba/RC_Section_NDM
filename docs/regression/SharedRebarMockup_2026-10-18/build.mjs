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
const data=Array.from({length:43},()=>Array(18).fill(null));
const write=(row,column,values)=>values.forEach((v,i)=>data[row-1][column-1+i]=v);
const mainRows=[27,28,29,30,31,32], extraRows=[38,39,40,41,42,43];
const sides=[
  ['Top','Верх',250,'Ширина верха B1.'],
  ['InnerRight','Внутр. вертикаль',550,'Высота стенки над полкой H1.'],
  ['Shoulder','Верх полки',null,'Разность ширины основания и верха: B2 − B1.'],
  ['Right','Справа',250,'Толщина нижней полки H2.'],
  ['Bottom','Низ',600,'Ширина основания B2.'],
  ['Left','Слева',null,'Полная высота: H1 + H2.'],
];
write(4,1,['Параметр','INPUT','INTERNAL','OUTPUT']);
write(5,1,['Length','mm','mm','m']);
write(5,8,['General','1. Общие настройки геометрии']);
write(6,8,['Ключ','Параметр','Значение','Ед.','Комментарий']);
write(7,8,['key','name','val','unit','note']);
write(8,8,['Type','Тип сечения','LSection','—','Постоянный код типа для ключей армирования.']);
write(9,8,['OffsetX','Смещение X',0,null,'Смещение всего сечения по X; плюс — вправо.']);
write(10,8,['OffsetY','Смещение Y',0,null,'Смещение всего сечения по Y; плюс — вверх.']);
write(12,8,['Geometry','2. Основная геометрия сечения']);
write(13,8,['Ключ стороны','Сторона','Длина','Ед.','Комментарий']);
write(14,8,['key','name','val','unit','note']);
sides.forEach((s,i)=>write(15+i,8,[s[0],s[1],s[2],null,s[3]]));
write(23,8,['Main','1. Основное армирование']);
write(24,8,['Ключ','Сторона','До оси','Диаметр','Количество','Отступ начала','Отступ конца','Комментарий']);
write(25,8,['key','side','as','d','n','s0','s1','note']);
write(26,8,['Units','Единицы',null,null,'шт',null,null,null]);
const count=[3,5,4,2,6,8];
const notes=['as — от края бетона до осей стержней.','d — диаметр основного ряда.','n — число стержней на выбранной стороне.','s0 и s1 — отступы вдоль линии размещения.','0 или пустой диаметр отключает ряд.','Ключ вычисляется по выбранной стороне геометрии.'];
sides.forEach((s,i)=>write(mainRows[i],8,[null,s[1],40,20,count[i],40,40,notes[i]]));
write(34,8,['Extra','2. Дополнительное армирование']);
write(35,8,['Ключ','Сторона','Диаметр 2','Смещение 2','Привязка 2','Положение 2','Диаметр 3','Смещение 3','Привязка 3','Положение 3','Комментарий']);
write(36,8,['key','side','d2','a2','bind2','pos2','d3','a3','bind3','pos3','note']);
write(37,8,['Units','Единицы',null,null,'—','—',null,null,'—','—',null]);
const extra=[
  [16,0,'EachBar','Stacked',12,0,'EverySecondBar','SideBySide'],
  [16,60,'EachBar','Stacked',null,0,'EachBar','Stacked'],
  [12,50,'EverySecondBar','Stacked',null,0,'EachBar','Stacked'],
  [null,0,'EachBar','Stacked',null,0,'EachBar','Stacked'],
  [16,0,'EachBar','SideBySide',12,0,'EverySecondBar','Stacked'],
  [16,70,'EachBar','Stacked',12,70,'EverySecondBar','SideBySide'],
];
const extraNotes=['Смещения 2 и 3 отсчитываются от основного ряда.','0 или пустой диаметр отключает дополнительный ряд.','EverySecondBar — у первого, третьего и следующих нечётных стержней.','Ключ и сторона связаны с основным армированием.','SideBySide — рядом вдоль линии размещения стержней.','Stacked — в глубину бетона от основного ряда.'];
sides.forEach((s,i)=>write(extraRows[i],8,[null,null,...extra[i],extraNotes[i]]));
rg('A1:R43').values=data;
rg('A1:U45').format.font={name:'Arial',size:9,color:ink};
rg('A1:U45').format.rowHeight=15;
rg('A1:G45').format.columnWidth=17;
rg('H1:H45').format.columnWidth=32;
rg('I1:U45').format.columnWidth=17;
rg('G1:G45').format.columnWidth=4;
rg('A1:U45').format.wrapText=false;
rg('A1:R43').format.horizontalAlignment='center';
rg('A1:R43').format.verticalAlignment='center';
rg('A1:R43').format.fill='#FFFFFF';
sh.showGridLines=false;
lists.showGridLines=false;
lists.getRange('A1:AE7').format.font={name:'Arial',size:9,color:ink};
function title(address,text){rg(address).merge();rg(address.split(':')[0]).values=[[text]];rg(address).format.font={name:'Arial',size:9,bold:true,color:ink};rg(address).format.horizontalAlignment='left';}
title('A3:D3','Единицы измерения');
title('H4:K4','Г-образное сечение');
title('H22:Q22','Общее армирование выбранного сечения');
// Нативные ссылки добавляются после экспорта: движок расчёта пакета не исполняет HYPERLINK.
rg('L4').values=[['Подробнее']];
rg('R22').values=[['Подробнее']];
for(const a of ['L4','R22'])rg(a).format.font={name:'Arial',size:9,color:'#0000FF',underline:'single'};
for(const [top,last,note] of [[5,10,'L'],[12,20,'L'],[23,32,'O'],[34,43,'R']]){
  const band=`H${top}:${note}${top}`;
  rg(band).format.fill=section;rg(band).format.font={name:'Arial',size:9,bold:true,color:ink};
  rg(`I${top}:${note}${top}`).merge();rg(`I${top}`).format.horizontalAlignment='left';
  rg(`H${top+1}:${note}${top+1}`).format.fill=header;
  rg(`H${top+1}:${note}${top+1}`).format.font={name:'Arial',size:9,bold:true,color:ink};
  rg(`H${top+2}:${note}${top+2}`).format.fill=section;
  rg(`H${top+2}:${note}${top+2}`).format.font={name:'Arial',size:9,italic:true,color:muted};
  rg(`${note}${top+1}:${note}${last}`).format.horizontalAlignment='right';
}
for(const r of [26,37])rg(`H${r}:R${r}`).format.font={name:'Arial',size:9,color:muted};
for(const address of ['A4:D5','H5:L20','H23:R43'])rg(address).format.borders={preset:'all',style:'dashed',color:ink};
rg('A4:D4').format.fill=header;rg('A4:D4').format.font={name:'Arial',size:9,bold:true,color:ink};
const help=[
  [9,'Параметры Г-сечения'],
  [10,'Четыре размера задаются вручную. Две длины вычисляются формулами.'],
  [12,'Как задаётся армирование'],
  [13,'Выберите сторону из списка и задайте её армирование.'],
  [14,'Ключ формируется из типа сечения и ключа стороны геометрии.'],
  [15,'Переименуйте ключ стороны — связанные ключи обновятся формулами.'],
  [16,'Дополнительные ряды относятся к той же стороне основного ряда.'],
  [17,'Пустая сторона оставляет ключ пустым.'],
  [18,'Все ключи вычисляются формулами Excel. VBA в этой книге нет.'],
];
for(const [r,text] of help){rg(`A${r}:F${r}`).merge();rg(`A${r}`).values=[[text]];rg(`A${r}:F${r}`).format.horizontalAlignment='left';}
for(const r of [9,12])rg(`A${r}`).format.font={name:'Arial',size:9,bold:true,color:ink};
wb.names.add('rngUnitSettings',"='Config'!$A$4:$D$5");
wb.names.add('InputLengthUnit',"='Config'!$B$5");
wb.names.add('rngLSectionGeometry',"='Config'!$H$5:$L$20");
wb.names.add('rngRebarInput',"='Config'!$H$23:$R$43");
wb.names.add('selGeometryType',"='Config'!$J$8");
wb.names.add('rngCurrentSideKeys',"='Config'!$H$15:$H$20");
wb.names.add('rngCurrentSideNames',"='Config'!$I$15:$I$20");
rg('J17').formulas=[['=J19-J15']];
rg('J20').formulas=[['=J16+J18']];
const unitFormula='=InputLengthUnit';
const unitCells=['K9','K10',...Array.from({length:6},(_,i)=>'K'+(15+i)),'J26','K26','M26','N26','J37','K37','N37','O37'];
unitCells.forEach(a=>rg(a).formulas=[[unitFormula]]);
const keyFormula=r=>`=IF(OR(selGeometryType="",I${r}=""),"",selGeometryType&"."&INDEX(rngCurrentSideKeys,MATCH(I${r},rngCurrentSideNames,0)))`;
for(let i=0;i<6;i++){
  rg('H'+mainRows[i]).formulas=[[keyFormula(mainRows[i])]];
  rg('H'+extraRows[i]).formulas=[['=H'+mainRows[i]+'&""']];
  rg('I'+extraRows[i]).formulas=[['=I'+mainRows[i]+'&""']];
}
const validations=[];
let nextColumn=0;
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
const inputCells=['J8','J9','J10','J15','J16','J18','J19','B5'];
for(let i=0;i<6;i++){
  const r=mainRows[i],e=extraRows[i];
  choice('I'+r,'lstSide'+(i+1),sides.map(s=>s[1]),sides.map((s,j)=>`='Config'!I${15+j}`));
  inputCells.push(...['I','J','K','L','M','N'].map(c=>c+r));
  for(const [c,kind,values] of [['L','Bind2',['EachBar','EverySecondBar']],['M','Pos2',['Stacked','SideBySide']],['P','Bind3',['EachBar','EverySecondBar']],['Q','Pos3',['Stacked','SideBySide']]]){
    choice(c+e,`lst${kind}_${i+1}`,values);inputCells.push(c+e);
  }
  inputCells.push(...['J','K','N','O'].map(c=>c+e));
  for(const c of ['J','K','M','N'])rg(c+r).dataValidation={rule:{type:'decimal',operator:'greaterThanOrEqual',formula1:0}};
  rg('L'+r).dataValidation={rule:{type:'whole',operator:'greaterThanOrEqual',formula1:0}};
  for(const c of ['J','K','N','O'])rg(c+e).dataValidation={rule:{type:'decimal',operator:'greaterThanOrEqual',formula1:0}};
}
inputCells.forEach(a=>{rg(a).format.fill='#FFFFFF';rg(a).format.horizontalAlignment='center';});
for(const a of ['J9','J10','J15','J16','J17','J18','J19','J20','J27:N32','J38:K43','N38:O43'])rg(a).setNumberFormat('General');
for(const r of mainRows)rg('L'+r).setNumberFormat('0');
wb.recalculate();
const expected=sides.map(s=>'LSection.'+s[0]);
const actual=rg('H27:H32').values.flat();
if(JSON.stringify(actual)!==JSON.stringify(expected))throw Error('Dynamic key formulas did not calculate: '+JSON.stringify(actual));
if(rg('J17').values[0][0]!==350||rg('J20').values[0][0]!==800)throw Error('L-section dimension formulas failed.');
console.log((await wb.inspect({kind:'region',sheetId:'Config',range:'H27:K32',maxChars:1600,tableMaxRows:6,tableMaxCols:4,include:'values,formulas'})).ndjson);
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!|not implemented',options:{useRegex:true,maxResults:20},summary:'Formula error scan'});
console.log(errors.ndjson);
if(errors.ndjson.includes('"kind":"match"'))throw Error('Unexpected formula errors.');
await fs.mkdir(out,{recursive:true});
const file=await SpreadsheetFile.exportXlsx(wb);await file.save(path.join(out,filename));
for(const [name,range] of [['Geometry','H4:L20'],['Rebar','H22:R43']]){
  const image=await wb.render({sheetName:'Config',range,scale:1.5,format:'png'});
  await fs.writeFile(name+'_Artifact.png',new Uint8Array(await image.arrayBuffer()));
}
await fs.writeFile('Manifest.json',JSON.stringify({file:path.join(out,filename),geometry:'H5:L20',rebar:'H23:R43',mainRows,extraRows,sides:sides.map(s=>({key:s[0],name:s[1]})),expectedKeys:expected,validations,inputCells,unitCells,mainWorkbookSHA256:'ABAA60FCB940F5C850571F515BF4C7C82DA474E7002B7DD122C54BB4E147C45C'},null,2));
console.log('SHARED_REBAR_MOCKUP: '+path.join(out,filename));
