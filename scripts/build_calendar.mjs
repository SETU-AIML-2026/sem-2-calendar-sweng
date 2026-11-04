import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from 'file:///C:/Users/fxwal/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';
const dir=new URL('.',import.meta.url).pathname.replace(/^\/([A-Z]:)/,'$1');
const raw=await fs.readFile(dir+'../calendar.yaml','utf8');
const weeks=[...raw.matchAll(/  - date: (\d{4}-\d{2}-\d{2})\s+week: (\d+|~)\s+topic: ([^\n]+)/g)];
if(weeks.length!==17)throw Error('Expected 17 calendar weeks');
const wb=Workbook.create(),s=wb.worksheets.add('calendar-2027');s.showGridLines=false;
s.getRange('A1:J24').format.font={name:'Arial',size:11,color:'#222222'};
s.getRange('A1:J24').format.verticalAlignment='center';
s.getRange('A1:A24').format.columnWidth=15;
s.getRange('B1:B24').format.columnWidth=17;
s.getRange('C1:I24').format.columnWidth=9;
s.getRange('J1:J24').format.columnWidth=42;
s.mergeCells('A1:B1');s.mergeCells('A2:B2');
s.getRange('A1').values=[[2027]];s.getRange('A1').format.font={size:16,bold:true,color:'#626262'};
s.getRange('A2:J2').values=[['Semester 2',null,'M','T','W','T','F','S','S','Modules']];
s.getRange('A2:J2').format={fill:'#F2F2F2',rowHeight:26,borders:{bottom:{style:'thin',color:'#444444'}}};
s.getRange('A2:B2').format.font={bold:true,color:'#626262'};
s.getRange('C2:J2').format.horizontalAlignment='center';
s.getRange('A3:J3').format.rowHeight=18;
const epoch=Date.UTC(1899,11,30);let prev='';
for(let i=0;i<weeks.length;i++){
 const [,date,num,topicRaw]=weeks[i],r=i+4;
 const topic=topicRaw.trim();
 const mon=new Date(date+'T00:00:00Z');
 // Show a month label in the first row containing a day in that month.
 const sunday=new Date(mon.getTime()+6*86400000);
 const month=sunday.toLocaleString('en-IE',{month:'long',timeZone:'UTC'});
 if(month!==prev){s.getRange(`A${r}`).values=[[month]];prev=month;}
 s.getRange(`A${r}:B${r}`).format.font={bold:true,color:'#626262'};
 s.getRange(`B${r}`).values=[[num==='~'?(topic==='Reading Week'?'reading-week':topic==='Easter Break'?'EASTER':null):Number(num)]];
 s.getRange(`B${r}`).format.horizontalAlignment='right';
 s.getRange(`B${r}`).format.font={bold:true,italic:true,color:'#626262'};
 s.getRange(`C${r}`).values=[[(mon.getTime()-epoch)/86400000]];
 for(let j=1;j<7;j++)s.getRange(`${String.fromCharCode(67+j)}${r}`).formulas=[[`=C${r}+${j}`]];
 s.getRange(`C${r}:I${r}`).setNumberFormat('d');
 s.getRange(`C${r}:I${r}`).format.horizontalAlignment='right';
 s.getRange(`C${r}:J${r}`).format.borders={preset:'all',style:'thin',color:'#333333'};
 s.getRange(`A${r}:J${r}`).format.rowHeight=25;
 if(num!=='~'){
  s.getRange(`J${r}`).values=[['Security & Front End Web App Dev']];
  s.getRange(`J${r}`).format.horizontalAlignment='center';
  s.getRange(`C${r}`).format.fill='#D6C6EA';
  s.getRange(`G${r}`).format.fill='#D6C6EA';
  s.getRange(`E${r}`).format.fill='#75BFFF';
 }
 if(num==='~' && !topic.startsWith('January On-Site'))s.getRange(`C${r}:J${r}`).format.fill=topic==='Reading Week'?'#FFF3D6':'#EEE8F5';
}
s.getRange('G4').format.fill='#DCEEE5';
s.getRange('G4').format.font={bold:true,color:'#20533D'};
s.getRange('J4').values=[['15 Jan: January On-Site']];
s.getRange('J4').format={fill:'#DCEEE5',wrapText:true,font:{bold:true,color:'#20533D'}};
s.getRange('A4:J4').format.rowHeight=36;
s.getRange('A22').values=[['Purple: Security (Mon / Fri). Blue: Front End Web App Dev (Wed). Breaks and on-site are shaded separately.']];
s.getRange('A23').values=[['Teaching: 18 January - 7 May 2027. Each row runs Monday to Sunday.']];
s.getRange('A24').values=[['Sources: calendar.yaml and 2027 Software Engineering delivery pattern. Times and assessments to follow.']];
s.getRange('A22:J24').format.font={size:10,color:'#626262'};
s.getRange('L1:N25').format.columnWidth=9;
s.getRange('L1:N25').format.font={name:'Arial',size:10,color:'#626262'};
s.getRange('L2').values=[['Delivery key']];
s.getRange('L2').format.font={bold:true,size:12};
s.getRange('L3').values=[['Purple: Mon / Fri']];
s.getRange('L14').values=[['Blue: Wedrsday']];
for(const [file,row] of [['security.png',3],['front-end.png',14]]){
 const bytes=await fs.readFile(dir+'assets/'+file);
 s.images.add({dataUrl:'data:image/png;base64,'+bytes.toString('base64'),anchor:{from:{row,col:11},extent:{widthPx:145,heightPx:136}}});
}
s.freezePanes.freezeRows(3);
wb.recalculate();
console.log((await wb.inspect({kind:'table',range:'calendar-2027!A4:J20',tableMaxRows:17,tableMaxCols:10,maxChars:2000})).ndjson);
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#VALUE!|#NAME\\?',options:{useRegex:true,maxResults:10}})).ndjson);
await(await SpreadsheetFile.exportXlsx(wb)).save(dir+'../unit-1-sem-2/Semester_2_2027_sweng.xlsx');
const img=await wb.render({sheetName:'calendar-2027',range:'A1:N26',scale:1.5,format:'png'});
await fs.writeFile(dir+'styled-calendar.png',new Uint8Array(await img.arrayBuffer()));

