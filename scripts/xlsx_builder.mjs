// Public Artifact Tool API only. Exact print/rich-text settings are patched by office.py.
import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const [input,output]=process.argv.slice(2);
const spec=JSON.parse(await fs.readFile(input,'utf8'));
const wb=Workbook.create();
function col(n){let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;}
for(const t of spec.sheets){
  const sh=wb.worksheets.add(t.name??'表格');sh.showGridLines=false;
  const offset=t.attachment_label?2:0, title=1+offset, hr=2+offset;
  const end=hr+t.rows.length+1, last=col(t.headers.length);
  const area=sh.getRange(`A1:${last}${end}`);
  area.format.font.name='仿宋_GB2312';area.format.font.size=16;
  area.format.font.color='#000000';area.format.horizontalAlignment='center';
  area.format.verticalAlignment='center';area.format.rowHeight=26;
  area.format.borders={preset:'all',style:'thin',color:'#000000'};
  const literal=v=>typeof v==='string'&&v.startsWith('=')?"'"+v:v;
  const values=[t.headers.map(literal),...t.rows.map(r=>r.map(v=>v&&typeof v==='object'?null:literal(v)))];
  sh.getRange(`A${hr}:${last}${hr+t.rows.length}`).values=values;
  for(let ri=0;ri<t.rows.length;ri++)for(let ci=0;ci<t.headers.length;ci++){
    const v=t.rows[ri][ci];
    if(v&&typeof v==='object')sh.getRange(`${col(ci+1)}${hr+ri+1}`).formulas=[[v.formula]];
    if(t.identifier_columns?.includes(ci))sh.getRange(`${col(ci+1)}${hr+ri+1}`).setNumberFormat('@');
  }
  sh.getRange(`A${title}`).values=[[t.title]];
  sh.getRange(`A${title}:${last}${title}`).merge();
  const titleWidth=[...t.title].reduce((s,c)=>s+(/[\u2e80-\uffff]/u.test(c)?22:12.1),0);
  sh.getRange(`A${title}:${last}${title}`).format.rowHeight=Math.max(35,Math.ceil(titleWidth/(t.widths_mm.reduce((a,b)=>a+b,0)*72/25.4-8))*28);
  sh.getRange(`A${hr}:${last}${hr}`).format.font.name='黑体';
  sh.getRange(`A${end}`).values=[[t.organization+' 制']];
  sh.getRange(`A${end}:${last}${end}`).merge();
  if(offset){sh.getRange('A1').values=[[t.attachment_label]];sh.getRange(`A1:${last}1`).merge();}
  for(let ci=0;ci<t.headers.length;ci++)sh.getRange(`${col(ci+1)}1:${col(ci+1)}${end}`).format.columnWidthPx=t.widths_mm[ci]*96/25.4;
  // Repeated-value grouping is applied only to explicit, non-identifier columns.
  for(const ci of t.merge_columns??[]){
    let a=0;
    while(a<t.rows.length){let b=a+1;while(b<t.rows.length&&t.rows[b][ci]===t.rows[a][ci]&&t.rows[b].slice(0,ci).every((v,k)=>v===t.rows[a][k]))b++;
      if(b-a>1)sh.getRange(`${col(ci+1)}${hr+a+1}:${col(ci+1)}${hr+b}`).merge();a=b;}
  }
  sh.freezePanes.freezeRows(hr);
}
wb.recalculate();
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:30}});
await fs.writeFile(output+'.formula-scan.json',errors.ndjson??'');
const blob=await SpreadsheetFile.exportXlsx(wb);await blob.save(output);
