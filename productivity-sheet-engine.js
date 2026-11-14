(() => {
"use strict";
const byId = id => document.getElementById(id);
const sheetDefaults = () => Array.from({length:12}, () => Array(8).fill(""));
if (!Array.isArray(data.sheets) || !data.sheets.length) {
  data.sheets = [{id:"sheet-" + Date.now(), name:"Sheet1", cells:Array.isArray(data.sheet)?data.sheet:sheetDefaults(), styles:{}}];
}
data.sheets = data.sheets.map((s,i) => ({
  id:String(s.id || ("sheet-"+i)), name:String(s.name || ("Sheet"+(i+1))),
  cells:Array.isArray(s.cells)?s.cells:sheetDefaults(), styles:s.styles && typeof s.styles==="object"?s.styles:{}
}));
let activeSheetId = data.activeSheetId && data.sheets.some(s=>s.id===data.activeSheetId) ? data.activeSheetId : data.sheets[0].id;
let selectedCell = {r:0,c:0}, sortDescending = false;
const activeSheet = () => data.sheets.find(s=>s.id===activeSheetId) || data.sheets[0];
const syncSheet = () => {
  const s=activeSheet(); data.sheet=s.cells;
  rows=Math.max(1,s.cells.length); cols=Math.max(1,...s.cells.map(r=>r.length));
  s.cells.forEach(r=>{while(r.length<cols)r.push("");});
  data.activeSheetId=activeSheetId;
};
syncSheet();
const colLetters = n => {let out="";for(n++;n;n=Math.floor((n-1)/26))out=String.fromCharCode(65+(n-1)%26)+out;return out;};
const parseRef = ref => {const m=/^([A-Z]+)([1-9]\d*)$/i.exec(ref);if(!m)throw Error("Bad reference");let c=0;for(const ch of m[1].toUpperCase())c=c*26+ch.charCodeAt(0)-64;return {r:Number(m[2])-1,c:c-1};};
const rawAt = (r,c) => String(activeSheet().cells[r]?.[c] ?? "");
const cellValue = (r,c,stack=[]) => {
  const raw=rawAt(r,c); if(!raw.startsWith("=")) return raw.trim()!=="" && Number.isFinite(Number(raw)) ? Number(raw) : raw;
  const key=r+":"+c; if(stack.includes(key))return "#CYCLE";
  try{return parseFormula(raw.slice(1),[...stack,key]);}catch{return "#ERROR";}
};
const tokenize = source => {
  const out=[]; let i=0;
  while(i<source.length){
    if(/\s/.test(source[i])){i++;continue;}
    const rest=source.slice(i);
    const m=rest.match(/^(?:(\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?|("(?:[^"]|"")*")|(\$?[A-Z]{1,3}\$?[1-9]\d*)|([A-Za-z_][A-Za-z0-9_]*)|(<=|>=|<>|!=|==|[+\-*\/^%(),:<>]))/);
    if(!m)throw Error("Unexpected character");
    const raw=m[0]; let type="op",value=raw;
    if(m[1]!==undefined){type="number";value=Number(raw);}
    else if(m[2]!==undefined){type="string";value=m[2].slice(1,-1).replace(/""/g,'"');}
    else if(m[3]!==undefined){type="ref";value=raw.replace(/\$/g,"").toUpperCase();}
    else if(m[4]!==undefined){type="name";value=raw.toUpperCase();}
    out.push({type,value});i+=raw.length;
  }
  return out;
};
const parseFormula = (source,stack) => {
  const t=tokenize(source);let p=0;
  const peek=()=>t[p], take=()=>t[p++];
  const accept=v=>peek()&&peek().value===v?(p++,true):false;
  const expect=v=>{if(!accept(v))throw Error("Expected "+v);};
  const flatten=x=>Array.isArray(x)?x.flat(Infinity):[x];
  const refVal=ref=>{const q=parseRef(ref);return cellValue(q.r,q.c,stack);};
  const compare=(a,op,b)=>{const x=typeof a==="number"?a:String(a).toLowerCase(),y=typeof b==="number"?b:String(b).toLowerCase();switch(op){case"=":case"==":return x===y;case"<>":case"!=":return x!==y;case">":return x>y;case"<":return x<y;case">=":return x>=y;case"<=":return x<=y;}};
  const primary=()=>{
    const token=take();if(!token)throw Error("Missing value");
    if(token.type==="number"||token.type==="string")return token.value;
    if(token.type==="ref"){
      if(accept(":")){const end=take();if(!end||end.type!=="ref")throw Error("Bad range");const a=parseRef(token.value),b=parseRef(end.value),values=[];for(let r=Math.min(a.r,b.r);r<=Math.max(a.r,b.r);r++)for(let c=Math.min(a.c,b.c);c<=Math.max(a.c,b.c);c++)values.push(cellValue(r,c,stack));return values;}
      return refVal(token.value);
    }
    if(token.type==="name"){
      if(token.value==="TRUE")return true;if(token.value==="FALSE")return false;
      if(!accept("("))throw Error("Unknown name");
      const args=[];if(!accept(")")){do{args.push(expression())}while(accept(","));expect(")");}
      const flat=args.flat(Infinity), nums=flat.filter(v=>typeof v==="number"&&Number.isFinite(v));
      const fn=token.value;
      switch(fn){
        case"SUM":return nums.reduce((a,b)=>a+b,0);
        case"AVERAGE":return nums.length?nums.reduce((a,b)=>a+b,0)/nums.length:0;
        case"MIN":return nums.length?Math.min(...nums):0;
        case"MAX":return nums.length?Math.max(...nums):0;
        case"COUNT":return nums.length;
        case"COUNTA":return flat.filter(v=>v!==""&&v!==null&&v!==undefined).length;
        case"IF":return args[0]?(args[0]?args[1]:args[2]??false):false;
        case"AND":return flat.every(Boolean);
        case"OR":return flat.some(Boolean);
        case"NOT":return !args[0];
        case"ROUND":return Number(Number(args[0]||0).toFixed(Math.max(0,Math.min(10,Number(args[1]||0)))));
        case"ABS":return Math.abs(Number(args[0]||0));
        case"SQRT":return Math.sqrt(Number(args[0]));
        case"POWER":return Math.pow(Number(args[0]),Number(args[1]));
        case"CONCAT":return flat.join("");
        case"LEN":return String(args[0]??"").length;
        case"UPPER":return String(args[0]??"").toUpperCase();
        case"LOWER":return String(args[0]??"").toLowerCase();
        default:throw Error("Unsupported function");
      }
    }
    if(token.value==="("){const v=expression();expect(")");return v;}
    if(token.value==="-"||token.value==="+"){const v=Number(unary());return token.value==="-"?-v:v;}
    throw Error("Invalid expression");
  };
  const unary=()=>primary();
  const power=()=>{let a=unary();if(accept("^"))a=Math.pow(Number(a),Number(power()));return a;};
  const product=()=>{let a=power();while(peek()&&["*","/","%"].includes(peek().value)){const op=take().value,b=Number(power());a=op==="*"?Number(a)*b:op==="/"?Number(a)/b:Number(a)%b;}return a;};
  const sum=()=>{let a=product();while(peek()&&["+","-"].includes(peek().value)){const op=take().value,b=Number(product());a=op==="+"?Number(a)+b:Number(a)-b;}return a;};
  const expression=()=>{let a=sum();while(peek()&&[">","<",">=","<=","=","==","<>","!="].includes(peek().value)){const op=take().value;a=compare(a,op,sum());}return a;};
  const result=expression();if(p!==t.length)throw Error("Unexpected token");return result;
};
const displayValue = (r,c) => {
  const v=cellValue(r,c);if(typeof v==="number"&&Number.isFinite(v))return String(Number(v.toPrecision(12)));return String(v??"");
};
const styleFor = (r,c) => activeSheet().styles[r+":"+c] || {};
const safeColor = v => /^#[0-9a-f]{6}$/i.test(v||"")?v:"";
const renderTabs = () => {
  const host=byId("sheetTabs");if(!host)return;host.innerHTML="";
  data.sheets.forEach(s=>{const b=document.createElement("button");b.type="button";b.textContent=s.name;b.className=s.id===activeSheetId?"active":"";b.setAttribute("aria-pressed",s.id===activeSheetId?"true":"false");b.onclick=()=>switchSheet(s.id);host.append(b);});
};
const updateFormulaBar = (r,c) => {
  selectedCell={r,c};const formula=byId("formulaInput");if(formula)formula.value=rawAt(r,c);
  const address=byId("cellAddress");if(address)address.textContent=colLetters(c)+(r+1);
  const st=styleFor(r,c);if(byId("fontBold"))byId("fontBold").classList.toggle("active",!!st.bold);
  if(byId("fontItalic"))byId("fontItalic").classList.toggle("active",!!st.italic);
  if(byId("fontColor"))byId("fontColor").value=safeColor(st.color)||"#172033";
  if(byId("fillColor"))byId("fillColor").value=safeColor(st.background)||"#ffffff";
};
const saveActive = () => {data.sheet=activeSheet().cells;persist();};
const setCell = (r,c,value) => {const s=activeSheet();while(s.cells.length<=r)s.cells.push(Array(cols).fill(""));while(s.cells[r].length<=c)s.cells[r].push("");s.cells[r][c]=value;saveActive();renderSheet();};
const renderSheet = () => {
  syncSheet();const s=activeSheet(),table=byId("sheetTable");if(!table)return;
  renderTabs();
  const filter=(byId("sheetFilter")?.value||"").trim().toLowerCase();
  let out='<thead><tr><th class="rowhead">#</th>'+Array.from({length:cols},(_,c)=>'<th>'+colLetters(c)+'</th>').join("")+'</tr></thead><tbody>';
  for(let r=0;r<rows;r++){
    const rowText=s.cells[r].map((v,c)=>String(cellValue(r,c))).join(" ").toLowerCase();
    const hidden=filter&&!rowText.includes(filter);
    out+='<tr data-row="'+r+'"'+(hidden?' style="display:none"':"")+'><th class="rowhead">'+(r+1)+'</th>';
    for(let c=0;c<cols;c++){
      const st=styleFor(r,c),css=[];
      if(st.bold)css.push("font-weight:700");if(st.italic)css.push("font-style:italic");
      if(st.align&&["left","center","right"].includes(st.align))css.push("text-align:"+st.align);
      if(safeColor(st.color))css.push("color:"+st.color);if(safeColor(st.background))css.push("background-color:"+st.background);
      const raw=rawAt(r,c),shown=raw.startsWith("=")?displayValue(r,c):raw;
      out+='<td><input data-r="'+r+'" data-c="'+c+'" aria-label="'+colLetters(c)+(r+1)+'" value="'+escapeHtml(shown)+'" style="'+css.join(";")+'" title="'+escapeHtml(raw)+'"></td>';
    }out+="</tr>";
  }out+="</tbody>";table.innerHTML=out;
  table.querySelectorAll("input[data-r]").forEach(input=>{
    const r=Number(input.dataset.r),c=Number(input.dataset.c);
    input.onfocus=()=>{updateFormulaBar(r,c);if(rawAt(r,c).startsWith("="))input.value=rawAt(r,c);};
    input.onchange=()=>setCell(r,c,input.value);
    input.onclick=()=>updateFormulaBar(r,c);
    input.onkeydown=e=>{if(e.key==="Enter"||e.key==="Tab"){e.preventDefault();const nr=Math.max(0,Math.min(rows-1,r+(e.key==="Enter"?1:0))),nc=Math.max(0,Math.min(cols-1,c+(e.key==="Tab"?1:0)));const next=table.querySelector('input[data-r="'+nr+'"][data-c="'+nc+'"]');if(next)next.focus();}};
  });
  const name=byId("sheetName");if(name)name.value=s.name;
  const status=byId("sheetStatus");if(status)status.textContent=rows+" × "+cols+" cells • "+data.sheets.length+" worksheet(s) • formulas enabled";
  updateFormulaBar(selectedCell.r<rows?selectedCell.r:0,selectedCell.c<cols?selectedCell.c:0);
};
const switchSheet = id => {activeSheetId=id;syncSheet();selectedCell={r:0,c:0};renderSheet();};
const addWorksheet = () => {const n=data.sheets.length+1,s={id:"sheet-"+Date.now()+"-"+Math.random().toString(36).slice(2,7),name:"Sheet"+n,cells:sheetDefaults(),styles:{}};data.sheets.push(s);switchSheet(s.id);persist();};
const deleteWorksheet = () => {if(data.sheets.length===1){alert("A workbook must keep at least one worksheet.");return;}if(!confirm("Delete worksheet '"+activeSheet().name+"'? This cannot be undone unless you have a backup."))return;data.sheets=data.sheets.filter(s=>s.id!==activeSheetId);activeSheetId=data.sheets[0].id;syncSheet();persist();renderSheet();};
const applyStyle = patch => {const s=activeSheet(),key=selectedCell.r+":"+selectedCell.c;s.styles[key]={...(s.styles[key]||{}),...patch};saveActive();renderSheet();};
const sortSelectedColumn = () => {const s=activeSheet(),c=selectedCell.c,decorated=s.cells.map((row,i)=>({row,index:i,style:Object.fromEntries(Object.entries(s.styles).filter(([k])=>k.startsWith(i+":")).map(([k,v])=>[k.slice(k.indexOf(":")+1),v]))}));sortDescending=!sortDescending;decorated.sort((a,b)=>{const x=cellSort(a.row[c]),y=cellSort(b.row[c]);return (x<y?-1:x>y?1:a.index-b.index)*(sortDescending?-1:1);});s.cells=decorated.map(x=>x.row);s.styles={};decorated.forEach((item,r)=>Object.entries(item.style).forEach(([col,st])=>s.styles[r+":"+col]=st));saveActive();renderSheet();};
const cellSort = row => {const v=String(row[selectedCell.c]??"");if(v.startsWith("="))return "";const n=Number(v);return v.trim()!==""&&Number.isFinite(n)?n:v.toLowerCase();};
const bind = () => {
  byId("addWorksheet").onclick=addWorksheet;byId("deleteWorksheet").onclick=deleteWorksheet;
  byId("renameWorksheet").onclick=()=>{const name=prompt("Worksheet name",activeSheet().name);if(name&&name.trim()){activeSheet().name=name.trim().slice(0,31);persist();renderSheet();}};
  byId("sheetName").onchange=e=>{activeSheet().name=e.target.value.trim().slice(0,31)||"Sheet";persist();renderTabs();};
  byId("formulaInput").onchange=e=>setCell(selectedCell.r,selectedCell.c,e.target.value);
  byId("fontBold").onclick=()=>applyStyle({bold:!styleFor(selectedCell.r,selectedCell.c).bold});
  byId("fontItalic").onclick=()=>applyStyle({italic:!styleFor(selectedCell.r,selectedCell.c).italic});
  byId("fontColor").oninput=e=>applyStyle({color:e.target.value});
  byId("fillColor").oninput=e=>applyStyle({background:e.target.value});
  byId("alignCells").onchange=e=>applyStyle({align:e.target.value});
  byId("sheetFilter").oninput=renderSheet;
  byId("sortColumn").onclick=sortSelectedColumn;
  byId("clearFilter").onclick=()=>{byId("sheetFilter").value="";renderSheet();};
  byId("addRow").onclick=()=>{activeSheet().cells.push(Array(cols).fill(""));saveActive();renderSheet();};
  byId("addCol").onclick=()=>{activeSheet().cells.forEach(r=>r.push(""));saveActive();renderSheet();};
  const oldCsv=byId("csvFile").onchange;
  byId("csvFile").onchange=async e=>{const f=e.target.files[0];if(!f)return;try{const parsed=csvParse(await f.text());if(!parsed.length)throw Error("No rows found");const ncols=Math.max(1,...parsed.map(r=>r.length));activeSheet().cells=parsed.map(r=>Array.from({length:ncols},(_,i)=>r[i]||""));activeSheet().styles={};saveActive();renderSheet();}catch(err){alert("Could not import CSV: "+err.message);}e.target.value="";};
  byId("csvExport").onclick=()=>{const csv=activeSheet().cells.map(r=>r.map(v=>'"'+String(v).replace(/"/g,'""')+'"').join(",")).join("\r\n");download((activeSheet().name||"sheet")+".csv",csv,"text/csv;charset=utf-8");};
};
if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",()=>{bind();renderSheet();},{once:true});else{bind();renderSheet();}
window.OpenCreativeSheets={getWorkbook:()=>data.sheets.map(s=>({name:s.name,cells:s.cells,styles:s.styles})),setWorkbook:sheets=>{if(!Array.isArray(sheets)||!sheets.length)throw Error("Workbook needs at least one sheet");data.sheets=sheets.map((s,i)=>({id:"import-"+i+"-"+Date.now(),name:String(s.name||"Sheet"+(i+1)).slice(0,31),cells:s.cells,styles:s.styles||{}}));activeSheetId=data.sheets[0].id;syncSheet();persist();renderSheet();},setImportedStyles:(sheetIndex,styles)=>{if(data.sheets[sheetIndex])data.sheets[sheetIndex].styles=styles||{};persist();renderSheet();}};
})();