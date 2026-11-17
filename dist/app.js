const $ = s => document.querySelector(s);
let base = null, sources = [];
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const colName = i => XLSX.utils.encode_col(i);
const cleanName = name => name.replace(/\.[^.]+$/, '');
const norm=v=>String(v??'').normalize('NFKC').toLowerCase().replace(/ё/g,'е').replace(/\s+/g,' ').trim();
function autoKeys(s){const h=(s.rows[s.header]||[]).map(norm);s.code=h.findIndex(v=>/^(код|артикул|sku)/.test(v));s.name=h.findIndex(v=>/^(наименование|название|номенклатура|товар)/.test(v));}
function keyFields(s){const h=s.rows[s.header]||[];const options='<option value="-1">Не выбран</option>'+Array.from({length:maxCols(s)},(_,i)=>'<option value="'+i+'">'+colName(i)+' · '+esc(h[i]||'Без заголовка')+'</option>').join('');return '<label>Артикул / код<select data-key="code">'+options+'</select></label><label>Название товара<select data-key="name">'+options+'</select></label>';}

const message = s => { $('#message').textContent = s; };
function setSheet(item, reset = false) {
  const ws = item.wb.Sheets[item.sheet];
  item.rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:'', blankrows:true});
  item.header = reset ? 0 : Math.max(0, Math.min(item.header || 0, item.rows.length - 1));
  if (reset) item.selected = [];
  autoKeys(item);
}
async function read(file) {
  if (file.size > 40 * 1024 * 1024) throw Error('Файл больше 40 МБ: ' + file.name);
  const wb = XLSX.read(await file.arrayBuffer(), {type:'array', cellStyles:true, cellNF:true, cellDates:true});
  if (!wb.SheetNames.length) throw Error('В файле нет листов: ' + file.name);
  const item = {filename:file.name, wb, sheet:wb.SheetNames[0], header:0, selected:[]};
  setSheet(item, true);
  return item;
}
function maxCols(s) {
  const ws = s.wb.Sheets[s.sheet];
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  return Math.max(range.e.c + 1, ...s.rows.slice(0, Math.min(s.rows.length, 50)).map(r => r.length), 1);
}
function config(s, isBase, idx) {
  const head = s.rows[s.header] || [];
  const count = maxCols(s);
  return '<div class="config" data-base="'+isBase+'" data-index="'+idx+'">'+
    '<div class="config-top"><strong>▤ '+esc(s.filename)+'</strong><button data-remove aria-label="Удалить файл">×</button></div>'+
    '<div class="fields"><label>Лист<select data-sheet>'+s.wb.SheetNames.map(n=>'<option '+(n===s.sheet?'selected':'')+'>'+esc(n)+'</option>').join('')+'</select></label>'+
    '<label>Строка заголовков<input data-header type="number" min="1" max="'+Math.max(1,s.rows.length)+'" value="'+(s.header+1)+'"></label>'+keyFields(s)+
    (isBase ? '<p class="base-note">Все существующие столбцы исходного файла сохраняются. Новые добавляются справа.</p>' :
    '<div class="multi-price"><strong>Выберите столбцы для переноса</strong><div class="price-checks">'+
    Array.from({length:count},(_,i)=>'<label class="price-check"><input type="checkbox" data-col="'+i+'" '+(s.selected.includes(i)?'checked':'')+'><span>'+colName(i)+' · '+esc(head[i] || 'Без заголовка')+'</span></label>').join('')+
    '</div></div>')+'</div></div>';
}
function render() {
  $('#baseConfig').innerHTML = base ? config(base,true,0) : '';
  $('#pricesConfig').innerHTML = sources.map((s,i)=>config(s,false,i)).join('');
  document.querySelectorAll('.config').forEach(el=>{
    const isBase=el.dataset.base==='true', idx=Number(el.dataset.index), s=isBase?base:sources[idx];
    el.querySelector('[data-remove]').onclick=()=>{if(isBase)base=null;else sources.splice(idx,1);render();};
    el.querySelector('[data-sheet]').onchange=e=>{s.sheet=e.target.value;setSheet(s,true);render();};
    el.querySelector('[data-header]').onchange=e=>{s.header=Math.max(0,Math.min(s.rows.length-1,Number(e.target.value)-1));autoKeys(s);render();};
    el.querySelectorAll('[data-key]').forEach(input=>{input.value=String(s[input.dataset.key]);input.onchange=()=>{s[input.dataset.key]=Number(input.value);update();};});
    el.querySelectorAll('[data-col]').forEach(input=>input.onchange=()=>{
      const col=Number(input.dataset.col);
      s.selected=input.checked?[...s.selected,col].sort((a,b)=>a-b):s.selected.filter(x=>x!==col);
      update();
    });
  });
  update();
}
function update() {
  const n=sources.reduce((a,s)=>a+s.selected.length,0);
  $('#run').disabled=!base||!n;
  $('#selection').textContent=n ? 'Выбрано столбцов: '+n : 'Отметьте столбцы, которые нужно добавить';
}
async function load(files,isBase) {
  try {
    message('Читаем Excel…');
    const items=[];
    for (const f of files) {items.push(await read(f));if(isBase)break;}
    if(isBase)base=items[0]||base;else sources.push(...items);
    message('');render();
  } catch(e) {message('Ошибка чтения: '+e.message);}
}
for (const [input,drop,isBase] of [['#baseInput','#baseDrop',true],['#pricesInput','#pricesDrop',false]]) {
  $(input).onchange=e=>{load(e.target.files,isBase);e.target.value='';};
  $(drop).ondragover=e=>{e.preventDefault();$(drop).classList.add('drag');};
  $(drop).ondragleave=()=>$(drop).classList.remove('drag');
  $(drop).ondrop=e=>{e.preventDefault();$(drop).classList.remove('drag');load(e.dataTransfer.files,isBase);};
}
function buildIndex(s,mode){
  const byCode=new Map(),byName=new Map();
  const add=(map,key,row)=>{if(!key)return;if(!map.has(key))map.set(key,[]);map.get(key).push(row);};
  for(let row=s.header+1;row<s.rows.length;row++){
    const data=s.rows[row]||[];
    if(mode!=='name'&&s.code>=0)add(byCode,norm(data[s.code]),row);
    if(mode!=='code'&&s.name>=0)add(byName,norm(data[s.name]),row);
  }
  return {byCode,byName};
}
function findSourceRow(data,b,s,index,mode){
  const code=b.code>=0?norm(data[b.code]):'';
  const name=b.name>=0?norm(data[b.name]):'';
  if(mode!=='name'&&code&&s.code>=0){
    const rows=index.byCode.get(code)||[];
    if(rows.length===1)return rows[0];
    if(rows.length>1)return -1;
  }
  if(mode!=='code'&&name&&s.name>=0){
    const rows=index.byName.get(name)||[];
    if(rows.length===1)return rows[0];
  }
  return -1;
}
function appendColumns(wb,b,sources,mode){
  const ws=wb.Sheets[b.sheet],range=XLSX.utils.decode_range(ws['!ref']||'A1');
  const originalEnd=range.e.c;
  let nextCol=originalEnd+1,filled=0,empty=0;
  for(const s of sources){
    if(!s.selected.length)continue;
    if(mode==='code'&&(b.code<0||s.code<0))throw Error('Выберите колонки артикулов: '+s.filename);
    if(mode==='name'&&(b.name<0||s.name<0))throw Error('Выберите колонки названий: '+s.filename);
    const index=buildIndex(s,mode),sourceWs=s.wb.Sheets[s.sheet];
    for(const sourceCol of s.selected){
      const targetCol=nextCol++;
      const label=cleanName(s.filename)+(s.selected.length>1?' ('+colName(sourceCol)+')':'');
      ws[XLSX.utils.encode_cell({r:b.header,c:targetCol})]={t:'s',v:label};
      for(let row=b.header+1;row<=range.e.r;row++){
        const sourceRow=findSourceRow(b.rows[row]||[],b,s,index,mode);
        if(sourceRow<0){empty++;continue;}
        const cell=sourceWs[XLSX.utils.encode_cell({r:sourceRow,c:sourceCol})];
        if(!cell||cell.t==='z'||cell.v===undefined||cell.v===''){empty++;continue;}
        const copy={...cell};delete copy.f;delete copy.F;
        if(cell.f){copy.t='s';copy.v=cell.w??String(cell.v??'');}
        ws[XLSX.utils.encode_cell({r:row,c:targetCol})]=copy;filled++;
      }
    }
  }
  range.e.c=Math.max(originalEnd,nextCol-1);ws['!ref']=XLSX.utils.encode_range(range);
  return {columns:nextCol-originalEnd-1,filled,empty};
}
$('#run').onclick=()=>{
  try {
    if(!base)throw Error('Загрузите основной файл.');
    if(!sources.some(s=>s.selected.length))throw Error('Выберите хотя бы один столбец.');
    const wb=XLSX.read(XLSX.write(base.wb,{type:'array',bookType:'xlsx'}),{type:'array',cellStyles:true,cellNF:true,cellDates:true});
    const n=appendColumns(wb,base,sources,$('#mode').value);
    XLSX.writeFile(wb,cleanName(base.filename)+'_с_добавленными_столбцами.xlsx');
    message('Готово! Добавлено столбцов: '+n.columns+'. Заполнено: '+n.filled+'. Пустых: '+n.empty+'. Исходный файл не перезаписан.');
  } catch(e) {message('Ошибка: '+e.message);}
};
