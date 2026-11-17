const $ = s => document.querySelector(s);
let base = null, sources = [];
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const colName = i => XLSX.utils.encode_col(i);
const cleanName = name => name.replace(/\.[^.]+$/, '');
const message = s => { $('#message').textContent = s; };
function setSheet(item, reset = false) {
  const ws = item.wb.Sheets[item.sheet];
  item.rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:'', blankrows:true});
  item.header = reset ? 0 : Math.max(0, Math.min(item.header || 0, item.rows.length - 1));
  if (reset) item.selected = [];
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
    '<label>Строка заголовков<input data-header type="number" min="1" max="'+Math.max(1,s.rows.length)+'" value="'+(s.header+1)+'"></label>'+
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
    el.querySelector('[data-header]').onchange=e=>{s.header=Math.max(0,Math.min(s.rows.length-1,Number(e.target.value)-1));render();};
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
function appendColumns(wb, baseItem, supplierItems) {
  const ws=wb.Sheets[baseItem.sheet];
  const range=XLSX.utils.decode_range(ws['!ref']||'A1');
  const originalEnd=range.e.c;
  let nextCol=originalEnd+1;
  let maxRow=range.e.r;
  for (const s of supplierItems) {
    const sourceWs=s.wb.Sheets[s.sheet];
    for (const sourceCol of s.selected) {
      const targetCol=nextCol++;
      const name=cleanName(s.filename)+(s.selected.length>1?' ('+colName(sourceCol)+')':'');
      const headerAddr=XLSX.utils.encode_cell({r:baseItem.header,c:targetCol});
      ws[headerAddr]={t:'s',v:name};
      maxRow=Math.max(maxRow,baseItem.header);
      const sourceRange=XLSX.utils.decode_range(sourceWs['!ref']||'A1');
      for (let row=s.header+1;row<=sourceRange.e.r;row++) {
        const sourceCell=sourceWs[XLSX.utils.encode_cell({r:row,c:sourceCol})];
        if (!sourceCell || sourceCell.t==='z') continue;
        const targetRow=baseItem.header+1+(row-s.header-1);
        const dest=XLSX.utils.encode_cell({r:targetRow,c:targetCol});
        const copy={...sourceCell};
        delete copy.f; delete copy.F;
        if (sourceCell.f) {
          copy.t='s';copy.v=sourceCell.w ?? String(sourceCell.v ?? '');
        }
        ws[dest]=copy;
        maxRow=Math.max(maxRow,targetRow);
      }
    }
  }
  range.e.c=Math.max(originalEnd,nextCol-1);
  range.e.r=Math.max(range.e.r,maxRow);
  ws['!ref']=XLSX.utils.encode_range(range);
  return nextCol-originalEnd-1;
}
$('#run').onclick=()=>{
  try {
    if(!base)throw Error('Загрузите основной файл.');
    if(!sources.some(s=>s.selected.length))throw Error('Выберите хотя бы один столбец.');
    const wb=XLSX.read(XLSX.write(base.wb,{type:'array',bookType:'xlsx'}),{type:'array',cellStyles:true,cellNF:true,cellDates:true});
    const n=appendColumns(wb,base,sources);
    XLSX.writeFile(wb,cleanName(base.filename)+'_с_добавленными_столбцами.xlsx');
    message('Готово! Добавлено столбцов: '+n+'. Исходный файл не перезаписан.');
  } catch(e) {message('Ошибка: '+e.message);}
};
