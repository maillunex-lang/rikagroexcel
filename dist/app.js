const $ = s => document.querySelector(s);
let base = null, sources = [], manualChoices=new Map(), reviewItems=[], reviewCursor=0;
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
  manualChoices.clear();reviewItems=[];$('#review').hidden=true;
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
function tokens(v){return [...new Set(norm(v).split(/[^\p{L}\p{N}]+/u).filter(x=>x.length>1))];}
function similarity(a,b){
  const x=norm(a),y=norm(b);
  if(!x||!y)return 0;
  if(x===y)return 1;
  const tx=tokens(x),ty=tokens(y);
  const common=tx.filter(t=>ty.includes(t)).length;
  const wordScore=common/Math.max(tx.length,ty.length,1);
  const grams=v=>{const z=' '+v+' ',g=new Map();for(let i=0;i<z.length-1;i++){const t=z.slice(i,i+2);g.set(t,(g.get(t)||0)+1);}return g;};
  const gx=grams(x),gy=grams(y);
  let shared=0,total=0;
  for(const n of gx.values())total+=n;
  for(const n of gy.values())total+=n;
  for(const [k,n] of gx)shared+=Math.min(n,gy.get(k)||0);
  return 0.55*(2*shared/Math.max(total,1))+0.45*wordScore;
}
function reviewKey(si,row){return si+':'+row;}
function collectReviews(){
  const mode=$('#mode').value,items=[];
  if(!base)return items;
  const end=XLSX.utils.decode_range(base.wb.Sheets[base.sheet]['!ref']||'A1').e.r;
  sources.forEach((s,si)=>{
    if(!s.selected.length||base.name<0||s.name<0)return;
    const index=buildIndex(s,mode);
    const candidates=[];
    for(let j=s.header+1;j<s.rows.length;j++){
      const name=s.rows[j]?.[s.name];
      if(norm(name))candidates.push({row:j,name:String(name)});
    }
    for(let row=base.header+1;row<=end;row++){
      const data=base.rows[row]||[];
      if(findSourceRow(data,base,s,index,mode)>=0)continue;
      const name=data[base.name];
      if(!norm(name))continue;
      const ranked=candidates.map(c=>({...c,score:similarity(name,c.name)}))
        .filter(c=>c.score>=0.38).sort((a,b)=>b.score-a.score).slice(0,4);
      if(ranked.length)items.push({si,row,name:String(name),code:base.code>=0?String(data[base.code]??''):'',ranked});
    }
  });
  return items;
}
function renderReview(){
  const item=reviewItems[reviewCursor];
  if(!item){$('#review').hidden=true;return;}
  const s=sources[item.si],key=reviewKey(item.si,item.row),selected=manualChoices.get(key);
  $('#reviewProgress').textContent=(reviewCursor+1)+' из '+reviewItems.length;
  $('#reviewSource').textContent=s.filename+' · строка '+(item.row+1);
  $('#reviewProduct').textContent=item.name+(item.code?' · '+item.code:'');
  $('#reviewOptions').innerHTML='<label class="suggestion"><input type="radio" name="suggest" value="-1" '+(selected===undefined||selected===-1?'checked':'')+'><span>Оставить пустым</span></label>'+
    item.ranked.map(c=>'<label class="suggestion"><input type="radio" name="suggest" value="'+c.row+'" '+(selected===c.row?'checked':'')+'><span><b>'+esc(c.name)+'</b><small>Строка '+(c.row+1)+' · схожесть '+Math.round(c.score*100)+'%</small></span></label>').join('');
  $('#reviewPrev').disabled=reviewCursor===0;
  $('#reviewNext').textContent=reviewCursor===reviewItems.length-1?'Готово':'Следующий →';
}
function openReview(){
  if(!base||!sources.some(s=>s.selected.length))throw Error('Выберите файлы и столбцы.');
  reviewItems=collectReviews();reviewCursor=0;
  if(!reviewItems.length){message('Похожих названий для проверки не найдено. Несовпавшие ячейки останутся пустыми.');return;}
  $('#review').hidden=false;renderReview();
}
function saveReviewChoice(){
  const item=reviewItems[reviewCursor];
  if(!item)return;
  const checked=document.querySelector('input[name="suggest"]:checked');
  manualChoices.set(reviewKey(item.si,item.row),checked?Number(checked.value):-1);
}
$('#reviewPrev').onclick=()=>{saveReviewChoice();reviewCursor=Math.max(0,reviewCursor-1);renderReview();};
$('#reviewNext').onclick=()=>{saveReviewChoice();if(reviewCursor+1===reviewItems.length){$('#review').hidden=true;message('Ручная проверка завершена. Можно скачать Excel.');}else{reviewCursor++;renderReview();}};
$('#reviewClose').onclick=()=>{$('#review').hidden=true;};
$('#reviewButton').onclick=()=>{try{openReview();}catch(e){message(e.message);}};
$('#manual').onchange=()=>{$('#reviewButton').hidden=!$('#manual').checked;manualChoices.clear();reviewItems=[];$('#review').hidden=true;};
function appendColumns(wb,b,sources,mode){
  const ws=wb.Sheets[b.sheet],range=XLSX.utils.decode_range(ws['!ref']||'A1');
  const originalEnd=range.e.c;
  let nextCol=originalEnd+1,filled=0,empty=0;
  for(const s of sources){
    if(!s.selected.length)continue;
    if(mode==='code'&&(b.code<0||s.code<0))throw Error('Выберите колонки артикулов: '+s.filename);
    if(mode==='name'&&(b.name<0||s.name<0))throw Error('Выберите колонки названий: '+s.filename);
    const index=buildIndex(s,mode),sourceWs=s.wb.Sheets[s.sheet],si=sources.indexOf(s);
    for(const sourceCol of s.selected){
      const targetCol=nextCol++;
      const label=cleanName(s.filename)+(s.selected.length>1?' ('+colName(sourceCol)+')':'');
      ws[XLSX.utils.encode_cell({r:b.header,c:targetCol})]={t:'s',v:label};
      for(let row=b.header+1;row<=range.e.r;row++){
        const key=reviewKey(si,row);
        const sourceRow=manualChoices.has(key)?manualChoices.get(key):findSourceRow(b.rows[row]||[],b,s,index,mode);
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
    if($('#manual').checked&&!reviewItems.length){openReview();if(reviewItems.length)return;}
    if($('#manual').checked&&reviewItems.some(x=>!manualChoices.has(reviewKey(x.si,x.row)))){message('Завершите ручную проверку или отключите её галочкой.');$('#review').hidden=false;renderReview();return;}
    const wb=XLSX.read(XLSX.write(base.wb,{type:'array',bookType:'xlsx'}),{type:'array',cellStyles:true,cellNF:true,cellDates:true});
    const n=appendColumns(wb,base,sources,$('#mode').value);
    XLSX.writeFile(wb,cleanName(base.filename)+'_с_добавленными_столбцами.xlsx');
    message('Готово! Добавлено столбцов: '+n.columns+'. Заполнено: '+n.filled+'. Пустых: '+n.empty+'. Исходный файл не перезаписан.');
  } catch(e) {message('Ошибка: '+e.message);}
};
