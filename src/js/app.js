/* パソコン診断処方箋 Web版 ── 入力フォーム・反映プレビュー・Firebase保存 */
(function(){
'use strict';
const BP = window.BP;
const $ = s=>document.querySelector(s);
const $$ = s=>Array.prototype.slice.call(document.querySelectorAll(s));

/* ===================== モデル ===================== */
let model = {};        // 入力セル ref -> 値
let currentId = null;  // 編集中レコードID
let dirty = false;
let __L = {};          // りすと計算結果(実行時)

function todayStr(){
  const d=new Date();
  return d.getFullYear()+'/'+(d.getMonth()+1)+'/'+d.getDate();
}
// 既定値の上書き（Excel初期値から変更したいもの）
const DEFAULT_OVERRIDES = { 'F34':'なし', 'C35':'M.2(NVMe)' };
// Dドライブ区分の選択肢
const DDRIVE_OPTIONS = ['なし','内蔵SSDorHDD','外付けSSD','外付けHDD','USBメモリー'];
const DDRIVE_DETAIL_CELLS = ['F35','F36','F37','F38','F39','F40','F41','F42','F43'];
// CPUメーカー=apple のときのCPUシリーズ選択肢（Apple Mシリーズ型番）
const APPLE_M_SERIES = ['M1','M1 Pro','M1 Max','M1 Ultra',
  'M2','M2 Pro','M2 Max','M2 Ultra',
  'M3','M3 Pro','M3 Max','M3 Ultra',
  'M4','M4 Pro','M4 Max',
  'M5','M5 Pro','M5 Max'];
// メモリークロック選択肢(MHz)。DDR=4ならDDR4、=5ならDDR5、それ以外は手動入力
const DDR4_CLOCKS = ['4266','3600','3200','3000','2933','2666','2400','2133'];
const DDR5_CLOCKS = ['8000','7200','6400','6000','5600','5333','5200','5120','5000','4800','4400','4000','3600','3200'];
function initModel(){
  model = {};
  for(const ref in BP.INPUT_DEFAULTS){
    const v = BP.INPUT_DEFAULTS[ref];
    if(v && typeof v==='object'){ model[ref]=''; }   // 数式デフォルト(C2等)は後で計算
    else model[ref]=v;
  }
  for(const k in DEFAULT_OVERRIDES) model[k]=DEFAULT_OVERRIDES[k];
  model['C2']=todayStr();
  currentId=null; dirty=false;
}

/* ===================== 数式評価 ===================== */
const _cache={};
function compile(expr){
  if(_cache[expr]) return _cache[expr];
  let fn;
  try{ fn=new Function('I','L','H','N','TODAY','VLOOKUP_MACUPD','return ('+expr+');'); }
  catch(e){ fn=()=>''; }
  return _cache[expr]=fn;
}
const N = x=>{ const n=parseFloat(x); return isNaN(n)?0:n; };
const TODAY = ()=>todayStr();
function VLOOKUP_MACUPD(key){ return ''; } // 未使用(保険)

function evalExpr(expr, Ifn, Lfn, Hfn){
  try{
    let r=compile(expr)(Ifn,Lfn,Hfn,N,TODAY,VLOOKUP_MACUPD);
    return (r===undefined||r===null||(typeof r==='number'&&isNaN(r)))?'':r;
  }catch(e){ return ''; }
}

/* 入力シート内の数式(FXI: C2,D列,G列) を model に反映 */
function computeInput(){
  const I = ref => (model[ref]!==undefined && model[ref]!==null)? model[ref] : '';
  for(const ref in BP.FXI){
    model[ref]= evalExpr(BP.FXI[ref], I, ()=> '', ()=> '');
  }
}

/* りすと計算 (FXL) → __L */
function computeRisuto(){
  const L = Object.assign({}, BP.RISUTO_STATIC);
  const I = ref => (model[ref]!==undefined && model[ref]!==null)? model[ref] : '';
  const Lf = ref => (L[ref]!==undefined && L[ref]!==null)? L[ref] : '';
  // 数式は相互依存があるため複数パス
  for(let pass=0; pass<4; pass++){
    for(const ref in BP.FXL){
      L[ref]= evalExpr(BP.FXL[ref], I, Lf, ()=> '');
    }
  }
  __L = L;
  return L;
}

/* 反映計算 (FX) → H */
function computeHanei(L){
  const H = {};
  const I = ref => (model[ref]!==undefined && model[ref]!==null)? model[ref] : '';
  const Lf = ref => (L[ref]!==undefined && L[ref]!==null)? L[ref] : '';
  const Hf = ref => (H[ref]!==undefined && H[ref]!==null)? H[ref] : '';
  for(let pass=0; pass<3; pass++){
    for(const ref in BP.FX){
      H[ref]= evalExpr(BP.FX[ref], I, Lf, Hf);
    }
  }
  return H;
}

/* ===================== プレビュー描画 ===================== */
let scalerEl, a4El;
function buildPreview(){
  scalerEl = $('#scaler');
  scalerEl.innerHTML = BP.PRINT_HTML;
  a4El = $('#a4');
  const tbl = scalerEl.querySelector('table.hanei');
  const natW = tbl ? tbl.offsetWidth : 1078;
  const natH = tbl ? tbl.offsetHeight : 1758;
  drawColorUnderlays(tbl);   // 色付き領域の継ぎ目(方眼)を消す同色ベタ下地
  // .a4枠を小さめ(707x1000=約187x264mm)にし、Safariのヘッダ/フッタ余白が入っても1ページに収める
  const BOX_W=707, BOX_H=1000;        // .a4ボックス
  const PAGE_W=707, PAGE_H=976;       // コンテンツを収める領域（上下左右に内側余白）
  const scale = Math.min(PAGE_W/natW, PAGE_H/natH);
  const cw=natW*scale, ch=natH*scale;
  const offX = Math.max(0,(BOX_W - cw)/2); // 水平センタリング
  const offY = Math.max(0,(BOX_H - ch)/2); // 垂直センタリング
  scalerEl.style.transform='translate('+offX+'px,'+offY+'px) scale('+scale+')';
  window.__scale = scale; window.__offX=offX; window.__offY=offY;
  placeLogo();
}

// 色付き領域(CPU黄/メモリ水/記憶媒体桃)の背後に同色のベタ下地を敷き、セルの継ぎ目(方眼)を解消
function drawColorUnderlays(tbl){
  if(!tbl) return;
  tbl.style.position='relative'; tbl.style.zIndex='1';
  const cols=tbl.querySelectorAll('colgroup col');
  let xs=[0]; cols.forEach(c=>xs.push(xs[xs.length-1]+(parseFloat(c.style.width)||0)));
  const trs=tbl.querySelectorAll('tr');
  let ys=[0]; trs.forEach(r=>ys.push(ys[ys.length-1]+(parseFloat(r.style.height)||0)));
  // 領域: [col1,col2,row1,row2,color]（rowは絶対値, 表はrow2始まり）
  const regions=[[2,22,37,40,'#FFFFCC'],[2,22,41,44,'#CCFFFF'],[2,22,45,59,'#FFCCCC']];
  regions.forEach(([c1,c2,r1,r2,col])=>{
    const d=document.createElement('div');
    d.style.cssText='position:absolute;z-index:0;left:'+xs[c1-1]+'px;top:'+ys[r1-2]+'px;width:'+(xs[c2]-xs[c1-1])+'px;height:'+(ys[r2-1]-ys[r1-2])+'px;background:'+col;
    scalerEl.insertBefore(d, tbl);
  });
}
function placeLogo(){
  if(a4El.querySelector('.logo')) return;
  const img=document.createElement('img');
  img.className='logo'; img.src='img/logo.png'; img.alt='';
  // 反映の元画像アンカー(col4,row1付近)に小さく配置
  img.style.left='5%'; img.style.top='0.6%'; img.style.width='8%'; img.style.opacity='0.92';
  img.onerror=()=>img.remove();
  a4El.appendChild(img);
}
function fmt(v){
  if(v===undefined||v===null) return '';
  if(typeof v==='number'){ if(isNaN(v))return''; return String(v); }
  return String(v);
}
function bindPreview(){
  computeInput();
  const L = computeRisuto();
  const H = computeHanei(L);
  // 数式セル
  scalerEl.querySelectorAll('[data-f]').forEach(sp=>{
    sp.textContent = fmt(H[sp.getAttribute('data-f')]);
  });
  autofitCells();
  drawBars(L);
}

/* ストレージ使用率バー＋DiskInfo/DiskMark読み出し（図形再現） */
function drawBars(L){
  a4El.querySelectorAll('.usebar,.diskrow,.battbox,.ovline').forEach(e=>e.remove());
  const tbl = scalerEl.querySelector('table.hanei');
  if(!tbl) return;
  const cols = tbl.querySelectorAll('colgroup col');
  let xs=[0]; cols.forEach(c=>xs.push(xs[xs.length-1]+(parseFloat(c.style.width)||0)));
  const trs = tbl.querySelectorAll('tr');
  let ys=[0]; trs.forEach(r=>ys.push(ys[ys.length-1]+(parseFloat(r.style.height)||0)));
  const natW=xs[xs.length-1], natH=ys[ys.length-1];
  const trect=tbl.getBoundingClientRect(), arect=a4El.getBoundingClientRect();
  const ox=trect.left-arect.left, oy=trect.top-arect.top;
  function PX(col){ return ox + (xs[col-1]/natW)*trect.width; }
  function PW(c1,c2){ return ((xs[c2]-xs[c1-1])/natW)*trect.width; }
  function PY(row){ return oy + (ys[row-2]/natH)*trect.height; }
  function PHpx(r1,r2){ return (ys[r2-1]-ys[r1-2])/natH*trect.height; }
  // 使用率バー（薄め）
  function placeBar(col1,col2,row1,row2,used,label){
    const bar=document.createElement('div'); bar.className='usebar';
    bar.style.left=PX(col1)+'px'; bar.style.top=PY(row1)+'px';
    bar.style.width=PW(col1,col2)+'px'; bar.style.height=Math.max(11,PHpx(row1,row2))+'px';
    const u=Math.max(0,Math.min(1,used));
    bar.innerHTML='<div class="used" style="width:'+(u*100).toFixed(1)+'%"></div><div class="free" style="width:'+((1-u)*100).toFixed(1)+'%"></div>'+
      (label?'<span class="lab">'+esc(label)+'</span>':'');
    a4El.appendChild(bar);
  }
  // DiskInfo/DiskMark 4項目の読み出し
  function placeReadouts(col1,col2,row1,row2,items){
    const wrap=document.createElement('div'); wrap.className='diskrow';
    wrap.style.left=PX(col1)+'px'; wrap.style.top=PY(row1)+'px';
    wrap.style.width=PW(col1,col2)+'px'; wrap.style.height=Math.max(20,PHpx(row1,row2))+'px';
    items.forEach(it=>{
      const b=document.createElement('div'); b.className='diskbox';
      b.innerHTML='<span class="dl">'+esc(it.l)+'</span><span class="dv">'+esc(it.v||'—')+'</span>';
      wrap.appendChild(b);
    });
    a4El.appendChild(wrap);
  }
  const cv=(cell,unit)=>{ const v=model[cell]; return (v!==undefined&&v!==null&&String(v)!=='')? (v+unit):''; };
  const cItems=[{l:'電源投入数',v:cv('C38','回')},{l:'使用時間',v:cv('C39','時間')},{l:'読込速度',v:cv('C40','MB/s')},{l:'ランダム',v:cv('C41','MB/s')}];
  const dItems=[{l:'電源投入数',v:cv('F38','回')},{l:'使用時間',v:cv('F39','時間')},{l:'読込速度',v:cv('F40','MB/s')},{l:'ランダム',v:cv('F41','MB/s')}];
  // 使用率 = 1 - 空き率
  const cCap=N(L['Q18']), dCap=N(L['R18']);
  if(cCap>0) placeBar(3,22,47,48, N(L['Q19']), L['Q16']);
  placeReadouts(3,22,49,51,cItems);
  if(dCap>0) placeBar(3,22,55,56, N(L['R19']), L['R16']);
  placeReadouts(3,22,57,59,dItems);
  // バッテリー残量（物理故障診断エリア「無」の右隣。枠を小さめにし「無」と隙間を空けて重ならないように）
  const bv=model['C96'];
  const batt=document.createElement('div'); batt.className='battbox';
  batt.style.left=PX(35)+'px'; batt.style.top=(PY(61)+2)+'px';
  batt.style.width=PW(35,42)+'px'; batt.style.height=Math.max(13,PHpx(61,62)-5)+'px';
  batt.innerHTML='<span class="bl">バッテリー残量</span><span class="bv">'+
    ((bv!==undefined&&bv!==null&&String(bv)!=='')? esc(bv)+'％' : '—')+'</span>';
  a4El.appendChild(batt);
  // 「無」(物理故障 有無 AF61:AG62) の右側に縦線（元データに罫線が無いため補う）
  const vl=document.createElement('div'); vl.className='ovline';
  vl.style.left=PX(34)+'px'; vl.style.top=PY(61)+'px';
  vl.style.width='1px'; vl.style.height=PHpx(61,62)+'px';
  a4El.appendChild(vl);
}

/* 文字がセル幅/高さを超える箇所のフォントを自動縮小（見切れ防止） */
function autofitCells(){
  const tds=scalerEl.querySelectorAll('td');
  tds.forEach(td=>{
    if(!('origfs' in td.dataset)) td.dataset.origfs = td.style.fontSize || '';
    td.style.fontSize = td.dataset.origfs;       // 元サイズに戻してから判定
  });
  tds.forEach(td=>{
    const t=td.textContent; if(!t || !t.trim()) return;
    let guard=0;
    while((td.scrollWidth>td.clientWidth+1 || td.scrollHeight>td.clientHeight+2) && guard<16){
      const fs=parseFloat(getComputedStyle(td).fontSize);
      if(fs<=6) break;
      td.style.fontSize=(fs-0.5)+'px'; guard++;
    }
  });
}

/* ===================== フォーム描画 ===================== */
const TAGCLS={'必須':'req','選択':'sel','入力':'inp','appleのみ':'apl'};
const EDIT_TAGS={'必須':1,'選択':1,'入力':1,'appleのみ':1};
const MAJOR_HEADERS={'基本性能(スペック)診断':1};
function isHeaderRow(r){
  if(!r.edits.length){
    if(r.label && /^処方箋/.test(r.label)) return 'major';
    if(r.tag && !EDIT_TAGS[r.tag]) return MAJOR_HEADERS[r.tag]?'major':'sub';
  }
  return false;
}
// セル毎のサブラベル
const SUBLAB={'D31':'／ 全','C31':'空き','C62':'件数','D62':'アイテム','C82':'件数','D82':'アイテム'};
function isStorageRow(rw){ return rw>=34 && rw<=44; }

let winBlock=null, macBlock=null;
function renderForm(){
  const pane=$('#formPane'); pane.innerHTML='';
  let sec=null, body=null;
  function newSec(title, sub, mini){
    sec=document.createElement('div'); sec.className='sec'+(sub?' sub':'');
    const h=document.createElement('div'); h.className='sec-h';
    h.innerHTML=esc(title)+(mini?' <span class="mini">'+esc(mini)+'</span>':'');
    sec.appendChild(h); body=document.createElement('div'); sec.appendChild(body);
    pane.appendChild(sec); return sec;
  }
  newSec('処方箋（基本情報）',false);

  const rows=BP.FORM_ROWS;
  for(let i=0;i<rows.length;i++){
    const r=rows[i];
    // 相談内容ブロックは専用描画
    if(r.row===104){ renderSoudan(pane); // 105-144 をまとめて描画
      // skip until 145
      while(i<rows.length && rows[i].row<145) i++;
      i--; continue;
    }
    // ストレージは見出し行(34)にも入力があるため、ここで明示的に見出しを差し込む
    if(r.row===34){ newSec('ストレージ（Cドライブ ／ Dドライブ）', false, '各行 C:＝Cドライブ、D:＝Dドライブ'); }
    const hd=isHeaderRow(r);
    if(hd){
      let title = (r.label&&/^処方箋/.test(r.label))? r.label : (r.tag||r.label||'');
      let mini='';
      newSec(title, hd==='sub', mini);
      // Windows/Mac ブロックの開始を記録
      if(r.tag==='Windowsシステム診断'){ winBlock=sec; sec.dataset.block='win'; }
      if(r.tag==='Macシステム診断'){ macBlock=sec; sec.dataset.block='mac'; }
      continue;
    }
    if(!r.edits.length) continue;
    body.appendChild(makeRow(r));
  }
  applyWinMac();
  applyDDrive(false);
}

// メモリー欄: DDR(行29)をクロック(行28)より前に表示するよう並べ替え（1回だけ）
function setupFormRows(){
  const rows=BP.FORM_ROWS;
  const i28=rows.findIndex(r=>r.row===28), i29=rows.findIndex(r=>r.row===29);
  if(i28>=0 && i29>=0 && i28<i29){ const t=rows[i28]; rows[i28]=rows[i29]; rows[i29]=t; }
}
// メモリークロック(C28)のコントロールをDDR選択に応じて作り直す（選択⇔手動入力の切替）
function rebuildClockControl(){
  const old=document.querySelector('[data-cell="C28"]');
  if(old){ const fresh=makeControl({cell:'C28'}); old.replaceWith(fresh); }
}

// Dドライブ区分が「なし」のときD:詳細欄を隠す（doClear=trueなら値もクリア）
function applyDDrive(doClear){
  const none = (model['F34']==='なし' || model['F34']==null || model['F34']==='');
  if(none && doClear){
    DDRIVE_DETAIL_CELLS.forEach(c=>{ model[c]=''; const el=document.querySelector('[data-cell="'+c+'"]'); if(el) el.value=''; });
  }
  document.querySelectorAll('.ddrive-field').forEach(el=>{ el.style.display = none ? 'none' : ''; });
}

function makeRow(r){
  const row=document.createElement('div'); row.className='frow'; row.dataset.row=r.row;
  const head=document.createElement('div'); head.className='rowhead';
  // 行34はストレージ見出し行: タグ('ストレージ C:')を出さず、Dドライブ区分のラベルにする
  let rtag=r.tag, rlab=r.label;
  if(r.row===34){ rtag=''; rlab='Dドライブ区分'; }
  const tag=document.createElement('span');
  tag.className='tag '+(TAGCLS[rtag]||'non'); tag.textContent=rtag||'';
  if(!rtag) tag.style.visibility='hidden';
  const lab=document.createElement('span'); lab.className='lab'; lab.textContent=cleanLabel(rlab)||'　';
  head.appendChild(tag); head.appendChild(lab);
  const ctrls=document.createElement('div'); ctrls.className='ctrls';
  r.edits.forEach(e=>{
    if(e.cell==='B34') return; // 見出し行に紛れた不要なドロップダウンは除外
    const drive = isStorageRow(r.row) ? (e.cell[0]==='F'?'D':'C') : null;
    const sub = SUBLAB[e.cell] || '';
    // Dドライブ詳細(行35-43のF列)は、区分「なし」で一括非表示にできるようまとめる
    if(drive==='D' && r.row!==34){
      const w=document.createElement('span'); w.className='ddrive-field';
      const c=document.createElement('span'); c.className='drivechip d'; c.textContent='D:'; w.appendChild(c);
      w.appendChild(makeControl(e));
      ctrls.appendChild(w);
      return;
    }
    if(drive){ const c=document.createElement('span'); c.className='drivechip '+(drive==='D'?'d':'c'); c.textContent=drive+':'; ctrls.appendChild(c); }
    else if(sub){ const s=document.createElement('span'); s.className='sublabel'; s.textContent=sub; ctrls.appendChild(s);}
    ctrls.appendChild(makeControl(e));
    if(drive==='C'){ const br=document.createElement('span'); br.className='cbreak'; br.style.flexBasis='100%'; br.style.height='0'; ctrls.appendChild(br); } // C:とD:を改行で分ける
  });
  row.appendChild(head); row.appendChild(ctrls);
  return row;
}

function makeControl(e){
  const cell=e.cell;
  const opts = optionsFor(cell);
  if(opts){
    const sel=document.createElement('select'); sel.dataset.cell=cell;
    fillOptions(sel, opts, model[cell]);
    sel.addEventListener('change', ()=>{ setVal(cell, sel.value); });
    return sel;
  }
  const inp=document.createElement('input'); inp.className='f'; inp.dataset.cell=cell;
  inp.value = (model[cell]==null?'':model[cell]);
  if(cell[0]==='C'&&(''+cell).length<=4) {}
  inp.addEventListener('input', ()=>{ setVal(cell, inp.value); });
  return inp;
}
function fillOptions(sel, opts, val){
  sel.innerHTML='';
  const blank=document.createElement('option'); blank.value=''; blank.textContent='―'; sel.appendChild(blank);
  let found=false;
  opts.forEach(o=>{ if(o===''||o==null) return; const op=document.createElement('option'); op.value=o; op.textContent=o; if(String(o)===String(val))found=true; sel.appendChild(op);});
  if(val && !found){ const op=document.createElement('option'); op.value=val; op.textContent=val; sel.appendChild(op);}
  sel.value = (val==null?'':val);
}
function optionsFor(cell){
  if(cell==='F34') return DDRIVE_OPTIONS.slice();   // Dドライブ区分の選択肢を差し替え
  if(cell==='C28'){                                  // メモリークロック: DDRに応じて選択 or 手動入力
    const ddr=String(model['C29']||'');
    if(ddr.indexOf('5')>=0) return DDR5_CLOCKS.slice();
    if(ddr.indexOf('4')>=0) return DDR4_CLOCKS.slice();
    return null;                                     // それ以外は手動入力(text)
  }
  const dv = BP.DV[cell];
  if(!dv) return null;
  if(dv.src==='INDIRECT($C$19)'){
    if(model['C19']==='apple') return APPLE_M_SERIES.slice();   // apple → Mシリーズ型番
    return (BP.INDIRECT[model['C19']]||[]).slice();
  }
  if(!dv.static && !dv.dyn) return [];
  const out=(dv.static||[]).slice();
  if(dv.dyn){ dv.dyn.forEach(ref=>{ const v=__L[ref]; if(v && out.indexOf(v)<0) out.push(v); }); }
  return out;
}

/* 相談内容ブロック(105-144): 10行 × [相談入力→お困り選択 / 詳細入力→詳細選択] */
function renderSoudan(pane){
  const sec=document.createElement('div'); sec.className='sec';
  sec.innerHTML='<div class="sec-h">処方箋（お困りごと・相談内容）<span class="mini">入力欄に直接記入 → 右の選択で反映（プリセットも選択可）</span></div>';
  const body=document.createElement('div'); sec.appendChild(body); pane.appendChild(sec);
  for(let k=0;k<10;k++){
    const row=document.createElement('div'); row.className='frow'; row.style.gridTemplateColumns='30px 1fr';
    const num=document.createElement('div'); num.className='lab'; num.textContent=(k+1)+'.';
    const ctrls=document.createElement('div'); ctrls.className='ctrls'; ctrls.style.flexDirection='column'; ctrls.style.alignItems='stretch';
    // お困りごと: 入力 C105+k → 選択 C115+k
    ctrls.appendChild(relLine('お困りごと', 'C'+(105+k), 'C'+(115+k)));
    ctrls.appendChild(relLine('詳細', 'C'+(125+k), 'C'+(135+k)));
    row.appendChild(num); row.appendChild(ctrls); body.appendChild(row);
  }
}
function relLine(name, inputCell, selectCell){
  const wrap=document.createElement('div'); wrap.style.display='flex'; wrap.style.gap='6px'; wrap.style.alignItems='center'; wrap.style.margin='2px 0';
  const tag=document.createElement('span'); tag.className='tag inp'; tag.textContent=name; tag.style.minWidth='52px';
  const inp=document.createElement('input'); inp.className='f wide'; inp.placeholder='自由入力'; inp.dataset.cell=inputCell;
  inp.value=model[inputCell]==null?'':model[inputCell];
  const sel=document.createElement('select'); sel.dataset.cell=selectCell; sel.style.maxWidth='260px';
  const refresh=()=>fillOptions(sel, optionsFor(selectCell)||[], model[selectCell]);
  refresh();
  inp.addEventListener('input', ()=>{ setVal(inputCell, inp.value, true); refresh(); });
  sel.addEventListener('change', ()=>{ setVal(selectCell, sel.value); });
  wrap.appendChild(tag); wrap.appendChild(inp);
  const arrow=document.createElement('span'); arrow.className='sublabel'; arrow.textContent='→反映:'; wrap.appendChild(arrow);
  wrap.appendChild(sel);
  return wrap;
}

/* 値変更 */
let recomputeTimer=null;
function setVal(cell, val, light){
  model[cell]=val; dirty=true;
  // INDIRECT連動: CPUメーカー(C19)変更でシリーズ(C20)選択肢更新
  if(cell==='C19'){ const s=$('select[data-cell="C20"]'); if(s){ model['C20']=''; fillOptions(s, optionsFor('C20')||[], ''); } }
  if(cell==='C46'){ applyWinMac(); }
  if(cell==='F34'){ applyDDrive(true); }
  if(cell==='C29'){ rebuildClockControl(); }   // DDR変更でクロック欄を選択/手動に切替
  bindPreview();
  if(!light) refreshDynamicSelects();
  markSync('●未保存', '#caa23a');
}
// 動的選択肢(タイプ入力反映等)を再構築
function refreshDynamicSelects(){
  $$('select[data-cell]').forEach(sel=>{
    const cell=sel.dataset.cell; const dv=BP.DV[cell];
    if(!dv) return;
    if(dv.src==='INDIRECT($C$19)' || (dv.dyn&&dv.dyn.length)){
      fillOptions(sel, optionsFor(cell)||[], model[cell]);
    }
  });
}
function applyWinMac(){
  const isWin = (model['C46']==='Windows');
  $$('.sec[data-block="win"]').forEach(s=>s.style.display=isWin?'':'none');
  $$('.sec[data-block="mac"]').forEach(s=>s.style.display=isWin?'none':'');
}

/* ===================== ユーティリティ ===================== */
function esc(s){ return (s==null?'':String(s)).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function cleanLabel(s){ if(!s)return ''; return String(s).replace(/^→/,'').replace(/[★]/g,''); }
function toast(msg){ const t=$('#toast'); t.textContent=msg; t.classList.add('show'); clearTimeout(t._t); t._t=setTimeout(()=>t.classList.remove('show'),1800); }
function markSync(txt,col){ const s=$('#sync'); s.textContent=txt; if(col)s.style.color=col; }

/* ===================== Firebase ===================== */
let dbRef=null, allRecords={};
function initFirebase(){
  if(typeof firebase==='undefined'||!FIREBASE_CONFIG.apiKey){ markSync('⚪ ローカルのみ','#ccc'); return; }
  try{
    firebase.initializeApp(FIREBASE_CONFIG);
    const db=firebase.database();
    dbRef=db.ref(DB_NODE);
    db.ref('.info/connected').on('value', s=>{
      if(s.val()===true) markSync(dirty?'●未保存':'🟢 接続済','#7bdca0');
      else markSync('🔴 オフライン','#e88');
    });
    dbRef.on('value', snap=>{ allRecords = snap.val()||{}; renderList(); });
  }catch(e){ console.warn('Firebase init error', e); markSync('⚪ ローカルのみ','#ccc'); }
}
function recordTitle(m){
  return (m['C4']||'（無題）');
}
function save(){
  computeInput();
  const m=Object.assign({}, model);
  const id = currentId || ('rx_'+Date.now()+'_'+Math.floor(Math.random()*1e4));
  const rec={
    id, model:m,
    name:m['C4']||'', member:m['C3']||'', tester:m['C7']||'',
    date:m['C2']||'', updatedAt:new Date().toISOString()
  };
  currentId=id; dirty=false;
  if(dbRef){ dbRef.child(id).set(rec).then(()=>{ toast('💾 保存しました'); markSync('🟢 保存済','#7bdca0'); })
    .catch(e=>{ toast('保存失敗: '+e.message); }); }
  else { allRecords[id]=rec; localStorage.setItem('pc_rx_local', JSON.stringify(allRecords)); toast('💾 ローカル保存'); }
}
function loadRecord(id){
  const rec = allRecords[id]; if(!rec)return;
  model = Object.assign({}, BP.INPUT_DEFAULTS_FLAT(), rec.model);
  // ensure all keys exist
  for(const ref in BP.INPUT_DEFAULTS){ if(model[ref]===undefined){ const v=BP.INPUT_DEFAULTS[ref]; model[ref]=(v&&typeof v==='object')?'':v; } }
  currentId=id; dirty=false;
  renderForm(); bindPreview(); closeList(); toast('読み込みました');
}
function duplicateRecord(id){ const rec=allRecords[id]; if(!rec)return; loadRecord(id); currentId=null; toast('複製しました（新規として保存できます）'); }
function deleteRecord(id){ if(!confirm('この処方箋を削除しますか？'))return; if(dbRef)dbRef.child(id).remove(); else {delete allRecords[id]; localStorage.setItem('pc_rx_local',JSON.stringify(allRecords)); renderList();} toast('削除しました'); }

/* 一覧 */
function renderList(){
  const box=$('#reclist'); if(!box)return;
  const q=($('#search').value||'').trim().toLowerCase();
  const arr=Object.values(allRecords).sort((a,b)=>(b.updatedAt||'').localeCompare(a.updatedAt||''));
  const f=arr.filter(r=>{ if(!q)return true; return [r.name,r.member,r.tester,r.date].join(' ').toLowerCase().includes(q); });
  if(!f.length){ box.innerHTML='<div class="empty">保存済みの処方箋はありません</div>'; return; }
  box.innerHTML='';
  f.forEach(r=>{
    const el=document.createElement('div'); el.className='rec';
    el.innerHTML='<div class="meta"><b>'+esc(r.name||'（無題）')+'</b>'+
      '<small>会員: '+esc(r.member||'-')+' ／ 点検者: '+esc(r.tester||'-')+' ／ '+esc((r.updatedAt||'').slice(0,10))+'</small></div>'+
      '<div class="acts"><button class="open">開く</button><button class="dup">複製</button><button class="del">削除</button></div>';
    el.querySelector('.open').onclick=()=>loadRecord(r.id);
    el.querySelector('.dup').onclick=()=>duplicateRecord(r.id);
    el.querySelector('.del').onclick=()=>deleteRecord(r.id);
    box.appendChild(el);
  });
}
function openList(){ $('#listModal').classList.add('open'); renderList(); }
function closeList(){ $('#listModal').classList.remove('open'); }

/* ===================== 起動 ===================== */
BP.INPUT_DEFAULTS_FLAT=function(){ const o={}; for(const ref in BP.INPUT_DEFAULTS){ const v=BP.INPUT_DEFAULTS[ref]; o[ref]=(v&&typeof v==='object')?'':v; } for(const k in DEFAULT_OVERRIDES) o[k]=DEFAULT_OVERRIDES[k]; o['C2']=todayStr(); return o; };

function newDoc(){ if(dirty && !confirm('保存していない変更があります。新規作成しますか？'))return; initModel(); renderForm(); bindPreview(); toast('新規作成'); }

function boot(){
  setupFormRows();
  initModel();
  buildPreview();
  renderForm();
  bindPreview();
  initFirebase();
  $('#btnNew').onclick=newDoc;
  $('#btnList').onclick=openList;
  $('#btnSave').onclick=save;
  $('#btnPrint').onclick=()=>window.print();
  $('#listClose').onclick=closeList;
  $('#listModal').addEventListener('click',e=>{ if(e.target.id==='listModal')closeList(); });
  $('#search').addEventListener('input', renderList);
  // local fallback load
  if(typeof firebase==='undefined'){ try{ allRecords=JSON.parse(localStorage.getItem('pc_rx_local')||'{}'); }catch(e){} }
  window.addEventListener('beforeunload', e=>{ if(dirty){ e.preventDefault(); e.returnValue=''; } });
}
document.addEventListener('DOMContentLoaded', boot);
})();
