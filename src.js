import pptxgen from 'pptxgenjs';
const $ = id => document.getElementById(id);
const board = $('board'), status = $('status');
const key = 'canvas-biblico-web-v1';
let items = [], selected = new Set(), original = false, scale = 1;
const id = () => globalThis.crypto?.randomUUID?.() || `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}-${Math.random().toString(36).slice(2)}`;
function report(message){ status.textContent = message; }
function persist(){try{localStorage.setItem(key,JSON.stringify({version:1,items}));}catch{report('Não foi possível salvar automaticamente. Salve o projeto em arquivo.');}}
function render(){
 board.replaceChildren();
 for(const item of items){
  if(original && item.kind !== 'word') continue;
  const el=document.createElement('div'); el.className=`item ${item.kind==='note'?'note':''} ${selected.has(item.id)&&!original?'selected':''}`;
  el.textContent=item.text; el.style.left=(original?item.ox:item.x)+'px'; el.style.top=(original?item.oy:item.y)+'px';
  el.style.color=original?'#202b38':item.color; el.style.fontWeight=!original&&item.bold?'bold':'normal';
  if(!original && item.highlight && item.kind==='word') el.style.backgroundColor=item.highlight;
  el.addEventListener('pointerdown',e=>startDrag(e,item)); board.append(el);
 }
}
function changed(message='Alterações salvas neste navegador'){persist();render();report(message);}
function fit(){const available=Math.max(250,$('viewport').clientWidth-48);scale=Math.min(1,available/1200);board.style.transform=`scale(${scale})`;board.style.marginBottom=`${-675*(1-scale)}px`;}
function startDrag(e,item){
 if(original)return; e.preventDefault(); e.stopPropagation();
 if(e.shiftKey){if(selected.has(item.id))selected.delete(item.id);else selected.add(item.id);}else if(!selected.has(item.id)){selected=new Set([item.id]);}
 render(); const sx=e.clientX, sy=e.clientY, positions=items.filter(i=>selected.has(i.id)).map(i=>({item:i,x:i.x,y:i.y}));
 const move=ev=>{for(const p of positions){p.item.x=Math.max(0,Math.min(1180,p.x+(ev.clientX-sx)/scale));p.item.y=Math.max(0,Math.min(640,p.y+(ev.clientY-sy)/scale));}render();};
 const up=()=>{window.removeEventListener('pointermove',move);window.removeEventListener('pointerup',up);persist();}; window.addEventListener('pointermove',move);window.addEventListener('pointerup',up,{once:true});
}
function mutate(fn){if(original){report('Volte ao quadro editável para fazer alterações.');return;}if(!selected.size){report('Selecione pelo menos um elemento.');return;}items.filter(i=>selected.has(i.id)).forEach(fn);changed();}
function download(name,content,type){const url=URL.createObjectURL(new Blob([content],{type}));const a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}
$('add').onclick=()=>{if(original)return report('Volte ao quadro editável para adicionar texto.');const text=$('input').value.trim();if(!text)return report('Cole um texto antes de adicioná-lo.');const ctx=document.createElement('canvas').getContext('2d');ctx.font='28px Arial';let x=45,y=items.length?Math.min(600,Math.max(...items.map(i=>i.y))+55):45;for(const line of text.split(/\r?\n/)){for(const word of line.trim().split(/\s+/).filter(Boolean)){const width=ctx.measureText(word).width;if(x+width>1155){x=45;y+=42;}items.push({id:id(),kind:'word',text:word,x,y,ox:x,oy:y,color:'#202b38',highlight:null,bold:false});x+=width+10;}x=45;y+=42;} $('input').value='';changed('Texto adicionado.');};
$('note').onclick=()=>{if(original)return report('Volte ao quadro editável para adicionar notas.');const text=$('note-text').value.trim();if(!text)return report('Escreva uma anotação.');items.push({id:id(),kind:'note',text,x:850,y:50,ox:850,oy:50,color:'#202b38'});$('note-text').value='';changed('Anotação adicionada.');};
$('apply-color').onclick=()=>mutate(i=>i.color=$('color').value);$('apply-highlight').onclick=()=>mutate(i=>{if(i.kind==='word')i.highlight=$('highlight').value;});$('unhighlight').onclick=()=>mutate(i=>i.highlight=null);$('bold').onclick=()=>mutate(i=>i.bold=!i.bold);
$('delete').onclick=()=>{if(original)return;if(!selected.size)return report('Selecione pelo menos um elemento.');items=items.filter(i=>!selected.has(i.id));selected.clear();changed('Seleção excluída.');};
$('original').onclick=()=>{original=!original;selected.clear();$('original').setAttribute('aria-pressed',String(original));$('original').textContent=original?'Voltar ao quadro':'Texto original';render();};
$('fit').onclick=fit;window.addEventListener('resize',fit);board.addEventListener('pointerdown',()=>{selected.clear();render();});
$('new').onclick=()=>{if(items.length&&!confirm('Criar um novo projeto? Salve o projeto atual antes de continuar.'))return;items=[];selected.clear();original=false;$('original').setAttribute('aria-pressed','false');$('original').textContent='Texto original';changed('Novo projeto criado.');};
$('save').onclick=()=>download('canvas-biblico.canvasbiblico',JSON.stringify({format:'canvas-biblico',version:1,items},null,2),'application/json');
$('open').onclick=()=>$('file').click();$('file').onchange=async e=>{const file=e.target.files[0];if(!file)return;try{const data=JSON.parse(await file.text());if(data.format!=='canvas-biblico'||data.version!==1||!Array.isArray(data.items)||!data.items.every(i=>i&&typeof i.id==='string'&&['word','note'].includes(i.kind)&&typeof i.text==='string'&&['x','y','ox','oy'].every(k=>Number.isFinite(i[k]))))throw Error();if(items.length&&!confirm('Abrir este projeto substituirá o quadro atual. Continuar?'))return;items=data.items;selected.clear();original=false;$('original').setAttribute('aria-pressed','false');$('original').textContent='Texto original';changed('Projeto aberto.');}catch{report('Arquivo inválido ou incompatível.');}finally{e.target.value='';}};
$('ppt').onclick=async()=>{if(!items.length)return report('Adicione conteúdo antes de exportar.');try{report('Gerando PowerPoint...');const ppt=new pptxgen();ppt.layout='LAYOUT_WIDE';ppt.author='Canvas Bíblico';ppt.subject='Estudo bíblico';for(const isOriginal of [true,false]){const slide=ppt.addSlide();slide.background={color:'FFFFFF'};for(const item of items){if(isOriginal&&item.kind!=='word')continue;const x=(isOriginal?item.ox:item.x)/1200*13.333,y=(isOriginal?item.oy:item.y)/675*7.5;const ctx=document.createElement('canvas').getContext('2d');ctx.font=`${!isOriginal&&item.bold?'bold ':''}${item.kind==='note'?20:28}px Arial`;const width=item.kind==='note'?240:Math.max(20,ctx.measureText(item.text).width+8);const h=item.kind==='note'?Math.min(300,Math.max(60,Math.ceil(item.text.length/22)*27)):40;let opts={x,y,w:width/1200*13.333,h:h/675*7.5,fontFace:'Arial',fontSize:item.kind==='note'?15:21,color:(isOriginal?'202B38':(item.color||'#202b38').replace('#','')),bold:!isOriginal&&!!item.bold,margin:0,breakLine:false,vertAlign:'mid'};if(!isOriginal&&(item.highlight||item.kind==='note'))opts.fill={color:(item.kind==='note'?'FFF4C4':item.highlight.replace('#',''))};slide.addText(item.text,opts);}}await ppt.writeFile({fileName:'canvas-biblico.pptx'});report('PowerPoint exportado: original e quadro organizado.');}catch(error){console.error(error);report('Não foi possível exportar o PowerPoint.');}};
try{const saved=JSON.parse(localStorage.getItem(key)||'null');if(saved?.version===1&&Array.isArray(saved.items))items=saved.items;}catch{}render();fit();
