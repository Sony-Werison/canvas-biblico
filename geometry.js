export function guidesFor(items, step) {
 const blocks=new Map(),guides=[];
 for(const w of items) if(w.kind==='word'&&w.blockId){if(!blocks.has(w.blockId))blocks.set(w.blockId,[]);blocks.get(w.blockId).push(w);}
 for(const words of blocks.values()){
  words.sort((a,b)=>a.y-b.y||a.x-b.x);const origin=words[0].originX,rows=[];
  for(const word of words){let row=rows[rows.length-1];if(!row||Math.abs(row.y-word.y)>1){row={x:word.x,y:word.y,height:35};rows.push(row);}row.x=Math.min(row.x,word.x);}
  for(const row of rows)row.level=Math.max(0,Math.min(32,Math.round((row.x-origin)/step)));
  const max=Math.max(...rows.map(r=>r.level));
  for(let level=0;level<max;level++)for(let i=0;i<rows.length;i++){
   if(rows[i].level<=level)continue;const first=i;while(i+1<rows.length&&rows[i+1].level>level)i++;
   const parent=first>0?rows[first-1]:null,top=parent?parent.y+parent.height+3:rows[first].y;
   const siblings=rows.slice(first,i+1).filter(r=>r.level===level+1);if(!siblings.length)continue;
   const bottom=siblings[siblings.length-1].y+17.5;if(bottom>top)guides.push({x:origin+level*step,top,bottom,ticks:siblings.map(r=>({y:r.y+17.5,endX:r.x-8}))});
  }
 }return guides;
}
export function connectionLine(connection,items,widthOf){
 const a=items.find(i=>i.id===connection.from),b=items.find(i=>i.id===connection.to);if(!a||!b)return null;
 const ax=a.x+widthOf(a)/2,ay=a.y+17.5,bx=b.x+widthOf(b)/2,by=b.y+17.5;
 const dx=bx-ax,dy=by-ay,d=Math.hypot(dx,dy)||1,ex=Math.min(widthOf(b)/2,Math.abs(dx)>0?widthOf(b)/2*d/Math.abs(dx):Infinity),ey=Math.abs(dy)>0?17.5*d/Math.abs(dy):Infinity,offset=Math.min(ex,ey);
 return {x1:ax,y1:ay,x2:bx-dx/d*offset,y2:by-dy/d*offset};
}

// Os deslocamentos de apresentação não alteram as posições armazenadas.
export function layoutReverses(items,widthOf){
 const result=items.map(i=>({...i}));
 for(const blockId of new Set(items.filter(i=>i.kind==='word'&&i.blockId).map(i=>i.blockId))){
  const words=items.filter(i=>i.kind==='word'&&i.blockId===blockId).sort((a,b)=>a.y-b.y||a.x-b.x),rows=[];
  for(const w of words){let row=rows[rows.length-1];if(!row||Math.abs(row.y-w.y)>1){row={y:w.y,words:[]};rows.push(row);}row.words.push(w);}
  let shiftY=0,previousBottom=-Infinity,previousNaturalBottom=-Infinity;
  for(const row of rows){
   let top=row.y,bottom=row.y+35;
   for(const w of row.words){const r=w.reverse;if(!r?.visible)continue;if(r.placement==='above')top=Math.min(top,w.y-32);if(r.placement==='below')bottom=Math.max(bottom,w.y+69);}
   const gap=Math.min(16,Math.max(0,row.y-previousNaturalBottom));shiftY=Math.max(shiftY,previousBottom+gap-top);previousBottom=bottom+shiftY;previousNaturalBottom=row.y+35;
   let shiftX=0;for(const w of row.words){const out=result.find(i=>i.id===w.id);out.x=w.x+shiftX;out.y=w.y+shiftY;const r=w.reverse;if(r?.visible&&r.placement==='replace'){const reverseWidth=widthOf({...w,text:r.text});shiftX+=Math.max(0,reverseWidth-widthOf(w));}}
  }
 }
 for(const note of result.filter(i=>i.kind==='note'&&i.anchorId)){
  const anchor=result.find(i=>i.id===note.anchorId&&i.kind==='word');
  if(anchor&&Number.isFinite(note.offsetX)&&Number.isFinite(note.offsetY)){note.x=anchor.x+note.offsetX;note.y=anchor.y+note.offsetY;}
 }
 return result;
}
