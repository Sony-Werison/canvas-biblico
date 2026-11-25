// Reorganização por passagem, adaptada do motor do Canvas Bíblico original.
export function fitRowsToSlides(items, updateOriginal=false){
 for(const block of new Set(items.filter(i=>i.kind==='word'&&i.blockId).map(i=>i.blockId))){
  const words=items.filter(i=>i.kind==='word'&&i.blockId===block).sort((a,b)=>a.order-b.order);
  let shift=0,lastY=null,rowShift=0;
  for(const word of words){
   const naturalY=word.y;
   if(lastY===null||Math.abs(naturalY-lastY)>1){let y=naturalY+shift;const page=Math.floor(Math.max(0,y)/675);if(y-page*675>635){shift=(page+1)*675+35-naturalY;y=naturalY+shift;}rowShift=shift;lastY=naturalY;}
   word.y=naturalY+rowShift;if(updateOriginal)word.oy=word.y;
  }
 }
}
export function reflowAfterDrag(items, before, movedIds, options, widthOf) {
  const gap = 10, lineGap = 16;
  const moved = new Set(movedIds);
  const blockIds = new Set(items.filter(w => w.kind === 'word' && moved.has(w.id)).map(w => w.blockId));
  for (const blockId of blockIds) {
    if (!blockId) continue; // Projetos antigos mantêm o arrasto livre.
    const words = items.filter(w => w.kind === 'word' && w.blockId === blockId).sort((a,b) => a.order-b.order);
    const changed = words.filter(w => moved.has(w.id));
    if (!changed.length) continue;
    const anchor = changed[0], oldAnchor = before.get(anchor.id);
    if (!oldAnchor) continue;
    if (changed.length === words.length) {
      const dx=anchor.x-oldAnchor.x,dy=anchor.y-oldAnchor.y;
      for(const w of words){w.originX+=dx;w.originY+=dy;w.ox+=dx;w.oy+=dy;}
      continue;
    }
    const first=words.indexOf(anchor), originX=anchor.originX, originY=anchor.originY;
    const previous=words[first-1], neighbor=words.filter(w=>!moved.has(w.id)&&w.order<anchor.order&&Math.abs(w.y-anchor.y)<Math.min(options.lineStep/2,18)&&w.x+widthOf(w)<=anchor.x+Math.min(options.indentStep/2,widthOf(anchor)/2)).sort((a,b)=>(b.x+widthOf(b))-(a.x+widthOf(a)))[0];
    let x=neighbor?neighbor.x+widthOf(neighbor)+gap:originX+Math.max(0,Math.round((anchor.x-originX)/options.indentStep))*options.indentStep;
    let y=neighbor?neighbor.y:originY+Math.round((anchor.y-originY)/options.lineStep)*options.lineStep;
    if(previous && Math.abs(y-previous.y)<0.1){let edge=x-gap;for(let i=first-1;i>=0&&Math.abs(words[i].y-y)<1;i--){const w=words[i],width=widthOf(w);if(w.x+width<=edge)break;w.x=edge-width;edge=w.x-gap;}}
    anchor.x=x;anchor.y=y;
    if(previous){anchor.breakBefore=Math.abs(y-previous.y)<1?0:Math.max(1,Math.round((y-previous.y)/options.lineStep));anchor.breakKind=anchor.breakBefore?'manual':'none';}
    x+=widthOf(anchor)+gap;
    for(let i=first+1;i<words.length;i++){
      const word=words[i],old=before.get(word.id),oldPrev=before.get(words[i-1].id);
      const previousBreak=old&&oldPrev&&old.y-oldPrev.y>18;
      const manual=word.breakKind==='manual'||word.breakBefore>0||(previousBreak&&word.breakKind!=='auto');
      if(manual){y+=previousBreak?old.y-oldPrev.y:Math.ceil((35+lineGap)/options.lineStep)*options.lineStep*Math.max(1,word.breakBefore||1);x=previousBreak?old.x:originX;word.breakBefore=Math.max(1,word.breakBefore||1);word.breakKind='manual';}
      else if(x+widthOf(word)>originX+(word.flowWidth||1000)+0.1){y+=Math.ceil((35+lineGap)/options.lineStep)*options.lineStep;x=originX;word.breakBefore=0;word.breakKind='auto';}
      else {word.breakBefore=0;word.breakKind='none';}
      word.x=x;word.y=y;x+=widthOf(word)+gap;
    }
  }
}
