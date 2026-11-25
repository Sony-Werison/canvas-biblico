import JSZip from 'jszip';
import {fadeAnimation} from './fade-animation.ts';
import {layoutReverses,connectionLine} from './geometry.js';

// Reservar margem para o texto, os versos e as anotações de cada faixa.
export const PAGE_HEIGHT = 675;
export function splitIntoPages(items,divisions=[]){
 const pages=new Map();
 for(const item of items.filter(i=>i.kind==='word')){const index=Math.max(0,Math.floor(item.y/PAGE_HEIGHT));if(!pages.has(index))pages.set(index,[]);pages.get(index).push(item);}
 const output=[],breaks=new Set(divisions);
 for(const [index,words] of [...pages.entries()].sort((a,b)=>a[0]-b[0])){
  words.sort((a,b)=>Math.abs(a.y-b.y)<1?a.x-b.x:a.y-b.y);let group=[];
  const flush=()=>{if(group.length){output.push({offset:output.length&&breaks.has(group[0].id)?Math.max(index*PAGE_HEIGHT,group[0].y-45):index*PAGE_HEIGHT,items:group});group=[];}};
  for(const word of words){if(breaks.has(word.id))flush();group.push(word);}flush();
 }
 for(const item of items.filter(i=>i.kind!=='word')){
  const target=output.find(p=>item.anchorId&&p.items.some(w=>w.id===item.anchorId))||output.find(p=>Math.floor(item.y/PAGE_HEIGHT)===Math.floor(p.items[0].y/PAGE_HEIGHT));
  if(target)target.items.push(item);else output.push({offset:Math.max(0,Math.floor(item.y/PAGE_HEIGHT))*PAGE_HEIGHT,items:[item]});
 }
 return output;
}

export function buildReverseLayouts(items,elements,widthOf){
 const owners=items.filter(i=>i.kind==='word'&&i.reverse).sort((a,b)=>Math.abs(a.y-b.y)<10?a.x-b.x:a.y-b.y);
 const opened=new Set(),layouts=[];
 for(const owner of owners){
  opened.add(owner.id);
  const scene=layoutReverses(items.map(i=>i.reverse?{...i,reverse:{...i.reverse,visible:opened.has(i.id)}}:i),widthOf);
  const byId=new Map(scene.map(i=>[i.id,i]));
  const updated=elements.map(e=>{
   const meta=e.customData?.canvasBiblico||{},word=byId.get(e.id)||byId.get(meta.highlightFor)||byId.get(meta.reverseFor);
   if(word){const offset=meta.role==='reverse-label'?(word.reverse.placement==='above'?-32:word.reverse.placement==='below'?37:0):0;return {...e,x:word.x,y:word.y+offset};}
   if(meta.role==='connector'){const line=connectionLine({from:meta.fromId,to:meta.toId},scene,widthOf);if(line)return {...e,x:Math.min(line.x1,line.x2),y:Math.min(line.y1,line.y2),width:Math.abs(line.x2-line.x1)||1,height:Math.abs(line.y2-line.y1)||1};}
   return e;
  });
  layouts.push({ownerId:owner.id,elements:updated,guides:[]});
 }
 return layouts;
}

// Mesmo formato de transição por objetos usado pelo exportador original.
export async function addMorphTransition(bytes, slideAnimations=[]){
 const zip=await JSZip.loadAsync(bytes);
 const morph='<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159"><p:transition spd="slow" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" p14:dur="1500"><p159:morph option="byObject"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>';
 for(const {slideNumber,elements,layouts} of slideAnimations){
  const path=`ppt/slides/slide${slideNumber}.xml`,file=zip.file(path);if(!file)throw new Error('Slide organizado não encontrado.');
  const xml=await file.async('string');if(!xml.includes('</p:sld>'))throw new Error('Slide organizado inválido.');
  const withTransition=xml.replace('</p:sld>',morph+'</p:sld>');
  zip.file(path,fadeAnimation(withTransition,elements,layouts,{scale:13.333/1200,ox:0,oy:0}));
 }
 return zip.generateAsync({type:'arraybuffer',compression:'DEFLATE'});
}
