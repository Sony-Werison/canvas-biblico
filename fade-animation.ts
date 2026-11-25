// Adaptado do exportador PowerPoint original do Canvas B?blico.
type ExportLayout={elements:any[];guides?:{x:number;top:number;bottom:number;ticks:{y:number;endX:number}[]}[]};
type ReverseLayout=ExportLayout&{ownerId:string};
type PresentationExport={guides?:ExportLayout['guides']};
const metadata=(e:any)=>e.customData?.canvasBiblico||{};
const animatedElement=(e:any)=>(!!metadata(e).highlightFor&&metadata(e).kind!=='font-text')||['note','note-box','connector','connector-head','connector-marker','reverse-background','reverse-label'].includes(metadata(e).role)||e.type==='arrow';
export function fadeAnimation(xml:string,elements:any[],layouts:ReverseLayout[]=[],projection={scale:1,ox:0,oy:0},guides:PresentationExport['guides']=[]):string {
    const byId=new Map(elements.map(e=>[e.id,e])),parent=new Map(elements.map(e=>[e.id,e.id]));
    const root=(id:string):string=>{const p=parent.get(id);if(!p || p===id)return id;const r=root(p);parent.set(id,r);return r;};
    const join=(a:string,b:string)=>{if(parent.has(a)&&parent.has(b))parent.set(root(a),root(b));};
    const groupOwner=new Map<string,string>();
    for(const e of elements){
        for(const group of e.groupIds||[]){const owner=groupOwner.get(group);if(owner)join(e.id,owner);else groupOwner.set(group,e.id);}
        const m=metadata(e);
        if(m.headFor)join(e.id,m.headFor);
        if(m.noteFor)join(e.id,m.noteFor);
        if(e.containerId)join(e.id,e.containerId);
    }
    // A rendered highlight is often split into one shape per word. Its range
    // identity, rather than those individual shapes, defines the logical click.
    const rangeOwner=new Map<string,string>();
    for(const e of elements.filter(animatedElement)){
        const owner=byId.get(metadata(e).highlightFor),ranges=owner?.customData?.bibleHighlights?.rangeIds;
        for(const range of Object.values(ranges||{})){
            const key=String(range),previous=rangeOwner.get(key);
            if(previous)join(e.id,previous);else rangeOwner.set(key,e.id);
        }
        if(metadata(e).rangeNextId){
            const nextWord=byId.get(metadata(e).rangeNextId);
            if(nextWord)for(const next of elements.filter(n=>metadata(n).highlightFor===nextWord.id && metadata(n).kind===metadata(e).kind))join(e.id,next.id);
        }
    }
    // Connectors may bind to either the note text or its container.
    for(const e of elements){const m=metadata(e);if(m.role!=='connector'&&e.type!=='arrow')continue;
        for(const id of [m.toId,m.fromId,e.startBinding?.elementId,e.endBinding?.elementId]){
            const target=byId.get(id);
            if(target && elements.some(note=>metadata(note).role==='note'&&root(note.id)===root(target.id)))join(e.id,id);
        }
    }
    const noteRoots=new Set(elements.filter(e=>metadata(e).role==='note').map(e=>root(e.id)));
    const candidates=elements.filter(e=>e.opacity!==0 && (animatedElement(e)||noteRoots.has(root(e.id))) &&
        !(e.type==='text'&&metadata(e).role==='passage'&&!metadata(e).highlightFor));
    const shapes=new Map<string,string[]>();
    for(const match of xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"[^>]*\bname="([^"]+)"/g)){
        const ids=shapes.get(match[2])||[];ids.push(match[1]);shapes.set(match[2],ids);
    }
    const batches=new Map<string,{ids:Set<string>;anchors:any[];fallbacks:any[]}>();
    for(const e of candidates){
        const m=metadata(e),identity=m.kind==='styled-text'?`format-${m.highlightFor}`:e.id;
        const ids=[...(shapes.get(`!!cb-${identity}`)||[])];
        for(const [name,parts] of shapes)if(name.startsWith(`!!cb-${e.id}-segment-`))ids.push(...parts);
        if(!ids.length)continue;
        const key=root(e.id),batch=batches.get(key)||{ids:new Set<string>(),anchors:[],fallbacks:[]};
        ids.forEach(id=>batch.ids.add(id));
        const anchors=[m.highlightFor,m.reverseFor,...(m.role==='note'?[]:(m.targetIds||[m.targetId])),m.fromId,m.toId]
            .map(id=>byId.get(id)).filter(target=>target && target.type==='text'&&!metadata(target).highlightFor&&metadata(target).role!=='note');
        batch.anchors.push(...(noteRoots.has(key)?[]:anchors));batch.fallbacks.push(e);batches.set(key,batch);
    }
    const ordered=[...batches.values()].map(batch=>({...batch,anchor:batch.fallbacks.find(e=>metadata(e).role==='note')||
        [...(batch.anchors.length?batch.anchors:batch.fallbacks)].sort((a,b)=>a.y-b.y||a.x-b.x)[0]}))
        .sort((a,b)=>a.anchor.y-b.anchor.y||a.anchor.x-b.anchor.x);
    // Row tolerance prevents a small border/arrow offset from changing reading order.
    const rows:typeof ordered[]=[];
    for(const batch of ordered){const row=rows.at(-1);if(row && Math.abs(batch.anchor.y-row[0].anchor.y)<=Math.min(batch.anchor.height,row[0].anchor.height)*.3)row.push(batch);else rows.push([batch]);}
    const steps=rows.flatMap(row=>row.sort((a,b)=>a.anchor.x-b.anchor.x))
        .sort((a,b)=>Number(a.fallbacks.some(e=>metadata(e).role==='note'))-Number(b.fallbacks.some(e=>metadata(e).role==='note')));
    if(!steps.length)return xml;
    // Animate one native group per logical element, keeping the animation pane
    // as concise as the click sequence while preserving every editable child.
    let shapeId=Math.max(...[...xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)].map(m=>Number(m[1])));
    const shapeXml=new Map<string,string>();
    for(const match of xml.matchAll(/<p:(sp|pic|cxnSp)>[\s\S]*?<\/p:\1>/g)){
        const id=match[0].match(/<p:cNvPr\b[^>]*\bid="(\d+)"/)?.[1];if(id)shapeXml.set(id,match[0]);
    }
    const geometry=(layout:ExportLayout)=>{
        const result=new Map<string,{x:number;y:number;width:number;height:number}>();
        for(const e of layout.elements)result.set(`!!cb-${metadata(e).kind==='styled-text'?`format-${metadata(e).highlightFor}`:e.id}`,e);
        for(const [i,g]of (layout.guides||[]).entries()){
            result.set(`!!cb-indent-${i}`,{x:g.x,y:g.top,width:0,height:g.bottom-g.top});
            for(const [j,t]of g.ticks.entries())result.set(`!!cb-indent-${i}-${j}`,{x:g.x,y:t.y,width:t.endX-g.x,height:0});
        }
        return result;
    };
    type Frame={x:number;y:number;width:number;height:number};
    const initialGeometry=geometry({elements,guides}),layoutGeometry=new Map(layouts.map(l=>[l.ownerId,geometry(l)]));
    const shapeGeometry=new Map<string,Frame>(),shapeNames=new Map<string,string>();
    for(const [name,ids]of shapes)for(const id of ids)shapeNames.set(id,name);
    for(const [id,shape]of shapeXml){
        const off=shape.match(/<a:off x="(-?\d+)" y="(-?\d+)"/),ext=shape.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
        if(off&&ext)shapeGeometry.set(id,{x:Number(off[1])/914400,y:Number(off[2])/914400,width:Number(ext[1])/914400,height:Number(ext[2])/914400});
    }
    const frameAt=(id:string,state:typeof initialGeometry):Frame|undefined=>{
        const shape=shapeGeometry.get(id),raw=shapeNames.get(id)||'',name=initialGeometry.has(raw)?raw:raw.replace(/-segment-\d+$/,'');
        const base=initialGeometry.get(name),target=state.get(name);if(!shape||!base||!target)return shape;
        const sx=base.width>0?target.width/base.width:1,sy=base.height>0?target.height/base.height:1;
        return {x:projection.ox+target.x*projection.scale+(shape.x-projection.ox-base.x*projection.scale)*sx,
            y:projection.oy+target.y*projection.scale+(shape.y-projection.oy-base.y*projection.scale)*sy,width:shape.width*sx,height:shape.height*sy};
    };
    const transformAt=(id:string,state:typeof initialGeometry)=>{
        const base=shapeGeometry.get(id)!,frame=frameAt(id,state)||base;
        return {dx:frame.x+frame.width/2-base.x-base.width/2,dy:frame.y+frame.height/2-base.y-base.height/2,
            sx:base.width?frame.width/base.width:1,sy:base.height?frame.height/base.height:1};
    };
    // Animate a whole native group. Child property animations use the group's
    // local coordinates in PowerPoint and can expose a future hidden verse.
    // Split only when children have different movement over the full timeline;
    // their entrance effects still share the same click.
    const motionTargets=new Map([...shapeGeometry.keys()].map(id=>[id,id]));
    const motionEntry=new Map<string,{step:number;transform:ReturnType<typeof transformAt>}>();
    let revealedGeometry=initialGeometry;
    const signature=(id:string)=>shapeGeometry.has(id)?JSON.stringify([...layoutGeometry.values()].map(state=>Object.values(transformAt(id,state)).map(n=>Math.round(n*1e6)/1e6))):'';
    const grouped=steps.map((batch,stepIndex)=>{
        const partitions=new Map<string,string[]>();
        for(const child of batch.ids){const key=signature(child),list=partitions.get(key)||[];list.push(child);partitions.set(key,list);}
        const members=batch.fallbacks,note=members.find(e=>metadata(e).role==='note'),reverse=members.find(e=>metadata(e).role==='reverse-label');
        const reverseOwner=reverse&&metadata(reverse).reverseFor;
        if(reverseOwner&&layoutGeometry.has(reverseOwner))revealedGeometry=layoutGeometry.get(reverseOwner)!;
        const entryGeometry=revealedGeometry;
        const words=[...new Map(batch.anchors.map(e=>[e.id,e])).values()]
            .sort((a,b)=>Math.abs(a.y-b.y)<=Math.min(a.height,b.height)*.3?a.x-b.x:a.y-b.y).map(e=>e.text).join(' ');
        const label=(note?'Nota':reverse?`Verso: ${reverse.text}`:`Realce: ${words||batch.anchor.text||''}`).slice(0,90).replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;');
        const ids=new Set<string>(),groups:string[]=[];
        for(const children of partitions.values()){
            const id=String(++shapeId),frames=children.map(child=>frameAt(child,entryGeometry)).filter((f):f is Frame=>!!f);
            const x=frames.length?Math.min(...frames.map(f=>f.x)):0,y=frames.length?Math.min(...frames.map(f=>f.y)):0;
            const w=frames.length?Math.max(...frames.map(f=>f.x+f.width))-x:13.333333,h=frames.length?Math.max(...frames.map(f=>f.y+f.height))-y:7.5;
            const emu=(n:number)=>Math.round(n*914400),cx=Math.max(1,emu(w)),cy=Math.max(1,emu(h));
            const content=children.map(child=>{
                const frame=frameAt(child,entryGeometry),part=shapeXml.get(child)||'';if(!frame)return part;
                return part.replace(/<a:off x="-?\d+" y="-?\d+"\/>/,`<a:off x="${emu(frame.x)}" y="${emu(frame.y)}"/>`)
                    .replace(/<a:ext cx="\d+" cy="\d+"\/>/,`<a:ext cx="${Math.max(1,emu(frame.width))}" cy="${Math.max(1,emu(frame.height))}"/>`);
            }).join('');
            groups.push(`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="${label}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${cx}" cy="${cy}"/><a:chOff x="${emu(x)}" y="${emu(y)}"/><a:chExt cx="${cx}" cy="${cy}"/></a:xfrm></p:grpSpPr>${content}</p:grpSp>`);
            for(const child of children){motionTargets.delete(child);const part=shapeXml.get(child);if(part)xml=xml.replace(part,'');}
            const representative=children.find(child=>shapeGeometry.has(child));if(representative){motionTargets.set(id,representative);motionEntry.set(id,{step:stepIndex,transform:transformAt(representative,entryGeometry)});}
            ids.add(id);
        }
        const replaceIds=new Set<string>(),fadeIds=new Set<string>();
        for(const e of members){const m=metadata(e);
            if(['formatted-text','styled-text'].includes(m.kind))fadeIds.add(m.highlightFor);
            if(m.role==='reverse-label'){
                const owner=byId.get(m.reverseFor),om=metadata(owner||{});
                if(!['above','below'].includes(om.reversePlacement))for(const wordId of om.reverseIds||[m.reverseFor]){replaceIds.add(wordId);replaceIds.add(`format-${wordId}`);}
            }
        }
        const hidden=[...replaceIds].flatMap(wordId=>shapes.get(`!!cb-${wordId}`)||[]);
        const faded=[...fadeIds].filter(wordId=>!replaceIds.has(wordId)).flatMap(wordId=>shapes.get(`!!cb-${wordId}`)||[]);
        return {ids,group:groups.join(''),hidden,faded,duration:fadeIds.size?250:500,reverseOwner};
    });
    xml=xml.replace('</p:spTree>',grouped.map(g=>g.group).join('')+'</p:spTree>');
    let nodeId=2;
    const next=()=>++nodeId;
    let currentGeometry=initialGeometry;
    const motion=(owner:string|undefined,stepIndex:number)=>{
        const target=owner&&layoutGeometry.get(owner);if(!target)return '';
        const effects:string[]=[];
        for(const [id,representative]of motionTargets){
            const entry=motionEntry.get(id);
            // A movement before the entrance makes PowerPoint treat the
            // object as initially visible. Hidden groups are already placed
            // at their entrance layout and only move after being revealed.
            if(entry&&entry.step>=stepIndex)continue;
            const from=transformAt(representative,currentGeometry),to=transformAt(representative,target);
            if(entry)for(const state of [from,to]){state.dx-=entry.transform.dx;state.dy-=entry.transform.dy;state.sx/=entry.transform.sx;state.sy/=entry.transform.sy;}
            if(['dx','dy','sx','sy'].every(key=>Math.abs((from as any)[key]-(to as any)[key])<1e-7))continue;
            const behavior=()=>`<p:cBhvr additive="base"><p:cTn id="${next()}" dur="300" fill="hold"/><p:tgtEl><p:spTgt spid="${id}"/></p:tgtEl></p:cBhvr>`;
            let movement='';
            if(Math.abs(from.dx-to.dx)>1e-7||Math.abs(from.dy-to.dy)>1e-7)movement+=`<p:animMotion origin="layout" path="M ${from.dx/13.333333} ${from.dy/7.5} L ${to.dx/13.333333} ${to.dy/7.5} E" pathEditMode="relative">${behavior()}</p:animMotion>`;
            if(Math.abs(from.sx-to.sx)>1e-7||Math.abs(from.sy-to.sy)>1e-7)movement+=`<p:animScale zoomContents="1">${behavior()}<p:from x="${Math.round(from.sx*100000)}" y="${Math.round(from.sy*100000)}"/><p:to x="${Math.round(to.sx*100000)}" y="${Math.round(to.sy*100000)}"/></p:animScale>`;
            effects.push(`<p:par><p:cTn id="${next()}" presetID="0" presetClass="path" presetSubtype="0" fill="hold" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${movement}</p:childTnLst></p:cTn></p:par>`);
        }
        currentGeometry=target;return effects.join('');
    };
    const sequence=grouped.map((batch,stepIndex)=>{
        const click=next(),parallel=next();
        const movement=motion(batch.reverseOwner,stepIndex);
        const exits=batch.hidden.map(shape=>`<p:par><p:cTn id="${next()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="${next()}" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="${shape}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="hidden"/></p:to></p:set></p:childTnLst></p:cTn></p:par>`).join('');
        // Keep the regular glyphs visible while their replacement fades in.
        // A native exit effect also tells PowerPoint that the base starts visible.
        const crossfades=batch.faded.map(shape=>`<p:par><p:cTn id="${next()}" presetID="10" presetClass="exit" presetSubtype="0" fill="hold" grpId="0" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="out" filter="fade"><p:cBhvr><p:cTn id="${next()}" dur="${batch.duration}" fill="hold"/><p:tgtEl><p:spTgt spid="${shape}"/></p:tgtEl></p:cBhvr></p:animEffect><p:set><p:cBhvr><p:cTn id="${next()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${batch.duration}"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${shape}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="hidden"/></p:to></p:set></p:childTnLst></p:cTn></p:par>`).join('');
        const effects=[...batch.ids].map((shape,index)=>`<p:par><p:cTn id="${next()}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="${index?'withEffect':'clickEffect'}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="${next()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${shape}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set><p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${next()}" dur="${batch.duration}"/><p:tgtEl><p:spTgt spid="${shape}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`).join('');
        return `<p:par><p:cTn id="${click}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${parallel}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${effects}${crossfades}${exits}${movement}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`;
    }).join('');
    const builds=grouped.flatMap(batch=>[...batch.ids].map(id=>`<p:bldP spid="${id}" grpId="0"/>`)).join('');
    const timing=`<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${sequence}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>${builds}</p:bldLst></p:timing>`;
    return xml.replace('</p:sld>',timing+'</p:sld>');
}
