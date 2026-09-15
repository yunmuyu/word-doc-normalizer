function isMostlyChinese(s){return (s.match(/[\u4e00-\u9fff]/g)||[]).length>(s.match(/[A-Za-z]/g)||[]).length}

function paragraphTextNodes(p){return Array.from(p.getElementsByTagNameNS(W,'t')).filter(t=>{const r=t.parentNode;return r&&r.namespaceURI===W&&r.localName==='r'&&!protectedRun(r)})}
function rewriteParagraphTextPreserveRuns(p,transformer){
  const nodes=paragraphTextNodes(p);if(!nodes.length)return false;
  const old=nodes.map(t=>t.textContent||'').join(''),neo=transformer(old);if(neo===old)return false;
  let pos=0;
  for(let i=0;i<nodes.length;i++){
    const t=nodes[i],oldLen=(t.textContent||'').length;
    const take=i===nodes.length-1?neo.slice(pos):neo.slice(pos,pos+Math.min(oldLen,Math.max(0,neo.length-pos)));
    t.textContent=take;pos+=take.length;
    if(/^\s|\s$/.test(take))t.setAttribute('xml:space','preserve');else t.removeAttribute('xml:space');
  }
  return true;
}
function insertPlainTextAtStart(p,text){
  const r=wEl(p.ownerDocument,'r'),base=cloneBaseRPr(p);if(base)r.appendChild(base);setBold(r,false);setBlack(r);
  const t=wEl(p.ownerDocument,'t');if(/^\s|\s$/.test(text))t.setAttribute('xml:space','preserve');t.textContent=text;r.appendChild(t);
  const pp=firstDirect(p,W,'pPr');p.insertBefore(r,pp?pp.nextSibling:p.firstChild);return r;
}
function appendPlainText(p,text){const r=wEl(p.ownerDocument,'r'),base=cloneBaseRPr(p);if(base)r.appendChild(base);setBold(r,false);setBlack(r);const t=wEl(p.ownerDocument,'t');if(/^\s|\s$/.test(text))t.setAttribute('xml:space','preserve');t.textContent=text;r.appendChild(t);p.appendChild(r);return r}
function paragraphNumPr(p){const pp=firstDirect(p,W,'pPr');return pp&&firstDirect(pp,W,'numPr')}
function removeParagraphNumbering(p){const pp=firstDirect(p,W,'pPr');if(pp)removeDirect(pp,'numPr')}
function numPrKey(p){const np=paragraphNumPr(p);if(!np)return'';const ni=firstDirect(np,W,'numId'),il=firstDirect(np,W,'ilvl');return `${ni?ni.getAttributeNS(W,'val')||'0':'0'}|${il?il.getAttributeNS(W,'val')||'0':'0'}`}

function collapseCitationNumbers(nums){
  const out=[];let i=0;
  while(i<nums.length){let j=i;while(j+1<nums.length&&nums[j+1]===nums[j]+1)j++;out.push(j>i?`${nums[i]}-${nums[j]}`:`${nums[i]}`);i=j+1}
  return out.join('，');
}
function normalizeCitationSequences(n,stats){
  const before=n;
  n=n.replace(/(?:\[\d+\]\s*){2,}/g,seq=>{const nums=Array.from(seq.matchAll(/\[(\d+)\]/g),m=>Number(m[1]));return nums.length>1?`[${collapseCitationNumbers(nums)}]`:seq});
  n=n.replace(/\[(\d+(?:\s*[,，]\s*\d+)+)\]/g,(m,inner)=>{const nums=inner.split(/[,，]/).map(x=>Number(x.trim())).filter(Number.isFinite);return nums.length>1?`[${collapseCitationNumbers(nums)}]`:m});
  if(stats&&n!==before)stats.citationMerged++;
  return n
}

const UNIT_TOKEN='(?:km|cm|mm|μm|um|nm|m|ms|μs|us|ns|s|min|h|Hz|kHz|MHz|GHz|V|mV|kV|A|mA|kA|W|mW|kW|MW|Pa|kPa|MPa|GPa|N|kN|J|kJ|K|℃|rad(?:\\/s)?|m\\/s(?:²|2)?|m²|m³|kg|mg|g|L|mL|dB)';
function groupedNumber(raw){
  const clean=raw.replace(/\s+/g,'');if(!/^\d+(?:\.\d+)?$/.test(clean))return raw;
  const [ip,fp]=clean.split('.');if(!fp&&ip.length===4){const y=Number(ip);if(y>=1900&&y<=2099)return raw}
  const gi=ip.length>3?ip.replace(/\B(?=(\d{3})+(?!\d))/g,' '):ip;
  const gf=fp&&fp.length>3?fp.replace(/(\d{3})(?=\d)/g,'$1 '):fp;
  return gf!==undefined?`${gi}.${gf}`:gi;
}
function normalizeNumbersAndUnits(n,kind){
  if(['affil','enAffil','classLine','reference','date','fund','title','enTitle','author','enAuthor'].includes(kind))return n;
  const unitRe=new RegExp(`(^|[^A-Za-z0-9_])([0-9]+(?:\\.[0-9]+)?)\\s*(${UNIT_TOKEN})(?=$|[^A-Za-z])`,'g');
  n=n.replace(unitRe,(m,pre,num,unit)=>`${pre}${groupedNumber(num)} ${unit}`);
  n=n.replace(/(^|[^A-Za-z0-9_.])([0-9]{4,}(?:\.[0-9]+)?|[0-9]+\.[0-9]{4,})(?=$|[^A-Za-z0-9_.])/g,(m,pre,num)=>`${pre}${groupedNumber(num)}`);
  return n;
}
function normalizeChineseQuotes(n,kind){if(!['body','cnAbstract','cnKeywords','caption','authorBio'].includes(kind))return n;return n.replace(/"([^"\r\n]*[\u4e00-\u9fff][^"\r\n]*)"/g,'“$1”')}

function normalizeString(s,kind,opt,stats){let n=s.replace(/\u00a0/g,' ');
 if(opt.refs)n=normalizeCitationSequences(n,stats);
 if(!opt.punct)return n;
 n=n.replace(/(\d+)\s*:\s*(\d+)\s*:\s*(\d+)/g,'$1∶$2∶$3');
 n=normalizeChineseQuotes(n,kind);
 n=normalizeNumbersAndUnits(n,kind);
 if(kind==='cnKeywords')n=n.replace(/;/g,'；');
 if(kind==='enAbstract')n=n.replace(/^Abstract：\s*/i,'Abstract: ');
 if(kind==='enKeywords')n=n.replace(/^(?:Keywords|Key\s*words)：\s*/i,m=>m.replace('：',': ')).replace(/；/g,';').replace(/;(?=\S)/g,'; ');
 if(kind==='fund')n=n.replace(/（No\.\s*/gi,'（').replace(/\bNo\.\s*/gi,'');
 if(kind==='affil')n=n.replace(/^\(/,'（').replace(/\)$/,'）').replace(/;/g,'；');
 if(['body','cnAbstract','cnKeywords','authorBio'].includes(kind)){
   n=n.replace(/^\s*[（(]\s*([1-9]\d*)\s*[）)]\s*/,'$1）');
   n=n.replace(/^\s*([1-9]\d*)[、．\.]\s*/,'$1）');
 }
 if(kind==='h1'){n=n.replace(/^(\d+)\s*/,'$1  ');if(/^(\d+)\s*(?:结\s*论(?:\s*与\s*展\s*望)?|总\s*结(?:\s*与\s*展\s*望)?)\s*$/.test(n)){const no=n.match(/^(\d+)/)[1];n=`${no}  结  论`;stats.finalHeadingFixed++}else{const m=n.match(/^(\d+)  ([\u4e00-\u9fff]{2})$/);if(m)n=`${m[1]}  ${m[2][0]}  ${m[2][1]}`}}
 if(kind==='h2')n=n.replace(/^(\d+\.\d+)\s*/,'$1  ');
 if(kind==='h3')n=n.replace(/^(\d+\.\d+\.\d+)\s*/,'$1  ');
 if(kind==='caption')n=normCaptionText(n);
 if(kind==='refHead')n='参考文献：';
 return n}
function normalizeTextNodes(p,kind,stats,opt){const changed=rewriteParagraphTextPreserveRuns(p,s=>normalizeString(s,kind,opt,stats));if(changed)stats.textFixes++}

function stylePrefixLabel(p,re){
  const full=pText(p),m=full.match(re);if(!m)return;let remain=m[0].length;
  for(const r of Array.from(allRuns(p))){
    if(protectedRun(r))continue;const tx=pText(r);if(!tx)continue;
    if(remain<=0){setBold(r,false);continue}
    if(remain>=tx.length){setBold(r,true);remain-=tx.length;continue}
    if(isSimpleRun(r)){
      const a=cloneRunWithText(r,tx.slice(0,remain)),b=cloneRunWithText(r,tx.slice(remain));setBold(a,true);setBold(b,false);r.parentNode.insertBefore(a,r);r.parentNode.insertBefore(b,r);rm(r);remain=0;
    }else{setBold(r,true);remain=0}
  }
}
function boldWholeParagraph(p){for(const r of allRuns(p))if(pText(r)&&!protectedRun(r))setBold(r,true)}
function boldClassLabels(p){for(const r of allRuns(p)){if(protectedRun(r))continue;const tx=pText(r);setBold(r,/中图分类号[:：]|文献标识码[:：]|文章编号[:：]/.test(tx))}}
function styleCrossRefs(p,stats,opt){if(!opt.refs)return;for(const r of Array.from(allRuns(p)))stats.citations+=decorateMatchesInRun(r,/\[\d+(?:[-–—,，]\d+)*\]/g,x=>{setSuper(x,true);if(opt.markup)setHighlight(x,'yellow')});if(opt.markup){for(const r of Array.from(allRuns(p)))stats.crossrefs+=decorateMatchesInRun(r,/(?:如|见|参见)?(?:图|表)\s*\d+(?:\s*所示)?/g,x=>setHighlight(x,'yellow'))}}
function applyParaGeometry(p,k){setSpacing(p);clearTabs(p);if(['title','author','affil','enTitle','enAuthor','enAffil','caption'].includes(k)){setJc(p,'center');setInd(p,null)}else if(k==='empty'){}else{setJc(p,'left');setInd(p,420)}}
function clearParagraphBold(p){for(const r of allRuns(p))if(pText(r)&&!protectedRun(r))setBold(r,false)}

async function readNumberingSpec(zip){
 const out=new Map(),f=zip.file('word/numbering.xml');if(!f)return out;
 const txt=await f.async('string'),d=new DOMParser().parseFromString(txt,'application/xml');if(d.getElementsByTagName('parsererror')[0])return out;
 const absMap=new Map();
 for(const abs of Array.from(d.getElementsByTagNameNS(W,'abstractNum'))){
   const aid=abs.getAttributeNS(W,'abstractNumId');const levels=new Map();
   for(const lvl of direct(abs,W,'lvl')){const il=lvl.getAttributeNS(W,'ilvl')||'0',st=firstDirect(lvl,W,'start'),lt=firstDirect(lvl,W,'lvlText'),fmt=firstDirect(lvl,W,'numFmt');levels.set(il,{start:Number(st&&st.getAttributeNS(W,'val')||1),text:lt&&lt.getAttributeNS(W,'val')||'%1',format:fmt&&fmt.getAttributeNS(W,'val')||'decimal'})}
   absMap.set(aid,levels);
 }
 for(const num of Array.from(d.getElementsByTagNameNS(W,'num'))){
   const numId=num.getAttributeNS(W,'numId'),aid=firstDirect(num,W,'abstractNumId'),levels=absMap.get(aid&&aid.getAttributeNS(W,'val'))||new Map();
   for(const [il,base] of levels){let spec={...base};const ov=direct(num,W,'lvlOverride').find(x=>(x.getAttributeNS(W,'ilvl')||'0')===il);if(ov){const so=firstDirect(ov,W,'startOverride');if(so)spec.start=Number(so.getAttributeNS(W,'val')||spec.start);const ol=firstDirect(ov,W,'lvl');if(ol){const st=firstDirect(ol,W,'start'),lt=firstDirect(ol,W,'lvlText'),fmt=firstDirect(ol,W,'numFmt');if(st)spec.start=Number(st.getAttributeNS(W,'val')||spec.start);if(lt)spec.text=lt.getAttributeNS(W,'val')||spec.text;if(fmt)spec.format=fmt.getAttributeNS(W,'val')||spec.format}}
     out.set(`${numId}|${il}`,spec)}
 }
 return out;
}
function plainListLabel(spec,n,inRefs=false){if(inRefs||/\[\s*%\d\s*\]/.test(spec&&spec.text||''))return `[${n}] `;return `${n}）`}
function flattenAutomaticLists(body,stats,specs=new Map()){
 const ps=Array.from(body.childNodes).filter(n=>n.nodeType===1&&n.namespaceURI===W&&n.localName==='p');
 const refHeadIndex=ps.findIndex(p=>/^参考文献[:：]?\s*$/.test(visibleText(p))),bioIndex=ps.findIndex(p=>/^作者简介[:：]/.test(visibleText(p)));
 const counters=new Map();
 for(let i=0;i<ps.length;i++){
   const p=ps[i],np=paragraphNumPr(p);if(!np)continue;const key=numPrKey(p)||'list',spec=specs.get(key)||{start:1,text:'%1'},current=counters.has(key)?counters.get(key)+1:spec.start;counters.set(key,current);
   const t=visibleText(p),inRefs=refHeadIndex>=0&&i>refHeadIndex&&(bioIndex<0||i<bioIndex);removeParagraphNumbering(p);
   if(inRefs){if(!/^\[\d+\]\s*/.test(t))insertPlainTextAtStart(p,plainListLabel(spec,current,true));stats.listsFlattened++;continue}
   if(/^\d+(?:\.\d+){0,2}\s+/.test(t)||isCaptionText(t)){stats.listsFlattened++;continue}
   if(!/^\s*(?:\d+）|\[\d+\])/.test(t))insertPlainTextAtStart(p,plainListLabel(spec,current,false));stats.listsFlattened++;
 }
}
function nextBodyParagraph(p){let n=p.nextSibling;while(n&&!(n.nodeType===1&&n.namespaceURI===W&&n.localName==='p')){if(n&&n.nodeType===1)return null;n=n.nextSibling}return n}
function mergeMicroHeadings(body,stats){
 const ps=Array.from(body.childNodes).filter(n=>n.nodeType===1&&n.namespaceURI===W&&n.localName==='p');
 for(const p of ps){if(!p.parentNode)continue;const t=visibleText(p),m=t.match(/^([1-9]\d*）)([^。！？；：]{2,40})$/);if(!m)continue;
   const nx=nextBodyParagraph(p);if(!nx)continue;const nt=visibleText(nx);if(!nt||/^(?:\d+(?:\.\d+){0,2}\s+|\d+）|图\s*\d+|表\s*\d+|参考文献[:：]|作者简介[:：]|收稿日期[:：]|基金项目[:：])/.test(nt))continue;
   appendPlainText(p,'。');
   for(const node of Array.from(nx.childNodes)){if(node.nodeType===1&&node.namespaceURI===W&&node.localName==='pPr')continue;p.appendChild(node)}
   rm(nx);stats.microHeadingsMerged++;
 }
}

function applyFormatting(body,opt,stats){const kinds=classifyBody(body);for(const p of Array.from(body.getElementsByTagNameNS(W,'p'))){setParagraphFontDefault(p);for(const r of allRuns(p)){setFont(r);setBlack(r)}}for(const [p,k] of kinds){if(!p.parentNode)continue;normalizeTextNodes(p,k,stats,opt);applyParaGeometry(p,k);clearParagraphBold(p);if(['title','enTitle'].includes(k))boldWholeParagraph(p);
 if(k==='cnAbstract')stylePrefixLabel(p,/^摘\s*要[:：]/);
 if(k==='cnKeywords')stylePrefixLabel(p,/^关键词[:：]/);
 if(k==='enAbstract')stylePrefixLabel(p,/^Abstract[:：]/i);
 if(k==='enKeywords')stylePrefixLabel(p,/^(?:Keywords|Key\s*words)[:：]/i);
 if(k==='date')stylePrefixLabel(p,/^收稿日期[:：]/);
 if(k==='fund')stylePrefixLabel(p,/^基金项目[:：]/);
 if(k==='authorBio')stylePrefixLabel(p,/^作者简介[:：]/);
 if(k==='refHead')boldWholeParagraph(p);
 if(k==='classLine')boldClassLabels(p);
 if(k==='caption'&&opt.markup)for(const r of allRuns(p))if(pText(r)&&!protectedRun(r))setHighlight(r,'yellow');
 if(k==='caption')stats.captionStyled++;
 if(k==='author'||k==='enAuthor'){for(const r of allRuns(p)){const rp=firstDirect(r,W,'rPr'),v=rp&&firstDirect(rp,W,'vertAlign');if(v&&v.getAttributeNS(W,'val')==='superscript'&&/^\d+$/.test(pText(r).trim())&&opt.markup)setHighlight(r,'yellow')}}
 if(['body','cnAbstract','cnKeywords','authorBio'].includes(k))styleCrossRefs(p,stats,opt)}
 for(const p of Array.from(body.getElementsByTagNameNS(W,'p')))for(const r of allRuns(p))setBlack(r)
}

function attrVal(el,local){return el?el.getAttributeNS(W,local):''}
function collectMentions(body,kind){const out=new Map();for(const p of Array.from(body.childNodes).filter(n=>n.nodeType===1&&n.namespaceURI===W&&n.localName==='p')){const t=visibleText(p);if(!t||isCaptionText(t)||/^参考文献[:：]?/.test(t)||/^作者简介[:：]/.test(t))continue;const re=new RegExp(`${kind}\\s*(\\d+)(?!\\d)`,'g');let m;while((m=re.exec(t))){const num=Number(m[1]);if(!out.has(num))out.set(num,[]);out.get(num).push(t.slice(Math.max(0,m.index-18),Math.min(t.length,m.index+m[0].length+18)))}}return out}
function explicitRedCount(body){let n=0;for(const c of Array.from(body.getElementsByTagNameNS(W,'color'))){const v=(attrVal(c,'val')||'').toUpperCase();if(v&&!['000000','AUTO'].includes(v))n++}return n}
function paragraphIsBold(p){const runs=allRuns(p).filter(r=>pText(r).trim()&&!protectedRun(r));return !!runs.length&&runs.every(r=>{const rp=firstDirect(r,W,'rPr');return !!(rp&&firstDirect(rp,W,'b'))})}
function collectUnitSpacingIssues(body){const out=[];const re=new RegExp(`\\d(?:\\.\\d+)?(${UNIT_TOKEN})\\b`);for(const p of Array.from(body.childNodes).filter(n=>n.nodeType===1&&n.namespaceURI===W&&n.localName==='p')){const t=visibleText(p);if(re.test(t))out.push(t.slice(0,90));if(out.length>=8)break}return out}
function collectAsciiChineseQuotes(body){const out=[];for(const p of Array.from(body.childNodes).filter(n=>n.nodeType===1&&n.namespaceURI===W&&n.localName==='p')){const t=visibleText(p);if(/"[^"\r\n]*[\u4e00-\u9fff][^"\r\n]*"/.test(t))out.push(t.slice(0,90));if(out.length>=8)break}return out}
function remainingDirectNumPr(body){return Array.from(body.childNodes).filter(n=>n.nodeType===1&&n.namespaceURI===W&&n.localName==='p'&&paragraphNumPr(n)).length}
function auditDocument(body,stats,sourceDate){
 const caps=captionMap(body),figMent=collectMentions(body,'图'),tblMent=collectMentions(body,'表');const figNums=[...caps.fig.keys()].sort((a,b)=>a-b),tblNums=[...caps.tbl.keys()].sort((a,b)=>a-b);const missingFig=figNums.filter(n=>!figMent.has(n)),missingTbl=tblNums.filter(n=>!tblMent.has(n));const orphanFig=[...figMent.keys()].filter(n=>!caps.fig.has(n)).sort((a,b)=>a-b),orphanTbl=[...tblMent.keys()].filter(n=>!caps.tbl.has(n)).sort((a,b)=>a-b);
 const ps=Array.from(body.childNodes).filter(n=>n.nodeType===1&&n.namespaceURI===W&&n.localName==='p');const refHead=ps.find(p=>/^参考文献[:：]?\s*$/.test(visibleText(p))),bio=ps.find(p=>/^作者简介[:：]/.test(visibleText(p))),date=ps.find(p=>/^收稿日期[:：]/.test(visibleText(p)));const h1s=ps.filter(p=>/^\d+\s+/.test(visibleText(p)));const lastH1=h1s.length?h1s[h1s.length-1]:null;const finalHeadingOK=!lastH1||/^\d+\s+结\s+论\s*$/.test(visibleText(lastH1));
 const bodyText=ps.filter(p=>p!==refHead&&(!refHead||ps.indexOf(p)<ps.indexOf(refHead))).map(visibleText).join('\n');const adjacent=[];for(const m of bodyText.matchAll(/\[(\d+)\]\s*\[(\d+)\]/g))adjacent.push(m[0]);const actualDrawings=body.getElementsByTagNameNS(W,'drawing').length;
 let bibliographyCount=0;if(refHead){const a=ps.indexOf(refHead),b=bio?ps.indexOf(bio):ps.length;for(let i=a+1;i<b;i++)if(visibleText(ps[i]))bibliographyCount++}
 const cites=[];for(const m of bodyText.matchAll(/\[([0-9]+(?:[-–—][0-9]+)?(?:，[0-9]+(?:[-–—][0-9]+)?)*)\]/g)){for(const part of m[1].split('，')){const ab=part.split(/[-–—]/).map(Number),a=ab[0],b=ab.length>1?ab[1]:a;if(Number.isFinite(a)&&Number.isFinite(b)&&b>=a&&b-a<100)for(let i=a;i<=b;i++)cites.push(i)}}const citedUnique=[...new Set(cites)].sort((a,b)=>a-b),citeMissingBib=citedUnique.filter(n=>bibliographyCount&&n>bibliographyCount);
 return{figNums,tblNums,missingFig,missingTbl,orphanFig,orphanTbl,actualDrawings,redRemaining:explicitRedCount(body),adjacent,refHeadOK:!!refHead&&visibleText(refHead)==='参考文献：',refHeadBold:!!refHead&&paragraphIsBold(refHead),bioOK:!!bio&&attrVal(firstDirect(ensurePPr(bio),W,'jc'),'val')!=='center',finalHeadingOK,finalHeading:lastH1?visibleText(lastH1):'',sourceDate,outputDate:date?visibleText(date):'',datePreserved:!sourceDate||!date||visibleText(date)===sourceDate,bibliographyCount,citedUnique,citeMissingBib,citationMerged:stats.citationMerged,remainingNumPr:remainingDirectNumPr(body),unitSpacingIssues:collectUnitSpacingIssues(body),asciiChineseQuotes:collectAsciiChineseQuotes(body)}
}
