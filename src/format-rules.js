function isMostlyChinese(s){return (s.match(/[\u4e00-\u9fff]/g)||[]).length>(s.match(/[A-Za-z]/g)||[]).length}
function groupStandaloneNumbers(s,kind){if(['affil','enAffil','classLine','reference','date','fund','title','enTitle','author','enAuthor'].includes(kind))return s;return s.replace(/(^|[^A-Za-z0-9_.-])(\d{4,5})(?=([^A-Za-z0-9]|$))/g,(m,pre,num)=>{const n=Number(num);if(n>=1900&&n<=2099)return m;if(/^\d{5}$/.test(num)&&/邮编|邮政/.test(s))return m;const g=num.replace(/\B(?=(\d{3})+(?!\d))/g,' ');return pre+g})}
function normalizeCitationSequences(n,stats){
 const before=n;
 n=n.replace(/(?:\[\d+\]\s*){2,}/g,seq=>{const nums=Array.from(seq.matchAll(/\[(\d+)\]/g),m=>Number(m[1]));if(nums.length<2)return seq;const ok=nums.every((v,i)=>i===0||v===nums[i-1]+1);if(!ok)return seq;return `[${nums[0]}-${nums[nums.length-1]}]`});
 n=n.replace(/\[(\d+)\s*[,，]\s*(\d+)\]/g,(m,a,b)=>Number(b)===Number(a)+1?`[${a}-${b}]`:m.replace('，',','));
 if(stats&&n!==before)stats.citationMerged++;
 return n
}
function normalizeString(s,kind,opt,stats){let n=s.replace(/\u00a0/g,' ');
 if(opt.refs)n=normalizeCitationSequences(n,stats);
 if(!opt.punct)return n;
 // V4 保守原则：不自动改日期连接符、普通数值范围、参考文献页码范围和千位分组。
 n=n.replace(/(\d+)\s*:\s*(\d+)\s*:\s*(\d+)/g,'$1∶$2∶$3');
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
function normalizeSimpleText(p,kind,stats,opt){const s=visibleText(p),n=normalizeString(s,kind,opt,stats);if(n!==s&&isSimpleParagraph(p)){setPlainParagraphText(p,n);stats.textFixes++}}
function normalizeTextNodes(p,kind,stats,opt){if(isSimpleParagraph(p)){normalizeSimpleText(p,kind,stats,opt);return}for(const t of Array.from(p.getElementsByTagNameNS(W,'t'))){let a=t.textContent||'',b=a;if(opt.refs)b=normalizeCitationSequences(b,stats);if(b!==a){t.textContent=b;stats.textFixes++}}}
function stylePrefixLabel(p,re){const full=pText(p);const m=full.match(re);if(!m)return;const label=m[0];if(isSimpleParagraph(p)){const base=cloneBaseRPr(p);const pp=firstDirect(p,W,'pPr');for(const n of Array.from(p.childNodes)){if(n.nodeType===1&&n!==pp&&!['bookmarkStart','bookmarkEnd','proofErr'].includes(n.localName))rm(n)}const add=(text,bold)=>{if(!text)return;const r=wEl(p.ownerDocument,'r');if(base)r.appendChild(base.cloneNode(true));setBold(r,bold);setBlack(r);const t=wEl(p.ownerDocument,'t');if(/^\s|\s$/.test(text))t.setAttribute('xml:space','preserve');t.textContent=text;r.appendChild(t);p.appendChild(r)};add(label,true);add(full.slice(label.length),false);return}let remain=label.length;for(const r of allRuns(p)){if(protectedRun(r))continue;const tx=pText(r);if(!tx)continue;setBold(r,remain>0);remain-=tx.length}}
function boldClassLabels(p){for(const r of allRuns(p)){if(protectedRun(r))continue;const tx=pText(r);setBold(r,/中图分类号[:：]|文献标识码[:：]|文章编号[:：]/.test(tx))}}
function styleCrossRefs(p,stats,opt){if(!opt.refs)return;for(const r of Array.from(allRuns(p)))stats.citations+=decorateMatchesInRun(r,/\[\d+(?:[-–—,，]\d+)*\]/g,x=>{setSuper(x,true);if(opt.markup)setHighlight(x,'yellow')});if(opt.markup){for(const r of Array.from(allRuns(p)))stats.crossrefs+=decorateMatchesInRun(r,/(?:如|见|参见)?(?:图|表)\s*\d+(?:\s*所示)?/g,x=>setHighlight(x,'yellow'))}}
function applyParaGeometry(p,k){setSpacing(p);clearTabs(p);if(['title','author','affil','enTitle','enAuthor','enAffil','caption'].includes(k)){setJc(p,'center');setInd(p,null)}else if(k==='empty'){}else{setJc(p,'left');setInd(p,420)}}
function clearParagraphBold(p){for(const r of allRuns(p))if(pText(r)&&!protectedRun(r))setBold(r,false)}
function applyFormatting(body,opt,stats){const kinds=classifyBody(body);for(const p of Array.from(body.getElementsByTagNameNS(W,'p'))){setParagraphFontDefault(p);for(const r of allRuns(p)){setFont(r);setBlack(r)}}for(const [p,k] of kinds){if(!p.parentNode)continue;normalizeTextNodes(p,k,stats,opt);applyParaGeometry(p,k);clearParagraphBold(p);if(['title','enTitle'].includes(k))for(const r of allRuns(p))if(pText(r)&&!protectedRun(r))setBold(r,true);
 if(k==='cnAbstract')stylePrefixLabel(p,/^摘\s*要[:：]/);
 if(k==='cnKeywords')stylePrefixLabel(p,/^关键词[:：]/);
 if(k==='enAbstract')stylePrefixLabel(p,/^Abstract[:：]/i);
 if(k==='enKeywords')stylePrefixLabel(p,/^(?:Keywords|Key\s*words)[:：]/i);
 if(k==='date')stylePrefixLabel(p,/^收稿日期[:：]/);
 if(k==='fund')stylePrefixLabel(p,/^基金项目[:：]/);
 if(k==='authorBio')stylePrefixLabel(p,/^作者简介[:：]/);
 if(k==='classLine')boldClassLabels(p);
 if(k==='caption'&&opt.markup)for(const r of allRuns(p))if(pText(r)&&!protectedRun(r))setHighlight(r,'yellow');
 if(k==='caption')stats.captionStyled++;
 if(k==='author'||k==='enAuthor'){for(const r of allRuns(p)){const rp=firstDirect(r,W,'rPr'),v=rp&&firstDirect(rp,W,'vertAlign');if(v&&v.getAttributeNS(W,'val')==='superscript'&&/^\d+$/.test(pText(r).trim())&&opt.markup)setHighlight(r,'yellow')}}
 if(['body','cnAbstract','cnKeywords','authorBio'].includes(k))styleCrossRefs(p,stats,opt)}
 // 最后再清一次显式颜色，确保正文、参考文献、作者简介无红字残留。
 for(const p of Array.from(body.getElementsByTagNameNS(W,'p')))for(const r of allRuns(p))setBlack(r)
}
function attrVal(el,local){return el?el.getAttributeNS(W,local):''}
function isHighlightedRun(r){const rp=firstDirect(r,W,'rPr'),h=rp&&firstDirect(rp,W,'highlight');return !!h&&attrVal(h,'val')==='yellow'}
function collectMentions(body,kind){const out=new Map();for(const p of Array.from(body.childNodes).filter(n=>n.nodeType===1&&n.namespaceURI===W&&n.localName==='p')){const t=visibleText(p);if(!t||isCaptionText(t)||/^参考文献[:：]?/.test(t)||/^作者简介[:：]/.test(t))continue;const re=new RegExp(`${kind}\\s*(\\d+)(?!\\d)`,'g');let m;while((m=re.exec(t))){const num=Number(m[1]);if(!out.has(num))out.set(num,[]);out.get(num).push(t.slice(Math.max(0,m.index-18),Math.min(t.length,m.index+m[0].length+18)))}}return out}
function explicitRedCount(body){let n=0;for(const c of Array.from(body.getElementsByTagNameNS(W,'color'))){const v=(attrVal(c,'val')||'').toUpperCase();if(v&&!['000000','AUTO'].includes(v))n++}return n}
function auditDocument(body,stats,sourceDate){const caps=captionMap(body),figMent=collectMentions(body,'图'),tblMent=collectMentions(body,'表');const figNums=[...caps.fig.keys()].sort((a,b)=>a-b),tblNums=[...caps.tbl.keys()].sort((a,b)=>a-b);const missingFig=figNums.filter(n=>!figMent.has(n)),missingTbl=tblNums.filter(n=>!tblMent.has(n));const orphanFig=[...figMent.keys()].filter(n=>!caps.fig.has(n)).sort((a,b)=>a-b),orphanTbl=[...tblMent.keys()].filter(n=>!caps.tbl.has(n)).sort((a,b)=>a-b);const ps=Array.from(body.childNodes).filter(n=>n.nodeType===1&&n.namespaceURI===W&&n.localName==='p');const refHead=ps.find(p=>/^参考文献[:：]?\s*$/.test(visibleText(p))),bio=ps.find(p=>/^作者简介[:：]/.test(visibleText(p))),date=ps.find(p=>/^收稿日期[:：]/.test(visibleText(p)));const h1s=ps.filter(p=>/^\d+\s+/.test(visibleText(p)));const lastH1=h1s.length?h1s[h1s.length-1]:null;const finalHeadingOK=!lastH1||/^\d+\s+结\s+论\s*$/.test(visibleText(lastH1));const bodyText=ps.filter(p=>p!==refHead&&(!refHead||ps.indexOf(p)<ps.indexOf(refHead))).map(visibleText).join('\n');const adjacent=[];for(const m of bodyText.matchAll(/\[(\d+)\]\s*\[(\d+)\]/g))if(Number(m[2])===Number(m[1])+1)adjacent.push(m[0]);const actualDrawings=body.getElementsByTagNameNS(W,'drawing').length;let bibliographyCount=0;if(refHead){const a=ps.indexOf(refHead),b=bio?ps.indexOf(bio):ps.length;for(let i=a+1;i<b;i++)if(visibleText(ps[i]))bibliographyCount++}const cites=[];for(const m of bodyText.matchAll(/\[(\d+)(?:[-–—](\d+))?\]/g)){let a=Number(m[1]),b=m[2]?Number(m[2]):a;if(b>=a&&b-a<100)for(let i=a;i<=b;i++)cites.push(i)}const citedUnique=[...new Set(cites)].sort((a,b)=>a-b),citeMissingBib=citedUnique.filter(n=>bibliographyCount&&n>bibliographyCount);return{figNums,tblNums,missingFig,missingTbl,orphanFig,orphanTbl,actualDrawings,redRemaining:explicitRedCount(body),adjacent,refHeadOK:!!refHead&&visibleText(refHead)==='参考文献：',bioOK:!!bio&&attrVal(firstDirect(ensurePPr(bio),W,'jc'),'val')!=='center',finalHeadingOK,finalHeading:lastH1?visibleText(lastH1):'',sourceDate,outputDate:date?visibleText(date):'',datePreserved:!sourceDate||!date||visibleText(date)===sourceDate,bibliographyCount,citedUnique,citeMissingBib,citationMerged:stats.citationMerged}}
