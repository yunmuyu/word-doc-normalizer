class UIController{
  constructor(){
    this.selectedFile=null;this.lastResult=null;this.lastUrl=null;
    this.fileInput=$('file');this.drop=$('drop');this.runBtn=$('run');
    this.bind();
  }
  bind(){
    this.drop.onclick=e=>{if(e.target!==this.fileInput)this.fileInput.click()};
    this.fileInput.onchange=()=>this.pick(this.fileInput.files[0]);
    for(const ev of ['dragenter','dragover'])this.drop.addEventListener(ev,e=>{e.preventDefault();this.drop.classList.add('drag')});
    for(const ev of ['dragleave','drop'])this.drop.addEventListener(ev,e=>{e.preventDefault();this.drop.classList.remove('drag')});
    this.drop.addEventListener('drop',e=>this.pick(e.dataTransfer.files[0]));
    $('reset').onclick=()=>this.reset();
    this.runBtn.onclick=()=>this.run();
    $('download').onclick=()=>this.download();
  }
  showStatus(msg,type=''){const s=$('status');s.textContent=msg;s.className='status show '+type}
  reset(){
    this.selectedFile=null;this.lastResult=null;if(this.lastUrl){URL.revokeObjectURL(this.lastUrl);this.lastUrl=null}
    this.fileInput.value='';$('fileLabel').textContent='点击或拖入原版 .docx';this.runBtn.disabled=true;$('download').disabled=true;
    $('status').className='status';$('log').className='log';$('audit').className='audit';
  }
  pick(f){
    if(!f)return;if(!/\.docx$/i.test(f.name)){this.showStatus('只支持 .docx 文件。','err');return}
    this.selectedFile=f;this.lastResult=null;if(this.lastUrl){URL.revokeObjectURL(this.lastUrl);this.lastUrl=null}
    $('download').disabled=true;$('fileLabel').textContent=f.name;this.runBtn.disabled=false;this.showStatus('已载入：'+f.name,'');
  }
  options(){return{layout:$('optLayout').checked,figures:$('optFigures').checked,punct:$('optPunct').checked,refs:$('optRefs').checked,markup:$('optMarkup').checked}}
  async run(){
    if(!this.selectedFile)return;this.runBtn.disabled=true;$('download').disabled=true;
    this.showStatus('V5 正在按分阶段规则引擎处理，并执行结构完整性检查……','');$('log').className='log';$('audit').className='audit';
    try{
      const out=await processDocx(this.selectedFile,this.options());this.lastResult=out;
      if(this.lastUrl)URL.revokeObjectURL(this.lastUrl);this.lastUrl=URL.createObjectURL(out.blob);$('download').disabled=false;
      this.showStatus(out.audit.integrity.severe?'处理完成，但完整性守卫发现高风险变化；请先看审核清单。':'处理完成。请先看变更记录与审核清单，再下载 Word。',out.audit.integrity.severe?'err':'ok');
      renderLog(out.stats,out.audit,out.changes);
    }catch(e){console.error(e);this.showStatus('处理失败：'+(e.message||e),'err')}
    finally{this.runBtn.disabled=false}
  }
  download(){if(!this.lastResult||!this.lastUrl)return;const a=document.createElement('a');a.href=this.lastUrl;a.download='规范化V5_'+this.selectedFile.name;a.click()}
}

function numsText(a,prefix){return a.length?a.map(n=>prefix+n).join('、'):'无'}
function deltaText(d){const pairs=Object.entries(d).filter(([,v])=>v);return pairs.length?pairs.map(([k,v])=>`${k}${v>0?'+':''}${v}`).join('，'):'未产生计数型变化'}
function renderLog(s,a,changes){
  $('log').className='log show';
  $('metrics').innerHTML=`<div class=metric><strong>${s.captionStyled}</strong><span>图表题标黄</span></div><div class=metric><strong>${s.citations}</strong><span>引文上标/标黄</span></div><div class=metric><strong>${s.citationMerged}</strong><span>连续引文合并</span></div><div class=metric><strong>${s.anchorsFixed}</strong><span>浮动图安全归位</span></div><div class=metric><strong>${s.formulas}</strong><span>公式编号修复</span></div><div class=metric><strong>${s.numberingFixed}</strong><span>自动编号前括号修正</span></div><div class=metric><strong>${a.redRemaining}</strong><span>剩余显式彩色文字</span></div>`;
  $('details').innerHTML=`<b>自动规则摘要 · V5 分阶段引擎</b><br>• 收稿日期连接符保持原样：${a.outputDate||'未发现收稿日期'}<br>• 普通数值范围、单位数值、参考文献页码范围：不自动改写。<br>• 末个一级标题固定为“N  结  论”；“参考文献：”固定标题；作者简介左对齐、首行缩进，只有标签加粗。<br>• 自动编号“（1）/（2）…”统一为“1）/2）…”，连续引文自动合并。<br>• 黄色只作为图表题、正文图表引用和参考文献引文校样标记。<br>• 图题：${numsText(a.figNums,'图')}；表题：${numsText(a.tblNums,'表')}；参考文献条目 ${a.bibliographyCount} 条。`;
  $('changes').innerHTML='<b>本次规则执行记录</b>'+changes.map(c=>`<div class="chg"><b>${c.label}</b><small>${deltaText(c.delta)}</small></div>`).join('');
  const warnings=[];
  if(a.missingFig.length)warnings.push(`有图题但正文没有出现对应图号：${numsText(a.missingFig,'图')}。请确认是否需要补“如图 n 所示/见图 n”。`);
  if(a.missingTbl.length)warnings.push(`有表题但正文没有出现对应表号：${numsText(a.missingTbl,'表')}。请确认是否需要补正文引用。`);
  if(a.orphanFig.length)warnings.push(`正文提到了图号但没找到对应图题：${numsText(a.orphanFig,'图')}。`);
  if(a.orphanTbl.length)warnings.push(`正文提到了表号但没找到对应表题：${numsText(a.orphanTbl,'表')}。`);
  if(a.adjacent.length)warnings.push(`仍发现相邻引文 ${a.adjacent.join('、')}，需要人工确认是否应合并。`);
  if(a.citeMissingBib.length)warnings.push(`正文引用了但参考文献列表中未找到：${a.citeMissingBib.map(n=>'['+n+']').join('、')}。`);
  if(!a.datePreserved)warnings.push(`收稿日期发生变化：原“${a.sourceDate}”→输出“${a.outputDate}”，请勿使用该结果并反馈。`);
  if(a.redRemaining)warnings.push(`仍有 ${a.redRemaining} 处显式非黑色文字，需人工检查。`);
  if(!a.finalHeadingOK)warnings.push(`最后一个一级标题当前为“${a.finalHeading||'未识别'}”，按样刊固定格式应为“N  结  论”。`);
  for(const w of a.integrity.warnings) warnings.push(`结构完整性守卫：${w}`);
  $('audit').className='audit show';
  $('autoChecks').innerHTML=`<div class="check ${a.integrity.severe?'bad':'ok'}">${a.integrity.severe?'!':'✓'} DOCX 结构完整性守卫</div><div class="guard ${a.integrity.warnings.length?'bad':'ok'}">${a.integrity.warnings.length?a.integrity.warnings.join('<br>'):'受保护对象数量未减少：'+a.integrity.ok.join('；')}</div><div class="check ${a.redRemaining===0?'ok':'bad'}">${a.redRemaining===0?'✓':'!'} 红色/彩色正文清理</div><div class="check ${a.datePreserved?'ok':'bad'}">${a.datePreserved?'✓':'!'} 收稿日期未被改写</div><div class="check ${a.finalHeadingOK?'ok':'bad'}">${a.finalHeadingOK?'✓':'!'} 固定末级一级标题“N  结  论”</div><div class="check ${a.refHeadOK?'ok':'bad'}">${a.refHeadOK?'✓':'!'} 固定标题“参考文献：”</div><div class="check ${a.bioOK?'ok':'bad'}">${a.bioOK?'✓':'!'} 作者简介非居中</div><div class="check ${a.adjacent.length===0?'ok':'bad'}">${a.adjacent.length===0?'✓':'!'} 相邻连续引文检查</div>`;
  $('warnings').innerHTML=warnings.length?warnings.map(x=>`<div class="issue">⚠ ${x}</div>`).join(''):'<div class="check ok">✓ 未发现图表引用缺失、孤立图表号、参考文献序号或结构完整性明显异常。</div>';
  $('manual').innerHTML=`<label><input type="checkbox"> 我已核对“图 n/表 n”在正文中的引用位置是否需要补写</label><label><input type="checkbox"> 我已核对图号、表号与实际图表顺序</label><label><input type="checkbox"> 我已核对公式、单位和数值范围（工具不自动改普通数值范围）</label><label><input type="checkbox"> 我已核对参考文献内容本身及作者简介内容</label>`;
}

new UIController();
