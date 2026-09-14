class UIController{
  constructor(){this.selectedFile=null;this.lastResult=null;this.lastUrl=null;this.fileInput=$('file');this.drop=$('drop');this.runBtn=$('run');this.bind()}
  bind(){
    this.drop.onclick=e=>{if(e.target!==this.fileInput)this.fileInput.click()};this.fileInput.onchange=()=>this.pick(this.fileInput.files[0]);
    for(const ev of ['dragenter','dragover'])this.drop.addEventListener(ev,e=>{e.preventDefault();this.drop.classList.add('drag')});
    for(const ev of ['dragleave','drop'])this.drop.addEventListener(ev,e=>{e.preventDefault();this.drop.classList.remove('drag')});
    this.drop.addEventListener('drop',e=>this.pick(e.dataTransfer.files[0]));$('reset').onclick=()=>this.reset();this.runBtn.onclick=()=>this.run();$('download').onclick=()=>this.download();
  }
  showStatus(msg,type=''){const s=$('status');s.textContent=msg;s.className='status show '+type}
  reset(){this.selectedFile=null;this.lastResult=null;if(this.lastUrl){URL.revokeObjectURL(this.lastUrl);this.lastUrl=null}this.fileInput.value='';$('fileLabel').textContent='点击或拖入原版 .docx';this.runBtn.disabled=true;$('download').disabled=true;$('status').className='status';$('log').className='log';$('audit').className='audit'}
  pick(f){if(!f)return;if(!/\.docx$/i.test(f.name)){this.showStatus('只支持 .docx 文件。','err');return}this.selectedFile=f;this.lastResult=null;if(this.lastUrl){URL.revokeObjectURL(this.lastUrl);this.lastUrl=null}$('download').disabled=true;$('fileLabel').textContent=f.name;this.runBtn.disabled=false;this.showStatus('已载入：'+f.name,'')}
  options(){return{layout:$('optLayout').checked,figures:$('optFigures').checked,punct:$('optPunct').checked,refs:$('optRefs').checked,markup:$('optMarkup').checked}}
  async run(){
    if(!this.selectedFile)return;this.runBtn.disabled=true;$('download').disabled=true;this.showStatus('V5.1 正在处理，并执行 OOXML 结构与输出包双重自检……','');$('log').className='log';$('audit').className='audit';
    try{
      const out=await processDocx(this.selectedFile,this.options());this.lastResult=out;if(this.lastUrl)URL.revokeObjectURL(this.lastUrl);this.lastUrl=URL.createObjectURL(out.blob);
      const safe=out.audit.integrity.warnings.length===0&&out.audit.outputValidation&&out.audit.outputValidation.ok;$('download').disabled=!safe;
      this.showStatus(safe?'处理完成，结构完整性与输出 DOCX 自检均通过。可以下载后继续在 Word 中编辑。':'处理完成，但自检未通过，已阻止下载。',safe?'ok':'err');renderLog(out.stats,out.audit,out.changes);
    }catch(e){console.error(e);this.showStatus('处理失败 / 已阻止导出：'+(e.message||e),'err')}
    finally{this.runBtn.disabled=false}
  }
  download(){if(!this.lastResult||!this.lastUrl||$('download').disabled)return;const a=document.createElement('a');a.href=this.lastUrl;a.download='规范化V5.1_'+this.selectedFile.name;a.click()}
}

function numsText(a,prefix){return a.length?a.map(n=>prefix+n).join('、'):'无'}
function deltaText(d){const pairs=Object.entries(d).filter(([,v])=>v);return pairs.length?pairs.map(([k,v])=>`${k}${v>0?'+':''}${v}`).join('，'):'未产生计数型变化'}
function renderLog(s,a,changes){
  $('log').className='log show';
  $('metrics').innerHTML=`<div class=metric><strong>${s.captionStyled}</strong><span>图表题标黄</span></div><div class=metric><strong>${s.citations}</strong><span>引文上标/标黄</span></div><div class=metric><strong>${s.citationMerged}</strong><span>引文序列规范</span></div><div class=metric><strong>${s.listsFlattened}</strong><span>自动列表转纯文字</span></div><div class=metric><strong>${s.microHeadingsMerged}</strong><span>小标题并入正文</span></div><div class=metric><strong>${s.anchorsFixed}</strong><span>浮动图安全归位</span></div><div class=metric><strong>${a.redRemaining}</strong><span>剩余显式彩色文字</span></div>`;
  $('details').innerHTML=`<b>自动规则摘要 · V5.1 编辑安全版</b><br>• 收稿日期连接符保持原样：${a.outputDate||'未发现收稿日期'}。<br>• [1][2]→[1-2]；[1][10]→[1，10]；非连续序号用中文逗号，连续序号用短横线。<br>• 中文语境中的英文直双引号自动改为中文双引号；如 150m、2ms 自动改为 150 m、2 ms。<br>• 多位数字按三位分节，如 1234.1588→1 234.158 8；日期、年份、邮编、基金编号、参考文献年份/页码不参与该规则。<br>• Word 自动列表改为纯文字编号；参考文献条目输出为“[1] 作者……”普通段落。<br>• “1）小标题”与紧随其后的说明段合并为“1）小标题。正文……”。<br>• 末个一级标题固定为“N  结  论”；“参考文献：”固定且加粗；作者简介左对齐、首行缩进，仅标签加粗。<br>• 黄色只用于图表题、正文图表引用和正文参考文献引文校样。`;
  $('changes').innerHTML='<b>本次规则执行记录</b>'+changes.map(c=>`<div class="chg"><b>${c.label}</b><small>${deltaText(c.delta)}</small></div>`).join('');
  const warnings=[];
  if(a.missingFig.length)warnings.push(`有图题但正文没有出现对应图号：${numsText(a.missingFig,'图')}。请确认是否需要补“如图 n 所示/见图 n”。`);
  if(a.missingTbl.length)warnings.push(`有表题但正文没有出现对应表号：${numsText(a.missingTbl,'表')}。请确认是否需要补正文引用。`);
  if(a.orphanFig.length)warnings.push(`正文提到了图号但没找到对应图题：${numsText(a.orphanFig,'图')}。`);
  if(a.orphanTbl.length)warnings.push(`正文提到了表号但没找到对应表题：${numsText(a.orphanTbl,'表')}。`);
  if(a.adjacent.length)warnings.push(`仍发现拆开的相邻引文 ${a.adjacent.join('、')}，需要人工确认。`);
  if(a.citeMissingBib.length)warnings.push(`正文引用了但参考文献列表中未找到：${a.citeMissingBib.map(n=>'['+n+']').join('、')}。`);
  if(!a.datePreserved)warnings.push(`收稿日期发生变化：原“${a.sourceDate}”→输出“${a.outputDate}”，该结果已视为异常。`);
  if(a.redRemaining)warnings.push(`仍有 ${a.redRemaining} 处显式非黑色文字，需人工检查。`);
  if(!a.finalHeadingOK)warnings.push(`最后一个一级标题当前为“${a.finalHeading||'未识别'}”，按固定格式应为“N  结  论”。`);
  if(a.remainingNumPr)warnings.push(`正文仍有 ${a.remainingNumPr} 个 Word 自动列表段落，需人工确认是否应转为纯文字。`);
  if(a.unitSpacingIssues.length)warnings.push(`仍发现疑似数字与单位未空格：${a.unitSpacingIssues.join('；')}`);
  if(a.asciiChineseQuotes.length)warnings.push(`仍发现中文语境中的英文直双引号：${a.asciiChineseQuotes.join('；')}`);
  for(const w of a.integrity.warnings)warnings.push(`结构完整性守卫：${w}`);if(a.outputValidation&&!a.outputValidation.ok)for(const w of a.outputValidation.errors)warnings.push(`输出 DOCX 自检：${w}`);
  const packageOK=a.outputValidation&&a.outputValidation.ok;
  $('audit').className='audit show';
  $('autoChecks').innerHTML=`<div class="check ${a.integrity.warnings.length?'bad':'ok'}">${a.integrity.warnings.length?'!':'✓'} OOXML 受保护对象数量检查</div><div class="guard ${a.integrity.warnings.length?'bad':'ok'}">${a.integrity.warnings.length?a.integrity.warnings.join('<br>'):'受保护对象数量未减少：'+a.integrity.ok.join('；')}</div><div class="check ${packageOK?'ok':'bad'}">${packageOK?'✓':'!'} 输出 DOCX 重新打开与媒体/嵌入对象/关系文件自检</div><div class="guard ${packageOK?'ok':'bad'}">${packageOK?a.outputValidation.checks.join('；'):(a.outputValidation?a.outputValidation.errors.join('；'):'未执行')}</div><div class="check ${a.redRemaining===0?'ok':'bad'}">${a.redRemaining===0?'✓':'!'} 红色/彩色正文清理</div><div class="check ${a.datePreserved?'ok':'bad'}">${a.datePreserved?'✓':'!'} 收稿日期完全保留</div><div class="check ${a.refHeadOK&&a.refHeadBold?'ok':'bad'}">${a.refHeadOK&&a.refHeadBold?'✓':'!'} “参考文献：”固定标题且加粗</div><div class="check ${a.remainingNumPr===0?'ok':'bad'}">${a.remainingNumPr===0?'✓':'!'} 正文自动列表已转纯文字</div><div class="check ${a.unitSpacingIssues.length===0?'ok':'bad'}">${a.unitSpacingIssues.length===0?'✓':'!'} 数字与单位空格检查</div><div class="check ${a.asciiChineseQuotes.length===0?'ok':'bad'}">${a.asciiChineseQuotes.length===0?'✓':'!'} 中文双引号检查</div>`;
  $('warnings').innerHTML=warnings.length?warnings.map(x=>`<div class="issue">⚠ ${x}</div>`).join(''):'<div class="check ok">✓ 当前自动检查未发现明显异常。</div>';
  $('manual').innerHTML=`<label><input type="checkbox"> 我已核对“图 n/表 n”正文引用是否需要补写</label><label><input type="checkbox"> 我已核对图号、表号与实际图表顺序</label><label><input type="checkbox"> 我已核对公式、数值范围与特殊单位表达（数值范围本身不自动改连接符）</label><label><input type="checkbox"> 我已核对参考文献内容、作者简介内容和文中技术数据</label>`;
}
new UIController();
