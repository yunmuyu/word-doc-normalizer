const APP_VERSION='5.0.0-refactor';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const M='http://schemas.openxmlformats.org/officeDocument/2006/math';

const Stats=()=>({sections:0,figures:0,anchorsFixed:0,tables:0,captions:0,captionStyled:0,formulas:0,headings:0,citations:0,crossrefs:0,textFixes:0,removed:0,movedMeta:0,citationMerged:0,numberingFixed:0,finalHeadingFixed:0});
const $=id=>document.getElementById(id);

class ChangeRecorder{
  constructor(){this.items=[]}
  record(ruleId,label,before,after,note=''){
    const delta={};
    for(const k of Object.keys(after)) if(typeof after[k]==='number'&&after[k]!==before[k]) delta[k]=after[k]-before[k];
    this.items.push({ruleId,label,delta,note});
  }
}

class DocumentSnapshot{
  static take(body){
    const count=(ns,name)=>body.getElementsByTagNameNS(ns,name).length;
    return {
      paragraphs:count(W,'p'), tables:count(W,'tbl'), drawings:count(W,'drawing'),
      math:count(M,'oMath')+count(M,'oMathPara'), objects:count(W,'object'), pict:count(W,'pict'),
      hyperlinks:count(W,'hyperlink'), bookmarks:count(W,'bookmarkStart'), fields:count(W,'fldChar'), instrText:count(W,'instrText')
    };
  }
}

class IntegrityGuard{
  static compare(before,after){
    const warnings=[],ok=[];
    const protectedKeys=['tables','drawings','math','objects','pict','hyperlinks','bookmarks','fields','instrText'];
    for(const k of protectedKeys){
      if(after[k]<before[k]) warnings.push(`${k}: ${before[k]} → ${after[k]}（数量减少）`);
      else ok.push(`${k}: ${after[k]}`);
    }
    return {before,after,warnings,ok,severe:warnings.some(x=>/math|objects|pict|fields|instrText/.test(x))};
  }
}

class DocxContext{
  constructor(file,options,zip,doc,body){
    this.file=file; this.options=options; this.zip=zip; this.doc=doc; this.body=body;
    this.stats=Stats(); this.changes=new ChangeRecorder();
    this.sourceDate=this.findDate();
    this.before=DocumentSnapshot.take(body); this.guard=null; this.audit=null;
  }
  findDate(){
    const p=Array.from(this.body.childNodes).find(n=>n.nodeType===1&&n.namespaceURI===W&&n.localName==='p'&&/^收稿日期[:：]/.test(visibleText(n)));
    return p?visibleText(p):'';
  }
  snapshotStats(){return {...this.stats}}
}

class RuleEngine{
  constructor(rules){this.rules=rules}
  async run(ctx){
    for(const phase of ['package','structure','placement','format','post']){
      for(const rule of this.rules.filter(r=>r.phase===phase)){
        if(rule.enabled&&!rule.enabled(ctx)) continue;
        const before=ctx.snapshotStats();
        await rule.apply(ctx);
        ctx.changes.record(rule.id,rule.label,before,ctx.snapshotStats(),rule.note||'');
      }
    }
  }
}

class AuditEngine{
  run(ctx){
    ctx.stats.headings=Array.from(ctx.body.childNodes).filter(n=>n.nodeType===1&&n.namespaceURI===W&&n.localName==='p'&&/^(?:\d+(?:\.\d+){0,2})\s+/.test(visibleText(n))).length;
    const after=DocumentSnapshot.take(ctx.body);
    ctx.guard=IntegrityGuard.compare(ctx.before,after);
    ctx.audit=auditDocument(ctx.body,ctx.stats,ctx.sourceDate);
    ctx.audit.integrity=ctx.guard;
    return ctx.audit;
  }
}

const RULES=[
  {id:'layout.sections',label:'页面与节版式',phase:'package',enabled:c=>c.options.layout,apply:c=>normalizeSections(c.doc,c.stats)},
  {id:'math.repair',label:'MathType/公式编号保护',phase:'structure',apply:c=>repairMathTypeNumbers(c.body,c.stats)},
  {id:'structure.cleanup',label:'结构清理与元数据归位',phase:'structure',apply:c=>{cleanSeparators(c.body,c.stats);moveMetadata(c.body,c.stats)}},
  {id:'caption.split',label:'图表题识别与拆分',phase:'structure',apply:c=>splitEmbeddedCaptions(c.body,c.stats)},
  {id:'figures.place',label:'图片安全归位',phase:'placement',enabled:c=>c.options.figures,apply:c=>{let caps=captionMap(c.body);placeFigures(c.body,caps,c.stats)}},
  {id:'tables.place',label:'表格安全归位',phase:'placement',enabled:c=>c.options.figures,apply:c=>{let caps=captionMap(c.body);placeTables(c.body,caps,c.stats)}},
  {id:'format.main',label:'正文与标题确定性格式',phase:'format',apply:c=>applyFormatting(c.body,c.options,c.stats)},
  {id:'color.black',label:'正文显式颜色归黑',phase:'post',apply:c=>forcePartColorsBlack(c.doc)},
  {id:'package.aux',label:'编号/样式等辅助部件清理',phase:'post',apply:async c=>normalizeAuxParts(c.zip,c.stats)}
];

async function loadDocx(file,options){
  const zip=await JSZip.loadAsync(await file.arrayBuffer());
  const f=zip.file('word/document.xml');
  if(!f) throw new Error('这不是有效 DOCX：缺少 word/document.xml');
  const xml=await f.async('string');
  const doc=new DOMParser().parseFromString(xml,'application/xml');
  if(doc.getElementsByTagName('parsererror')[0]) throw new Error('Word XML 解析失败');
  const body=doc.getElementsByTagNameNS(W,'body')[0];
  if(!body) throw new Error('找不到 Word 正文');
  return new DocxContext(file,options,zip,doc,body);
}

async function processDocx(file,options){
  const ctx=await loadDocx(file,options);
  await new RuleEngine(RULES).run(ctx);
  const audit=new AuditEngine().run(ctx);
  ctx.zip.file('word/document.xml',new XMLSerializer().serializeToString(ctx.doc));
  const blob=await ctx.zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',compression:'DEFLATE'});
  return {blob,stats:ctx.stats,audit,changes:ctx.changes.items,version:APP_VERSION};
}
