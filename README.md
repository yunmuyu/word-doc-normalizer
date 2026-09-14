# word-doc-normalizer

面向科技期刊编辑流程的 DOCX 本地规范化与质检工具。

## 当前版本

V5（重构核心）

- 文档模型 → 分阶段规则引擎 → 完整性守卫 → 质检引擎 → 输出
- 浏览器本地处理 DOCX，稿件内容不会上传到服务器
- 自动执行已确认的确定性格式规则
- 对图表正文引用、公式、数值范围、参考文献等不宜自动决定的项目给出人工确认提示
- 保护图片、表格、公式、MathType/OLE、域、书签等 OOXML 结构

## 使用

打开仓库根目录的 `index.html`。启用 GitHub Pages 后，可直接通过固定网页入口使用。

## 源码结构

```text
index.html
styles.css
src/
  docx-core.js      # OOXML/DOCX 结构操作
  format-rules.js   # 确定性格式规则与质检规则
  engine.js         # RuleEngine / IntegrityGuard / AuditEngine
  ui.js             # 浏览器交互

docs/
  architecture.md
```

## 隐私与测试数据

仓库只保存工具源码、规则与说明文档。未发表稿件、作者原稿、正确版 DOCX 和真实黄金回归样本不提交到公开仓库；这些测试材料仅用于本地回归。

## 依赖

当前网页从 jsDelivr 加载 JSZip 3.10.1；DOCX 文件本身仍只在浏览器本地内存中处理。
