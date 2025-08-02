# google-docs-docx-reformat
替换Google Docs导出的docx的排版格式

<br />

Google docs导出的docx文件往往使用它们自己的样式、页面规格和字体。如果对排版有要求则需要人工调整，工作量大。

尤其是Gemini Deep Research等工具可以快速生成大量文档的情况下，人工调整排版尤其不现实。

<br />

本脚本利用Xceed DocX工具以及一些脏招（因为DocX不支持直接访问一些Open XML内容）实现了简单的重新格式化，大概可以满足批量处理的需求。

Bug还很多，代码后续再调整罢
