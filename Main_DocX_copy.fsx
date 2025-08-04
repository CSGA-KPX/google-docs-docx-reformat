#r "nuget: DocX, 4.0.25105.5786"

open System
open System.Collections.Generic
open System.IO
open System.Reflection
open System.Xml
open System.Xml.Linq
open System.Xml.XPath

open Xceed.Document.NET
open Xceed.Words.NET


module Debug =
    let printPublicProperties (obj: obj) =
        let objType = obj.GetType()

        let properties =
            objType.GetProperties(BindingFlags.Public ||| BindingFlags.Instance)

        for prop in properties do
            let value = prop.GetValue(obj, null)
            printfn "Property: %s, Value: %A" prop.Name value

    let inline Pause () =
        Console.WriteLine("回车继续")
        Console.ReadLine() |> ignore

let inline centiMeterToPoints (cm: float32) = cm * 28.3464566929f
let inline lineSpacingConvert (lh: float32) = lh * 240.0f / 20.0f

let nsManager = XmlNamespaceManager(NameTable())
nsManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

let doc =
    let testFile = "sample/肾上腺衰老：标志与后果_.docx"
    let fileData = File.ReadAllBytes(testFile)
    let ms = new MemoryStream()
    ms.Write(fileData)
    DocX.Load(ms)

doc.MarginTop <- centiMeterToPoints 1.2f
doc.MarginBottom <- centiMeterToPoints 1.2f
doc.MarginLeft <- centiMeterToPoints 1.7f
doc.MarginRight <- centiMeterToPoints 0.8f
doc.MirrorMargins <- true

if isNull (doc.Footers.First) then
    doc.AddFooters()
    doc.MarginFooter <- centiMeterToPoints 0.5f
    let p = doc.Footers.Odd.InsertParagraph("内容由AI生成，内容仅供参考，请仔细甄别")
    p.Alignment <- Alignment.center

let styleDoc =
    [ for f in doc.ParagraphFormattings do
          f.StyleId, f ]
    |> readOnlyDict

let getParagraphFontSize (p: Paragraph) =
    let defaultFontSize = 22.0
    // 首先找w:rPr
    // 然后找p.StyleId

    let sz =
        let ret = doc.Xml.XPathSelectElement("w:pPr/w:rPr/w:sz", nsManager)

        if not <| isNull ret then
            XName.Get("val", ret.Name.NamespaceName) |> ret.Attribute |> float |> Some
        else
            None

    let szCs =
        let ret = doc.Xml.XPathSelectElement("w:pPr/w:rPr/w:szCs", nsManager)

        if not <| isNull ret then
            XName.Get("val", ret.Name.NamespaceName) |> ret.Attribute |> float |> Some
        else
            None

    match sz, szCs with
    | Some(a), Some(b) when a <> b -> failwithf $"sz szCs冲突：{p.Xml}"
    | Some(a), Some(_) -> a
    | Some(a), None -> a
    | None, Some(b) -> b
    | None, None when String.IsNullOrWhiteSpace(p.StyleId) -> defaultFontSize
    | None, None ->
        let succ, ret = styleDoc.TryGetValue(p.StyleId)

        if succ && ret.Size.HasValue then
            ret.Size.Value
        else
            defaultFontSize

let applyGrid (doc: DocX) =
    let sectPr = doc.Xml.XPathSelectElement("//w:sectPr", nsManager)
    let w = sectPr.Name.NamespaceName

    let docGrid =
        XElement(
            XName.Get("docGrid", w),
            XAttribute(XName.Get("type", w), "lines"),
            XAttribute(XName.Get("val", w), "309")
        )

    sectPr.Add(docGrid)

applyGrid (doc)

let changeStylesDirty (doc: DocX) =
    let styles: XDocument =
        let t = doc.GetType()

        let fieldInfo =
            t.GetField("_styles", BindingFlags.NonPublic ||| BindingFlags.Instance)

        if isNull fieldInfo then
            failwithf "未找到字段 _styles"
        else
            let value = fieldInfo.GetValue(doc)

            match value with
            | :? XDocument as xdoc -> xdoc
            | _ -> failwithf "字段类型不是 XDocument"

    // Dirty-hack 重写默认字体
    // Xceed.DocX不支持多语言字体
    let ret2 = styles.XPathSelectElements("//w:rFonts", nsManager)

    for attr in ret2.Attributes() |> Seq.toArray do
        match attr.Name.LocalName with
        | "ascii" -> attr.SetValue("Arial")
        | "eastAsia" -> attr.SetValue("宋体")
        | "hAnsi" -> attr.SetValue("Arial")
        | "cs" -> attr.Remove()
        | unk -> printfn $"notmatch attr : {unk}"

    // 取消Google预先设置的段间距
    let styleNodes =
        styles.XPathSelectElements(
            "//w:style[contains(@w:styleId, 'Heading') or contains(@w:styleId, 'itle')]",
            nsManager
        )
        |> Seq.toArray

    for styleNode in styleNodes do
        let spacingNode = styleNode.XPathSelectElement(".//w:spacing", nsManager)
        spacingNode.Remove()

changeStylesDirty (doc)

// 所有表格居中
for t in doc.Tables |> Seq.toArray do
    t.Alignment <- Alignment.center

for p in doc.Paragraphs |> Seq.toArray do
    // 删除空行，Google Docs常见
    if String.IsNullOrEmpty(p.Text) then
        p.Remove(false, RemoveParagraphFlags.None)
    else
        let fontElements = p.Xml.XPathSelectElements("//w:rFonts", nsManager)

        for attr in fontElements.Attributes() |> Seq.toArray do
            match attr.Name.LocalName with
            | "ascii" -> attr.SetValue("Arial")
            | "eastAsia" -> attr.SetValue("宋体")
            | "hAnsi" -> attr.SetValue("Arial")
            | "cs" -> attr.Remove()
            | unk -> printfn $"notmatch attr : {unk}"

        if
            // 给正文文本添加首行缩进
            //
            // 正文内容，而不是表格或者文本框里面的
            p.ParentContainer = ContainerType.Body
            // 而且不能是标题或者列表
            && (not <| (p.StyleId.Contains("Heading") || p.IsListItem))
        then
            p.IndentationFirstLine <-
                // 行首缩进，2个中文字符，约4个英语字符
                // 理论上应该乘以2，但我也不知道为什么不乘刚好
                getParagraphFontSize p |> float32

        if p.StyleId.Contains("Heading") then
            // 因为找不到让Word垂直居中的办法，手动设置前后间距
            let fontSize = getParagraphFontSize p
            let singleSizeSpacing = fontSize * (1.5 - 0.5) / 2.0 |> float32

            p.LineSpacingAfter <- singleSizeSpacing
            p.LineSpacingBefore <- singleSizeSpacing

doc.SaveAs("test_docx_copy.docx")
