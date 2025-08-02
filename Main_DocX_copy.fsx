#r "nuget: DocX, 4.0.25105.5786"

open System
open System.Collections.Generic
open System.IO
open System.Reflection
open System.Xml
open System.Xml.Linq
open System.Xml.XPath

open Xceed
open Xceed.Document.NET
open Xceed.Words.NET

let nsManager = XmlNamespaceManager(NameTable())
nsManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

let printPublicProperties (obj: obj) =
    let objType = obj.GetType()

    let properties =
        objType.GetProperties(BindingFlags.Public ||| BindingFlags.Instance)

    for prop in properties do
        let value = prop.GetValue(obj, null)
        printfn "Property: %s, Value: %A" prop.Name value

let inline Pause() =
    Console.WriteLine("回车继续")
    Console.ReadLine() |> ignore

let doc =
    let testFile = "sample/肾上腺衰老：标志与后果_.docx"
    let fileData = File.ReadAllBytes(testFile)
    let ms = new MemoryStream()
    ms.Write(fileData)
    DocX.Load(ms)

let inline centiMeterToPoints (cm : float32) = 
    cm * 28.3464566929f

doc.MarginTop <- centiMeterToPoints 1.2f
doc.MarginBottom <- centiMeterToPoints 1.2f
doc.MarginLeft <- centiMeterToPoints 1.7f
doc.MarginRight <- centiMeterToPoints 0.8f
doc.MirrorMargins <- true

let applyGrid (doc : DocX) = 
    let sectPr = doc.Xml.XPathSelectElement("//w:sectPr", nsManager)
    let w = sectPr.Name.NamespaceName
    let docGrid = XElement(XName.Get("docGrid", w), XAttribute(XName.Get("type", w), "lines"), XAttribute(XName.Get("val", w), "309"))
    sectPr.Add(docGrid)

applyGrid (doc)

let changeStylesDirty (doc : DocX) = 
    let styles : XDocument = 
        let t = doc.GetType()
        let fieldInfo = t.GetField("_styles", BindingFlags.NonPublic ||| BindingFlags.Instance)

        if isNull fieldInfo then
            failwithf "未找到字段 _styles"
        else
            let value = fieldInfo.GetValue(doc)
            match value with
            | :? XDocument as xdoc ->
                xdoc
            | _ ->
                failwithf "字段类型不是 XDocument"

    let nsManager = XmlNamespaceManager(NameTable())
    nsManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

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
        styles.XPathSelectElements("//w:style[contains(@w:styleId, 'Heading') or contains(@w:styleId, 'itle')]", nsManager)
        |> Seq.toArray

    for styleNode in styleNodes do
        let spacingNode = styleNode.XPathSelectElement(".//w:spacing", nsManager)
        spacingNode.Remove()

changeStylesDirty(doc)

let inline lineSpacingConvert (lh : float32) =
    lh * 240.0f / 20.0f

let styleDoc =
    [
        for f in doc.ParagraphFormattings do
            f.StyleId, f
    ] |> readOnlyDict

for t in doc.Tables |> Seq.toArray do
    // 所有表格尽量居中
    t.Alignment <- Alignment.center

for p in doc.Paragraphs |> Seq.toArray do
    
    if String.IsNullOrEmpty(p.Text) then
        p.Remove(false, RemoveParagraphFlags.None)
    else
        let ret2 = p.Xml.XPathSelectElements("//w:rFonts", nsManager)

        for attr in ret2.Attributes() |> Seq.toArray do
            match attr.Name.LocalName with
            | "ascii" -> attr.SetValue("Arial")
            | "eastAsia" -> attr.SetValue("宋体")
            | "hAnsi" -> attr.SetValue("Arial")
            | "cs" -> attr.Remove()
            | unk -> printfn $"notmatch attr : {unk}"

        if p.ParentContainer = ContainerType.Body && (not <| (p.StyleId.Contains("Heading") || p.IsListItem)) then
            //printfn "%A/%A = %s" p.IsListItem p.ListItemType p.Text
            p.IndentationFirstLine <- 24.0f

        if p.StyleId.Contains("Heading") then
            
            // 因为找不到让Word垂直居中的办法
            // 为了避免行距不好看，在前面添加一个空行
            let fontSize = styleDoc.[p.StyleId].Size
            let singleSizeSpacing = 
                fontSize.Value * (1.5 - 0.5) / 2.0
                |> float32

            p.LineSpacingAfter <- singleSizeSpacing
            p.LineSpacingBefore <- singleSizeSpacing

            //p.LineSpacing <- lineSpacingConvert 1.5f

            // remove w:before in w:p/w:pPr/w:spaceing
            let ret = p.Xml.XPathSelectElement("w:pPr/w:spacing", nsManager)

            let before = 
                XName.Get("before", ret.Name.NamespaceName)
                |> ret.Attribute

            if not <| isNull before then
                ()//before.Remove()

        ()

doc.SaveAs("test_docx_copy.docx")
