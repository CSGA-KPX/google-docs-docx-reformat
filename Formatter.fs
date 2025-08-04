module Formatter

open System
open System.Collections.Generic
open System.IO
open System.Reflection
open System.Xml
open System.Xml.Linq
open System.Xml.XPath

open Xceed.Document.NET
open Xceed.Words.NET


let inline private centiMeterToPoints (cm: float32) = cm * 28.3464566929f
let inline private lineSpacingConvert (lh: float32) = lh * 240.0f / 20.0f

let private nsManager =
    let ns = XmlNamespaceManager(NameTable())
    ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    ns

let private tryGetParagraphFontSize (doc: DocX) (styles: IReadOnlyDictionary<string, Formatting>) (p: Paragraph) =
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
    | Some(a), Some(_) -> Some a
    | Some(a), None -> Some a
    | None, Some(b) -> Some b
    | None, None when String.IsNullOrWhiteSpace(p.StyleId) -> None
    | None, None ->
        let succ, ret = styles.TryGetValue(p.StyleId)

        if succ && ret.Size.HasValue then
            Some ret.Size.Value
        else
            None

type DocumentReformattingOption() =

    member val PageMarginTopCm = 1.2f with get, set
    member val PageMarginBottomCm = 1.2f with get, set
    member val PageMarginLeftCm = 1.7f with get, set
    member val PageMarginRightCm = 0.8f with get, set
    member val PageMirrorMargins = true with get, set

    member val PageHeader: string option = None with get, set
    member val PageHeaderMargin = 0.5f with get, set
    member val PageFooter: string option = None with get, set
    member val PageFooterMargin = 0.5f with get, set

    /// 文档行网格高度，单位Dxa
    member val DocumentGridHeight: int = 309 with get, set

    member val DefaultFontSize = 22.0 with get, set
    /// 未实现：调整文档中预先设定的字号，不对DefaultFontSize生效
    member val FontSizeAdjust = 0.0 with get, set
    member val FontFamilyAscii = "Arial" with get, set
    member val FontFamilyAsia = "宋体" with get, set


let format (docFile: string) (formatting: DocumentReformattingOption) (savePath: string) =
    let fileData = File.ReadAllBytes(docFile)
    use ms = new MemoryStream()
    ms.Write(fileData)
    use doc = DocX.Load(ms)

    /// 样式库
    let styleDoc =
        [ for f in doc.ParagraphFormattings do
              f.StyleId, f ]
        |> readOnlyDict

    // 调整页边距
    doc.MarginTop <- centiMeterToPoints formatting.PageMarginTopCm
    doc.MarginBottom <- centiMeterToPoints formatting.PageMarginBottomCm
    doc.MarginLeft <- centiMeterToPoints formatting.PageMarginLeftCm
    doc.MarginRight <- centiMeterToPoints formatting.PageMarginRightCm
    doc.MirrorMargins <- formatting.PageMirrorMargins

    // 页眉
    if isNull (doc.Headers.First) && formatting.PageHeader.IsSome then
        doc.AddHeaders()
        doc.MarginHeader <- centiMeterToPoints formatting.PageHeaderMargin
        let p = doc.Headers.Odd.InsertParagraph(formatting.PageHeader.Value)
        p.Alignment <- Alignment.center

    // 页脚
    if isNull (doc.Footers.First) && formatting.PageFooter.IsSome then
        doc.AddFooters()
        doc.MarginFooter <- centiMeterToPoints formatting.PageFooterMargin
        let p = doc.Footers.Odd.InsertParagraph(formatting.PageFooter.Value)
        p.Alignment <- Alignment.center

    // 文档网格
    // 如果不设置的话，文本不会垂直居中，排版会比较奇怪
    let inline applyGrid (doc: DocX) =
        let sectPr = doc.Xml.XPathSelectElement("//w:sectPr", nsManager)
        let w = sectPr.Name.NamespaceName

        let docGrid =
            XElement(
                XName.Get("docGrid", w),
                XAttribute(XName.Get("type", w), "lines"),
                XAttribute(XName.Get("val", w), formatting.DocumentGridHeight)
            )

        sectPr.Add(docGrid)

    applyGrid (doc)

    // 调整默认样式
    // 因为DocX不直接暴露默认样式库，所以只能反射获取然后修改
    let inline changeStylesDirty (doc: DocX) =
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
            | "ascii" -> attr.SetValue(formatting.FontFamilyAscii)
            | "eastAsia" -> attr.SetValue(formatting.FontFamilyAsia)
            | "hAnsi" -> attr.SetValue(formatting.FontFamilyAscii)
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

    // 让所有表格居中
    for t in doc.Tables |> Seq.toArray do
        t.Alignment <- Alignment.center

    // 重新格式化段落
    for p in doc.Paragraphs |> Seq.toArray do
        // 删除空行，Google Docs常见
        if String.IsNullOrEmpty(p.Text) then
            p.Remove(false, RemoveParagraphFlags.None)
        else
            // 替换字体
            let fontElements = p.Xml.XPathSelectElements("//w:rFonts", nsManager)

            for attr in fontElements.Attributes() |> Seq.toArray do
                match attr.Name.LocalName with
                | "ascii" -> attr.SetValue(formatting.FontFamilyAscii)
                | "eastAsia" -> attr.SetValue(formatting.FontFamilyAsia)
                | "hAnsi" -> attr.SetValue(formatting.FontFamilyAscii)
                | "cs" -> attr.Remove()
                | unk -> printfn $"notmatch attr : {unk}"


            // 给正文文本添加首行缩进
            if
                // 正文内容，而不是表格或者文本框里面的
                p.ParentContainer = ContainerType.Body
                // 而且不能是标题或者列表
                && (not <| (p.StyleId.Contains("Heading") || p.IsListItem))
            then
                // 行首缩进，2个中文字符，约4个英语字符
                // 理论上应该乘以2，但我也不知道为什么不乘刚好
                p.IndentationFirstLine <-
                    tryGetParagraphFontSize doc styleDoc p
                    |> Option.defaultValue formatting.DefaultFontSize
                    |> float32

            if p.StyleId.Contains("Heading") then
                // 因为找不到让Word垂直居中的办法，手动设置前后间距
                let fontSize =
                    tryGetParagraphFontSize doc styleDoc p
                    |> Option.defaultValue formatting.DefaultFontSize

                let singleSizeSpacing = fontSize * (1.5 - 0.5) / 2.0 |> float32

                p.LineSpacingAfter <- singleSizeSpacing
                p.LineSpacingBefore <- singleSizeSpacing


    doc.SaveAs(savePath)
