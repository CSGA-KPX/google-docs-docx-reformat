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


let printPublicProperties (obj: obj) =
    let objType = obj.GetType()

    let properties =
        objType.GetProperties(BindingFlags.Public ||| BindingFlags.Instance)

    for prop in properties do
        let value = prop.GetValue(obj, null)
        printfn "Property: %s, Value: %A" prop.Name value

let doc =
    let testFile = "sample/肾上腺衰老：标志与后果_.docx"
    let fileData = File.ReadAllBytes(testFile)
    let ms = new MemoryStream()
    ms.Write(fileData)

    DocX.Load(ms)


let inline Pause() =
    Console.WriteLine("回车继续")
    Console.ReadLine() |> ignore

let inline centiMeterToPoints (cm : float32) = 
    cm * 28.3464566929f

let inline lineSpacingConvert (lh : float32) =
    lh * 240.0f / 20.0f

doc.MarginTop <- centiMeterToPoints 1.2f
doc.MarginBottom <- centiMeterToPoints 1.2f
doc.MarginLeft <- centiMeterToPoints 1.7f
doc.MarginRight <- centiMeterToPoints 0.8f
doc.MirrorMargins <- true

for f in doc.ParagraphFormattings do
    printfn $"{f.StyleId}"
    printPublicProperties f
    printfn "//"
//Console.ReadLine() |> ignore

let nsManager = XmlNamespaceManager(NameTable())
nsManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")


for p in doc.Paragraphs |> Seq.toArray do
    
    if String.IsNullOrEmpty(p.Text) then
        p.Remove(false, RemoveParagraphFlags.None)
    else
        if p.ParentContainer = ContainerType.Body && (not <| p.StyleId.Contains("Heading")) then
            p.IndentationFirstLine <- 24.0f

        let style = p.StyleId

        if style.Contains("Heading") then
            // remove w:before in w:p/w:pPr/w:spaceing
            let ret = p.Xml.XPathSelectElement("w:pPr/w:spacing", nsManager)

            let before = 
                XName.Get("before", ret.Name.NamespaceName)
                |> ret.Attribute

            if not <| isNull before then
                before.Remove()

            printfn $"{p.Xml.ToString()}"
            p.LineSpacing <- lineSpacingConvert 1.5f
            printfn $"{p.Xml.ToString()}"

            
            //Pause()


        let spaceing = (p.LineSpacing, p.LineSpacingAfter, p.LineSpacingBefore)
        //p.SpacingAfter(0) |> ignore
        printfn $"{style} %A{spaceing} {p.Text}"
        
        ()

doc.SaveAs("test_docx.docx")
