#r "nuget: OfficeIMO.Word, 1.0.6"

open System
open System.Collections.Generic
open System.IO

open OfficeIMO.Word


let doc = 
    let testFile = "sample/肾上腺衰老：标志与后果_.docx"
    let fileData = File.ReadAllBytes(testFile)
    let ms = new MemoryStream()
    ms.Write(fileData)
    WordDocument.Load(ms)


printfn $"This document has {doc.Sections.Count} sections"


// 页面边距

//doc.Margins.Type <- WordMargin.Mirrored
doc.Margins.Top <- Nullable()
doc.Margins.Bottom <- Nullable()
doc.Margins.TopCentimeters <- 1.2
doc.Margins.BottomCentimeters <- 1.2
doc.Margins.LeftCentimeters <- 1.2
doc.Margins.RightCentimeters <- 0.8

doc.Ma


(*
for p in doc.Paragraphs do
    let style = p.Style
    if style.HasValue then
        let s = style.Value
        
        //WordParagraphStyle.OverrideBuiltInStyle
        

        match style.Value with
        | WordParagraphStyles.Custom -> 
            printfn "Custom"
        | _ -> 
            printfn $"Test: resetting {style.Value} to Normal"
            //p.Style <- WordParagraphStyles.Normal
            //p.SetStyleId



    printfn $"test = {p.Text}"

*)




let t = 
    for p in doc.Paragraphs do
        if String.IsNullOrEmpty(p.Text) then p.Remove()

doc.SaveAs("test.docx")