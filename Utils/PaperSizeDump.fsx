#r "nuget: System.Drawing.Common"

open System
open System.IO

open System.Drawing.Printing


let textOut = Text.StringBuilder()

textOut.AppendLine("module PaperSizes").AppendLine().AppendLine() |> ignore

(fun () ->
    let startNeedle = "type PaperSize ="
    let endNeedle = "}"
    let f = File.ReadAllText(__SOURCE_FILE__)

    let startIdx = f.LastIndexOf(startNeedle)
    let endIdx = f.IndexOf(endNeedle, startIdx)

    textOut.AppendLine(f.[startIdx..endIdx]).AppendLine()) ()

type PaperSize =
    { Name: string
      WidthPt: int
      HeightPt: int }

let ps = PrinterSettings()

for s in ps.PaperSizes do
    match s.Kind with
    | PaperKind.Custom -> ()
    | kind ->
        let uKind = kind.ToString()

        let lKind =
            let t = kind.ToString().ToCharArray()
            t.[0] <- Char.ToLowerInvariant(t.[0])
            String(t)

        let widthPt = s.Width * 72 / 100
        let heightPt = s.Height * 72 / 100

        let test =
            $"let %O{uKind} = {{ Name = \"{s.PaperName}\"; WidthPt = {widthPt}; HeightPt = {heightPt} }}"

        textOut.AppendLine(test) |> ignore

    printfn $"{s.PaperName} {s.Kind} {s.Width}x{s.Height}"

File.WriteAllText("PaperSize.fsx", textOut.ToString())
