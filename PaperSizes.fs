module PaperSizes


type PaperSize =
    { Name: string
      WidthPt: int
      HeightPt: int }

let Letter = { Name = "Letter/ANSI A"; WidthPt = 612; HeightPt = 792 }
let Tabloid = { Name = "Tabloid/ANSI B"; WidthPt = 792; HeightPt = 1224 }
let CSheet = { Name = "ANSI C"; WidthPt = 1224; HeightPt = 1584 }
let DSheet = { Name = "ANSI D"; WidthPt = 1584; HeightPt = 2448 }
let ESheet = { Name = "ANSI E"; WidthPt = 2448; HeightPt = 3168 }
let Legal = { Name = "US Legal"; WidthPt = 612; HeightPt = 1008 }
let Statement = { Name = "Statement"; WidthPt = 396; HeightPt = 612 }
let Executive = { Name = "Executive"; WidthPt = 522; HeightPt = 756 }
let A2 = { Name = "A2"; WidthPt = 1190; HeightPt = 1684 }
let A3 = { Name = "A3"; WidthPt = 841; HeightPt = 1190 }
let A4 = { Name = "A4"; WidthPt = 595; HeightPt = 841 }
let A5 = { Name = "A5"; WidthPt = 419; HeightPt = 595 }
let IsoB4 = { Name = "B4 (ISO)"; WidthPt = 708; HeightPt = 1000 }
let B4 = { Name = "B4 (JIS)"; WidthPt = 728; HeightPt = 1031 }
let B5 = { Name = "B5 (JIS)"; WidthPt = 516; HeightPt = 728 }
let Folio = { Name = "Folio"; WidthPt = 612; HeightPt = 936 }
let Quarto = { Name = "Quarto"; WidthPt = 609; HeightPt = 779 }
let Note = { Name = "Note"; WidthPt = 612; HeightPt = 792 }
let Number9Envelope = { Name = "Envelope #9"; WidthPt = 278; HeightPt = 638 }
let Number10Envelope = { Name = "Envelope #10"; WidthPt = 297; HeightPt = 684 }
let Number11Envelope = { Name = "Envelope #11"; WidthPt = 324; HeightPt = 746 }
let Number14Envelope = { Name = "Envelope #14"; WidthPt = 360; HeightPt = 828 }
let DLEnvelope = { Name = "Envelope DL"; WidthPt = 311; HeightPt = 623 }
let C3Envelope = { Name = "Envelope C3"; WidthPt = 918; HeightPt = 1298 }
let C4Envelope = { Name = "Envelope C4"; WidthPt = 649; HeightPt = 918 }
let C5Envelope = { Name = "Envelope C5"; WidthPt = 459; HeightPt = 649 }
let C6Envelope = { Name = "Envelope C6"; WidthPt = 323; HeightPt = 459 }
let C65Envelope = { Name = "Envelope C65"; WidthPt = 323; HeightPt = 649 }
let B4Envelope = { Name = "Envelope B4"; WidthPt = 708; HeightPt = 1000 }
let B5Envelope = { Name = "Envelope B5"; WidthPt = 498; HeightPt = 708 }
let B6Envelope = { Name = "Envelope B6"; WidthPt = 498; HeightPt = 354 }
let MonarchEnvelope = { Name = "Envelope Monarch"; WidthPt = 278; HeightPt = 540 }
let PersonalEnvelope = { Name = "Envelope Personal"; WidthPt = 261; HeightPt = 468 }
