open System
open System.IO

open Formatter


[<EntryPoint>]
let main args = 
    let format = DocumentReformattingOption()
    format.PageFooter <- Some "内容由AI生成，内容仅供参考，请仔细甄别"
    Debug.printPublicProperties format
    
    let files = 
        [|
            for path in args do
                if not <| File.Exists (path) then
                    raise <| FileNotFoundException(path)

                let targetPath = 
                    let dir = Path.GetDirectoryName(path)
                    let name = Path.GetFileNameWithoutExtension(path)
                    let ext = Path.GetExtension(path)

                    let newName = $"{name}_reformat{ext}"
                    if String.IsNullOrEmpty(dir) then
                        newName
                    else
                        Path.Combine(dir, newName)

                path, targetPath
        |]
    
    for (src, tgt) in files do 
        Formatter.format src format tgt

    Debug.pause()
    0