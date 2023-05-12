open System.IO

let files = 
    [
        {| InFile =  "../Notebooks/Tutorial.dib"; OutFile = "../../README.md" |}
        {| InFile = "../Notebooks/ExcelTableTutorial.dib"; OutFile = "../../ExcelTableTutorial.md" |}
    ]

files
|> Seq.iter (fun file ->
    let mutable inMeta : bool = false
    let mutable firstLine = true
    let mutable inCode = false
    let mutable skip = false
    
    let markDown = 
        file.InFile
        |> File.ReadAllLines
        |> Array.choose (fun line ->
            if line.StartsWith "#!meta" then
                inMeta <- true
            if line.StartsWith "#!fsharp" then
                inMeta <- false
                inCode <- true
                skip <- true
                Some "```fsharp"
            elif line.StartsWith "#!markdown" then
                inMeta <- false
                skip <- true
                if inCode then
                    inCode <- false
                    Some "```"
                else
                    if firstLine then
                        None
                    else
                        Some "---"          
            else
                if inMeta then
                    None
                else
                    firstLine <- false
                    if skip then
                        skip <- false
                        None
                    else
                        Some line
        )
    File.WriteAllLines(file.OutFile, markDown))
