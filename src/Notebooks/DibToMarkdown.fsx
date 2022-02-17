open System.IO

let inFile = "./src/Notebooks/Tutorial.dib"
let outFile = "./README.md"

let mutable firstLine = true
let mutable inCode = false
let mutable skip = false
 
let markDown = 
    inFile
    |> File.ReadAllLines
    |> Array.choose (fun line ->
        if line.StartsWith "#!fsharp" then
            inCode <- true
            skip <- true
            Some "```fsharp"
        elif line.StartsWith "#!markdown" then
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
            firstLine <- false
            if skip then
                skip <- false
                None
            else
                Some line
    )

File.WriteAllLines(outFile, markDown)