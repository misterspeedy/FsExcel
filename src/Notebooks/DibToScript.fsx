open System.IO

let inFile = "../Notebooks/Tutorial.dib"
let outFile = "../Tests/RegressionTests/CreateRegressionTestActuals.fsx"

let mutable inCode = false
let mutable inTest = false
let mutable testNumber = 1
 
let code = 
    [
        "#r \"nuget: ClosedXML\""
        "#r \"../../FsExcel/bin/Debug/net5.0/FsExcel.dll\""

        "let savePath = \"../Tests/RegressionTests/Actual\""

        for line in File.ReadAllLines inFile do
            if line.StartsWith "#r \"nuget: FsExcel\"" || line.StartsWith "let savePath =" || line.TrimStart().StartsWith "//" then
                ()
            elif line.StartsWith "<!-- Test -->" then
                inTest <- true
                $"module Test{testNumber} ="
                testNumber <- testNumber + 1
            elif line.StartsWith "#!fsharp" then
                inCode <- true
            elif line.StartsWith "#!markdown" then
                inCode <- false
                inTest <- false
            else
                if inCode && inTest then
                    $"    {line}"
    ]

File.WriteAllLines(outFile, code)