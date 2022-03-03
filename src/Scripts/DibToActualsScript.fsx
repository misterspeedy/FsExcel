open System.IO

let inFile = Path.Combine(__SOURCE_DIRECTORY__, "../Notebooks/Tutorial.dib")
let outFile = Path.Combine(__SOURCE_DIRECTORY__, "CreateRegressionTestActuals.fsx")

let mutable inCode = false
let mutable inTest = false
let mutable testNumber = 1
 
let code = 
    [
        "#r \"nuget: ClosedXML\""
        "#r \"../FsExcel/bin/Debug/net5.0/FsExcel.dll\""

        "let savePath = __SOURCE_DIRECTORY__ + \"/../Tests/RegressionTests/Actual\""

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
            elif inCode && inTest then
                    $"    {line}"
    ]

File.WriteAllLines(outFile, code)