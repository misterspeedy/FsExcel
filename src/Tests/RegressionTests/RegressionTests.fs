module Tests

open System
open System.IO
open Xunit
open Xunit.Abstractions
open ClosedXML.Excel

let expectedsPath = __SOURCE_DIRECTORY__ + "/Expected"
let actualsPath =  __SOURCE_DIRECTORY__ + "/Actual"


let isWindowsPlatform =
    if Environment.OSVersion.Platform = PlatformID.Win32NT
       || Environment.OSVersion.Platform = PlatformID.Win32S
       || Environment.OSVersion.Platform = PlatformID.Win32Windows
       || Environment.OSVersion.Platform = PlatformID.WinCE
    then
        true
    else
        false
        
module Check =

    let fromFilename(output : ITestOutputHelper) (filename : string) =
        let expected = new XLWorkbook(Path.Combine(expectedsPath, filename))
        let actual = new XLWorkbook(Path.Combine(actualsPath, filename))
        try
            Assert.Workbook.Equal(expected, actual, filename, output)
            Ok()
        with e -> Error (filename, e)

type Tests(output : ITestOutputHelper) =

    [<Fact>]
    member _.``RegressionTests`` () =
        
        let results =
            expectedsPath
            |> Directory.EnumerateFiles
            |> Seq.map (Path.GetFileName >> (Check.fromFilename output))
        
        let messages =
            results
            |> Seq.choose (function
                | Ok () -> None
                | Error ("FontNameSize.xlsx", _) when not isWindowsPlatform -> None // the font list is different
                | Error (_, e) -> Some (string e)
                )
            |> String.concat "\n"
        
        if messages.Length > 0 then
            failwith $"%s{messages}"