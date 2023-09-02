module Tests

open System.IO
open System.Runtime.InteropServices
open Xunit
open Xunit.Abstractions
open ClosedXML.Excel

let expectedsPath = "../../../Expected"
let actualsPath = "../../../Actual"

module Check =

    let fromFilename(output : ITestOutputHelper) (filename : string) =
        let actual = new XLWorkbook(Path.Combine(actualsPath, filename))
        let expected = new XLWorkbook(Path.Combine(expectedsPath, filename))
        Assert.Workbook.Equal(expected, actual, filename, output)    

type Tests(output : ITestOutputHelper) =
    do
        if RuntimeInformation.IsOSPlatform(OSPlatform.Linux) then
            LoadOptions.DefaultGraphicEngine <- ClosedXML.Graphics.DefaultGraphicEngine("DejaVu Sans")

    static member files =
        expectedsPath
        |> Directory.EnumerateFiles
        |> Seq.map (Path.GetFileName >> Array.singleton)

    [<Theory>]
    [<MemberData("files")>]
    member _.``RegressionTests`` (fileName:string) : unit = Check.fromFilename output fileName
