module Tests

open System.IO
open Xunit
open Xunit.Abstractions
open ClosedXML.Excel

let expectedsPath = "../../../Expected"
let actualsPath = "../../../Actual"

module Check =

    let fromFilename(output : ITestOutputHelper) (filename : string) =
        let expected = new XLWorkbook(Path.Combine(expectedsPath, filename))
        let actual = new XLWorkbook(Path.Combine(actualsPath, filename))
        Assert.Workbook.Equal(expected, actual, filename, output)    

type Tests(output : ITestOutputHelper) =

    [<Fact>]
    member _.``RegressionTests`` () =
        expectedsPath
        |> Directory.EnumerateFiles
        |> Seq.iter (Path.GetFileName >> (Check.fromFilename output))
