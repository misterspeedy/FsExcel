module Tests

open System.IO
open Xunit
open ClosedXML.Excel

let expectedsPath = "../../../Expected"
let actualsPath = "../../../Actual"

module Check =

    let fromFilename (filename : string) =
        let expected = new XLWorkbook(Path.Combine(expectedsPath, filename))
        let actual = new XLWorkbook(Path.Combine(actualsPath, filename))
        Assert.Workbook.Equal(expected, actual)    

[<Fact>]
let ``RegressionTests`` () =
    expectedsPath
    |> Directory.EnumerateFiles
    |> Seq.iter (Path.GetFileName >> Check.fromFilename)
