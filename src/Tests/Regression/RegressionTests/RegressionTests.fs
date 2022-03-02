module Tests

open System
open System.IO
open Xunit
open FsExcel
open ClosedXML.Excel

let savePath = "testtemp"
do Directory.CreateDirectory(savePath) |> ignore

[<Fact>]
let ``HelloWorld`` () =

    let filename = "HelloWorld.xlsx"
    [
        Cell [ String "Hello world!" ]
    ]
    |> Render.AsFile (Path.Combine(savePath, filename))

    let expected = new XLWorkbook(Path.Combine("../../../Expected", filename))
    let actual = new XLWorkbook(Path.Combine(savePath, filename))

    Assert.Workbook.Equal(expected, actual)    

[<Fact>]
let ``DataTypes`` () =

    let filename = "DataTypes.xlsx"
    [
        Cell [ String "String"]; Cell [ String "string" ]
        Go NewRow
        Cell [ String "Integer" ]; Cell [ Integer 42 ]
        Go NewRow
        Cell [ String "Number" ]; Cell [ Float Math.PI ]
        Go NewRow
        Cell [ String "Boolean" ]; Cell [ Boolean false ]
        Go NewRow
        Cell [ String "DateTime" ]; Cell [ DateTime (System.DateTime(1903, 12, 17)) ]
        Go NewRow
        Cell [ String "TimeSpan" ]
        Cell [ 
            TimeSpan (System.TimeSpan(hours=1, minutes=2, seconds=3)) 
            FormatCode "hh:mm:ss"
        ]
    ]
    |> Render.AsFile (Path.Combine(savePath, filename))

    let expected = new XLWorkbook(Path.Combine("../../../Expected", filename))
    let actual = new XLWorkbook(Path.Combine(savePath, filename))

    Assert.Workbook.Equal(expected, actual)
