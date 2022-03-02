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
    |> Render.AsFile (Path.Combine(savePath, "HelloWorld.xlsx"))

    let expected = new XLWorkbook(Path.Combine("../../../Expected", filename))
    let actual = new XLWorkbook(Path.Combine(savePath, "HelloWorld.xlsx"))

    // TODO we need to do this both ways round because either
    // side might have used cells that are not used on the other side.
    for ews in expected.Worksheets do
        match actual.TryGetWorksheet(ews.Name) with
        | true, aws ->
            for ec in ews.CellsUsed() do
                let ac = aws.Cell(ec.Address)
                Assert.Equal(ec, ac)
        | false, _ ->
            () // TODO
