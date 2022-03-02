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

    for ews in expected.Worksheets do
        match actual.TryGetWorksheet(ews.Name) with
        | true, aws ->

            // We combine the CellsUsed sequences from both sides because
            // the cells that are populated in each don't necessarily overlap perfectly:
            let allPopulatedAddresses =
                (ews.CellsUsed())
                |> Seq.append (aws.CellsUsed())
                |> Seq.distinctBy (fun c -> c.Address.ColumnNumber, c.Address.RowNumber)
                |> Seq.map (fun c -> c.Address)

            for address in allPopulatedAddresses do
                let ec = ews.Cell(address)
                let ac = aws.Cell(address)
                Assert.Equal(ec, ac)
        | false, _ ->
            () // TODO
