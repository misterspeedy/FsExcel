module Assert

open Xunit
open ClosedXML.Excel

module Cell = 

    let Equal (expected : IXLCell, actual : IXLCell) =
        if expected.DataType <> actual.DataType then
            raise (Xunit.Sdk.NotEqualException($"{expected.DataType}", $"{actual.DataType}"))
        else
            match expected.DataType with
            | XLDataType.Text ->
                let e = expected.GetString()
                let a = actual.GetString()
                Assert.Equal(e, a)
            | XLDataType.Number ->
                let e = expected.GetDouble()
                let a = actual.GetDouble()
                Assert.Equal(e, a)
            | XLDataType.Boolean ->
                let e = expected.GetBoolean()
                let a = actual.GetBoolean()
                Assert.Equal(e, a)
            | XLDataType.DateTime ->
                let e = expected.GetDateTime()
                let a = actual.GetDateTime()
                Assert.Equal(e, a)
            | XLDataType.TimeSpan ->
                let e = expected.GetTimeSpan()
                let a = actual.GetTimeSpan()
                Assert.Equal(e, a)
            | _ -> 
                raise <| System.NotImplementedException()

        Assert.Equal(expected.Style.Border.TopBorder, actual.Style.Border.TopBorder)
        Assert.Equal(expected.Style.Border.RightBorder, actual.Style.Border.RightBorder)
        Assert.Equal(expected.Style.Border.BottomBorder, actual.Style.Border.BottomBorder)
        Assert.Equal(expected.Style.Border.LeftBorder, actual.Style.Border.LeftBorder)

        Assert.Equal(expected.Style.Font.Bold, actual.Style.Font.Bold)
        Assert.Equal(expected.Style.Font.Italic, actual.Style.Font.Italic)
        Assert.Equal(expected.Style.Font.Underline, actual.Style.Font.Underline)
        // TODO There seems to be a bug (in ClosedXml or Excel) which means if a spreadsheet
        // is opened and then saved without any changes, and there is any strikethrough, this
        // test will start failing.  Need to look at CloseXml source - may be reading the
        // Strikethrough property incorrectly?
        Assert.Equal(expected.Style.Font.Strikethrough, actual.Style.Font.Strikethrough)

        Assert.Equal(expected.Style.Font.FontName, actual.Style.Font.FontName)
        Assert.Equal(expected.Style.Font.FontSize, actual.Style.Font.FontSize)

module Workbook = 

    let Equal (expected : IXLWorkbook, actual : IXLWorkbook) =

        // TODO should explicitly compare worksheet names first

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
                    Cell.Equal(ec, ac)
            | false, _ ->
                raise <| System.Exception($"Could not open sheet {ews.Name}")     
