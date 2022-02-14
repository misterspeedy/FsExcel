# FsExcel

An F# Excel spreadsheet generator, based on [ClosedXML](https://www.nuget.org/packages/ClosedXML/).

*This is still in early beta.  Anything might change!*

## Example code

```fsharp
open FsExcel
open ClosedXML.Excel

[
    Cell [
        Content(String "Hello world")
        Next(DownBy 1)
    ]
    Cell [
        Content(Number System.Math.PI)
        HorizontalAlignment(Center)
        FontEmphasis(Bold)
    ]

    Go(DownBy 3)

    Cell [
        Content(Number System.Math.E)
        FontEmphasis(Bold)
        FontEmphasis(Italic)
        FormatCode "0.00"
        Next(DownBy 1)
    ]

    Go(RC(5, 3))

    for i in 1..10 do
        Cell [
            Content(Number(float i))
            HorizontalAlignment(Left)
            Next(RightBy 1)
        ]

    Go(RC(7, 2))

    for m in 1..12 do
        Cell [
            Content(String(System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(m)))
            FontEmphasis(Italic)
            Border(TopBorder(XLBorderStyleValues.Medium))
            Border(RightBorder(XLBorderStyleValues.DashDotDot))
            Border(BottomBorder(XLBorderStyleValues.Thick))
            Border(LeftBorder(XLBorderStyleValues.SlantDashDot))
            HorizontalAlignment(Right)
            Next(DownBy 1)
        ]
] 
|> render
|> fun wb -> 
    match wb.Worksheets.TryGetWorksheet("Sheet 1") with
    | true, ws -> 
        ws.SheetView.FreezeRows(1)
        ws.Columns().AdjustToContents() |> ignore
    | false, _ ->
        ()
    wb.SaveAs(@"/temp/spreadsheet.xlsx")
```
