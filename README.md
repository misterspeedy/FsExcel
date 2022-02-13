# FsExcel

An F# Excel spreadsheet generator, based on [SpreadsheetLight](https://www.nuget.org/packages/SpreadsheetLight/).

*This is still in early beta.  Anything might change!*

## Example code

```fsharp
open FsExcel
open DocumentFormat.OpenXml
open System.Drawing

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
            Border(TopBorder(Spreadsheet.BorderStyleValues.Medium, Color.Red))
            Border(RightBorder(Spreadsheet.BorderStyleValues.DashDotDot, Color.Orange))
            Border(BottomBorder(Spreadsheet.BorderStyleValues.Thick, Color.Green))
            Border(LeftBorder(Spreadsheet.BorderStyleValues.SlantDashDot, Color.Blue))
            HorizontalAlignment(Right)
            Next(DownBy 1)
        ]
] 
|> render
|> fun ss -> 
    ss.FreezePanes(1, 0)
    ss.AutoFitColumn(1, 10)
    ss.SaveAs(@"/temp/spreadsheet.xlsx")
```
