<img src="https://raw.githubusercontent.com/misterspeedy/FsExcel/main/assets/logo.png"
     alt="FsExcel Logo"
     style="width: 150px;" />
     
[![Twitter URL](https://img.shields.io/twitter/url/https/twitter.com/fsexcel.svg?style=social&label=Twitter%20%40FsExcel)](https://twitter.com/fsexcel)
[![Nuget](https://img.shields.io/nuget/v/Fsexcel)](https://www.nuget.org/packages/FsExcel/)


## Welcome!

Welcome to FsExcel, a library for generating Excel spreadsheets using very simple code.

FsExcel is based on [ClosedXML](https://github.com/ClosedXML/ClosedXML) but abstracts away many of the complications of building spreadsheets cell by cell.

---
**This tutorial is also available as an <a href="https://raw.githubusercontent.com/misterspeedy/FsExcel/main/src/Notebooks/Tutorial.dib" download="Tutorial.dib">interactive notebook</a>. Download it, open in Visual Studio Code, and start generating spreadsheets for real!**

---
## Hello World

Here's the complete code to generate a spreadsheet with a single cell containing a string.

Run this and you should find a spreadsheet called `HelloWorld.xlsx` in your `/temp` folder. (Change the path to suit.)
<!-- Test -->

```fsharp
// For scripts only; for programs, use NuGet to install FsExcel:
#r "nuget: FsExcel"

let savePath = "/temp"

open System.IO
open FsExcel

[
    Cell [ String "Hello world!" ]
]
|> Render.AsFile (Path.Combine(savePath, "HelloWorld.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/HelloWorld.PNG?raw=true"
     alt="Hello World example"
     style="width: 120px;" />

This example embodies the main stages of building a spreadsheet using FsExcel:

1) Build a list using a list comprehension: `[ ... ]`
2) In the list make cells using `Cell`
3) Each cell gets a list of properties, in this case just the cell content, which here is a string: `String "Hello world!"`

If you've used `Fable.React`, or a similar library, you'll already be familiar with the concepts so far.

4) Send the resulting list to `FsExcel.Render.AsFile`, providing a path.

---
## Multiple Cells
<!-- Test -->

```fsharp
open System.IO
open FsExcel

[
    for i in 1..10 do
        Cell [ Integer i ]
]
|> Render.AsFile (Path.Combine(savePath, "MultipleCells.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/MultipleCells.PNG?raw=true"
     alt="Multiple Cells example"
     style="width: 500px;" />

Here we use a `for...` comprehension to build multiple cells. (Don't panic: we could have used `List.map` instead!)

By default each new cell is put on the right of its predecessor.

---
## Vertical Movement

If you want the next cell to be rendered below instead of to the right, you can add a `Next(DownBy 1)` property to the cell:

<!-- Test -->

```fsharp
open System.IO
open System.Globalization
open FsExcel

[
    for m in 1..12 do
        let monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(m)
        Cell [
            String monthName
            Next(DownBy 1)
        ]
]
|> Render.AsFile (Path.Combine(savePath, "VerticalMovement.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/VerticalMovement.PNG?raw=true"
     alt="Vertical Movement example"
     style="width: 100px;" />

---
The `Next` property overrides the default behaviour of rendering each successive cell one to the right. In this case we override it with a 'go down by 1' behaviour.

But what if you want a table of cells? Use the default behaviour for each cell in a row except the last. In the last cell use `Next NewRow`. This causes the next cell to be rendered in column 1 of the next row.

<!-- Test -->

```fsharp
open System.IO
open System.Globalization
open FsExcel

[
    for m in 1..12 do
        let monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(m)
        Cell [
            String monthName
        ]
        Cell [
            Integer monthName.Length
            Next NewRow
        ]
]
|> Render.AsFile (Path.Combine(savePath, "Rows.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/Rows.PNG?raw=true"
     alt="Rows example"
     style="width: 150px;" />

---
Maybe you don't like the idea of saying where to go next in the properties of a cell. No problem, you can have standalone position-control with the `Go` instruction:
<!-- Test -->

```fsharp
open System.IO
open System.Globalization
open FsExcel

[
    for m in 1..12 do
        let monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(m)
        Cell [ String monthName ]
        Cell [ Integer monthName.Length ]
        Go NewRow
]
|> Render.AsFile (Path.Combine(savePath, "RowsGo.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/RowsGo.PNG?raw=true"
     alt="Rows Go example"
     style="width: 150px;" />

---
## Indentation

Maybe you want a series of rows that don't start in column 1.  Use `Indent`:
<!-- Test -->

```fsharp
open System.IO
open System.Globalization
open FsExcel

[
    Go(Indent 2)

    for m in 1..12 do
        let monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(m)
        Cell [ String monthName ]
        Cell [ Integer monthName.Length ]
        Go NewRow
]
|> Render.AsFile (Path.Combine(savePath, "Indentation.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/Indentation.PNG?raw=true"
     alt="Indentation example"
     style="width: 220px;" />

Now each row begins at column 2.

Indents apply to all `NewRow` operations until some other indent value is set using `Go(Indent n)`. Specify no indenting with `Go(Indent 1)`.

You can specify indents relative to the current indent level using `Go(IndentBy n)` where _n_ can be a positive or negative integer.

---
## Border and Font Styling

You can add border styling and font emphasis (bold, italic, underline or strikethrough) styling using `Border (...)` and `FontEmphasis ...` cell properties.

The border style values are in `ClosedXML.Excel.XLBorderStyleValues` and the underline values are in `ClosedXML.Excel.XLFontUnderlineValues`.
<!-- Test -->

```fsharp
open System.IO
open System.Globalization
open FsExcel
open ClosedXML.Excel

[
    for heading in ["Month"; "Letter Count"] do
        Cell [
            String heading
            Border (Border.Bottom XLBorderStyleValues.Medium)
            FontEmphasis Bold
            FontEmphasis Italic
        ]
    Go NewRow
    
    for m in 1..12 do
        let monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(m)
        Cell [ 
            String monthName
            FontEmphasis (Underline XLFontUnderlineValues.DoubleAccounting)
            if monthName = "May" then
                FontEmphasis StrikeThrough
        ]
        Cell [ Integer monthName.Length ]
        Go NewRow
]
|> Render.AsFile (Path.Combine(savePath, "Styling.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/Styling.PNG?raw=true"
     alt="Styling example"
     style="width: 150px;" />

---
As they are just list items, styles can be composed and applied together as a list. You'll need a `yield!` to include these multiple elements in your cell property list.
<!-- Test -->

```fsharp
open System.IO
open System.Globalization
open FsExcel
open ClosedXML.Excel

let headingStyle = 
    [
        Border(Border.Bottom XLBorderStyleValues.Medium)
        FontEmphasis Bold
        FontEmphasis Italic 
    ]

[
    for heading in ["Month"; "Letter Count"] do
        Cell [
            String heading
            yield! headingStyle
        ]
    Go NewRow
    
    for m in 1..12 do
        let monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(m)
        Cell [ String monthName ]
        Cell [ Integer monthName.Length ]
        Go NewRow
]
|> Render.AsFile (Path.Combine(savePath, "ComposedStyling.xlsx"))

```
## Font Name and Size

You can set the font name using `FontName` and the size using `FontSize`:

<!-- Test -->

```fsharp
open System.IO
open System.Globalization
open FsExcel
open ClosedXML.Excel

[
    for i, fontName in ["Arial"; "Bahnschrift"; "Calibri"; "Cambria"; "Comic Sans MS"; "Consolas"; "Constantia"] |> List.indexed do
        Cell [
            String fontName
            FontName fontName
            FontSize (10 + (i * 2) |> float)
        ]
        Go NewRow
]
|> Render.AsFile (Path.Combine(savePath, "FontNameSize.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/FontNameSize.PNG?raw=true"
     alt="Number Format and Alignment example"
     style="width: 250px;" />

---
## Number Formatting and Alignment

Number styling can be applied using standard Excel format strings.  You can also apply horizontal alignment.
<!-- Test -->

```fsharp
open System
open System.IO
open FsExcel
open ClosedXML.Excel

module PseudoRandom =

    let mutable state = 1u
    let mangle (n : UInt64) = (n &&& (0x7fffffff |> uint64)) + (n >>> 31)

    let nextDouble() =
        state <- (state |> uint64) * 48271UL |> mangle |> mangle |> uint32
        (float state) / (float Int32.MaxValue)

let headingStyle = 
    [
        Border(Border.Bottom XLBorderStyleValues.Medium)
        FontEmphasis Bold
        FontEmphasis Italic 
    ]

[
    for heading, alignment in ["Stock Item", Left; "Price", Right ; "Count", Right] do
        Cell [
            String heading
            yield! headingStyle
            HorizontalAlignment alignment
        ]
    
    Go NewRow

    for item in ["Apples"; "Oranges"; "Pears"] do
        Cell [
            String item
        ]
        Cell [
            Float ((PseudoRandom.nextDouble()*1000.))
            FormatCode "$0.00"
        ]
        Cell [
            Integer (int (PseudoRandom.nextDouble()*100.))
            FormatCode "#,##0"
        ]
        Go NewRow
]
|> Render.AsFile (Path.Combine(savePath, "NumberFormatAndAlignment.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/NumberFormatAndAlignment.PNG?raw=true"
     alt="Number Format and Alignment example"
     style="width: 250px;" />

---
## Formulae

You can add a formula to a cell using `FormulaA1(...)`.  

Currently only the `A1` style of cell referencing is supported, meaning that you will need to keep track of the column letter and row number you want to refer to:
<!-- Test -->

```fsharp
open System
open System.IO
open FsExcel
open ClosedXML.Excel

module PseudoRandom =

    let mutable state = 1u
    let mangle (n : UInt64) = (n &&& (0x7fffffff |> uint64)) + (n >>> 31)

    let nextDouble() =
        state <- (state |> uint64) * 48271UL |> mangle |> mangle |> uint32
        (float state) / (float Int32.MaxValue)

let headingStyle = 
    [
        Border(Border.Bottom XLBorderStyleValues.Medium)
        FontEmphasis Bold
        FontEmphasis Italic 
    ]

[
    for heading, alignment in ["Stock Item", Left; "Price", Right ; "Count", Right; "Total", Right] do
        Cell [
            String heading
            yield! headingStyle
            HorizontalAlignment alignment
        ]
    
    Go NewRow

    for index, item in ["Apples"; "Oranges"; "Pears"] |> List.indexed do
        Cell [
            String item
        ]
        Cell [
            Float (PseudoRandom.nextDouble()*1000.)
            FormatCode "$0.00"
        ]
        Cell [
            Integer (int (PseudoRandom.nextDouble()*1000.))
            FormatCode "#,##0"
        ]
        Cell [
            FormulaA1 $"=B{index+2}*C{index+2}"
            FormatCode "$#,##0.00"
        ]
        Go NewRow
]
|> Render.AsFile (Path.Combine(savePath, "Formulae.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/Formulae.PNG?raw=true"
     alt="Styling example"
     style="width: 300px;" />

---
## Color

Set the font color with `FontColor` and the background color with the `BackgroundColor` property.  Set the border color with `BorderColor`.

The color values and color creation functions are in `ClosedXml.Excel.XLColor`.
<!-- Test -->

```fsharp
open System.IO
open FsExcel
open ClosedXML.Excel

[
    let values = [0..32..224] @ [255]
    for r in values do
        for g in values do
            for b in values do
                // N.B. the API refuses to fill a cell with black if its font is black
                // so the very first cell won't be colored.
                let backgroundColor = ClosedXML.Excel.XLColor.FromArgb(0, r, g, b)
                let fontColor = ClosedXML.Excel.XLColor.FromArgb(0, b, r, g)
                let borderColor = ClosedXML.Excel.XLColor.FromArgb(0, g, b, r)
                Cell [
                    String $"R={r};G={g};B={b}"
                    FontColor fontColor
                    BackgroundColor backgroundColor
                    Border (Border.Top XLBorderStyleValues.Thick)
                    Border (Border.Right XLBorderStyleValues.Thick)
                    Border (Border.Bottom XLBorderStyleValues.Thick)
                    Border (Border.Left XLBorderStyleValues.Thick)
                    // Could also have used Border.All:
                    // Border (Border.All XLBorderStyleValues.Thick)
                    BorderColor (BorderColor.Top borderColor)
                    BorderColor (BorderColor.Right borderColor)
                    BorderColor (BorderColor.Bottom borderColor)
                    BorderColor (BorderColor.Left borderColor)
                    // Could also have used BorderColor.All:
                    // BorderColor (BorderColor.All borderColor)
                ]
            Go NewRow
        Go NewRow

]
|> Render.AsFile (Path.Combine(savePath, "Color.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/Color.PNG?raw=true"
     alt="Color example"
     style="width: 400px;" />

---
## Range Styles

You can apply any properties to all cells from a point in your code using `Style [ prop; prop...]`. Don't forget to reset style with `Style []` afterwards.
<!-- Test -->

```fsharp
open System
open System.IO
open FsExcel
open ClosedXML.Excel

module PseudoRandom =

    let mutable state = 1u
    let mangle (n : UInt64) = (n &&& (0x7fffffff |> uint64)) + (n >>> 31)

    let nextDouble() =
        state <- (state |> uint64) * 48271UL |> mangle |> mangle |> uint32
        (float state) / (float Int32.MaxValue)

[
    Style [
        Border(Border.Bottom XLBorderStyleValues.Medium)
        FontEmphasis Bold
        FontEmphasis Italic 
    ]
    for heading in ["Stock Item"; "Price"; "Count"] do
        Cell [ String heading ]
    Style []
    
    Go NewRow
    for item in ["Apples"; "Oranges"; "Pears"] do
        Cell [
            String item
        ]
        Style [ FontEmphasis Italic ]        
        Cell [
            Float ((PseudoRandom.nextDouble()*1000.))
            FormatCode "$0.00"
        ]
        Cell [
            Integer (int (PseudoRandom.nextDouble()*100.))
            FormatCode "#,##0"
        ]
        Style []
        Go NewRow
]
|> Render.AsFile (Path.Combine(savePath, "RangeStyle.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/RangeStyle.PNG?raw=true"
     alt="Range Style example"
     style="width: 250px;" />

---
## Absolute Positioning

FsExcel is designed to save you from having to keep track of absolute row- and column-numbers. However sometimes you might want to position a cell at an absolute row or column position - or both.

After the explicitly-positioned cell, subsequent cells are by default rendered to the right again.
<!-- Test -->

```fsharp
open System.IO
open FsExcel
open ClosedXML.Excel

[
    Go (Col 3)
    Cell [ String "Col 3"]
    Go (Row 4)
    Cell [ String "Row 4"]
    Go (RC(6, 5))
    Cell [ String "R6C5"]
    Cell [ String "R6C6"]
]
|> Render.AsFile (Path.Combine(savePath, "AbsolutePositioning.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/AbsolutePositioning.PNG?raw=true"
     alt="Absolute Positioning example"
     style="width: 350px;" />    

---
Remember that, by default, successive cells are placed to the right of their predecessors? Sometimes (rarely) you might want to suppress that behaviour completely. To do that use `Next Stay`.
<!-- Test -->

```fsharp
open System.IO
open FsExcel

[
    for i in 1..5 do
        Cell [
            Integer i
            Next Stay
        ]
        Go(DownBy i)
]
|> Render.AsFile (Path.Combine(savePath, "Stay.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/Stay.PNG?raw=true"
     alt="Stay example"
     style="width: 150px;" />

---
## Worksheets (Tabs)

By default, all cells are placed into a worksheet (tab) called "Sheet1".  You can override this, and create additional worksheets, using `Worksheet ...`.

If you do not want a "Sheet1" tab you'll need to use `Worksheet` to create an explicitly named sheet - before creating any cells.

Each new worksheet starts at the top-left cell, has an indent setting of 1 (no indent), and has an empty list as its current `Style [...]` value.
<!-- Test -->

```fsharp
open System.IO
open FsExcel
open System.Globalization

[
    Worksheet CultureInfo.CurrentCulture.NativeName
    for m in 1..12 do
        let monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(m)
        Cell [ String monthName ]
        Cell [ Integer monthName.Length ]
        Go NewRow

    let ukrainianCulture = CultureInfo.GetCultureInfoByIetfLanguageTag("uk")
    Worksheet ukrainianCulture.NativeName
    for m in 1..12 do
        let monthName = ukrainianCulture.DateTimeFormat.GetMonthName(m)
        Cell [ String monthName ]
        Cell [ Integer monthName.Length ]
        Go NewRow
]
|> Render.AsFile (Path.Combine(savePath, "Worksheets.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/Worksheets.PNG?raw=true"
     alt="Workseets example"
     style="width: 350px;" />

---
## Autofitting

You can set the widths of columns to fit their contents using ``AutoFit AllCols``. You can auto fit a range of columns with ``AutoFit (ColRange(<c1>, <c2>))``.

You can autofit heights of rows with ``AutoFit AllRows`` and ``AutoFit (RowRange(<r1>,<r2>))``.

You can autofit all columns *and* all rows with ``AutoFit All``.

Perform ``AutoFit`` operations *after* the cells have been populated!
<!-- Test -->

```fsharp
open System.IO
open System.Globalization
open FsExcel
open ClosedXML.Excel

let headingStyle = 
    [
        Border(Border.Bottom XLBorderStyleValues.Medium)
        FontEmphasis Bold
        FontEmphasis Italic 
    ]

[
    for heading in ["Month"; "Letter Count"] do
        Cell [
            String heading
            yield! headingStyle
        ]
    Go NewRow
    
    for m in 1..12 do
        let monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(m)
        Cell [ String monthName ]
        Cell [ Integer monthName.Length ]
        Go NewRow

    AutoFit AllCols
]
|> Render.AsFile (Path.Combine(savePath, "AutosizeColumns.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/AutosizeColumns.PNG?raw=true"
     alt="Autosize Columns example"
     style="width: 200px;" />

---
## Tables from types

You can create a table of cells from an instance or a sequence of any type having serializable fields - for example a record type.

Use `Table.fromInstance` or `Table.fromSeq` and provide

- an orientation (`Table.Direction.Horizontal` or `Table.Direction.Vertical`)
- a function which, given an index and a field name, returns a list of properties for styling. (This style can be an empty list.)
- the instance or sequence.

In horizontal tables, the values for each record appear beside one another.  In vertical tables the values for a record appear below one another.

Calls to the cell style function are given 0 for the header, 1 for the first (or only) data row, 2 for the next and so on.

Tables don't automatically autofit - you'll have to do that (if you want) after the table is built.
<!-- Test -->

```fsharp
open System
open System.IO
open ClosedXML.Excel
open FsExcel

type JoiningInfo =  {
    Name : string
    Age : int
    Fees : decimal
    DateJoined : DateTime
}

// This works just as well if these are anonymous record instances,
// eg. {| Name = "..."; ... |}
let records = [
    { Name = "Jane Smith"; Age = 32; Fees = 59.25m; DateJoined = System.DateTime(2022, 3, 12) }
    { Name = "Michael Nguyễn"; Age = 23; Fees = 61.2m; DateJoined = System.DateTime(2022, 3, 13) }
    { Name = "Sofia Hernández"; Age = 58; Fees = 59.25m; DateJoined = System.DateTime(2022, 3, 15) }
]

let cellStyleVertical index name =
    if index = 0 then
        [ FontEmphasis Bold ]
    elif name = "Fees" then
        [ FormatCode "$0.00" ]
    else
        []

let cellStyleHorizontal index name =
    if index = 0 then
        [
            Border(Border.Bottom XLBorderStyleValues.Medium)
            FontEmphasis Bold
        ]
    elif name = "Fees" then
        [ FormatCode "$0.00" ]
    else
        []

records
|> Table.fromSeq Table.Direction.Vertical cellStyleVertical
|> fun cells -> cells @ [ AutoFit All ]
|> Render.AsFile (Path.Combine(savePath, "RecordSequenceVertical.xlsx"))

records
|> Table.fromSeq Table.Direction.Horizontal cellStyleHorizontal
|> fun cells -> cells @ [ AutoFit All ]
|> Render.AsFile (Path.Combine(savePath, "RecordSequenceHorizontal.xlsx"))

records
|> Seq.tryHead
|> Option.iter (fun r ->

    r 
    |> Table.fromInstance Table.Direction.Vertical cellStyleVertical
    |> fun cells -> cells @ [ AutoFit All ]
    |> Render.AsFile (Path.Combine(savePath, "RecordInstanceVertical.xlsx"))

    r 
    |> Table.fromInstance Table.Direction.Horizontal cellStyleHorizontal
    |> fun cells -> cells @ [ AutoFit All ]
    |> Render.AsFile (Path.Combine(savePath, "RecordInstanceHorizontal.xlsx")))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/RecordSequenceVertical.PNG?raw=true"
     alt="Table example - vertical record sequence"
     style="width: 450px;" />

<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/RecordSequenceHorizontal.PNG?raw=true"
     alt="Table example - horizontal record sequence"
     style="width: 320px;" />

<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/RecordInstanceVertical.PNG?raw=true"
     alt="Table example - vertical record instance"
     style="width: 200px;" />
     
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/RecordInstanceHorizontal.PNG?raw=true"
     alt="Table example - horizontal record instance"
     style="width: 280px;" />

---
## Rendering in Fable Elmish and similar web applications

You can use `Render.AsStream <stream> <items>` to render to a pre-existing stream, or `Render.AsStreamBytes <items>` to render as a byte array. 

`Render.AsStreamBytes` is useful for Fable-based and other web app scenarios. Render to a byte array on the server, and transfer the bytes to the client using Fable Remoting.  On the client use the `SaveFileAs` extension function to start a browser download.  Make sure you have opened the `Fable.Remoting.Client` to get the `SaveFileAs` method of a byte array.

There are few more details here: https://zaid-ajaj.github.io/Fable.Remoting/src/upload-and-download.html

```fsharp
open FsExcel

[
    Cell [ String "Hello world!" ]
]
|> Render.AsStreamBytes
|> fun bytes ->
    $"Bytes length: {bytes.Length}"

```
## Data Types

FsExcel supports the following data types for cell content:

- String
- Integer
- Float
- Boolean
- DateTime
- TimeSpan
<!-- Test -->

```fsharp
open System
open System.IO
open FsExcel

[
    Cell [ String "String"]; Cell [ String "string" ]
    Go NewRow
    Cell [ String "Integer" ]; Cell [ Integer 42 ]
    Go NewRow
    Cell [ String "Number" ]; Cell [ Float Math.PI ]
    Go NewRow
    Cell [ String "Boolean" ]; Cell [ Boolean false  ]
    Go NewRow
    Cell [ String "DateTime" ]; Cell [ DateTime (System.DateTime(1903, 12, 17)) ]
    Go NewRow
    Cell [ String "TimeSpan" ]
    Cell [ 
        TimeSpan (System.TimeSpan(hours=1, minutes=2, seconds=3)) 
        FormatCode "hh:mm:ss"
    ]
]
|> Render.AsFile (Path.Combine(savePath, "DataTypes.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/DataTypes.PNG?raw=true"
     alt="Data Types example"
     style="width: 200px;" />
