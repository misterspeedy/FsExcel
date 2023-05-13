<img src="https://raw.githubusercontent.com/misterspeedy/FsExcel/main/assets/logo.png"
     alt="FsExcel Logo"
     style="width: 150px;" />
     
## Excel Tables

FsExcel can create [Excel Tables](https://support.microsoft.com/en-us/office/overview-of-excel-tables-7ab0bb7d-3a9e-4b56-a3c9-6c94334e492c). You can apply table styles, add a totals row, and use [structured references](https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e) to include values from the table in formulae.

To create a table you'll need a list of record instances (or class instances). The fields of the instances will become columns in the table.

Let's start with a small dataset listing the contributions to UK electricity generation by various classes of energy source.
<!-- TestSetup -->

```fsharp
#r "nuget: ClosedXML"
#r "nuget: FsExcel"

let savePath = "/temp"

let yearStats =
    [
        // Year, Coal, Oil, Natural Gas, Nuclear, Hydro, Wind/Solar, Other, Total
        1996, 33.67, 3.87, 17.36, 22.18, 0.29, 0.04, 2.14, 79.55
        1997, 28.30, 1.85, 21.57, 21.98, 0.38, 0.06, 2.29, 76.43
        1998, 29.94, 1.70, 23.02, 23.44, 0.44, 0.08, 2.52, 81.14
        1999, 25.51, 1.54, 27.13, 22.22, 0.46, 0.07, 2.79, 79.72
        2000, 28.67, 1.55, 27.91, 19.64, 0.44, 0.08, 2.93, 81.21
        2001, 31.61, 1.42, 26.87, 20.77, 0.35, 0.08, 2.91, 84.01
        2002, 29.63, 1.29, 28.33, 20.10, 0.41, 0.11, 3.13, 83.00
        2003, 32.54, 1.19, 27.85, 20.04, 0.28, 0.11, 3.93, 85.95
        2004, 31.31, 1.10, 29.25, 18.16, 0.42, 0.17, 4.15, 84.57
        2005, 32.58, 1.31, 28.52, 18.37, 0.42, 0.25, 5.23, 86.68
        2006, 35.94, 1.43, 26.78, 17.13, 0.40, 0.36, 5.02, 87.06
        2007, 32.92, 1.16, 30.60, 14.04, 0.44, 0.46, 4.68, 84.28
        2008, 29.96, 1.58, 32.40, 11.91, 0.44, 0.61, 4.67, 81.58
        2009, 24.66, 1.51, 30.90, 15.23, 0.45, 0.80, 4.87, 78.42
        2010, 25.56, 1.18, 32.43, 13.93, 0.31, 0.89, 5.11, 79.41
        2011, 26.03, 0.78, 26.58, 15.63, 0.49, 1.39, 5.62, 76.52
        2012, 34.33, 0.73, 18.62, 15.21, 0.46, 1.82, 6.07, 77.24
        2013, 31.33, 0.59, 17.70, 15.44, 0.40, 2.61, 6.45, 74.52
        2014, 24.01, 0.55, 18.73, 13.85, 0.51, 3.10, 7.73, 68.48
        2015, 18.34, 0.61, 18.28, 15.48, 0.54, 4.11, 9.36, 66.72
        2016, 7.53, 0.58, 25.63, 15.41, 0.46, 4.09, 9.96, 63.67
        2017, 5.55, 0.54, 24.60, 15.12, 0.51, 5.25, 10.13, 61.71
        2018, 4.24, 0.49, 23.51, 14.06, 0.47, 5.98, 11.13, 59.88
        2019, 1.85, 0.39, 23.45, 12.09, 0.51, 6.56, 11.35, 56.19
        2020, 1.47, 0.36, 19.98, 10.72, 0.59, 7.61, 11.65, 52.38
        2021, 1.67, 0.41, 21.83, 9.90, 0.47, 6.60, 12.12, 53.01
    ] 

type YearStats =
    {
        Year : int
        Coal : float
        Oil : float
        NaturalGas : float
        Nuclear : float
        Hydro : float
        WindSolar : float
        Other : float
        Total : float
    }

module YearStats =

    let fromValues = 
        List.map (fun (y, c, o, ng, n, h, ws, ot, t) ->
            {
                Year = y
                Coal = c
                Oil = o
                NaturalGas = ng
                Nuclear = n
                Hydro = h
                WindSolar = ws
                Other = ot
                Total = t
            })

    
let yearStatsRecords = yearStats |> YearStats.fromValues

```
### A simple table

In the example below we start by adding a title using a basic FsExcel `Cell` item - this does not form part of the Excel table.

After the title we use `Go NewRow` twice to create a little empty space.

Then we use `Table` to create the table.

Like `Cell`, `Table` takes a list of properties.  The most important of these are:

- `TableName` - gives the Excel Table a name. If you omit this, your tables will be called "Table1", "Table2" etc. Any spaces in the name you provide will be removed.
- `TableItems` - takes a list of record or class instances, each of which will become a row in the table, using fields from the instance as columns in the table.

*The instance fields must be of simple types like `float`, `int`, `string` and `bool`.*

*`DateTime` and `DateTimeOffset` also work, but you should be very wary of formatting and timezone issues when using these. More complex types, including Discriminated Unions, collections, classes and entire records will not work.*

Finally we add another `Cell` item to provide a footer. Note how the 'current cell' after inserting a table is always the cell below and at the left of the added table.
<!-- Test -->

```fsharp
open System.IO
open FsExcel

[
    Cell [ 
        String "UK Electricity Energy Contributions by Fuel Type 1996-2021"
        FontEmphasis Bold
        FontSize 15
    ]
    Go NewRow; Go NewRow
    Table [
        TableName "UK Electricity"
        TableItems yearStatsRecords
    ]
    Cell [
        String "Source: https://www.gov.uk/government/statistical-data-sets/historical-electricity-data"
        FontEmphasis Italic
        FontSize 9
    ]
]
|> Render.AsFile (System.IO.Path.Combine(savePath, "ExcelTableSimple.xlsx"))

```
*Some of the rows in this and following screenshots have been hidden to save space.*

---
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/ExcelTableSimple.PNG?raw=true"
     alt="Excel table - simple example"
     style="width: 6
     00px;" />

---
### Adding a totals row

Excel tables can have a 'totals row', which can be populated with a value based on one of a number of standard functions - `SUM`, `AVERAGE` and so forth.

Add these using `Totals`, which takes a list of column names (each *must* be the name of one of the table columns), and a an item which specifies what to put in the total cell for these columns.

- To include a standard function (`Average`, `Count`, `CountNumbers`...) use `Function <totalsRowFunction>`. The available standard functions are enumerated in `ClosedXML.Excel.XLTotalsRowFunction`.
- To include a label use `Label <string>`.

*Note that the value `ClosedXML.Excel.XLTotalsRowFunction.Custom` is not currently supported and does nothing.*
<!-- Test -->

```fsharp
open System.IO
open FsExcel
open ClosedXML.Excel

[
    Cell [ 
        String "UK Electricity Energy Contributions by Fuel Type 1996-2021"
        FontEmphasis Bold
        FontSize 15
    ]
    Go NewRow; Go NewRow
    Table [
        TableName "UK Electricity"
        TableItems yearStatsRecords
        Totals(["Year"], Label "Average:")
        Totals(["Coal"; "Oil"; "NaturalGas"; "Nuclear"; "Hydro"; "WindSolar"; "Other"; "Total"], Function XLTotalsRowFunction.Average)
    ]
]
|> Render.AsFile (System.IO.Path.Combine(savePath, "ExcelTableTotals.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/ExcelTableTotals.PNG?raw=true"
     alt="Excel table - column totals example"
     style="width: 600px;" />

---
### Adding a number format for a table column

In the previous example, the totals rows (showing averages) have more decimal places than the values used to compute them.  You can set the number format for an entire table column, including the total cell if displayed) with `ColFormatCodes`.
<!-- Test -->

```fsharp
open System.IO
open FsExcel
open ClosedXML.Excel

[
    Cell [ 
        String "UK Electricity Energy Contributions by Fuel Type 1996-2021"
        FontEmphasis Bold
        FontSize 15
    ]
    Go NewRow; Go NewRow
    Table [
        TableName "UK Electricity"
        TableItems yearStatsRecords

        Totals(["Year"], Label "Average:")

        let statsColumns = ["Coal"; "Oil"; "NaturalGas"; "Nuclear"; "Hydro"; "WindSolar"; "Other"; "Total"]
        Totals(statsColumns, Function XLTotalsRowFunction.Average)
        ColFormatCodes(statsColumns, "0.00")
    ]
]
|> Render.AsFile (System.IO.Path.Combine(savePath, "ExcelTableColumnFormat.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/ExcelTableColumnFormat.PNG?raw=true"
     alt="Excel table - column format example"
     style="width: 600px;" />

---
### Adding formula-based columns

Some of the columns in your table can be calculated from other values in the table, or from values elsewhere in the spreadsheet.

To achieve this:
- Include the additional columns as fields in the item records or classes. Use dummy values when instantiating the instances - these will be overwritten in the table.
- Use `ColFormula` to populate the table cells with the required values.

`ColFormula` takes:

- A list of column names - each *must* be the name of one of the table columns.
- A string containing the required formula - e.g. `"=[Coal]+[Oil]+[NaturalGas]"`.

Typically the formula will use [structured references](https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e) to other items in the table. In this example, `[Coal]`, `[Oil]` and `[NaturalGas]` are structured references to the columns with these names.
<!-- Test -->

```fsharp
open System.IO
open FsExcel

type YearStatsWithCategoryTotals =
    {
        Year : int
        Coal : float
        Oil : float
        NaturalGas : float
        Nuclear : float
        Hydro : float
        WindSolar : float
        Other : float
        
        // These columns will be populated using formulae:
        TotalFossil : float
        TotalSustainable : float
        TotalSustainableNonNuclear : float

        Total : float
    }

module YearStatsWithCategoryTotals =

    let fromYearStatsRecords (yearStats : List<YearStats>) =
        yearStatsRecords
        |> List.map (fun ys ->
            {
                Year = ys.Year
                Coal = ys.Coal
                Oil = ys.Oil
                NaturalGas = ys.NaturalGas
                Nuclear = ys.Nuclear
                Hydro = ys.Hydro
                WindSolar = ys.WindSolar
                Other = ys.Other
                
                // Values are overwritten in the table by formulae:
                TotalFossil = -1.0
                TotalSustainable = -1.0
                TotalSustainableNonNuclear = -1.0

                Total = ys.Total
            }
        )

[
    Cell [ 
        String "UK Electricity Energy Contributions by Fuel Type 1996-2021"
        FontEmphasis Bold
        FontSize 15
    ]
    Go NewRow; Go NewRow
    Table [
        TableName "UK Electricity"
        TableItems (yearStatsRecords |> YearStatsWithCategoryTotals.fromYearStatsRecords)
        ColFormula ("TotalFossil", "=[Coal]+[Oil]+[NaturalGas]")
        ColFormula ("TotalSustainable", "=[Nuclear]+[Hydro]+[WindSolar]")
        ColFormula ("TotalSustainableNonNuclear", "=[Hydro]+[WindSolar]")
    ]
]
|> Render.AsFile (System.IO.Path.Combine(savePath, "ExcelTableColumnFormulae.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/ExcelTableColumnFormulae.PNG?raw=true"
     alt="Excel table - column formulae example"
     style="width: 600px;" />

---
### Using a class instead of a record type

Although tables are normally generated from F# record instances, class types are also supported.  The example below uses a class instead of a record type.

*Note that although it is also possible to use F# __anonymous__ record instances as table items, the F# compiler returns anonymous record fields in alphabetical order rather than declaration order, so your columns are likely to appear in the wrong order.*
<!-- Test -->

```fsharp
open System.IO
open FsExcel

type YearStatsClass(year : int, coal : float, oil : float, naturalGas : float, nuclear : float, hydro : float, windSolar : float, other : float, total : float) =
    member _.Year = year
    member _.Coal = coal
    member _.Oil = oil
    member _.NaturalGas = naturalGas
    member _.Nuclear = nuclear
    member _.Hydro = hydro
    member _.WindSolar = windSolar
    member _.Other = other
    member _.Total = total
    
let yearStatsClasses =
    yearStats
    |> List.map YearStatsClass

[
    Cell [ 
        String "UK Electricity Energy Contributions by Fuel Type 1996-2021"
        FontEmphasis Bold
        FontSize 15
    ]
    Go NewRow; Go NewRow
    Table [
        TableName "UK Electricity"
        TableItems yearStatsClasses
    ]
]
|> Render.AsFile (System.IO.Path.Combine(savePath, "ExcelTableClass.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/ExcelTableClass.PNG?raw=true"
     alt="Excel table - class example"
     style="width: 600px;" />

---
### Table styles

You can apply various kinds of styling to a table:

- To set a standard theme, use `Theme <theme>` where `<theme>` is one of the values in `ClosedXML.Excel.XLTableTheme`.
- To show row and column stripes, use `ShowRowStripes true` and `ShowColumnStripes true`.
- To emphasize the first and last columns, use `EmphasizeFirstColumn true` and `EmphasizeLastColumn true`.
- You can hide the header row with `ShowHeaderRow false`. We don't give an example here as our example table only makes sense with headers.
<!-- Test -->

```fsharp
open System.IO
open FsExcel
open ClosedXML.Excel

[
    Cell [ 
        String "UK Electricity Energy Contributions by Fuel Type 1996-2021"
        FontEmphasis Bold
        FontSize 15
    ]
    Go NewRow; Go NewRow
    Table [
        TableName "UK Electricity"
        TableItems yearStatsRecords
        Theme XLTableTheme.TableStyleDark9
        ShowRowStripes true
        ShowColumnStripes true
        EmphasizeFirstColumn true
        EmphasizeLastColumn true
    ]
]
|> Render.AsFile (System.IO.Path.Combine(savePath, "ExcelTableStyle.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/ExcelTableStyle.PNG?raw=true"
     alt="Excel table - table style example"
     style="width: 600px;" />

---
### Structured references from outside the table

Cells outside the table can use [structured references](https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e) to use values from within the table. This makes most sense when the formula involved reduces all the values in a column to a single value, e.g. `MIN`, `MAX` or `AVERAGE`.

In this example we use `MIN`, `MAX` and structured references, together with some string concatenation using `&`, to get the year range for the title from the table data.
<!-- Test -->

```fsharp
open System.IO
open FsExcel

[
    Cell [ 
        FormulaA1 "=\"UK Electricity Energy Contributions by Fuel Type \"&MIN(UKElectricity[Year])&\"-\"&MAX(UKElectricity[Year])"
        FontEmphasis Bold
        FontSize 15
    ]
    Go NewRow; Go NewRow
    Table [
        TableName "UK Electricity"
        TableItems yearStatsRecords
    ]
]
|> Render.AsFile (System.IO.Path.Combine(savePath, "ExcelTableStructuredReference.xlsx"))

```
<img src="https://github.com/misterspeedy/FsExcel/blob/main/assets/ExcelTableStructuredReference.PNG?raw=true"
     alt="Excel table - structured reference example"
     style="width: 700px;" />
