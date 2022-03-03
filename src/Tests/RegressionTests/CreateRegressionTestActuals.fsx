#r "nuget: ClosedXML"
#r "../../FsExcel/bin/Debug/net5.0/FsExcel.dll"
let savePath = "../Tests/RegressionTests/Expected"
module Test1 =
    
    
    
    open System.IO
    open FsExcel
    
    [
        Cell [ String "Hello world!" ]
    ]
    |> Render.AsFile (Path.Combine(savePath, "HelloWorld.xlsx"))
    
module Test2 =
    
    open System.IO
    open FsExcel
    
    [
        for i in 1..10 do
            Cell [ Integer i ]
    ]
    |> Render.AsFile (Path.Combine(savePath, "MultipleCells.xlsx"))
    
module Test3 =
    
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
    
module Test4 =
    
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
    
module Test5 =
    
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
    
module Test6 =
    
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
    
module Test7 =
    
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
    
module Test8 =
    
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
    
module Test9 =
    
    open System.IO
    open FsExcel
    open ClosedXML.Excel
    
    let r = System.Random()
    
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
                Float ((r.NextDouble()*1000.))
                FormatCode "$0.00"
            ]
            Cell [
                Integer (int (r.NextDouble()*100.))
                FormatCode "#,###"
            ]
            Go NewRow
    ]
    |> Render.AsFile (Path.Combine(savePath, "NumberFormatAndAlignment.xlsx"))
    
module Test10 =
    
    open System.IO
    open FsExcel
    open ClosedXML.Excel
    
    let r = System.Random()
    
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
                Float ((r.NextDouble()*1000.))
                FormatCode "$0.00"
            ]
            Cell [
                Integer (int (r.NextDouble()*100.))
                FormatCode "#,###"
            ]
            Cell [
                FormulaA1 $"=B{index+2}*C{index+2}"
                FormatCode "$#,##0.00"
            ]
            Go NewRow
    ]
    |> Render.AsFile (Path.Combine(savePath, "Formulae.xlsx"))
    
module Test11 =
    
    open System.IO
    open FsExcel
    open ClosedXML.Excel
    
    [
        let values = [0..32..224] @ [255]
        for r in values do
            for g in values do
                for b in values do
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
                        BorderColor (BorderColor.Top borderColor)
                        BorderColor (BorderColor.Right borderColor)
                        BorderColor (BorderColor.Bottom borderColor)
                        BorderColor (BorderColor.Left borderColor)
                    ]
                Go NewRow
            Go NewRow
    
    ]
    |> Render.AsFile (Path.Combine(savePath, "Color.xlsx"))
    
module Test12 =
    
    open System.IO
    open FsExcel
    open ClosedXML.Excel
    
    let r = System.Random()
    
    [
        Style [
            Border(Border.Bottom XLBorderStyleValues.Medium)
            FontEmphasis Bold
            FontEmphasis Italic 
        ]
        for heading, alignment in ["Stock Item", Left; "Price", Right ; "Count", Right] do
            Cell [ String heading ]
        Style []
        
        Go NewRow
        for item in ["Apples"; "Oranges"; "Pears"] do
            Cell [
                String item
            ]
            Style [ FontEmphasis Italic ]        
            Cell [
                Float ((r.NextDouble()*1000.))
                FormatCode "$0.00"
            ]
            Cell [
                Integer (int (r.NextDouble()*100.))
                FormatCode "#,###"
            ]
            Style []
            Go NewRow
    ]
    |> Render.AsFile (Path.Combine(savePath, "RangeStyle.xlsx"))
    
module Test13 =
    
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
    ]
    |> Render.AsFile (Path.Combine(savePath, "AbsolutePositioning.xlsx"))
    
module Test14 =
    
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
    
module Test15 =
    
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
    
module Test16 =
    
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
    
module Test17 =
    
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
    
module Test18 =
    
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
    