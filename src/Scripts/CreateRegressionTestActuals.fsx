#r "nuget: ClosedXML"
#r "../FsExcel/bin/Debug/netstandard2.1/FsExcel.dll"
let savePath = __SOURCE_DIRECTORY__ + "/../Tests/RegressionTests/Actual"
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
            let monthName = CultureInfo.GetCultureInfoByIetfLanguageTag("en-GB").DateTimeFormat.GetMonthName(m)
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
            let monthName = CultureInfo.GetCultureInfoByIetfLanguageTag("en-GB").DateTimeFormat.GetMonthName(m)
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
            let monthName = CultureInfo.GetCultureInfoByIetfLanguageTag("en-GB").DateTimeFormat.GetMonthName(m)
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
            let monthName = CultureInfo.GetCultureInfoByIetfLanguageTag("en-GB").DateTimeFormat.GetMonthName(m)
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
            let monthName = CultureInfo.GetCultureInfoByIetfLanguageTag("en-GB").DateTimeFormat.GetMonthName(m)
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
            let monthName = CultureInfo.GetCultureInfoByIetfLanguageTag("en-GB").DateTimeFormat.GetMonthName(m)
            Cell [ String monthName ]
            Cell [ Integer monthName.Length ]
            Go NewRow
    ]
    |> Render.AsFile (Path.Combine(savePath, "ComposedStyling.xlsx"))
    
module Test9 =
    
    open System.IO
    open System.Globalization
    open FsExcel
    open ClosedXML.Excel
    
    let fontNames = 
        SixLabors.Fonts.SystemFonts.Collection.Families
        |> Seq.map (fun font -> font.Name)
        |> Seq.sort
        |> Seq.truncate 20
    
    [
        for i, fontName in fontNames |> Seq.indexed do
            Cell [
                String fontName
                FontName fontName
                FontSize (10 + (i * 2) |> float)
            ]
            Go NewRow
    ]
    |> Render.AsFile (Path.Combine(savePath, "FontNameSize.xlsx"))
    
module Test10 =
    
    open System.IO
    open FsExcel
    open ClosedXML.Excel
    
    [
        Cell [ String "Without wrap text:"
               HorizontalAlignment Center
               VerticalAlignment Middle
               CellSize (ColWidth 16) ]
        Cell [ String "The quick brown fox jumps over the lazy dog."
               HorizontalAlignment Center
               VerticalAlignment Middle ]
        Go NewRow
        Cell [ String "With wrap text:"
               HorizontalAlignment Center
               VerticalAlignment Middle 
               CellSize (ColWidth 16) ]
        Cell [ String "The quick brown fox jumps over the lazy dog."
               HorizontalAlignment Center
               VerticalAlignment Middle
               WrapText true ]
    ]
    |> Render.AsFile (Path.Combine(savePath, "WrapText.xlsx"))
    
module Test11 =
    
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
    
module Test12 =
    
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
    
module Test13 =
    
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
    
module Test14 =
    
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
    
module Test15 =
    
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
    
module Test16 =
    
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
    
module Test17 =
    
    open System.IO
    open FsExcel
    
    [
        Cell [ 
            String "JohnDoe"
            Name "Username" ]
        Cell [ 
            String "john.doe@company.com"
            ScopedName ("Email", NameScope.Workbook) ]
    ]
    |> Render.AsFile (Path.Combine(savePath, "NamedCells.xlsx"))
    
module Test18 =
    
    open System.IO
    open FsExcel
    
    let britishCultureNativeName = "English (United Kingdom)"
    let ukrainianCultureNativeName = "українська"
    
    let britishCultureDateTimeFormatGetMonthName =
        [ "January"; "February"; "March"; "April"; "May"; "June"; "July";
           "August"; "September"; "October"; "November"; "December" ]
    
    let britishCultureDateTimeFormatAbbreviatedMonthNames =
        [ "Jan"; "Feb"; "Mar"; "Apr"; "May"; "Jun"; "Jul"; "Aug"; "Sep"; "Oct";
          "Nov"; "Dec" ]
    
    let ukrainianCultureDateTimeFormatGetMonthName =
        [ "січень"; "лютий"; "березень"; "квітень"; "травень"; "червень";
          "липень"; "серпень"; "вересень"; "жовтень"; "листопад"; "грудень" ]
    
    let ukrainianCultureDateTimeFormatAbbreviatedMonthNames =
        [ "січ"; "лют"; "бер"; "кві"; "тра"; "чер"; "лип"; "сер"; "вер"; "жов";
          "лис"; "гру" ]
    
    [
        Worksheet britishCultureNativeName
        for m in 0..11 do
            let monthName = britishCultureDateTimeFormatGetMonthName.[m]
            Cell [ String monthName ]
            Cell [ Integer monthName.Length ]
            Go NewRow
    
        Worksheet ukrainianCultureNativeName
        for m in 0..11 do
            let monthName = ukrainianCultureDateTimeFormatGetMonthName.[m]
            Cell [ String monthName ]
            Cell [ Integer monthName.Length ]
            Go NewRow
    
        Worksheet britishCultureNativeName // Switch back to the first worksheet
        Go (RC(13, 1))
        for m in 0..11 do
            let monthAbbreviation = britishCultureDateTimeFormatAbbreviatedMonthNames.[m]
            Cell [ String monthAbbreviation ]
            Cell [ Integer monthAbbreviation.Length ]
            Go NewRow
    
        Worksheet ukrainianCultureNativeName // Switch back to the second worksheet
        Go (RC(13, 1))
        for m in 0..11 do
            let monthAbbreviation = ukrainianCultureDateTimeFormatAbbreviatedMonthNames.[m]
            Cell [ String monthAbbreviation ]
            Cell [ Integer monthAbbreviation.Length ]
            Go NewRow
    ]
    |> Render.AsFile (Path.Combine(savePath, "Worksheets.xlsx"))
    
module Test19 =
    
    open System.IO
    open ClosedXML.Excel
    open FsExcel
    
    let workbook = new XLWorkbook(Path.Combine(savePath, "Worksheets.xlsx"))
    
    let britishCultureNativeName = "English (United Kingdom)"
    let ukrainianCultureNativeName = "українська"
    
    let altMonthNames = [| "Vintagearious"; "Fogarious"; "Frostarious"; "Snowous"; "Rainous"; "Windous"; "Buddal"; "Floweral"; "Meadowal"; "Reapidor"; "Heatidor"; "Fruitidor" |]
    
    [
        Workbook workbook
        Worksheet ukrainianCultureNativeName
        Go(RC(1,3))
        Cell [FormulaA1 $"='{britishCultureNativeName}'!B1*2" ]
        Worksheet britishCultureNativeName
        InsertRowsAbove 12 // The cell reference in the  formula above will be updated to B13
        for m in 0..11 do
            Cell [ String altMonthNames[m] ]
            Cell [ Integer altMonthNames[m].Length ]
            Go NewRow
    ]
    |> Render.AsFile (Path.Combine(savePath, "WorksheetsRevised.xlsx"))
    
module Test20 =
    
    open System.IO
    open System.Globalization
    open FsExcel
    
    [
        for x in 1..12 do
            for y in 0..12 do
                Cell [ Integer (x * y) ]
            Go NewRow
    
        SizeAll (ColWidth 5)
        SizeAll (RowHeight 20)
    ]
    |> Render.AsFile (Path.Combine(savePath, "ColumnWidthRowHeight.xlsx"))
    
module Test21 =
    
    open System.IO
    open System
    open ClosedXML.Excel
    open FsExcel
    
    [   Go NewRow
        for heading, colWidth in ["ID", 3.22; "Car Name", 10.33; "Car Description", 49.33; "Car Registration", 16.89 ] do
            Cell [
                String heading
                FontEmphasis Bold
                FontName "Calibri"
                FontSize 11
                HorizontalAlignment Center
                FontColor (XLColor.FromArgb(0, 255, 255, 255))
                BackgroundColor (XLColor.FromArgb(0, 68, 114, 196))
                Border(Border.All XLBorderStyleValues.Thin)
                CellSize (ColWidth colWidth)
            ]
        Go NewRow
        Cell [  Integer 1
                HorizontalAlignment Center] 
        Cell [  String "Ford Fiesta"]
        Cell [  String "Car Technical Details..."] 
        Cell [  String "AB12 CDE" 
                HorizontalAlignment Center]
    ]
    |> Render.AsFile (Path.Combine(savePath, "IndividualCellSize.xlsx"))
    
module Test22 =
    
    open System.IO
    open System.Globalization
    open FsExcel
    open ClosedXML.Excel
    
    open System.Runtime.InteropServices
    if not (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) then
        LoadOptions.DefaultGraphicEngine <- new ClosedXML.Graphics.DefaultGraphicEngine("Liberation Sans") 
    
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
            let monthName = CultureInfo.GetCultureInfoByIetfLanguageTag("en-GB").DateTimeFormat.GetMonthName(m)
            Cell [ String monthName ]
            Cell [ Integer monthName.Length ]
            Go NewRow
    
        AutoFit AllCols
    ]
    |> Render.AsFile (Path.Combine(savePath, "AutosizeColumns.xlsx"))
    
module Test23 =
    
    open System.IO
    open System
    open ClosedXML.Excel
    open FsExcel
    
    [   Go NewRow
        for heading, colWidth in ["ID", 3.22; "Car Name", 10.33; "Car Description", 49.33; "Car Registration", 16.89 ] do
            Cell [
                String heading
                FontEmphasis Bold
                FontName "Calibri"
                FontSize 11
                HorizontalAlignment Center
                FontColor (XLColor.FromArgb(0, 255, 255, 255))
                BackgroundColor (XLColor.FromArgb(0, 68, 114, 196))
                Border(Border.All XLBorderStyleValues.Thin)
                CellSize (ColWidth colWidth)
            ]
        Go NewRow
        Cell [  Integer 1
                HorizontalAlignment Left
                VerticalAlignment TopMost
                Name "ID" ] 
        Cell [  String "Ford Fiesta"
                HorizontalAlignment Center
                VerticalAlignment Middle ] 
        Cell [  String "Car Technical Details:"
                Next (DownBy 1) ]
        Cell [  String "Technical Detail 1"
                Next (DownBy 1) ]
        Cell [  String "Technical Detail 2"
                Next (DownBy 1)]
        Cell [  String "Technical Detail 3"
                Name "LastL" ]
        Go (RC (3, 4))
        Cell [  String "AB12 CDE" 
                HorizontalAlignment Right
                VerticalAlignment Base
                Name "Reg" ]
        Go (RC (6, 4))
        Cell [Name "RegEnd"]
        Go (RC (7, 3))
        Cell [  String "Another Technical Detail"
                FontEmphasis Italic
                VerticalAlignment Middle
                Name "TD" 
                Next Stay]
        Go (DownBy 1)
        Cell [ Name "info"]
    
        MergeCells ((ColRowLabel ("B", 3), ColRowLabel ("B", 6)))
        MergeCells ((NamedCell "ID", ColRowLabel ("A", 6)))
        MergeCells ((ColRowLabel ("C", 7), NamedCell "info")) 
        MergeCells ((NamedCell "Reg", NamedCell "RegEnd")) 
        
        Go (RC (10, 1))
        Cell [  String "Merging from a starting cell given a depth and span"
                BackgroundColor (XLColor.FromArgb(0, 80, 180, 220))
                FontEmphasis Bold
                HorizontalAlignment Center ] 
        MergeCells ((ColRowLabel ("A", 10), ColRowLabel ("D", 10)))
    
        Go (RC (12, 2))
        Cell [  String "The components that make up a car are: "
                Name "components" 
                HorizontalAlignment Left
                VerticalAlignment TopMost
                Border(Border.All XLBorderStyleValues.MediumDashDot)]
        Go (RC (12, 4))
        Cell [ Border(Border.All XLBorderStyleValues.MediumDashDot)]
        Go (RC (14, 4))
        Cell [ Border(Border.All XLBorderStyleValues.MediumDashDot)]
    
        Go (RC (15, 2))
        Cell [  String "Road Tax"
                HorizontalAlignment Center
                VerticalAlignment Middle
                Border(Border.All XLBorderStyleValues.SlantDashDot)]
        Go (RC (16, 2))
        Cell [ Border(Border.All XLBorderStyleValues.SlantDashDot)]
    
        MergeCells ((NamedCell "components", SpanDepth (3, 3)))
        MergeCells ((ColRowLabel ("B", 15), SpanDepth (1, 2))) 
    
        Go (RC (17, 4))
        Cell [  String "Insurance"
                Name "insurance" // NamedCells cannot begin with a number
                Border(Border.All XLBorderStyleValues.Dashed) ]
        Go (RC (17, 3))
        Cell [ Border(Border.All XLBorderStyleValues.Dashed)]
        Go (RC (17, 2))
        Cell [ Border(Border.All XLBorderStyleValues.Dashed)] 
       
        Go (RC (16, 4))
        Cell [  String "Signature"]
    
        MergeCells ((SpanDepth (3, 1), NamedCell "insurance")) 
        MergeCells ((SpanDepth (2, 2), ColRowLabel ("D", 16))) 
    ]
    |> Render.AsFile (Path.Combine(savePath, "MergeCellsWithVerticalAlignment.xlsx"))
    
module Test24 =
    
    open System
    open System.IO
    open ClosedXML.Excel
    open FsExcel
    
    type JoiningInfo = {
        Name : string
        Age : int
        Fees : decimal
        DateJoined : string
    }
    
    
    let records = [
        { Name = "Jane Smith"; Age = 32; Fees = 59.25m; DateJoined = "2022-03-12" } // Excel will treat these strings as dates
        { Name = "Michael Nguyễn"; Age = 23; Fees = 61.2m; DateJoined = "2022-03-13" }
        { Name = "Sofia Hernández"; Age = 58; Fees = 59.25m; DateJoined = "2022-03-15" }
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
    
module Test25 =
    
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
    
module Test26 =
    
    open System
    open System.IO
    open FsExcel
    
    let headings =
        [ Cell [ String "StringCol"; HorizontalAlignment Center ]
          Cell [ String "IntCol"; HorizontalAlignment Center ]
          Cell [ String "FloatCol"; HorizontalAlignment Center ]
          Cell [ String "DateTimeCol"; HorizontalAlignment Center ]
          Cell [ String "BooleanCol"; HorizontalAlignment Center ]
          Go NewRow ]
    
    let rows =
        [ 1 .. 5 ]
        |> Seq.map(fun i ->
            [ Cell [ String $"String{i}" ]
              Cell [ Integer i ]
              Cell [ Float ((i |> float) + 0.1) ]
              Cell [ DateTime (DateTime.Parse("15-July-2017 05:33:00").AddMinutes(i)) ]
              Cell [ Boolean (i % 2 |> Convert.ToBoolean) ]
              Go NewRow ])
        |> Seq.collect id
        |> List.ofSeq
    
    headings @ rows @ [ AutoFit All; AutoFilter [ EnableOnly RangeUsed ] ]
    |> Render.AsFile (Path.Combine(savePath, "AutoFilterEnableOnly.xlsx"))
    
module Test27 =
    
    open System
    open System.IO
    open FsExcel
    
    let headings =
        [ Cell [ String "StringCol"; HorizontalAlignment Center ]
          Cell [ String "IntCol"; HorizontalAlignment Center ]
          Cell [ String "FloatCol"; HorizontalAlignment Center ]
          Cell [ String "DateTimeCol"; HorizontalAlignment Center ]
          Cell [ String "BooleanCol"; HorizontalAlignment Center ]
          Go NewRow ]
    
    let rows =
        [ 1 .. 5 ]
        |> Seq.map(fun i ->
            [ Cell [ String $"String{i}" ]
              Cell [ Integer i ]
              Cell [ Float ((i |> float) + 0.1) ]
              Cell [ DateTime (DateTime.Parse("15-July-2017 05:33:00").AddMinutes(i)) ]
              Cell [ Boolean (i % 2 |> Convert.ToBoolean) ]
              Go NewRow ])
        |> Seq.collect id
        |> List.ofSeq
    
    headings @ rows @ [ AutoFit All; AutoFilter [ GreaterThanInt (RangeUsed, 2, 3); EqualToBool (RangeUsed, 5, true) ] ]
    |> Render.AsFile (Path.Combine(savePath, "AutoFilterCompound.xlsx"))
    
