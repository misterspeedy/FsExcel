namespace FsExcel

open System
open ClosedXML.Excel

type Position =
    | Row of int
    | Col of int
    | RC of int * int
    | RightBy of int
    | DownBy of int
    | LeftBy of int
    | UpBy of int
    | Indent of int
    | IndentBy of int
    | NewRow
    | Stay

type FontEmphasis =
    | Bold
    | Italic
    | Underline of XLFontUnderlineValues
    | StrikeThrough

type Border =
    | Top of XLBorderStyleValues
    | Right of XLBorderStyleValues
    | Bottom of XLBorderStyleValues
    | Left of XLBorderStyleValues
    | All of XLBorderStyleValues

type BorderColor =
    | Top of XLColor
    | Right of XLColor
    | Bottom of XLColor
    | Left of XLColor
    | All of XLColor

type HorizontalAlignment =
    | Left
    | Center
    | Right

type NameScope =
    | Worksheet
    | Workbook

type CellProp =
    | String of string
    | Float of float
    | Integer of int
    | Boolean of bool
    | DateTime of DateTime
    | TimeSpan of TimeSpan
    | FormulaA1 of string
    | Next of Position
    | FontEmphasis of FontEmphasis
    | FontName of string
    | FontSize of float
    | FontColor of XLColor
    | Border of Border
    | BorderColor of BorderColor
    | BackgroundColor of XLColor
    | HorizontalAlignment of HorizontalAlignment
    | FormatCode of string
    | Name of string
    | ScopedName of name: string * scope: NameScope

module CellProps =

    let hasNext (props : CellProp list) =
        props
        |> List.exists (function | Next _ -> true | _ -> false)

    let sort (props : CellProp list) =
        props
        |> List.sortBy (function
            | Next _ -> 1
            | _ -> 0)

type AutoFit =
    | All
    | ColRange of int * int
    | RowRange of int * int
    | AllCols
    | AllRows

type Size =
    | ColWidth of float
    | RowHeight of float

/// Represents the area of a worksheet to be filtered.
type AutoFilterRange =
    /// The entire range used in the worksheet.
    | RangeUsed
    // The current region around a spcified cell.
    | CurrentRegion of string
    // A specified range.
    | Range of string

type AutoFilter =
    | EnableOnly of AutoFilterRange
    | Clear of AutoFilterRange

    | EqualToString of AutoFilterRange * column : int * value : string
    | EqualToInt of AutoFilterRange * column : int * value : int
    | EqualToFloat of AutoFilterRange * column : int * value : float
    | EqualToDateTime of AutoFilterRange * column : int * value : DateTime
    | EqualToBool of AutoFilterRange * column : int * value : bool

    | NotEqualToString of AutoFilterRange * column : int * value : string
    | NotEqualToInt of AutoFilterRange * column : int * value : int
    | NotEqualToFloat of AutoFilterRange * column : int * value : float
    | NotEqualToDateTime of AutoFilterRange * column : int * value : DateTime
    | NotEqualToBool of AutoFilterRange * column : int * value : bool

    | BetweenInt of AutoFilterRange * column : int * value1 : int * value2 : int
    | BetweenFloat of AutoFilterRange * column : int * value1 : float * value2 : float
    // BetweenDateTime works, but reapplying the filter (CTRL+Alt+L) clears it
    // When looking at the filter in Excel both values are: 07/01/1900
    | BetweenDateTime of AutoFilterRange * column : int * value1 : DateTime * value2 : DateTime

    | NotBetweenInt of AutoFilterRange * column : int * value1 : int * value2 : int
    | NotBetweenFloat of AutoFilterRange * column : int * value1 : float * value2 : float
    | NotBetweenDateTime of AutoFilterRange * column : int * value1 : DateTime * value2 : DateTime

    | ContainsString of AutoFilterRange * column : int * value : string
    | NotContainsString of AutoFilterRange * column : int * value : string

    | BeginsWithString of AutoFilterRange * column : int * value : string
    | NotBeginsWithString of AutoFilterRange * column : int * value : string

    | EndsWithString of AutoFilterRange * column : int * value : string
    | NotEndsWithString of AutoFilterRange * column : int * value : string

    | Top of AutoFilterRange * column : int * value : int * bottomType : XLTopBottomType
    | Bottom of AutoFilterRange * column : int * value : int * bottomType : XLTopBottomType

    | GreaterThanInt of AutoFilterRange * column : int * value : int
    | GreaterThanFloat of AutoFilterRange * column : int * value : float
    | GreaterThanDateTime of AutoFilterRange * column : int * value : DateTime

    | LessThanInt of AutoFilterRange * column : int * value : int
    | LessThanFloat of AutoFilterRange * column : int * value : float
    | LessThanDateTime of AutoFilterRange * column : int * value : DateTime

    | EqualOrGreaterThanInt of AutoFilterRange * column : int * value : int
    | EqualOrGreaterThanFloat of AutoFilterRange * column : int * value : float
    | EqualOrGreaterThanDateTime of AutoFilterRange * column : int * value : DateTime

    | EqualOrLessThanInt of AutoFilterRange * column : int * value : int
    | EqualOrLessThanFloat of AutoFilterRange * column : int * value : float
    | EqualOrLessThanDateTime of AutoFilterRange * column : int * value : DateTime

    | AboveAverage of AutoFilterRange * column : int
    | BelowAverage of AutoFilterRange * column : int

type Item =
    | Cell of props : CellProp list
    | Style of props : CellProp list
    | Go of Position
    | Worksheet of string
    | AutoFit of AutoFit
    | Workbook of XLWorkbook
    | InsertRowsAbove of int
    | Size of Size
    | AutoFilter of AutoFilter list

module Render =

    /// Renders the provided items and returns a ClosedXml XLWorkbook instance.
    let AsWorkBook (items : Item list) =

        let mutable indent = 1
        let mutable r = 1
        let mutable c = 1
        let mutable style : CellProp list = []

        let reset() =
            indent <- 1
            r <- 1
            c <- 1
            style <- []

        let mutable wb = new XLWorkbook()
        let mutable currentWorksheet : IXLWorksheet option = None

        let getCurrentWorksheet() =
            currentWorksheet
            |> Option.defaultWith (fun _ ->
                let newWorksheet = wb.Worksheets.Add("Sheet1")
                currentWorksheet <- newWorksheet |> Some
                newWorksheet)

        let go = function
            | Row row ->
                r <- row |> max 1
            | Col col ->
                c <- col |> max 1
            | RC (row, col) ->
                r <- row |> max 1
                c <- col |> max 1
            | RightBy n ->
                c <- c+n
            | DownBy n ->
                r <- r+n
            | UpBy n ->
                r <- r-n |> max 1
            | LeftBy n ->
                c <- c-n |> max 1
            | Indent n ->
                indent <- n |> max 1
                c <- indent
            | IndentBy n ->
                indent <- indent + n |> max 1
                c <- indent
            | NewRow ->
                r <- r + 1
                c <- indent
            | Stay ->
                ()

        // let processAutoFilter (ws : IXLWorksheet) (item : Item) =
        let processAutoFilter (ws : IXLWorksheet) (autoFilters : AutoFilter list) =

            let getRange (autoFilterRange : AutoFilterRange) =
                match autoFilterRange with
                | RangeUsed ->
                    ws.RangeUsed()
                | CurrentRegion c ->
                    ws.Cell(c).CurrentRegion
                | Range r ->
                    ws.Range(r)

            let doIt = function
                | EnableOnly a ->
                    (getRange a).SetAutoFilter() |> ignore
                | Clear a ->
                    (getRange a).SetAutoFilter().Clear() |> ignore

                | EqualToString (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).EqualTo(c) |> ignore
                | EqualToInt (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).EqualTo(c) |> ignore
                | EqualToFloat (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).EqualTo(c) |> ignore
                | EqualToDateTime (a, b, c) ->
                    // This is needed for dates
                    // https://github.com/ClosedXML/ClosedXML/issues/701
                    // TODO: I'm not sure grouping by seconds would work in all cases.
                    // TODO: To be on the safe side the grouping would have to be passed in.
                    // TODO: This would not be nice.
                    (getRange a).SetAutoFilter().Column(b).AddDateGroupFilter(c, XLDateTimeGrouping.Second) |> ignore
                | EqualToBool (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).EqualTo((c.ToString())) |> ignore

                | NotEqualToString (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).NotEqualTo(c) |> ignore
                | NotEqualToInt (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).NotEqualTo(c) |> ignore
                | NotEqualToFloat (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).NotEqualTo(c) |> ignore
                | NotEqualToDateTime (a, b, c) ->
                    // This is needed for dates
                    // https://github.com/ClosedXML/ClosedXML/issues/701
                    // TODO: Does not work!
                    (getRange a).SetAutoFilter().Column(b).NotEqualTo(c) |> ignore
                | NotEqualToBool (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).NotEqualTo((c.ToString())) |> ignore

                | BetweenInt (a, b, c, d) ->
                    (getRange a).SetAutoFilter().Column(b).Between(c, d) |> ignore
                | BetweenFloat (a, b, c, d) ->
                    (getRange a).SetAutoFilter().Column(b).Between(c, d) |> ignore
                | BetweenDateTime (a, b, c, d) ->
                    (getRange a).SetAutoFilter().Column(b).Between(c, d) |> ignore

                | NotBetweenInt (a, b, c, d) ->
                    (getRange a).SetAutoFilter().Column(b).NotBetween(c, d) |> ignore
                | NotBetweenFloat (a, b, c, d) ->
                    (getRange a).SetAutoFilter().Column(b).NotBetween(c, d) |> ignore
                | NotBetweenDateTime (a, b, c, d) ->
                    (getRange a).SetAutoFilter().Column(b).NotBetween(c, d) |> ignore

                | ContainsString (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).Contains(c) |> ignore
                | NotContainsString (a, b, c) ->
                    // Works but appears as a Contains filter
                    (getRange a).SetAutoFilter().Column(b).NotContains(c) |> ignore

                | BeginsWithString (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).BeginsWith(c) |> ignore
                | NotBeginsWithString (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).NotBeginsWith(c) |> ignore

                | EndsWithString (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).EndsWith(c) |> ignore
                | NotEndsWithString (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).NotEndsWith(c) |> ignore

                // Top does not work with String, DateTime and Boolean
                | Top (a, b, c, d) ->
                    (getRange a).SetAutoFilter().Column(b).Top(c, d) |> ignore
                | Bottom (a, b, c, d) ->
                    (getRange a).SetAutoFilter().Column(b).Bottom(c, d) |> ignore

                | GreaterThanInt (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).GreaterThan(c) |> ignore
                | GreaterThanFloat (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).GreaterThan(c) |> ignore
                | GreaterThanDateTime (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).GreaterThan(c) |> ignore

                | LessThanInt (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).LessThan(c) |> ignore
                | LessThanFloat (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).LessThan(c) |> ignore
                | LessThanDateTime (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).LessThan(c) |> ignore

                | EqualOrGreaterThanInt (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).EqualOrGreaterThan(c) |> ignore
                | EqualOrGreaterThanFloat (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).EqualOrGreaterThan(c) |> ignore
                | EqualOrGreaterThanDateTime (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).EqualOrGreaterThan(c) |> ignore

                | EqualOrLessThanInt (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).EqualOrLessThan(c) |> ignore
                | EqualOrLessThanFloat (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).EqualOrLessThan(c) |> ignore
                | EqualOrLessThanDateTime (a, b, c) ->
                    (getRange a).SetAutoFilter().Column(b).EqualOrLessThan(c) |> ignore

                | AboveAverage (a, b) ->
                    (getRange a).SetAutoFilter().Column(b).AboveAverage() |> ignore
                | BelowAverage (a, b) ->
                    (getRange a).SetAutoFilter().Column(b).BelowAverage() |> ignore

            autoFilters |> List.iter doIt

        for item in items do

            match item with
            | Workbook workbook ->
                if currentWorksheet.IsNone
                then
                    wb <- workbook
                    currentWorksheet <- workbook.Worksheet(1) |> Some // This is the first worksheet (one-based array)
                    reset()
            | Worksheet name ->
                if wb.Worksheets.Contains(name)
                then currentWorksheet <- wb.Worksheet(name) |> Some
                else currentWorksheet <- wb.Worksheets.Add(name) |> Some
                reset()
            | InsertRowsAbove rs ->
                currentWorksheet.Value.Row(r).InsertRowsAbove(rs)   |> ignore
            | Go p ->
                go p
            | Cell props ->

                let ws = getCurrentWorksheet()

                let props =
                    if props |> CellProps.hasNext |> not then
                        Next(RightBy 1) :: props
                    else
                        props
                    |> fun ps -> style @ ps
                    // Ensure Next() props are applied after filling content.
                    |> CellProps.sort

                for prop in props do

                    let cell = ws.Cell(r, c)

                    match prop with
                    | String s ->
                        cell.Value <- s
                    | Float f ->
                        cell.Value <- f
                    | Integer i ->
                        cell.Value <- i
                    | Boolean b ->
                        cell.Value <- b
                    | DateTime dt ->
                        cell.Value <- dt
                    | TimeSpan ts ->
                        cell.Value <- ts
                    | FormulaA1 s ->
                        cell.FormulaA1 <- s
                    | Next p ->
                        go p
                    | FontEmphasis fe ->
                        match fe with
                        | FontEmphasis.Bold ->
                            cell.Style.Font.Bold <- true
                        | FontEmphasis.Italic ->
                            cell.Style.Font.Italic <- true
                        | FontEmphasis.Underline v ->
                            cell.Style.Font.Underline <- v
                        | FontEmphasis.StrikeThrough ->
                            cell.Style.Font.Strikethrough <- true
                    | FontName s ->
                        cell.Style.Font.FontName <- s
                    | FontSize x ->
                        cell.Style.Font.FontSize <- x
                    | BorderColor bc ->
                        match bc with
                        | BorderColor.Top c ->
                            cell.Style.Border.TopBorderColor <- c
                        | BorderColor.Right c ->
                            cell.Style.Border.RightBorderColor <- c
                        | BorderColor.Bottom c ->
                            cell.Style.Border.BottomBorderColor <- c
                        | BorderColor.Left c ->
                            cell.Style.Border.LeftBorderColor <- c
                        | BorderColor.All c ->
                            cell.Style.Border.TopBorderColor <- c
                            cell.Style.Border.RightBorderColor <- c
                            cell.Style.Border.BottomBorderColor <- c
                            cell.Style.Border.LeftBorderColor <- c
                    | Border b ->
                        match b with
                        | Border.Top style ->
                            cell.Style.Border.TopBorder <- style
                        | Border.Right style ->
                            cell.Style.Border.RightBorder <- style
                        | Border.Bottom style ->
                            cell.Style.Border.BottomBorder <- style
                        | Border.Left style ->
                            cell.Style.Border.LeftBorder <- style
                        | Border.All style ->
                            cell.Style.Border.TopBorder <- style
                            cell.Style.Border.RightBorder <- style
                            cell.Style.Border.BottomBorder <- style
                            cell.Style.Border.LeftBorder <- style
                    | BackgroundColor c ->
                        cell.Style.Fill.BackgroundColor <- c
                    | FontColor c ->
                        cell.Style.Font.FontColor <- c
                    | HorizontalAlignment h ->
                        match h with
                        | Left ->
                            cell.Style.Alignment.Horizontal <- XLAlignmentHorizontalValues.Left
                        | Center ->
                            cell.Style.Alignment.Horizontal <- XLAlignmentHorizontalValues.Center
                        | Right ->
                            cell.Style.Alignment.Horizontal <- XLAlignmentHorizontalValues.Right
                    | FormatCode fc ->
                        cell.Style.NumberFormat.Format <- fc
                    | Name name ->
                        cell.AddToNamed(name, XLScope.Worksheet) |> ignore
                    | ScopedName (name, scope) ->
                        let xlScope =
                            match scope with
                            | NameScope.Worksheet -> XLScope.Worksheet
                            | NameScope.Workbook -> XLScope.Workbook
                        cell.AddToNamed(name, xlScope) |> ignore
            | AutoFit af ->
                let ws = getCurrentWorksheet()

                match af with
                | All ->
                    ws.Columns().AdjustToContents() |> ignore
                    ws.Rows().AdjustToContents() |> ignore
                | ColRange (a, b) ->
                    ws.Columns(a, b).AdjustToContents() |> ignore
                | RowRange (a, b) ->
                    ws.Rows(a, b).AdjustToContents() |> ignore
                | AllCols ->
                    ws.Columns().AdjustToContents() |> ignore
                | AllRows ->
                    ws.Rows().AdjustToContents() |> ignore
            | Size s ->
                let ws = getCurrentWorksheet()

                match s with
                | ColWidth width ->
                    ws.Columns().Width <- width
                | RowHeight height ->
                    ws.Rows().Height <- height
            | Style s ->
                style <- s
            | AutoFilter autoFilter ->
                let ws = getCurrentWorksheet()

                processAutoFilter ws autoFilter

        wb

    /// Renders the provided items and saves the resulting workbook as a file. The provided path
    /// must have the extension '.xlsx'.
    let AsFile (path : string) (items : Item list) =
        items
        |> AsWorkBook
        |> fun wb -> wb.SaveAs path

    /// Renders the provided items and writes the result into the provided stream.
    let AsStream (stream : IO.Stream) (items : Item list) =
        items
        |> AsWorkBook
        |> fun wb ->
            wb.SaveAs(stream)

    /// Renders the provided items and returns the resulting Excel workbook as an array of bytes. This
    /// array can be provided as a browser download in Web App scenarios.
    let AsStreamBytes (items : Item list) =
        use stream = new IO.MemoryStream()
        items |> AsStream stream
        let bytes = stream.ToArray()
        bytes

    ///  Renders a workbook as a set of HTML tables.
    ///  This is primarily for use in Dotnet Interactive Notebooks, where you can use the `HTML()` helper
    ///  method to display the resulting HTML.
    //
    // TODO
    // - Alignment
    // - Formatting
    // - Use worksheet tabs for something
    let private buildStyle (cell : IXLCell) =
        let fontWeight = if cell.Style.Font.Bold then "bold" else "normal"
        let fontStyle = if cell.Style.Font.Italic then "italic" else "normal"
        // TODO could be more faithful here
        let borderLeft = if cell.Style.Border.LeftBorder <> XLBorderStyleValues.None then "thin solid" else "none"
        let borderTop = if cell.Style.Border.TopBorder <> XLBorderStyleValues.None then "thin solid" else "none"
        let borderRight = if cell.Style.Border.RightBorder <> XLBorderStyleValues.None then "thin solid" else "none"
        let borderBottom = if cell.Style.Border.BottomBorder <> XLBorderStyleValues.None then "thin solid" else "none"
        let textDecoration = if cell.Style.Font.Underline <> XLFontUnderlineValues.None then "underline" else "none"
        $"\"font-weight: {fontWeight}; font-style: {fontStyle}; text-decoration: {textDecoration}; border-left : {borderLeft}; border-top: {borderTop}; border-right: {borderRight}; border-bottom: {borderBottom};\""

    let AsHtml isHeader (items : Item list) =
        items
        |> AsWorkBook
        |> fun wb ->
        [
            for ws in wb.Worksheets do
                $"<h3>{ws.Name}</h3>"
                "<table>"
                for rowIndex, row in [(ws.FirstRowUsed().RowNumber())..(ws.LastRowUsed().RowNumber())] |> List.indexed do
                    "<tr>"
                    for colIndex, col in [(ws.FirstColumnUsed().ColumnNumber())..(ws.LastColumnUsed().ColumnNumber())] |> List.indexed do
                        let cell = ws.Cell(row, col)
                        let style = cell |> buildStyle
                        let isHeader = isHeader rowIndex colIndex
                        if isHeader then
                            $"<th style = {style}>"
                        else
                            $"<td style = {style}>"
                        cell.GetFormattedString()
                        if isHeader then
                            "</th>"
                        else
                            "</td>"
                    "</tr>"
                "</table>"
        ]
        |> fun strings ->
            String.Join(Environment.NewLine, strings)

