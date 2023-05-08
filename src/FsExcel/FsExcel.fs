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

type VerticalAlignment =
    | Base
    | Middle
    | TopMost

type NameScope =
    | Worksheet
    | Workbook

type Size =
    | ColWidth of float
    | RowHeight of float

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
    | VerticalAlignment of VerticalAlignment
    | TextRotation of degrees:int
    | WrapText of bool
    | FormatCode of string
    | Name of string
    | ScopedName of name: string * scope: NameScope
    | CellSize of Size

module CellProps =

    let hasNext (props : CellProp list) =
        props
        |> List.exists (function | Next _ -> true | _ -> false)

    let sort (props : CellProp list) =
        props
        |> List.sortBy (function
            | Next _ -> 1
            | _ -> 0)

/// There are three ways to reference a cell. By its label e.g A12 or name e.g. "apple" (if the cell has been previously named) or by a cell's span and depth
type CellLabel = 
    /// Column and Row label
    | ColRowLabel of Col:string * Row:int
    /// This is used for referencing named cells
    | NamedCell of string
    /// This identifies the column span (e.g. 2 columns wide) and row depth (e.g. 3 rows deep) of a merged cell
    | SpanDepth of ColSpan:int * RowDepth: int

// no need to specify RowMerge (span vertically), ColumnMerge (span horizontally), RowColumnMerge (box)
// Merge() + range takes care of it, no seperate definitions needed

type AutoFit =
    | All
    | ColRange of int * int
    | RowRange of int * int
    | AllCols
    | AllRows

/// Represents the area of a worksheet to be filtered.
type AutoFilterRange =
    /// The entire range used in the worksheet.
    | RangeUsed
    /// The current region around a spcified cell.
    | CurrentRegion of string
    /// A specified range.
    | Range of string

/// The filters available to be used with AutoFilter.
type AutoFilter =
    /// Enable AutoFilter but do not apply any filters.
    | EnableOnly of AutoFilterRange
    /// Clear an existing AutoFilter.
    | Clear of AutoFilterRange

    /// <summary>
    /// Filter the range by the column being equal to the string value.
    ///
    /// Example:
    ///
    /// EqualToString (RangeUsed, 1, "String3")
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A string used to filter the range.</param>
    | EqualToString of range : AutoFilterRange * column : int * value : string
    /// <summary>
    /// Filter the range by the column being equal to the integer value.
    ///
    /// Example:
    ///
    /// EqualToInt (CurrentRegion, 2, 42)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">An integer used to filter the range.</param>
    | EqualToInt of range : AutoFilterRange * column : int * value : int
    /// <summary>
    /// Filter the range by the column being equal to the float value.
    ///
    /// Example:
    ///
    /// EqualToFloat ("A1:E6", 3, 4.2)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A float used to filter the range.</param>
    | EqualToFloat of range : AutoFilterRange * column : int * value : float
    /// <summary>
    /// Filter the range by the column being equal to the DateTime value.
    ///
    /// Example:
    ///
    /// EqualToDateTime (RangeUsed, 4, DateTime.Parse("15-July-2017 05:34:00"))
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A DateTime used to filter the range.</param>
    | EqualToDateTime of range : AutoFilterRange * column : int * value : DateTime
    /// <summary>
    /// Filter the range by the column being equal to the boolean value.
    ///
    /// Example:
    ///
    /// EqualToBool (CurrentRegion, 5, true)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A boolean value used to filter the range.</param>
    | EqualToBool of range : AutoFilterRange * column : int * value : bool

    /// <summary>
    /// Filter the range by the column being not equal to the string value.
    ///
    /// Example:
    ///
    /// NotEqualToString ("A1:E6", 1, "String3")
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A string value used to filter the range.</param>
    | NotEqualToString of range : AutoFilterRange * column : int * value : string
    /// <summary>
    /// Filter the range by the column being not equal to the integer value.
    ///
    /// Example:
    ///
    /// NotEqualToInt (RangeUsed, 2, 42)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">An integer value used to filter the range.</param>
    | NotEqualToInt of range : AutoFilterRange * column : int * value : int
    /// <summary>
    /// Filter the range by the column being not equal to the float value.
    ///
    /// Example:
    ///
    /// NotEqualToFloat (CurrentRegion, 3, 4.2)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A float value used to filter the range.</param>
    | NotEqualToFloat of range : AutoFilterRange * column : int * value : float
    /// <summary>
    /// Filter the range by the column being not equal to the DateTime value.
    ///
    /// Example:
    ///
    /// NotEqualToDateTime ("A1:E6", 4, DateTime.Parse("15-July-2017 05:34:00"))
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A DateTime value used to filter the range.</param>
    | NotEqualToDateTime of range : AutoFilterRange * column : int * value : DateTime
    /// <summary>
    /// Filter the range by the column being not equal to the boolean value.
    ///
    /// Example:
    ///
    /// NotEqualToBool (RangeUsed, 5, true)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A boolean value used to filter the range.</param>
    | NotEqualToBool of range : AutoFilterRange * column : int * value : bool

    /// <summary>
    /// Filter the range by the column being between the integer values.
    ///
    /// Example:
    ///
    /// BetweenInt (CurrentRegion, 2, 5, 10)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="min">An integer value used to filter the range.</param>
    /// <param name="max">An integer value used to filter the range.</param>
    | BetweenInt of range : AutoFilterRange * column : int * min : int * max : int
    /// <summary>
    /// Filter the range by the column being between the float values.
    ///
    /// Example:
    ///
    /// BetweenFloat ("A1:E6", 3, 1.5, 6.3)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="min">A float value used to filter the range.</param>
    /// <param name="max">A float value used to filter the range.</param>
    | BetweenFloat of range : AutoFilterRange * column : int * min : float * max : float
    // BetweenDateTime works, but reapplying the filter (CTRL+Alt+L) clears it
    // When looking at the filter in Excel both values are: 07/01/1900
    /// <summary>
    /// Filter the range by the column being between the DateTime values.
    ///
    /// Example:
    ///
    /// let dtFrom = DateTime.Parse("15-July-2017")
    ///
    /// let dtTo = DateTime.Parse("14-July-2018")
    ///
    /// BetweenDateTime ("A1:E6", 5, dtFrom, dtTo)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="min">A DateTime value used to filter the range.</param>
    /// <param name="max">A DateTime value used to filter the range.</param>
    | BetweenDateTime of range : AutoFilterRange * column : int * min : DateTime * max : DateTime

    /// <summary>
    /// Filter the range by the column being not between the integer values.
    ///
    /// Example:
    ///
    /// NotBetweenInt (RangeUsed, 2, 5, 10)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="min">An integer value used to filter the range.</param>
    /// <param name="max">An integer value used to filter the range.</param>
    | NotBetweenInt of range : AutoFilterRange * column : int * min : int * max : int
    /// <summary>
    /// Filter the range by the column being not between the float values.
    ///
    /// Example:
    ///
    /// NotBetweenFloat (CurrentRegion, 3, 1.5, 6.3)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="min">A float value used to filter the range.</param>
    /// <param name="max">A float value used to filter the range.</param>
    | NotBetweenFloat of range : AutoFilterRange * column : int * min : float * max : float
    /// <summary>
    /// Filter the range by the column being not between the DateTime values.
    ///
    /// Example:
    ///
    /// let dtFrom = DateTime.Parse("15-July-2017")
    ///
    /// let dtTo = DateTime.Parse("14-July-2018")
    ///
    /// NotBetweenDateTime ("A1:E6", 5, dtFrom, dtTo)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="min">A DateTime value used to filter the range.</param>
    /// <param name="max">A DateTime value used to filter the range.</param>
    | NotBetweenDateTime of range : AutoFilterRange * column : int * min : DateTime * max : DateTime

    /// <summary>
    /// Filter the range by the column containing the string value.
    ///
    /// Example:
    ///
    /// ContainsString (RangeUsed, 1, "and")
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A string value used to filter the range.</param>
    | ContainsString of range : AutoFilterRange * column : int * value : string
    /// <summary>
    /// Filter the range by the column not containing the string value.
    ///
    /// Example:
    ///
    /// NotContainsString (RangeUsed, 1, "and")
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A string value used to filter the range.</param>
    | NotContainsString of range : AutoFilterRange * column : int * value : string

    /// <summary>
    /// Filter the range by the column beginning with the string value.
    ///
    /// Example:
    ///
    /// BeginsWithString (CurrentRegion, 1, "Start")
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A string value used to filter the range.</param>
    | BeginsWithString of range : AutoFilterRange * column : int * value : string
    /// <summary>
    /// Filter the range by the column not beginning with the string value.
    ///
    /// Example:
    ///
    /// NotBeginsWithString ("A1:E6", 1, "Start")
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A string value used to filter the range.</param>
    | NotBeginsWithString of range : AutoFilterRange * column : int * value : string

    /// <summary>
    /// Filter the range by the column ending with the string value.
    ///
    /// Example:
    ///
    /// EndsWithString (RangeUsed, 1, "ending")
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A string value used to filter the range.</param>
    | EndsWithString of range : AutoFilterRange * column : int * value : string
    /// <summary>
    /// Filter the range by the column not ending with the string value.
    ///
    /// Example:
    ///
    /// NotEndsWithString (CurrentRegion, 1, "ending")
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A string value used to filter the range.</param>
    | NotEndsWithString of range : AutoFilterRange * column : int * value : string

    /// <summary>
    /// Filter the range by the columns top n values.
    ///
    /// Examples:
    ///
    /// Top ("A1:E6", 2, 5, XLTopBottomType.Items)
    ///
    /// Top ("A1:E6", 2, 20, XLTopBottomType.Percent)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">An integer representing the number/percent of rows.</param>
    /// <param name="topType">An XLTopBottomType value used to filter the range.</param>
    | Top of range : AutoFilterRange * column : int * value : int * topType : XLTopBottomType
    /// <summary>
    /// Filter the range by the columns bottom n values.
    ///
    /// Examples:
    ///
    /// Bottom (RangeUsed, 2, 5, XLTopBottomType.Items)
    ///
    /// Bottom (RangeUsed, 2, 20, XLTopBottomType.Percent)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">An integer representing the number/percent of rows.</param>
    /// <param name="bottomType">An XLTopBottomType value used to filter the range.</param>
    | Bottom of range : AutoFilterRange * column : int * value : int * bottomType : XLTopBottomType

    /// <summary>
    /// Filter the range by the column being greater than the integer value.
    ///
    /// Example:
    ///
    /// GreaterThanInt (CurrentRegion, 2, 3)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">An integer value used to filter the range.</param>
    | GreaterThanInt of range : AutoFilterRange * column : int * value : int
    /// <summary>
    /// Filter the range by the column being greater than the float value.
    ///
    /// Example:
    ///
    /// GreaterThanFloat (CurrentRegion, 3, 3.5)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A float value used to filter the range.</param>
    | GreaterThanFloat of range : AutoFilterRange * column : int * value : float
    /// <summary>
    /// Filter the range by the column being greater than the DateTime value.
    ///
    /// Example:
    ///
    /// GreaterThanDateTime ("A1:E6", 4, DateTime.Parse("15-July-2017 05:36:00"))
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A DateTime value used to filter the range.</param>
    | GreaterThanDateTime of range : AutoFilterRange * column : int * value : DateTime

    /// <summary>
    /// Filter the range by the column being less than the integer value.
    ///
    /// Example:
    ///
    /// LessThanInt (CurrentRegion, 2, 3)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">An integer value used to filter the range.</param>
    | LessThanInt of range : AutoFilterRange * column : int * value : int
    /// <summary>
    /// Filter the range by the column being less than the float value.
    ///
    /// Example:
    ///
    /// LessThanFloat (CurrentRegion, 3, 3.5)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A float value used to filter the range.</param>
    | LessThanFloat of range : AutoFilterRange * column : int * value : float
    /// <summary>
    /// Filter the range by the column being less than the DateTime value.
    ///
    /// Example:
    ///
    /// LessThanDateTime ("A1:E6", 4, DateTime.Parse("15-July-2017 05:36:00"))
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A DateTime value used to filter the range.</param>
    | LessThanDateTime of range : AutoFilterRange * column : int * value : DateTime

    /// <summary>
    /// Filter the range by the column being greater than or equal to the integer value.
    ///
    /// Example:
    ///
    /// EqualOrGreaterThanInt (CurrentRegion, 2, 3)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">An integer value used to filter the range.</param>
    | EqualOrGreaterThanInt of range : AutoFilterRange * column : int * value : int
    /// <summary>
    /// Filter the range by the column being greater than or equal to the float value.
    ///
    /// Example:
    ///
    /// EqualOrGreaterThanFloat (CurrentRegion, 3, 3.5)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A float value used to filter the range.</param>
    | EqualOrGreaterThanFloat of range : AutoFilterRange * column : int * value : float
    /// <summary>
    /// Filter the range by the column being greater than or equal to the DateTime value.
    ///
    /// Example:
    ///
    /// EqualOrGreaterThanDateTime ("A1:E6", 4, DateTime.Parse("15-July-2017 05:36:00"))
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A DateTime value used to filter the range.</param>
    | EqualOrGreaterThanDateTime of range : AutoFilterRange * column : int * value : DateTime

    /// <summary>
    /// Filter the range by the column being less than or equal to the integer value.
    ///
    /// Example:
    ///
    /// EqualOrLessThanInt (CurrentRegion, 2, 3)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">An integer value used to filter the range.</param>
    | EqualOrLessThanInt of range : AutoFilterRange * column : int * value : int
    /// <summary>
    /// Filter the range by the column being less than or equal to the float value.
    ///
    /// Example:
    ///
    /// EqualOrLessThanFloat (CurrentRegion, 3, 3.5)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A float value used to filter the range.</param>
    | EqualOrLessThanFloat of range : AutoFilterRange * column : int * value : float
    /// <summary>
    /// Filter the range by the column being less than or equal to the DateTime value.
    ///
    /// Example:
    ///
    /// EqualOrLessThanDateTime ("A1:E6", 4, DateTime.Parse("15-July-2017 05:36:00"))
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    /// <param name="value">A DateTime value used to filter the range.</param>
    | EqualOrLessThanDateTime of range : AutoFilterRange * column : int * value : DateTime

    /// <summary>
    /// Filter the range by the column being above the average of the integer/float values.
    ///
    /// Example:
    ///
    /// AboveAverage (RangeUsed, 2)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    | AboveAverage of range : AutoFilterRange * column : int
    /// <summary>
    /// Filter the range by the column being below the average of the integer/float values.
    ///
    /// Example:
    ///
    /// BelowAverage (RangeUsed, 2)
    /// </summary>
    /// <param name="range">The range to be filtered.</param>
    /// <param name="column">The column number within the range to be filtered.</param>
    | BelowAverage of range : AutoFilterRange * column : int

type StyleMergedCell =
    | BorderType of Border
    | ColorBorder of BorderColor

type TotalsRowItem =
    | Label of string
    | Function of XLTotalsRowFunction
    | CustomA1 of string

type TableProperty =
    | TableName of string
    | Items of obj list
    | Theme of XLTableTheme 
    | ShowHeaderRow of bool
    | ShowRowStripes of bool
    | ShowColumnStripes of bool
    | EmphasizeFirstColumn of bool
    | EmphasizeLastColumn of bool
    | ShowAutoFilter of bool
    | Totals of List<string * TotalsRowItem>
    | ColFormulas of List<string * string>
    | ColFormulae of List<string * string>

[<AutoOpen>]
module TableProperty =
    let TableItems<'T> (items :'T list) =
        items |> List.map box |> TableProperty.Items

type Item =
    | Cell of props : CellProp list
    | Style of props : CellProp list
    | BorderMergedCell of borderProps : StyleMergedCell list
    | Go of Position
    | Worksheet of string
    | AutoFit of AutoFit
    | Workbook of XLWorkbook
    | InsertRowsAbove of int
    | SizeAll of Size 
    | MergeCells of c1:CellLabel * c2:CellLabel
    | AutoFilter of AutoFilter list
    | Table of TableProperty list

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

            let getRange (range : AutoFilterRange) =
                match range with
                | RangeUsed ->
                    ws.RangeUsed()
                | CurrentRegion c ->
                    ws.Cell(c).CurrentRegion
                | Range r ->
                    ws.Range(r)

            let setFilter = function
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
                    (getRange a).SetAutoFilter().Column(b).EqualTo((c.ToString().ToUpper())) |> ignore

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
                    (getRange a).SetAutoFilter().Column(b).NotEqualTo((c.ToString().ToUpper())) |> ignore

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

            autoFilters |> List.iter setFilter

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
                    | VerticalAlignment v ->
                         match v with
                         | Base ->
                             cell.Style.Alignment.Vertical <- XLAlignmentVerticalValues.Bottom
                         | Middle ->
                            cell.Style.Alignment.Vertical <- XLAlignmentVerticalValues.Center
                         | TopMost ->
                            cell.Style.Alignment.Vertical <- XLAlignmentVerticalValues.Top
                    | TextRotation degrees ->
                        if degrees >= -90 && degrees <= 90 then
                            cell.Style.Alignment.TextRotation <- degrees
                        else
                            raise <| ArgumentException $"Invalid TextRotation: {degrees}"
                    | WrapText wt->
                        cell.Style.Alignment.WrapText <- wt
                    | FormatCode fc ->
                        cell.Style.NumberFormat.Format <- fc
                    | CellProp.Name name ->
                        cell.AddToNamed(name, XLScope.Worksheet) |> ignore
                    | ScopedName (name, scope) ->
                        let xlScope =
                            match scope with
                            | NameScope.Worksheet -> XLScope.Worksheet
                            | NameScope.Workbook -> XLScope.Workbook
                        cell.AddToNamed(name, xlScope) |> ignore
                    | CellSize s ->
                        match s with
                        | ColWidth width ->
                            cell.WorksheetColumn().Width <- width
                        | RowHeight height ->
                            cell.WorksheetRow().Height <- height
            | MergeCells (c1, c2) ->
                let ws = getCurrentWorksheet()
                let crToStr (c : string, r : int) = c + string(r)            
                match c1, c2 with
                | (ColRowLabel (cSt, rSt), ColRowLabel (cE, rE)) -> 
                    let range = crToStr (cSt, rSt) + ":" + crToStr (cE, rE)
                    ws.Range(range).Merge() |> ignore
                | (ColRowLabel (cSt, rSt), NamedCell cell2) -> 
                    let range = crToStr (cSt, rSt) + ":" + crToStr (CellReference.namedCellToCR cell2 ws)
                    ws.Range(range).Merge() |> ignore
                | (NamedCell cell1, ColRowLabel (cE, rE)) -> 
                    let range = crToStr (CellReference.namedCellToCR cell1 ws) + ":" + crToStr (cE, rE)
                    ws.Range(range).Merge() |> ignore 
                | (NamedCell cell1, NamedCell cell2) ->
                    let range = crToStr (CellReference.namedCellToCR cell1 ws) + ":" + crToStr (CellReference.namedCellToCR cell2 ws)
                    ws.Range(range).Merge() |> ignore
                | (NamedCell cell1, SpanDepth (colSpan, rowDepth)) ->
                    let cell = (CellReference.namedCellToCR cell1 ws)
                    let range = crToStr (CellReference.namedCellToCR cell1 ws) + ":" + crToStr (CellReference.spanDepthToCellReference cell colSpan rowDepth)
                    ws.Range(range).Merge() |> ignore
                | (ColRowLabel (cSt, rSt), SpanDepth (colSpan, rowDepth)) ->
                    let range = crToStr (cSt, rSt) + ":" + crToStr (CellReference.spanDepthToCellReference (cSt, rSt) colSpan rowDepth)
                    ws.Range(range).Merge() |> ignore
                | (SpanDepth (colSpan, rowDepth), NamedCell cell2) ->
                    let cell = (CellReference.namedCellToCR cell2 ws)
                    let range = crToStr (CellReference.cellReverseSpanDepthToCR cell colSpan rowDepth) + ":" + crToStr (CellReference.namedCellToCR cell2 ws)
                    ws.Range(range).Merge() |> ignore
                | (SpanDepth (colSpan, rowDepth), ColRowLabel (cE, rE)) ->
                    let range = crToStr (CellReference.cellReverseSpanDepthToCR (cE, rE) colSpan rowDepth) + ":" + crToStr (cE, rE)
                    ws.Range(range).Merge() |> ignore
                | (SpanDepth (span, depth), SpanDepth (colSpan, rowDepth)) ->
                    ws |> ignore // ignore this case: cannot merge between two arbitary 
                // TODO: ideally, want to ignore the incomplete pattern match altogether to prevent user trying this option

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
            | SizeAll s ->
                let ws = getCurrentWorksheet()
                match s with
                | ColWidth width ->
                    ws.Columns().Width <- width
                | RowHeight height ->
                    ws.Rows().Height <- height
            | Style s ->
                style <- s
            | BorderMergedCell style ->
                let ws = getCurrentWorksheet()
                for borderStyle in style do
                    match borderStyle with
                    | BorderType borderstyle ->
                        match borderstyle with
                            | Border.Top style ->
                                ws.Cells().Style.Border.TopBorder <- style
                            | Border.Right style ->
                                ws.Cells().Style.Border.RightBorder <- style
                            | Border.Bottom style ->
                                ws.Cells().Style.Border.BottomBorder <- style
                            | Border.Left style ->
                                ws.Cells().Style.Border.LeftBorder <- style
                            | Border.All style ->
                                ws.Cells().Style.Border.TopBorder <- style
                                ws.Cells().Style.Border.RightBorder <- style
                                ws.Cells().Style.Border.BottomBorder <- style
                                ws.Cells().Style.Border.LeftBorder <- style
                    | ColorBorder bordercolour ->
                        match bordercolour with
                            | BorderColor.Top c ->
                                ws.Cells().Style.Border.TopBorderColor <- c
                            | BorderColor.Right c ->
                                ws.Cells().Style.Border.RightBorderColor <- c
                            | BorderColor.Bottom c ->
                                ws.Cells().Style.Border.BottomBorderColor <- c
                            | BorderColor.Left c ->
                                ws.Cells().Style.Border.LeftBorderColor <- c
                            | BorderColor.All c ->
                                ws.Cells().Style.Border.TopBorderColor <- c
                                ws.Cells().Style.Border.RightBorderColor <- c
                                ws.Cells().Style.Border.BottomBorderColor <- c
                                ws.Cells().Style.Border.LeftBorderColor <- c
            
            | AutoFilter autoFilter ->
                let ws = getCurrentWorksheet()

                processAutoFilter ws autoFilter

            | Table properties ->
                // TODO does this have to be repeated so much?:
                let ws = getCurrentWorksheet()

                let name =
                    properties
                    |> List.rev // "Obey the last order first"
                    |> List.tryPick (function | TableName name -> Some name | _ -> None)
                    |> Option.defaultValue null

                let items =
                    properties
                    |> List.tryPick (function | Items items -> Some items | _ -> None)
                    |> Option.defaultValue List.empty

                let cell = ws.Cell(r, c)
                let table = cell.InsertTable(items, name, true)
                let mutable includesTotalsRow = false
                properties
                |> List.iter (function
                    | TableName _
                    | Items _ ->
                        ()
                    | Theme theme -> 
                        table.Theme <- theme
                    | ShowHeaderRow b -> 
                        table.ShowHeaderRow <- b
                    | Totals items ->
                        // Latch includesTotalsRow on in case we have two separate TotalsRowItems passed in:
                        includesTotalsRow <- items.Length > 0 || includesTotalsRow
                        table.ShowTotalsRow <- includesTotalsRow
                        // TODO custom totals row formulae
                        items
                        |> List.iter (fun (name, item) -> 
                            let field = table.Field(name)
                            match item with
                            | Label label -> 
                                field.TotalsRowLabel <- label
                            | Function f when f = XLTotalsRowFunction.Custom -> 
                                // Do this via CustomA1
                                ()
                            | Function f ->
                                field.TotalsRowFunction <- f
                            | CustomA1 s -> field.TotalsRowFormulaA1 <- s)
                    | ColFormulas items 
                    | ColFormulae items ->
                        items
                        |> List.iter (fun (name, item) ->
                            let field = table.Field(name)
                            field.DataCells
                            |> Seq.iter (fun cell -> cell.FormulaA1 <- item))
                    | ShowRowStripes b -> 
                        table.ShowRowStripes <- b
                    | ShowColumnStripes b -> 
                        table.ShowColumnStripes <- b
                    | EmphasizeFirstColumn b -> 
                        table.EmphasizeFirstColumn <- b
                    | EmphasizeLastColumn b -> 
                        table.EmphasizeLastColumn <- b
                    | ShowAutoFilter b -> table.ShowAutoFilter <- b)

                r <- r + items.Length + if includesTotalsRow then 2 else 1

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

