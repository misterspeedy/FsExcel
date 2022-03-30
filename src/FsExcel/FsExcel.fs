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

type Item =
    | Cell of props : CellProp list
    | Style of props : CellProp list
    | Go of Position
    | Worksheet of string
    | AutoFit of AutoFit
    | Workbook of XLWorkbook
    | Size of Size

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

