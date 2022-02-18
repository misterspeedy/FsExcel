module FsExcel

open ClosedXML.Excel
open DocumentFormat.OpenXml.Spreadsheet

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

type Border =
    | Top of XLBorderStyleValues
    | Right of XLBorderStyleValues
    | Bottom of XLBorderStyleValues
    | Left of XLBorderStyleValues

type HorizontalAlignment =
    | Left
    | Center
    | Right

type CellProp =
    | String of string
    | Float of float
    | Integer of int
    | Next of Position
    | FontEmphasis of FontEmphasis
    | Border of Border
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

type Item =
    | Cell of props : CellProp list
    | Style of props : CellProp list
    | Go of Position

let render (sheetName : string) (items : Item list) =
    let mutable indent = 1
    let mutable r = 1
    let mutable c = 1
    let mutable style : CellProp list = []
    let wb = new XLWorkbook()
    let ws = wb.Worksheets.Add(sheetName)

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
        | Cell props ->

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
                | Next p ->
                    go p
                | FontEmphasis fe -> 
                    match fe with
                    | FontEmphasis.Bold ->
                        cell.Style.Font.Bold <- true
                    | FontEmphasis.Italic ->
                        cell.Style.Font.Italic <- true
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
                    // TODO border color
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
        | Style s ->
            style <- s        
        | Go p ->
            go p

    wb