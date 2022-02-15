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
    | TopBorder of XLBorderStyleValues
    | RightBorder of XLBorderStyleValues
    | BottomBorder of XLBorderStyleValues
    | LeftBorder of XLBorderStyleValues

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

type Item =
    | Cell of props : CellProp list
    | Go of Position

let render (sheetName : string) (items : Item list) =
    let mutable indent = 1
    let mutable r = 1
    let mutable c = 1
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

    let hasNext (props : CellProp list) =
        props
        |> List.exists (function | Next _ -> true | _ -> false)

    for item in items do

        match item with
        | Cell props ->

            let props = 
                if props |> hasNext |> not then
                    Next(RightBy 1) :: props
                else
                    props

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
                | CellProp.FontEmphasis fe -> 
                    match fe with
                    | FontEmphasis.Bold ->
                        cell.Style.Font.Bold <- true
                    | FontEmphasis.Italic ->
                        cell.Style.Font.Italic <- true
                | CellProp.Border b ->
                    match b with
                    | TopBorder style ->
                        cell.Style.Border.TopBorder <- style
                    | RightBorder style ->
                        cell.Style.Border.RightBorder <- style
                    | BottomBorder style ->
                        cell.Style.Border.BottomBorder <- style
                    | LeftBorder style ->
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
                | CellProp.FormatCode fc ->
                    cell.Style.NumberFormat.Format <- fc

        | Go p ->
            go p

    wb