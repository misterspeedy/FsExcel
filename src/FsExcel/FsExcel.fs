module FsExcel

// https://lukelowrey.com/use-github-actions-to-publish-nuget-packages/

open ClosedXML.Excel
open DocumentFormat.OpenXml.Spreadsheet

type Position =
    | RC of int * int
    | RightBy of int
    | DownBy of int
    | LeftBy of int
    | UpBy of int

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

type Content =
    | String of string
    | Number of float
    // TODO add DateTime

type CellProp =
    | Content of Content
    | Next of Position
    | FontEmphasis of FontEmphasis
    | Border of Border
    | HorizontalAlignment of HorizontalAlignment
    | FormatCode of string

type Item =
    | Cell of props : CellProp list
    | Go of Position

let render (items : Item list) =
    let mutable r = 1
    let mutable c = 1
    let wb = new XLWorkbook()
    // TODO - allow naming
    let ws = wb.Worksheets.Add("Sheet 1")

    let go = function
        | RC (row, col) ->
            r <- row
            c <- col
        | RightBy n ->
            c <- c+n
        | DownBy n ->
            r <- r+n
        | UpBy n ->
            r <- r-n |> max 1
        | LeftBy n ->
            c <- c-n |> max 1        

    for item in items do

        match item with
        | Cell props ->

            for prop in props do 

                let cell = ws.Cell(r, c)
                
                match prop with
                | Content con ->
                    match con with 
                    | String s ->
                        cell.Value <- s
                    | Number n ->
                        cell.Value <- n
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