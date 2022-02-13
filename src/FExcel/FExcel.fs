module FExcel

// https://lukelowrey.com/use-github-actions-to-publish-nuget-packages/

open SpreadsheetLight
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
    | TopBorder of BorderStyleValues * System.Drawing.Color
    | RightBorder of BorderStyleValues * System.Drawing.Color
    | BottomBorder of BorderStyleValues * System.Drawing.Color
    | LeftBorder of BorderStyleValues * System.Drawing.Color

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
    let ss = new SLDocument()

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

            let cellStyle = ss.CreateStyle()

            for prop in props do 
                
                match prop with
                | Content con ->
                    match con with 
                    | String s ->
                        ss.SetCellValue(r, c, s) |> ignore
                    | Number n ->
                        ss.SetCellValueNumeric(r, c, string n) |> ignore
                | Next p ->
                    go p
                | CellProp.FontEmphasis fe -> 
                    match fe with
                    | FontEmphasis.Bold ->
                        cellStyle.SetFontBold(true)
                    | FontEmphasis.Italic ->
                        cellStyle.SetFontItalic(true)
                | CellProp.Border b ->
                    match b with
                    | TopBorder(style, color) ->
                        cellStyle.SetTopBorder(style, color) |> ignore
                    | RightBorder(style, color) ->
                        cellStyle.SetRightBorder(style, color) |> ignore
                    | BottomBorder(style, color) ->
                        cellStyle.SetBottomBorder(style, color) |> ignore
                    | LeftBorder(style, color) ->
                        cellStyle.SetLeftBorder(style, color) |> ignore
                | HorizontalAlignment h ->
                    match h with
                    | Left ->
                        cellStyle.Alignment.Horizontal <- HorizontalAlignmentValues.Left
                    | Center ->
                        cellStyle.Alignment.Horizontal <- HorizontalAlignmentValues.Center
                    | Right ->
                        cellStyle.Alignment.Horizontal <- HorizontalAlignmentValues.Right
                | CellProp.FormatCode fc ->
                    cellStyle.FormatCode <- fc

                ss.SetCellStyle(r, c, cellStyle) |> ignore
        | Go p ->
            go p

    ss