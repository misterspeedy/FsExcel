namespace FsExcel

module CellReference = 

    open ClosedXML.Excel

    /// Returns the column letter and row number of a named cell given a named cell
    let namedCellToCR (cellName : string) (worksheet : IXLWorksheet) = 
        (worksheet.Cell(cellName).WorksheetColumn().ColumnLetter(), worksheet.Cell(cellName).WorksheetRow().RowNumber())

    // TODO: There are probably more efficient ways to map between alphabetic column headings and numeric indices, using simple calculations.
    let private alphabet = ['A'..'Z'] 

    let private doubleAlphabet = 
        List.allPairs alphabet alphabet
        |> List.map (fun (a1, a2) -> sprintf "%c%c" a1 a2)

    // One past the largest alphabetic column ref Excel supports - XFD:
    let private tripleAlphabetLimit = "XFE"
    let private tripleAlphabet =
        seq {
            for c1 in alphabet do
                for c2 in alphabet do
                    for c3 in alphabet do
                        sprintf "%c%c%c" c1 c2 c3
        }
        |> Seq.takeWhile ((<>) tripleAlphabetLimit)
        |> List.ofSeq

    let private colHeadings = 
        List.concat
            [
                alphabet |> List.map string
                doubleAlphabet
                tripleAlphabet
            ]

    let private colIndexLetters =
        colHeadings
        |> List.indexed
        |> List.map (fun (index, letters) -> (index + 1, letters))

    let private colIndexLettersMap = colIndexLetters |> Map.ofList
    
    let private colLettersToIndexMap =
        colIndexLetters
        |> List.map (fun (index, letters) -> (letters, index))
        |> Map.ofList

    let private tryColLettersToColIndex colLabel =
        colLettersToIndexMap |> Map.tryFind colLabel

    let private tryColIndexToColLetters colIndex =
        colIndexLettersMap |> Map.tryFind colIndex

    /// Returns the column letter and row number of the cell to which to merge to given the starting named or specific cell. Span and depth are integers of minimum value = 1.
    let spanDepthToCellReference (cellReference : (string * int)) (span : int) (depth : int) = 
        let colLabel, row = cellReference
        colLabel 
        |> tryColLettersToColIndex
        |> Option.map (fun colIndex ->
            let newColIndex = colIndex + span - 1
            let newRowIndex = row + depth - 1
            if newRowIndex > 1_048_576 then
                raise <| System.ArgumentException($"Depth would lead to an out of range row index")
            else
                newColIndex 
                |> tryColIndexToColLetters
                |> Option.map (fun _ ->
                    (colIndexLettersMap.[newColIndex], newRowIndex))
                |> Option.defaultWith (fun _ ->
                    raise <| System.ArgumentException($"Span would lead to an out of range reference")))
        |> Option.defaultWith (fun _ ->
            raise <| System.ArgumentException($"Out of range column reference: {colLabel}"))

    /// Returns the column letter and row number of the starting cell to which to merge to given the ending named cell. Span and depth are integers of minimum value = 1.
    let cellReverseSpanDepthToCR (cell : (string * int)) (span : int) (depth : int) = 
        let colLetter = 
            colIndexLettersMap.TryFind (colLettersToIndexMap.[cell |> fst] - (span - 1))
            |> Option.defaultValue "A" // if user tries to reverse merge beyond column A, they will only be able to reverse merge until column A
            
        let rowNum = 
            if ((cell |> snd) - (depth - 1)) > 0 then
                ((cell |> snd) - (depth - 1))
            else
                1 // if user tries to reverse merge beyond row 1, they will only be able to reverse merge until row 1
        (colLetter, rowNum)