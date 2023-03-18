module Tests

open System
open Xunit
open FsExcel

[<Theory>]
[<InlineData("A", 1, 1, 1, "A", 1)>]
[<InlineData("A", 1, 2, 1, "B", 1)>]
[<InlineData("A", 1, 1, 2, "A", 2)>]
[<InlineData("A", 1, 2, 2, "B", 2)>]
[<InlineData("Z", 1, 2, 2, "AA", 2)>]
[<InlineData("AA", 1, 1, 1, "AA", 1)>]
[<InlineData("AA", 1, 2, 1, "AB", 1)>]
[<InlineData("ZZ", 1, 1, 1, "ZZ", 1)>]
[<InlineData("ZZ", 1, 2, 1, "AAA", 1)>]
// Max column heading is "XFD":
[<InlineData("XFC", 1, 2, 1, "XFD", 1)>]
let ``Can translate from a span depth to a cell reference`` (startCol, startRow, span, depth, expectedCol, expectedRow) =
    let expected = expectedCol, expectedRow
    let actual = CellReference.spanDepthToCellReference (startCol, startRow) span depth
    Assert.Equal(expected, actual)

[<Theory>]
[<InlineData("A", 1, 16_385, 1)>]
// Max column heading is "XFD":
[<InlineData("XFD", 1, 2, 1)>]
let ``Spans beyond right limit of spreadsheet lead to an exception``(startCol, startRow, span, depth) =
    Assert.Throws<ArgumentException>(fun _ ->
        CellReference.spanDepthToCellReference (startCol, startRow) span depth |> ignore)

[<Theory>]
[<InlineData("A", 1, 1, 1_048_577)>]
[<InlineData("A", 1_048_577, 1, 2)>]
let ``Depths beyond bottom limit of spreadsheet lead to an exception``(startCol, startRow, span, depth) =
    Assert.Throws<ArgumentException>(fun _ ->
        CellReference.spanDepthToCellReference (startCol, startRow) span depth |> ignore)

