module Tests

open System
open System.IO
open Xunit
open FsExcel

let tempPath = "/temp"

[<Fact>]
let ``Volume test 1`` () =
    let rows = 1000
    let cols = 100
    [
        for row in 1..rows do
            for col in 1..cols do
                Cell [ Integer (row*col)]
            Go NewRow
    ]
    |> FsExcel.Render.AsFile(Path.Combine(tempPath, "VolumeTest1.xlsx")) 
    Assert.True(true)
