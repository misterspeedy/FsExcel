module Assert

open Xunit
open ClosedXML.Excel

let Equal (expected : IXLCell, actual : IXLCell) =
    if expected.DataType <> actual.DataType then
        raise (Xunit.Sdk.NotEqualException($"{expected.DataType}", $"{actual.DataType}"))
    else
        match expected.DataType with
        | XLDataType.Text ->
            let e = expected.GetString()
            let a = actual.GetString()
            Assert.Equal(e, a)
        | XLDataType.Number ->
            let e = expected.GetDouble()
            let a = expected.GetDouble()
            Assert.Equal(e, a)
        | _ -> ()
    // TODO
    // Boolean
    // DateTime
    // TimeSpan