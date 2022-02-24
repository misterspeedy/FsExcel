namespace FsExcel

module Table =

    open System
    open Microsoft.FSharp.Reflection
    open System.Collections.Concurrent
    open System.Runtime.Serialization

    type Direction = Vertical | Horizontal

    module private Fields = 

        let cache = ConcurrentDictionary<System.Type, Reflection.PropertyInfo[]>()
        
        let getSerializableFields t = cache.GetOrAdd(t, fun _ -> 
            let fields = FSharpType.GetRecordFields(t)
            // TODO Is this really the best way of determining if a record 
            // field should be serialized? Especially the matching-by-"Name+@" part.
            let serializableNames = 
                FormatterServices.GetSerializableMembers(t)
                |> Seq.map (fun m -> m.Name)
                |> Set.ofSeq
            fields
            |> Array.filter (fun f -> serializableNames.Contains (sprintf "%s@" f.Name)))

    module private Cells = 

        open FsExcel

        let body r =
            let t = r.GetType()
            let fields = Fields.getSerializableFields t
            fields
            |> Array.map (fun f -> FSharpValue.GetRecordField(r, f))
            |> Array.map (fun f ->
                match f with
                | :? String as s ->
                    String s
                | :? DateTimeOffset as dto ->
                    // TODO handle dates explictly
                    String (dto.ToString("u"))
                | :? DateTime as dt ->
                    String (dt.ToString("u"))
                | :? int as i ->
                    Integer i
                | :? float as f ->
                    Float f
                | :? float32 as f ->
                    Float (float f)
                | :? decimal as d ->
                    Float (float d)
                | _ -> 
                    String (string f))
            |> Array.map (fun content ->
                Cell [ content; Next Stay])
            |> List.ofArray

        let header<'T>() =
            let t = typeof<'T>
            let fields = Fields.getSerializableFields t
            fields
            |> Array.map (fun f -> f.Name)
            |> Array.map (fun s -> Cell [ String s; Next Stay ])
            |> List.ofArray

    let fromInstance<'T> (direction : Direction) (headingStyle : CellProp list) (bodyStyle : CellProp list) (x : 'T) =
        let headerCells = Cells.header<'T>()
        let bodyCells = x |> Cells.body
        match direction with
        | Horizontal ->
            [
                Style headingStyle
                for headerCell in headerCells do
                    headerCell
                    Go (RightBy 1)
                Go NewRow
                Style bodyStyle
                for bodyCell in bodyCells do
                    bodyCell
                    Go (RightBy 1)
            ]
        | Vertical ->
            [
                for heading, value in List.zip headerCells bodyCells do
                    Style headingStyle
                    heading
                    Go (RightBy 1)
                    Style  bodyStyle
                    value
                    Go (DownBy 1)
                    Go (LeftBy 1)
            ]

    let fromSeq<'T> (direction : Direction) (headingStyle : CellProp list) (bodyStyle : CellProp list) (xs : 'T seq) =
        let xs = xs |> Array.ofSeq
        let headerCells = Cells.header<'T>()

        match direction with
        | Vertical ->   
            [
                let depth = xs.Length+1
                Style headingStyle
                for headerCell in headerCells do
                    headerCell
                    Go (DownBy 1)
                Go (UpBy depth)
                Go (RightBy 1)
                Style bodyStyle
                for x in xs do
                    for bodyCell in x |> Cells.body do
                        bodyCell
                        Go (DownBy 1)
                    Go (UpBy depth)
                    Go (RightBy 1)
            ]
        | Horizontal ->
            [
                Style headingStyle
                for headerCell in headerCells do
                    headerCell
                    Go (RightBy 1)
                Go NewRow
                Style bodyStyle
                for x in xs do
                    for bodyCell in x |> Cells.body do
                        bodyCell
                        Go (RightBy 1)
                    Go NewRow
            ]
