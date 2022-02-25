namespace FsExcel

module Table =

    open System
    open Microsoft.FSharp.Reflection
    open System.Collections.Concurrent
    open System.Runtime.Serialization

    type Direction = Vertical | Horizontal

    module private Fields = 

        let cache = ConcurrentDictionary<System.Type, Reflection.PropertyInfo[]>()
        
        let serializable t = cache.GetOrAdd(t, fun _ -> 
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

        let body getCellStyle r =
            let t = r.GetType()
            let fields = Fields.serializable t
            fields
            |> Array.map (fun f -> f.Name, FSharpValue.GetRecordField(r, f))
            |> Array.map (fun (name, value) ->

                let style = getCellStyle name

                let content = 
                    match value with
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
                        String (string value)
                style, content)
            |> Array.map (fun (style, content) ->
                Cell [ content; yield! style; Next Stay])
            |> List.ofArray

        let header<'T>(getCellStyle) =
            let t = typeof<'T>
            t
            |> Fields.serializable
            |> Array.map (fun f -> f.Name)
            |> Array.map (fun name -> 
                let style = getCellStyle name
                Cell [ String name; yield! style; Next Stay ])
            |> List.ofArray

    type CellStyleGetter = int -> string -> CellProp list        

    let fromInstance<'T> (direction : Direction) (getCellStyle : CellStyleGetter) (x : 'T) =
        let headerCells = Cells.header<'T> (getCellStyle 0)
        let bodyCells = x |> Cells.body (getCellStyle 1)
        match direction with
        | Horizontal ->
            [
                for headerCell in headerCells do
                    headerCell
                    Go (RightBy 1)
                Go NewRow
                for bodyCell in bodyCells do
                    bodyCell
                    Go (RightBy 1)
            ]
        | Vertical ->
            [
                for heading, value in List.zip headerCells bodyCells do
                    heading
                    Go (RightBy 1)
                    value
                    Go (DownBy 1)
                    Go (LeftBy 1)
            ]

    type CellStyleGetterSeq = int -> string -> CellProp list

    let fromSeq<'T> (direction : Direction) (getCellStyle : CellStyleGetter) (xs : 'T seq) =
        let xs = xs |> Array.ofSeq
        let headerCells = Cells.header<'T> (getCellStyle 0)

        match direction with
        | Vertical ->   
            [
                let depth = xs.Length+1
                for headerCell in headerCells do
                    headerCell
                    Go (DownBy 1)
                Go (UpBy depth)
                Go (RightBy 1)
                for i, x in xs |> Seq.indexed do
                    for bodyCell in x |> Cells.body (getCellStyle (i+1)) do
                        bodyCell
                        Go (DownBy 1)
                    Go (UpBy depth)
                    Go (RightBy 1)
            ]
        | Horizontal ->
            [
                for headerCell in headerCells do
                    headerCell
                    Go (RightBy 1)
                Go NewRow
                for i, x in xs |> Seq.indexed do
                    for bodyCell in x |> Cells.body (getCellStyle (i+1)) do
                        bodyCell
                        Go (RightBy 1)
                    Go NewRow
            ]
