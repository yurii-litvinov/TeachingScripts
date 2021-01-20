module RoomData

open LessPainfulGoogleSheets

type RoomRecord =
    {
        semester: int
        discipline: string
        workType: string
        rooms: string list
    }

type RoomData(sheetId: string) =
    let service = openGoogleSheet "AssignmentMatcher"
    let getData () =
        let mutable data: Map<int, RoomRecord list> = Map.empty
        for page in [1..(if Config.programStartYear = 2019 then 4 else 8)] do
            let rawData = readGoogleSheet service sheetId $"{page} семестр" "A" "C" 1
            if rawData <> Seq.empty then
                let parsedData = 
                    rawData 
                    |> Seq.filter (fun row -> (Seq.length row) >= 3)
                    |> Seq.map
                        (fun row -> 
                            {
                                semester = page
                                discipline = row |> Seq.head
                                workType = row |> Seq.item 1
                                rooms = row |> Seq.item 2 |> (fun s -> if s.StartsWith "актовый" then [s] else s.Split [|' '|] |> Seq.toList)
                            }
                        )
                data <- data.Add(page, parsedData |> List.ofSeq)
        let data = data |> Map.map (fun _ v -> v |> Seq.groupBy (fun r -> r.discipline) |> Map.ofSeq)
        let data = data |> Map.map (fun _ v -> v |> Map.map (fun _ v -> v |> Seq.groupBy (fun r -> r.workType) |> Map.ofSeq))

        let checkAndMap (r: RoomRecord seq) =
            if r |> Seq.length <> 1 then
                failwith $"Ambuguous room data for {(r |> Seq.head).discipline}"
            (r |> Seq.head).rooms

        data |> Map.map (fun _ v -> v |> Map.map (fun _ v -> v |> Map.map (fun _ v -> checkAndMap v)))

    let data = getData ()

    member _.Rooms semester discipline =
        if data.ContainsKey semester && data.[semester].ContainsKey discipline then
            data.[semester].[discipline]
        else
            Map.empty
