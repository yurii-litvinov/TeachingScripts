module Utils

open System.Collections.Generic

open DataModel

let incMap (map: Dictionary<'k, int>) key  =
    if not <| map.ContainsKey key then
        map.Add (key, 1)
    else
        map.Item key <- (map.Item key) + 1

let countTotalDays (schedule: Day list) =
    let totalDays = Dictionary<CommissionMember, int>()
    schedule 
    |> List.iter (fun day -> day.commission |> List.iter (fun cm -> incMap totalDays cm))

    schedule 
    |> List.iter (fun day -> incMap totalDays day.chair)

    totalDays

let dayWorks day = 
    let sittingWorks = function
    | None -> []
    | Some sitting -> sitting.works

    let first = sittingWorks day.firstSitting
    let second = sittingWorks day.secondSitting
    first @ second

let countAdvisorTotalDays (schedule: Day list) =
    let totalDays = Dictionary<string, int>()
    schedule 
    |> List.iter (fun day -> dayWorks day |> List.iter (fun w -> incMap totalDays w.advisor))
    
    totalDays

let getTotalDays (map: Dictionary<_, _>) person =
    if map.ContainsKey person then
        map.Item person
    else 
        0

let available (map: Dictionary<_, _>) (person: CommissionMember) =
    person.maxDays > getTotalDays map person

let notInconvenientDate date commissionMember =
    not <| List.contains date commissionMember.inconvenientDates

let pick (random: System.Random) (seq: 'a seq) =
    if Seq.length seq = 0 then
        None
    else
        let pickNumber = random.Next(0, Seq.length seq)
        Some(seq |> Seq.item pickNumber)

let getParticipantId (str: string) = 
    let parts = str.Split(" ")
    if parts |> Seq.length < 2 then
        str
    else
        $"{parts.[0]} {parts.[1].ToCharArray().[0]}"
