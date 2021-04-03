module InitialScheduleGenerator

open System.Collections.Generic

open DataModel
open Utils

let random = System.Random()

let pickCommission (random: System.Random) chair commissionMembers date (totalDays: Dictionary<_, _>) = 
    let mutable available = 
        commissionMembers 
        |> List.filter (fun m -> getTotalDays totalDays m < m.maxDays)
        |> List.filter (notInconvenientDate date)
        |> List.filter (fun m -> m.company <> chair.company)
        |> Set.ofList
    let count = random.Next(4, 8)
    let mutable result = []
    for _ in 0..count do
        if not available.IsEmpty then
            let pick = pick random (available |> Set.toSeq)
            if pick.IsSome then
                available <- available.Remove (pick.Value)
                available <- available |> Set.filter (fun m -> m.company <> pick.Value.company)
                incMap totalDays (pick.Value)
                result <- pick.Value :: result
    result

let pickChair (random: System.Random) chairs date (totalDays: Dictionary<_, _>) =
    let chair = 
        chairs 
        |> List.filter (available totalDays)
        |> List.filter (notInconvenientDate date)
        |> (pick random)

    if chair.IsSome then
        incMap totalDays chair.Value

    chair

let generateSitting
    (random: System.Random)
    (slot: option<SittingSlot>) 
    (availableWorks: HashSet<Diploma>) 
    chair 
    (commission: list<CommissionMember>) 
    (linkedWorks: seq<list<Diploma>>)
    =
    match slot with
    | None -> None
    | Some slot ->
        let pickWorks () =
            let mutable result = []
            for _ in 1..8 do
                if availableWorks.Count > 0 && result.Length < 8 then
                    let work = 
                        availableWorks 
                        |> Seq.filter (fun w -> w.program = slot.program) 
                        |> Seq.filter (fun w -> w.reviewer <> chair.name)
                        |> Seq.filter (fun w -> commission |> List.forall (fun c -> w.reviewer <> c.name))
                        |> (pick random)

                    if work.IsSome then
                        let work = work.Value
                        availableWorks.Remove work |> ignore
                        result <- work :: result

                        let connect (works: Diploma list) =
                            if works |> List.contains work then
                                let remaining = works |> List.filter ((<>) work)
                                for rw in remaining do
                                    result <- rw :: result
                                    availableWorks.Remove rw |> ignore
                        
                        linkedWorks |> Seq.iter connect

            result

        Some { slot = slot
               works = pickWorks () }

let addAvailableCommissionMembers (random: System.Random) commissionMembers (schedule: Day list) =
    let totalDays = countTotalDays schedule
    let available = commissionMembers |> List.filter (fun m -> totalDays.ContainsKey m |> not)
    let mutable schedule = schedule
    available |> List.iter
        (fun m ->
            let day = schedule |> (pick random) |> Option.get
            schedule <- schedule |> List.map (fun d -> if d = day then {d with commission = m :: d.commission} else d)
        )
    schedule

let generateSchedule works commissionMembers chairs (sittings: list<SittingSlot>) linkedWorks =
    let totalDays = Dictionary<CommissionMember, int>()
    let availableWorks = HashSet<Diploma>()
    works |> List.iter (fun d -> availableWorks.Add d |> ignore) 

    let generateDay (date: Date) =
        let getSlot time = sittings |> List.filter (fun s -> s.date = date && s.time = time) |> List.tryHead
        let morningSlot = getSlot Morning
        let daySlot = getSlot Day

        let chair = pickChair random chairs date totalDays

        match chair with
        | None -> 
            //printfn "Failed to pick a chair for day %A" date
            None
        | Some chair ->
            let commission = pickCommission random chair commissionMembers date totalDays
            Some { date = date
                   chair = chair
                   commission = commission
                   firstSitting = generateSitting random morningSlot availableWorks chair commission linkedWorks
                   secondSitting = generateSitting random daySlot availableWorks chair commission linkedWorks }

    let schedule = 
        sittings
        |> List.map (fun s -> s.date)
        |> List.distinct
        |> List.map generateDay

    if schedule |> List.exists Option.isNone then
        //printfn "Failed to generate schedule, empty day"
        None
    elif availableWorks.Count > 0 then
        //printfn "Failed to assign every work. List of unassigned works:"
        //availableWorks |> Seq.iter (fun w -> printfn "%s" w.theme)
        None
    else
        let schedule = schedule |> List.map Option.get
        let schedule = addAvailableCommissionMembers random commissionMembers schedule
        Some { schedule = schedule 
               totalDays = countTotalDays schedule
               advisorTotalDays = countAdvisorTotalDays schedule }
