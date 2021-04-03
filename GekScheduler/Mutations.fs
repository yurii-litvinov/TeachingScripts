module Mutations

open DataModel
open Utils

let exchangeWorks (random: System.Random) (linkedWorks: seq<list<Diploma>>) (schedule: Schedule) =
    let sourceDay = 
        schedule.schedule
        |> List.filter (fun d -> dayWorks d |> List.length |> (<>) 0) 
        |> (pick random)
        |> Option.get

    let sourceSitting =
        if sourceDay.firstSitting.IsSome && sourceDay.secondSitting.IsSome then
            let isMorning = random.Next(0, 2) = 1
            if isMorning then sourceDay.firstSitting.Value else sourceDay.secondSitting.Value
        elif sourceDay.firstSitting.IsSome then
            sourceDay.firstSitting.Value
        else 
            sourceDay.secondSitting.Value

    let program = sourceSitting.slot.program

    let isDayWithProgramSlot day program =
        let sittingProgram = function
        | None -> []
        | Some sitting -> [sitting.slot.program]
        sittingProgram day.firstSitting @ sittingProgram day.secondSitting |> List.contains program

    let targetDay = 
        schedule.schedule 
        |> List.except [sourceDay]
        |> List.filter (fun d -> isDayWithProgramSlot d program)
        |> (pick random)

    if targetDay.IsNone then
        schedule
    else
        let targetDay = targetDay.Value
        let targetSitting = 
            if targetDay.firstSitting.IsSome && targetDay.firstSitting.Value.slot.program = program 
               && targetDay.secondSitting.IsSome && targetDay.secondSitting.Value.slot.program = program 
            then
                if random.Next(0, 2) = 1 then targetDay.firstSitting.Value else targetDay.secondSitting.Value
            elif targetDay.firstSitting.IsSome && targetDay.firstSitting.Value.slot.program = program then
                targetDay.firstSitting.Value
            else
                targetDay.secondSitting.Value

        let notLinked w = linkedWorks |> Seq.concat |> Seq.contains w |> not
        let sourceWork = sourceSitting.works |> List.filter notLinked |> (pick random)
        let targetWork = targetSitting.works |> List.filter notLinked |> (pick random)

        if sourceWork.IsNone || targetWork.IsNone then
            schedule
        else
            let newSchedule =
                schedule.schedule
                |> List.map 
                    (fun d -> 
                        if d = sourceDay then 
                            if d.firstSitting.IsSome && d.firstSitting.Value = sourceSitting then
                                { d with
                                    firstSitting = 
                                        Some { slot = sourceSitting.slot
                                               works = targetWork.Value :: (sourceSitting.works |> List.except [sourceWork.Value]) }
                                }
                            else
                                { d with
                                    secondSitting = 
                                        Some { slot = sourceSitting.slot
                                               works = targetWork.Value :: (sourceSitting.works |> List.except [sourceWork.Value]) }
                                }
                        else 
                            d
                    )
                |> List.map 
                    (fun d -> 
                        if d = targetDay then 
                            if d.firstSitting.IsSome && d.firstSitting.Value = targetSitting then
                                { d with
                                    firstSitting = 
                                        Some { slot = targetSitting.slot
                                               works = sourceWork.Value :: (targetSitting.works |> List.except [targetWork.Value]) }
                                }
                            else
                                { d with
                                    secondSitting = 
                                        Some { slot = targetSitting.slot
                                               works = sourceWork.Value :: (targetSitting.works |> List.except [targetWork.Value]) }
                                }
                        else 
                            d
                    )
            { schedule = newSchedule
              totalDays = schedule.totalDays
              advisorTotalDays = countAdvisorTotalDays newSchedule }

let moveWork (random: System.Random) (linkedWorks: seq<list<Diploma>>) (schedule: Schedule) =
    let sourceDay = 
        schedule.schedule
        |> List.filter (fun d -> dayWorks d |> List.length |> (<>) 0) 
        |> (pick random)
        |> Option.get

    let sourceSitting =
        if sourceDay.firstSitting.IsSome && sourceDay.secondSitting.IsSome then
            let isMorning = random.Next(0, 2) = 1
            if isMorning then sourceDay.firstSitting.Value else sourceDay.secondSitting.Value
        elif sourceDay.firstSitting.IsSome then
            sourceDay.firstSitting.Value
        else 
            sourceDay.secondSitting.Value

    let program = sourceSitting.slot.program

    let isDayWithProgramSlot day program =
        let sittingProgram = function
        | None -> []
        | Some sitting -> [sitting.slot.program]
        sittingProgram day.firstSitting @ sittingProgram day.secondSitting |> List.contains program

    let targetDay = 
        schedule.schedule 
        |> List.except [sourceDay]
        |> List.filter (fun d -> isDayWithProgramSlot d program)
        |> List.filter (fun d -> dayWorks d |> List.length < 16)
        |> pick random

    if targetDay.IsNone then
        schedule
    else
        let targetDay = targetDay.Value
        let targetSitting = 
            if targetDay.firstSitting.IsSome && targetDay.firstSitting.Value.slot.program = program 
               && targetDay.secondSitting.IsSome && targetDay.secondSitting.Value.slot.program = program 
            then
                if random.Next(0, 2) = 1 then targetDay.firstSitting.Value else targetDay.secondSitting.Value
            elif targetDay.firstSitting.IsSome && targetDay.firstSitting.Value.slot.program = program then
                targetDay.firstSitting.Value
            else
                targetDay.secondSitting.Value

        let notLinked w = linkedWorks |> Seq.concat |> Seq.contains w |> not
        let sourceWork = sourceSitting.works |> List.filter notLinked |> (pick random)

        if sourceWork.IsNone then
            schedule
        else
            let newSchedule =
                schedule.schedule
                |> List.map 
                    (fun d -> 
                        if d = sourceDay then 
                            if d.firstSitting.IsSome && d.firstSitting.Value = sourceSitting then
                                { d with
                                    firstSitting = 
                                        Some { slot = sourceSitting.slot
                                               works = sourceSitting.works |> List.except [sourceWork.Value] }
                                }
                            else
                                { d with
                                    secondSitting = 
                                        Some { slot = sourceSitting.slot
                                               works = sourceSitting.works |> List.except [sourceWork.Value] }
                                }
                        else 
                            d
                    )
                |> List.map 
                    (fun d -> 
                        if d = targetDay then 
                            if d.firstSitting.IsSome && d.firstSitting.Value = targetSitting then
                                { d with
                                    firstSitting = 
                                        Some { slot = targetSitting.slot
                                               works = sourceWork.Value :: targetSitting.works }
                                }
                            else
                                { d with
                                    secondSitting = 
                                        Some { slot = targetSitting.slot
                                               works = sourceWork.Value :: targetSitting.works }
                                }
                        else 
                            d
                    )
            { schedule = newSchedule
              totalDays = schedule.totalDays
              advisorTotalDays = countAdvisorTotalDays newSchedule }

let exchangeCommissionMembers (random: System.Random) (schedule: Schedule) =
    let sourceDay = schedule.schedule |> (pick random) |> Option.get
    let targetDay = schedule.schedule |> List.except [sourceDay] |> (pick random) |> Option.get

    let sourceMember = sourceDay.commission |> (pick random) |> Option.get
    let targetMember = targetDay.commission |> (pick random) |> Option.get

    if targetDay.commission |> List.contains sourceMember || sourceDay.commission |> List.contains targetMember then
        schedule
    else
        let newSchedule = 
            schedule.schedule
            |> List.map 
                (fun d -> 
                    if d = sourceDay then 
                        { d with commission = targetMember :: (d.commission |> List.except [sourceMember])}
                    else 
                        d
                )
            |> List.map 
                (fun d -> 
                    if d = targetDay then 
                        { d with commission = sourceMember :: (d.commission |> List.except [targetMember])}
                    else 
                        d
                )
        {
            schedule = newSchedule
            totalDays = countTotalDays newSchedule
            advisorTotalDays = schedule.advisorTotalDays
        }

let addCommissionMember (random: System.Random) commissionMembers (schedule: Schedule) =
    let day = schedule.schedule |> (pick random) |> Option.get

    let totalDays = schedule.totalDays
    let commissionMember = 
        commissionMembers 
        |> List.filter (available totalDays)
        |> List.filter (notInconvenientDate day.date)
        |> List.filter (fun m -> day.commission |> List.contains m |> not)
        |> (pick random)

    if commissionMember.IsNone then 
        schedule
    else
        let newSchedule =
            schedule.schedule
            |> List.map (fun d -> if d = day then { d with commission = commissionMember.Value :: d.commission} else d)

        {
            schedule = newSchedule
            totalDays = countTotalDays newSchedule
            advisorTotalDays = schedule.advisorTotalDays
        }

let removeCommissionMember (random: System.Random) commissionMembers (schedule: Schedule) =
    let day = schedule.schedule |> (pick random) |> Option.get

    let commissionMember = 
        commissionMembers 
        |> List.filter (fun m -> getTotalDays schedule.totalDays m > 1)
        |> (pick random)

    if commissionMember.IsNone then 
        schedule
    else
        let newSchedule = 
            schedule.schedule
            |> List.map (fun d -> if d = day then { d with commission = d.commission |> List.except [commissionMember.Value]} else d)
        
        {
            schedule = newSchedule
            totalDays = countTotalDays newSchedule
            advisorTotalDays = schedule.advisorTotalDays
        }

let exchangeChairs (random: System.Random) (schedule: Schedule) =
    let sourceDay = schedule.schedule |> (pick random) |> Option.get
    let targetDay = schedule.schedule |> List.except [sourceDay] |> (pick random) |> Option.get

    let newSchedule = 
        schedule.schedule
        |> List.map (fun d -> if d = sourceDay then { d with chair = targetDay.chair} else d)
        |> List.map (fun d -> if d = targetDay then { d with chair = sourceDay.chair} else d)
    {
        schedule = newSchedule
        totalDays = countTotalDays newSchedule
        advisorTotalDays = schedule.advisorTotalDays
    }

let mutate linkedWorks commissionMembers schedule =
    let random = System.Random(int System.DateTime.Now.Ticks + System.Threading.Thread.CurrentThread.ManagedThreadId)

    let mutations = 
        [ exchangeWorks random linkedWorks
          moveWork random linkedWorks
          exchangeCommissionMembers random
          (addCommissionMember random commissionMembers)
          (removeCommissionMember random commissionMembers)
          exchangeChairs random ]

    let mutationsCount = random.Next(1, 4)
    let mutable result = schedule
    for _ in 1..mutationsCount do
        result <- (pick random mutations).Value result

    result
