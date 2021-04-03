module Constraints

open DataModel
open Utils

let isNoReviewerInCommission (day: Day) =
    let notAReviewer (commissionMember: CommissionMember) =
        dayWorks day |> List.forall (fun work -> work.reviewer <> commissionMember.name)

    day.commission |> List.forall notAReviewer
    && notAReviewer day.chair

let isNoAdvisorInCommission (day: Day) =
    let notAnAdvisor (commissionMember: CommissionMember) =
        dayWorks day |> List.forall (fun work -> getParticipantId work.advisor <> getParticipantId commissionMember.name)

    day.commission |> List.forall notAnAdvisor
    && notAnAdvisor day.chair

let everyoneIsOkayWithDates (day: Day) =
    day.commission |> List.forall (fun cm -> cm.inconvenientDates |> List.contains day.date |> not)
    && day.chair.inconvenientDates |> List.contains day.date |> not

let noCommissionMemberDuplicatesAllowed (day: Day) =
    if not ((List.distinct day.commission) = day.commission) then
        failwith "Commission duplicated"
    true

let commissionMustHaveAdequateSize (day: Day) =
    day.commission.Length >= 5 && day.commission.Length <= 12

let commissionMembersFromSameCompanyNotAllowed (day: Day) =
    let commissionMembers = day.chair :: day.commission 
    let companies = commissionMembers |> List.map (fun cm -> cm.company)
    (List.distinct companies) = companies

let thereShallBeNoMoreThan8WorksInASitting (day: Day) =
    let checkSitting = function
    | None -> true
    | Some sitting -> sitting.works.Length <= 8
    checkSitting day.firstSitting && checkSitting day.secondSitting

let allSittingWorksHaveCorrectProgram (day: Day) =
    let checkSitting = function
    | None -> true
    | Some sitting -> sitting.works |> List.forall (fun w -> w.program = sitting.slot.program)
    checkSitting day.firstSitting && checkSitting day.secondSitting

let totalDaysLessThanMax (schedule: Schedule) =
    schedule.totalDays
    |> Seq.forall (fun pair -> pair.Key.maxDays >= pair.Value)

let allCommissionMembersUsed (commissionMembers: list<CommissionMember>) (chairs: list<CommissionMember>) (schedule: Schedule) =
    schedule.totalDays.Count = commissionMembers.Length + chairs.Length

let allChairsUsed (chairs: list<CommissionMember>) (schedule: Schedule) =
    let isPresent name =
        (schedule.totalDays |> Seq.exists (fun pair -> pair.Key.name = name))

    let areChairsUsed = chairs |> List.forall (fun c -> isPresent c.name)
    if areChairsUsed then
        true
    else
        failwith "Failed to use all chairs"
    
let allWorksShallBeDefended works (schedule: Schedule) =
    let allWorks = schedule.schedule |> List.map dayWorks |> List.concat
    works |> List.forall (fun w -> if not (List.contains w allWorks) then printfn "Failed to schedule %s" w.author; false else true)
