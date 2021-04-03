module Fitness

open DataModel
open Constraints
open Utils

let commissionMemberDateConvenienceScore = 10
let chairDateConvenienceScore = 30
let commissionMemberTotalDaysPenalty = 7
let studentDateConvenienceScore = 1
let advisorDateConvenienceScore = 3
let advisorTotalDaysPenalty = 3
let themeScore = 2
let commissionSizeBalancePenalty = 10

let chairDateConvenience day = 
    if List.contains day.date day.chair.convenientDates then chairDateConvenienceScore else 0

let commissionDateConvenience day = 
    day.commission 
    |> List.map (fun m -> if List.contains day.date m.convenientDates then 1 else 0)
    |> List.sum
    |> (*) commissionMemberDateConvenienceScore

let studentDateConvenience day =
    dayWorks day
    |> List.map 
        (
            fun w -> 
                if List.contains day.date w.convenientDates then 
                    1 
                elif List.contains day.date w.inconvenientDates then
                    -3
                else
                    0
        )
    |> List.sum
    |> (*) studentDateConvenienceScore

let advisorDateConvenience (advisorPreferences: Map<_, _>) day =
    dayWorks day
    |> List.map 
        (fun w -> 
                let id = getParticipantId w.advisor
                if advisorPreferences.ContainsKey id && fst advisorPreferences.[id] |> List.contains day.date then 
                    1
                elif advisorPreferences.ContainsKey id && snd advisorPreferences.[id] |> List.contains day.date then
                    -5
                else
                    0
        )
    |> List.sum
    |> (*) advisorDateConvenienceScore

let themeConvenience day =
    dayWorks day
    |> List.map (fun work ->
        let commissionScore = 
            day.commission 
            |> List.map (fun m -> if m.interestingThemes |> List.contains work.theme then themeScore else 0)
            |> List.sum
        commissionScore
        )
    |> List.sum

let totalDaysCommissionPenalty schedule =
    schedule.totalDays
    |> Seq.map (fun pair -> if pair.Key.involvementPreference = Min then (pown commissionMemberTotalDaysPenalty (pair.Value - 1)) - 1  else -1)
    |> Seq.sum

let totalDaysAdvisorPenalty schedule =
    schedule.advisorTotalDays
    |> Seq.map (fun pair -> (pair.Value - 1) * advisorTotalDaysPenalty)
    |> Seq.sum

let commissionBalance schedule =
    let e = float (schedule.schedule |> Seq.map (fun d -> d.commission.Length) |> Seq.sum) / (float schedule.schedule.Length)
    let d = schedule.schedule |> Seq.map (fun d -> d.commission.Length) |> Seq.map (fun l -> pown (e - float l) 2) |> Seq.sum
    int (-d * (float commissionSizeBalancePenalty))

let fitness (advisorPreferences: Map<_, _>) (schedule: Schedule) =
    { commissionDates = schedule.schedule |> List.map commissionDateConvenience |> List.sum
      chairDates = schedule.schedule |> List.map chairDateConvenience |> List.sum
      studentDates = schedule.schedule |> List.map studentDateConvenience |> List.sum
      advisorDates = schedule.schedule |> List.map (advisorDateConvenience advisorPreferences) |> List.sum
      themes = schedule.schedule |> List.map themeConvenience |> List.sum
      commissionBalance = commissionBalance schedule
      totalDaysCommissionPenalty = totalDaysCommissionPenalty schedule
      totalDaysAdvisorPenalty = totalDaysAdvisorPenalty schedule }

let totalFitness fitness =
    fitness.commissionDates
    + fitness.chairDates
    + fitness.studentDates 
    + fitness.advisorDates
    + fitness.themes 
    + fitness.commissionBalance
    - fitness.totalDaysCommissionPenalty
    - fitness.totalDaysAdvisorPenalty

let calculateFitness advisorPreferences commissionMembers chairs works (schedule: Schedule option): Fitness option =
    if schedule.IsNone then
        None
    else
        let localConstraints = 
            [ isNoReviewerInCommission
              isNoAdvisorInCommission
              commissionMustHaveAdequateSize
              everyoneIsOkayWithDates
              thereShallBeNoMoreThan8WorksInASitting
              commissionMembersFromSameCompanyNotAllowed
              noCommissionMemberDuplicatesAllowed
              allSittingWorksHaveCorrectProgram ]

        let localConstraintsHolding =
            let checkConstraint schedule predicate =
                schedule |> List.forall (fun sitting -> predicate sitting)
            localConstraints
            |> List.forall (checkConstraint schedule.Value.schedule)

        let globalConstraints = 
            [ totalDaysLessThanMax
              allCommissionMembersUsed commissionMembers chairs
              allWorksShallBeDefended works
              allChairsUsed chairs ]

        let globalConstraintsHolding =
            globalConstraints |> List.forall (fun c -> c schedule.Value)

        if (not localConstraintsHolding) || (not globalConstraintsHolding) then
            None
        else
            Some (fitness advisorPreferences schedule.Value)
