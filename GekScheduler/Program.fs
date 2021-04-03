open InitialScheduleGenerator
open DataModel
open Fitness
open Mutations
open System
open System.IO
open FSharp.Collections.ParallelSeq

let printTimeSlot = function
    | Morning -> "10:00"
    | Day -> "14:00"

let printSitting writer (sitting: option<Sitting>) =
    match sitting with
    | None -> ()
    | Some sitting ->
        fprintfn writer "    Время: %s" (printTimeSlot sitting.slot.time)
        fprintfn writer "    Работы:"
        for i in 0..sitting.works.Length - 1 do
            let work = sitting.works |> List.item i
            fprintfn writer "%s\t%s\t%s"work.author work.theme work.advisor

let printDay writer (day: Day) =
    let (Date date) = day.date
    fprintfn writer "Дата заседания: %s" date
    fprintfn writer "    Председатель: %s" day.chair.name
    fprintfn writer "    Члены комиссии:"
    let sortedCommission = day.commission |> List.sortBy (fun cm -> cm.name)
    for i in 0..sortedCommission.Length - 1 do
        fprintfn writer "%s" (sortedCommission |> List.item i).name

    printSitting writer day.firstSitting
    printSitting writer day.secondSitting

    fprintfn writer ""

let printFitness writer fitness =
    fprintfn writer "Удобство дат для членов комиссии: %i" fitness.commissionDates
    fprintfn writer "Удобство дат для председателей: %i" fitness.chairDates
    fprintfn writer "Удобство дат для студентов: %i" fitness.studentDates
    fprintfn writer "Удобство дат для научников: %i" fitness.advisorDates
    fprintfn writer "Интересность тем: %i" fitness.themes
    fprintfn writer "Штраф за общее количество дней комиссии: %i" fitness.totalDaysCommissionPenalty
    fprintfn writer "Штраф за общее количество дней научника: %i" fitness.totalDaysAdvisorPenalty
    fprintfn writer "Сбалансированность размеров комиссий: %i" fitness.commissionBalance
    fprintfn writer "Итого: %i" (totalFitness fitness)

let printTotalDays writer schedule =
    schedule.totalDays
    |> Seq.iteri (fun i pair -> fprintfn writer "%i. %s --- %i" i pair.Key.name pair.Value)

let initParallel generator count =
    Seq.replicate 10 (async { return Seq.init (count / 10) generator |> Seq.toList }) 
    |> Async.Parallel 
    |> Async.RunSynchronously 
    |> Seq.concat 
    |> Seq.toList

[<EntryPoint>]
let main _ =
    let sittingSlots = DataCollector.getSlots ()
    let works = DataCollector.getWorks ()
    let commissionMembers = DataCollector.getCommissionMembers ()
    let linkedWorks = DataCollector.getLinkedWorks ()
    let linkedWorks = 
        linkedWorks
        |> Seq.map (fun lw -> 
            lw |> List.map (fun author -> 
                works |> List.filter (fun w -> w.author.StartsWith author) |> List.exactlyOne))

    let advisorPreferences = DataCollector.getAdvisorPreferences ()

    let chairs = commissionMembers |> List.filter (fun c -> c.isChair)
    let commissionMembers = commissionMembers |> List.filter (fun c -> not c.isChair)

    let calculateFitness s = calculateFitness advisorPreferences commissionMembers chairs works s
    let generateSchedule () = generateSchedule works commissionMembers chairs sittingSlots linkedWorks
    let mutate s = mutate linkedWorks commissionMembers s

    let generationSize = 300
    let replicationCount = 10
    let wildMutations = 100
    let generations = 150

    let startTime = System.DateTime.Now

    let zeroGeneration = initParallel (fun _ -> generateSchedule ()) generationSize
    let failedToGenerate = zeroGeneration |> List.filter Option.isNone |> List.length
    
    let mutable generation = zeroGeneration |> List.filter Option.isSome |> Seq.ofList

    generation |> Seq.map (fun s -> (s, calculateFitness s)) |> Seq.toList |> ignore

    for i in 0..generations do
        generation <-
            generation
            |> Seq.map (fun s -> List.replicate replicationCount s)
            |> Seq.concat
            |> Seq.map Option.get
            |> PSeq.map mutate |> PSeq.toList
            |> Seq.map Some
            |> Seq.append generation
            |> Seq.append (initParallel (fun _ -> generateSchedule ()) wildMutations)
            |> Seq.distinct

        let generationWithFitness = generation |> PSeq.map (fun s -> (s, calculateFitness s)) |> PSeq.toList
        let diedOff = (generationWithFitness |> Seq.filter (fun (_, f) -> Option.isNone f) |> Seq.length)
        
        let sorted = 
            generationWithFitness 
            |> Seq.filter (fun (_, f) -> Option.isSome f) 
            |> Seq.toList
            |> List.sortByDescending (fun (_, f) -> totalFitness f.Value)

        generation <- sorted |> List.map fst |> (fun l -> if l.Length < generationSize then l else List.take generationSize l) |> Seq.ofList

        if not sorted.IsEmpty then
            let _, topFitness = sorted.Head
            printfn "Поколение: %i из %i, размер: %i, лучший результат %i, померло %i" i generations (Seq.length generation) (totalFitness topFitness.Value) diedOff

    let schedules =
        let generationWithFitness = generation |> Seq.map (fun s -> (s, calculateFitness s))
        if (generationWithFitness |> Seq.length) = 0 then
            Seq.empty
        else
            generationWithFitness |> Seq.take 5

    let endTime = System.DateTime.Now

    if schedules |> Seq.isEmpty then
        printfn "Расписание не найдено"
    else
        printfn ""
        printfn ""
        let mutable i = 0
        for schedule, fitness in schedules do
            let myDocsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) 
            let fullPath = Path.Combine(myDocsPath, $"schedule-{i}.txt")
            use writer = new StreamWriter(path = fullPath)

            let schedule = schedule.Value
            schedule.schedule |> List.iter (printDay writer)
        
            fprintfn writer ""
        
            printFitness writer fitness.Value

            fprintfn writer ""
            printTotalDays writer schedule

            i <- i + 1
        
        printfn ""
        printfn "Сгенерировалось некорректных расписаний: %i" failedToGenerate
        printfn "Время генерации: %f с" (endTime - startTime).TotalSeconds

    0
