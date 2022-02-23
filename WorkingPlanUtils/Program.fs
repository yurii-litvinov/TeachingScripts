open CurriculumParser
open System.IO

let plansFolder = "../WorkingPlans"

let planNameToCode fileName =
    FileInfo(fileName).Name.Substring(3, "9999-2084".Length)

let planCodeToFileName planCode =
    Directory.EnumerateFiles (System.AppDomain.CurrentDomain.BaseDirectory + "/../../../" + plansFolder)
    |> Seq.find (fun f -> planNameToCode f = planCode)

let comparePlans oldPlan newPlan =
    let oldCurriculum = DocxCurriculum(planCodeToFileName oldPlan)
    let newCurriculum = DocxCurriculum(planCodeToFileName newPlan)

    let oldPlanDisciplines = oldCurriculum.Disciplines 
    let newPlanDisciplines = newCurriculum.Disciplines 

    let sharedDisciplines = oldPlanDisciplines |> Seq.filter (fun d -> newPlanDisciplines |> Seq.exists (fun nd -> d.Code = nd.Code))

    let str (discipline: Discipline) = sprintf "[%s] '%s'" discipline.Code discipline.RussianName

    printfn "Дисциплины плана %s, которых нет в %s:" oldPlan newPlan
    oldPlanDisciplines |> Seq.filter (fun d -> sharedDisciplines |> Seq.contains d |> not) |> Seq.iter (printfn "%s" << str)
    printfn ""

    printfn "Дисциплины плана %s, которых нет в %s:" newPlan oldPlan
    newPlanDisciplines |> Seq.filter (fun d -> sharedDisciplines |> Seq.exists (fun sd -> sd.Code = d.Code) |> not) |> Seq.iter (printfn "%s" << str)
    printfn ""

    let matchingPairs = sharedDisciplines |> Seq.map (fun d -> (d, newPlanDisciplines |> Seq.find (fun nd -> d.Code = nd.Code)))

    let laborIntensity (discipline: Discipline) = discipline.Implementations |> Seq.sumBy (fun i -> i.LaborIntensity)
    let attestationTypes (discipline: Discipline) = discipline.Implementations |> Seq.map (fun i -> i.MonitoringTypes) |> Seq.toList
    let semesters (discipline: Discipline) = discipline.Implementations |> Seq.map (fun i -> i.Semester) |> Seq.toList
    let hours (discipline: Discipline) = discipline.Implementations |> Seq.map (fun i -> i.WorkHours) |> Seq.toList

    printfn ""
    printfn "Разница в трудоёмкости:"
    matchingPairs 
    |> Seq.filter (fun (o, n) -> (laborIntensity o) <> (laborIntensity n))
    |> Seq.iter (fun (o, n) -> printfn "%s: трудоёмкость в %s: %d, в %s: %d" (str o) oldPlan (laborIntensity o) newPlan (laborIntensity n))

    printfn ""
    printfn "Разница в семестрах реализации:"
    matchingPairs 
    |> Seq.filter (fun (o, n) -> (semesters o) <> (semesters n))
    |> Seq.iter (fun (o, n) -> printfn "%s: семестры в %s: %A, в %s: %A" (str o) oldPlan (semesters o) newPlan (semesters n))

    printfn ""
    printfn "Разница в форме аттестации:"
    matchingPairs 
    |> Seq.filter (fun (o, n) -> (attestationTypes o) <> (attestationTypes n))
    |> Seq.iter (fun (o, n) -> printfn "%s: формы аттестации в %s: %A, в %s: %A" (str o) oldPlan (attestationTypes o) newPlan (attestationTypes n))

    printfn ""
    printfn "Разница в часах:"
    matchingPairs 
    |> Seq.filter (fun (o, n) -> (hours o) <> (hours n))
    |> Seq.iter (fun (o, n) -> printfn "%s: часы в %s: %A, в %s: %A" (str o) oldPlan (hours o) newPlan (hours n))

    printfn ""
    let oldPlanCompetences = oldCurriculum.Competences
    let newPlanCompetences = newCurriculum.Competences
    let sharedCompetences = oldPlanCompetences |> Seq.filter (fun c -> newPlanCompetences |> Seq.exists (fun nc -> c.Code = nc.Code && c.Description = nc.Description))

    printfn "Компетенции плана %s, которых нет в %s:" oldPlan newPlan
    oldPlanCompetences |> Seq.filter (fun с -> sharedCompetences |> Seq.contains с |> not) |> Seq.iter (fun c -> printfn "%s %s" c.Code c.Description)
    printfn ""

    printfn "Компетенции плана %s, которых нет в %s:" newPlan oldPlan
    newPlanCompetences |> Seq.filter (fun c -> sharedCompetences |> Seq.exists (fun oc -> oc.Code = c.Code && oc.Description = c.Description) |> not) |> Seq.iter (fun c -> printfn "%s %s" c.Code c.Description)
    printfn ""

    ()

let printBySemester plan semester =
    let curriculum = DocxCurriculum(planCodeToFileName plan)
    curriculum.Disciplines
    |> Seq.filter (fun d -> d.Implementations |> Seq.exists (fun i -> i.Semester = int semester))
    |> Seq.sortBy (fun d -> d.Code)
    |> Seq.iter (fun d -> printfn "[%s] %s" d.Code d.RussianName)


let printHelp () =
    printfn "Добро пожаловать в Working Plan Utilities, набор инструментов для работы с учебными планами."
    printfn "Использование:"
    printfn "\tdotnet run -- -p <номер рабочего плана и год> <номер семестра>"
    printfn "\t\t- распечатать отсортированный список дисциплин в данном семестре."
    printfn "\tdotnet run -- -c <номер рабочего плана и год 1> <номер рабочего плана и год 2>"
    printfn "\t\t- сравнить два учебных плана."
    printfn ""
    printfn "Имеющиеся учебные планы:"
    Directory.EnumerateFiles(plansFolder) 
    |> Seq.map (fun p -> FileInfo(p).Name.Substring(3, "9999-2084".Length))
    |> Seq.iter (printf "%s ")

[<EntryPoint>]
let main argv =
    if argv.Length < 2 then
        printHelp ()
    else
        let call f =
            if argv.Length < 3 then
                printHelp ()
            else
                f argv.[1] argv.[2]

        match argv.[0] with
        | "--print-by-semester" | "-p" -> call printBySemester
        | "--compare" | "-c" -> call comparePlans
        | _ -> printHelp ()
    0
