module ComparePlans

open CommonUtils

open CurriculumParser

let compareToSemester oldPlan newPlan semester =
    let oldCurriculum = DocxCurriculum(planCodeToFileName oldPlan)
    let newCurriculum = DocxCurriculum(planCodeToFileName newPlan)

    let filterBySemester (disciplines: Discipline seq) =
        disciplines |> Seq.filter (fun d -> d.Implementations |> Seq.exists(fun i -> i.Semester <= semester))

    let oldPlanDisciplines = oldCurriculum.Disciplines |> filterBySemester
    let newPlanDisciplines = newCurriculum.Disciplines |> filterBySemester

    let sharedDisciplines = oldPlanDisciplines |> Seq.filter (fun d -> newPlanDisciplines |> Seq.exists (fun nd -> d.Code = nd.Code))

    let str (discipline: Discipline) = sprintf "[%s] '%s'" discipline.Code discipline.RussianName

    if semester < 8 then
        printfn "Сравнение учебных планов %s и %s до семестра %i. \n\n" oldPlan newPlan semester

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

let comparePlans oldPlan newPlan =
    compareToSemester oldPlan newPlan 8