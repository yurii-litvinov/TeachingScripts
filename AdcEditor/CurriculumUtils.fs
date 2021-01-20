module CurriculumUtils

open CurriculumParser

let lectionWorkTypes = ["Лекции"; "Консультации"; "Коллоквиумы"; "Промежуточная аттестация (экз)"]
let practiceWorkTypes = [
    "Семинары" 
    "Практические занятия"
    "Лабораторные работы"
    "Контрольные работы"
    "Текущий контроль (ауд)"
    "Промежуточная аттестация (зач)"
    "Под руководством преподавателя"
    "В присутствии преподавателя"]

type Curriculum(fileName: string) =
    let curriculum = DocxCurriculum(fileName)

    member _.GetWorkTypes disciplineName semester =
        let curriculumDiscipline = curriculum.Disciplines |> Seq.find(fun d -> d.RussianName = disciplineName)
        let impl = curriculumDiscipline.Implementations |> Seq.find(fun d -> d.Semester = semester)
        let workHours = impl.WorkHours.Split [|' '|] |> Seq.rev |> Seq.skip 4 |> Seq.rev |> Seq.map int

        let attestationTypes = impl.MonitoringTypes
        let attestationWorkType = 
            match attestationTypes with
            | "зачёт" -> ["Промежуточная аттестация (зач)"]
            | "экзамен" -> ["Промежуточная аттестация (экз)"]
            | "текущий контроль" -> [""]
            | "зачёт, экзамен" -> ["Промежуточная аттестация (зач)"; "Промежуточная аттестация (экз)"]
            | _ -> failwith (sprintf "Unknown monitoring type: %s" attestationTypes)

        let workTypes = 
            [
                "Лекции"; 
                "Семинары"; 
                "Консультации"; 
                "Практические занятия"; 
                "Лабораторные работы"; 
                "Контрольные работы"; 
                "Коллоквиумы"; 
                "Текущий контроль (ауд)"
            ] @ attestationWorkType @ [
                "Под руководством преподавателя"; 
                "В присутствии преподавателя"
            ] 

        if List.length attestationWorkType = 1 then
            Seq.zip workTypes workHours |> Seq.filter (fun wt -> snd wt > 0) |> Seq.toList
        else
            let attestationHours = workHours |> Seq.item 8
            let workHours = seq {yield! Seq.take 8 workHours; yield attestationHours / 2; yield attestationHours / 2; yield! Seq.skip 9 workHours}
            Seq.zip workTypes workHours |> Seq.filter (fun wt -> snd wt > 0) |> Seq.toList

    member v.GetLectionWorkTypes disciplineName semester =
        let workTypes = v.GetWorkTypes disciplineName semester
        workTypes |> List.filter (fun wt -> List.contains (fst wt) lectionWorkTypes)

    member v.GetPracticeWorkTypes disciplineName semester =
        let workTypes = v.GetWorkTypes disciplineName semester
        workTypes |> List.filter (fun wt -> List.contains (fst wt) practiceWorkTypes)