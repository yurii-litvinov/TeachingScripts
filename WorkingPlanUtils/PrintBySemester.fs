module PrintBySemester

open CommonUtils

open CurriculumParser

let printBySemester plan semester =
    let printDiscipline (discipline: Discipline) =
        let implementation = discipline.Implementations |> Seq.find (fun i -> i.Semester = int semester)
        let workHours = implementation.WorkHours.Split(' ')
        let lectures = int workHours[0]
        let practices = [1; 3; 4; 5] |> List.map (fun i -> int workHours[i]) |> List.sum
        let homework = int workHours[9] + int workHours[11]
        printfn "[%s] %s\t%i\t%i\t%i" discipline.Code discipline.RussianName lectures practices homework

    let curriculum = DocxCurriculum(planCodeToFileName plan)
    curriculum.Disciplines
    |> Seq.filter (fun d -> d.Implementations |> Seq.exists (fun i -> i.Semester = int semester))
    |> Seq.sortBy (fun d -> d.Code)
    |> Seq.iter (fun d -> printDiscipline d)