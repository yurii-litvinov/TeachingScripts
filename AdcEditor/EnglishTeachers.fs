module EnglishTeachers

open LessPainfulXlsx

type EnglishTeacherRecord =
    {
        name: string
        semester: int
        workingPlan: string
        group: string
        trajectory: string
    }

type EnglishTeachers (workPlan: string) =

    let getData () =
        let sheet = openXlsxSheet "../../Английский язык. 03.12.2020.xlsx" "Лист1"

        let semester = readColumnByName sheet "Семестр курса" |> Seq.toList
        let workingPlan = readColumnByName sheet "Код учебного плана" |> Seq.toList
        let group = readColumnByName sheet "Группа" |> Seq.toList
        let trajectory =  readColumnByName sheet "Траектория" |> Seq.toList
        let surname =  readColumnByName sheet "Фамилия преподавателя" |> Seq.toList
        let name =  readColumnByName sheet "Имя, отчество преподавателя" |> Seq.toList

        let fullNames = Seq.zip surname name |> Seq.map (fun (s, n) -> $"{s} {n}")
        let works = Seq.zip3 semester workingPlan trajectory
        let records = 
            Seq.zip3 fullNames works group
            |> Seq.map (fun (n, (s, w, t), g) -> {
                                                name = n
                                                semester = s.Substring(2) |> int
                                                workingPlan = w
                                                group = g
                                                trajectory = t.Substring(0, "Траектория 9".Length)
                                              })
        let records = records |> Seq.filter (fun r -> r.workingPlan = workPlan) |> Seq.filter (fun r -> r.group.Contains("MATH"))

        let data = records |> Seq.groupBy (fun r -> r.semester) |> Map.ofSeq
        let data = data |> Map.map (fun _ v -> v |> Seq.groupBy (fun r -> r.trajectory) |> Map.ofSeq)
        let data = data |> Map.map (fun _ m -> m |> Map.map (fun _ v -> v |> Seq.map (fun r -> r.name) |> List.ofSeq))
        data

    let data = getData ()

    member _.Teachers semester trajectory =
        data.[semester].[trajectory]

    member _.HaveData semester trajectory =
        data.ContainsKey semester && data.[semester].ContainsKey trajectory
