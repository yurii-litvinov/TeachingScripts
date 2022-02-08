open CurriculumParser
open System.IO

[<EntryPoint>]
let main _ =

    let VM_5666 = ([ 2, "../WorkingPlans/ВМ.5666-2020.docx"
                     4, "../WorkingPlans/ВМ.5666-2019.docx" ],
                     ("1VYwkYbHPxkNIWkk-JDvFzrd6H3n1tmwKzAlupp4YxGw", "Sheet1"))

    let SV_5080 = ([ 2, "../WorkingPlans/СВ.5080-2020.docx"
                     4, "../WorkingPlans/СВ.5080-2019.docx"
                     6, "../WorkingPlans/СВ.5080-2018.docx"
                     8, "../WorkingPlans/СВ.5080-2017.docx" ],
                    ("1J9LP41sq55Sy9v1jFbOpdUJbChlRUQOGfHltfmci9e0", "Sheet1"))

    let workingPlans = fst SV_5080
    let target = snd SV_5080

    let curriculums = workingPlans |> List.map (fun (sem, plan) -> (sem, DocxCurriculum(plan)))

    let disciplines = 
        curriculums
        |> Seq.map (
            fun (sem, curriculum) -> 
                (sem, curriculum.Disciplines 
                |> Seq.filter (
                    fun d -> 
                        d.Implementations 
                        |> Seq.exists (fun i -> i.Semester = sem))))
        |> Seq.map (fun (sem, ds) -> (sem, ds |> Seq.sortBy (fun d -> d.Code)))

    let columnA = seq {
        // yield "Рег. № и наименование дисциплины"
        for sem in disciplines do
            let semesterName = 
                workingPlans 
                |> Map.ofList 
                |> fun m -> Path.GetFileNameWithoutExtension(FileInfo(m.[fst sem]).Name)

            yield $"семестр {fst sem} ({semesterName})"
            for d in snd sem do
                yield $"[{d.Code}] {d.RussianName}"
    }

    let columnB = seq {
        // yield "Форма ПА (указать: зачет/ экзамен)"
        for sem in disciplines do
            yield ""
            for d in snd sem do
                let impl = d.Implementations |> Seq.find (fun i -> i.Semester = fst sem)
                yield impl.MonitoringTypes
    }

    // let columnC = seq {
    //     yield "Результаты ПА (процентное соотношение оценок обучающихся по дисциплине из отчета Учебного Управления)"
    // }

    // let columnD = seq {
    //     yield "Замечания по ФОС"
    // }

    // let columnE = seq {
    //     yield "Выводы и рекомендации"
    // }

    let service = LessPainfulGoogleSheets.openGoogleSheet "AssignmentMatcher"

    LessPainfulGoogleSheets.writeGoogleSheetColumn service (fst target) (snd target) "A" 2 columnA
    LessPainfulGoogleSheets.writeGoogleSheetColumn service (fst target) (snd target) "B" 2 columnB
    // LessPainfulGoogleSheets.writeGoogleSheetColumn service (fst target) (snd target) "C" 1 columnC
    // LessPainfulGoogleSheets.writeGoogleSheetColumn service (fst target) (snd target) "D" 1 columnD
    // LessPainfulGoogleSheets.writeGoogleSheetColumn service (fst target) (snd target) "E" 1 columnE

    0