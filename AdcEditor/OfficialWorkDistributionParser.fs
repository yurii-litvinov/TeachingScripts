module OfficialWorkDistributionParser

open LessPainfulXlsx
open WorkDistributionTypes
open LessPainfulGoogleSheets
open HardcodedData

type WorkDistributionRecord =
    {
        teacher: string
        discipline: string
        semester: int
        workType: string
    }

type OfficialWorkDistributionParser() =

    let teacherData () =
        let mutable teachers: string list = []
        let service = openGoogleSheet "AssignmentMatcher"
        let pages = [
            "ТП, 1-2 семестры"
            "Матобес, 3-4 семестры"
            "Матобес, 5 семестр" 
            "Матобес, 6 семестр" 
            "Матобес, 7 семестр" 
            "Матобес, 8 семестр"]

        for page in pages do
            let rawData = 
                readGoogleSheet service "1aLSgrsDnWAiGD7nMbAts6SxHA9PybMaNlWOPiNgsk7w" page "A" "C" 2
                |> Seq.map (fun row -> if Seq.length row > 2 then row |> Seq.skip 2 |> Seq.head else "")
                |> Seq.filter ((<>) "")
        
            let parseWorkloadInfo (workloadString: string) = 
                let strings = workloadString.Split [|'\n'|]
                if strings.[0].StartsWith("Лекции") then
                    let lecturer = strings.[0].Replace("Лекции: ", "").Trim()
                    let practicioners = strings |> List.ofArray |> List.skip 2 |> List.map (fun s -> s.Trim())
                    lecturer :: practicioners
                else
                    let practicioners = strings |> List.ofArray |> List.map (fun s -> s.Trim())
                    practicioners

            teachers <- teachers @ (rawData |> Seq.map parseWorkloadInfo |> Seq.concat |> Seq.distinct |> List.ofSeq)

        teachers |> List.distinct

    let teacherDataFromFirstCourse () =
        let sheet = openXlsxSheet "../../МатОбес РПП бак маг.xlsx" "17_5006_1 курс"
        let surnameColumn = readColumnByName sheet "Фамилия преподавателя (заполнить)" |> Seq.map (fun s -> s.Trim()) |> Seq.toList
        let nameColumn = readColumnByName sheet "Имя (заполнить)" |> Seq.map (fun s -> s.Trim()) |> Seq.toList
        let middleNameColumn = readColumnByName sheet "Отчество (заполнить)" |> Seq.map (fun s -> s.Trim()) |> Seq.toList

        let names = Seq.zip3 surnameColumn nameColumn middleNameColumn |> Seq.map (fun (s, n, m) -> $"{s} {n} {m}")
        names |> Seq.map (fun s -> if s = "Нестеров В В" then "Нестеров Владимир Викторович" else s) |> Seq.toList

    let getFirstCourseData () =
        let sheet = openXlsxSheet "../../МатОбес РПП бак маг.xlsx" "17_5006_1 курс"
        let disciplineColumn = readColumnByName sheet "Наименование дисц" |> Seq.toList
        let semesterColumn = readColumnByName sheet "Период обучения" |> Seq.toList
        let workTypeColumn = readColumnByName sheet "Вид учебной работы" |> Seq.toList
        let surnameColumn = readColumnByName sheet "Фамилия преподавателя (заполнить)" |> Seq.toList
        let nameColumn = readColumnByName sheet "Имя (заполнить)" |> Seq.toList
        let middleNameColumn = readColumnByName sheet "Отчество (заполнить)" |> Seq.toList

        let names = Seq.zip3 surnameColumn nameColumn middleNameColumn |> Seq.map (fun (s, n, m) -> $"{s} {n} {m}")
        let works = Seq.zip3 disciplineColumn semesterColumn workTypeColumn
        let records = 
            Seq.zip names works 
            |> Seq.map (fun (n, (d, s, wt)) -> {
                                                teacher = n; 
                                                discipline = d; 
                                                semester = s.Replace("Семестр ", "") |> int; 
                                                workType = wt
                                               })
        let firstSemester = records |> Seq.filter (fun r -> r.semester = 1)
        let secondSemester = records |> Seq.filter (fun r -> r.semester = 2)
        (firstSemester, secondSemester)

    let getOtherCoursesData () =
        let correctYear (_, (_, semester, _), year) =
            year = "2020-2021 уч. год" && (semester = 7 || semester = 8)
            || year = "2019-2020 уч. год" && (semester = 5 || semester = 6)
            || year = "2018-2019 уч. год" && (semester = 3 || semester = 4)

        let sheet = openXlsxSheet "../../МатОбес РПП бак маг.xlsx" "5006"
        let disciplineColumn = readColumnByName sheet "Дисциплина" |> Seq.toList
        let semesterColumn = 
            readColumnByName sheet "Учебный период" 
            |> Seq.map (fun s -> s.Replace("Семестр ", "") |> int ) 
            |> Seq.toList
        let workTypeColumn = readColumnByName sheet "Виды учебной работы" |> Seq.toList
        let yearColumn = readColumnByName sheet "Учебный год" |> Seq.toList
        let teacherColumn = 
            readColumnByName sheet "Преподаватель" 
            |> Seq.map (fun t -> t.Remove(t.IndexOf(','))) 
            |> Seq.toList 

        let works = Seq.zip3 disciplineColumn semesterColumn workTypeColumn
        let records = 
            Seq.zip3 teacherColumn works yearColumn
            |> Seq.filter correctYear
            |> Seq.map (fun (t, w, _) -> t, w)
            |> Seq.map (fun (n, (d, s, wt)) -> {
                                                teacher = n
                                                discipline = d
                                                semester = s
                                                workType = wt
                                               })
            |> Seq.filter (fun r -> r.semester <> 1)
            |> Seq.filter (fun r -> r.semester <> 2)

        let knownTeachers = (teacherData ()) @ (teacherDataFromFirstCourse ()) @ additionalTeacherFullNames |> List.distinct

        let records =
            records 
            |> Seq.map (fun r ->
                let rawTeacher = r.teacher
                let rawTeacherSurname = rawTeacher.Substring(0, rawTeacher.IndexOf('.'))
                let candidates = knownTeachers |> List.filter (fun t -> t.StartsWith rawTeacherSurname)
                let fullName = 
                    if candidates = [] then
                        failwith (sprintf "Unknown teacher: %s" rawTeacher)
                    elif candidates.Length > 1 then
                        failwith (sprintf "Ambiguous teacher: %s" rawTeacher)
                    else 
                        candidates.[0]
                {r with teacher = fullName}
                ) |> Seq.toList

        let semester3 = records |> Seq.filter (fun r -> r.semester = 3)
        let semester4 = records |> Seq.filter (fun r -> r.semester = 4)
        let semester5 = records |> Seq.filter (fun r -> r.semester = 5)
        let semester6 = records |> Seq.filter (fun r -> r.semester = 6)
        let semester7 = records |> Seq.filter (fun r -> r.semester = 7)
        let semester8 = records |> Seq.filter (fun r -> r.semester = 8)
        (semester3, semester4, semester5, semester6, semester7, semester8)

    let teacherOverride (discipline: WorkDistributionRecord) teacher =
        if teacher = "Пименов Александр Александрович" && (discipline.semester = 7 || discipline.semester = 8) then
            "Смирнов Михаил Николаевич"
        else
            teacher

    let foldDiscipline (discipline: WorkDistributionRecord seq) =
        let lecturers = 
            discipline 
            |> Seq.fold (
                fun lecturers r ->
                    if r.workType = "Лекции" then 
                       (teacherOverride r r.teacher) :: lecturers
                    else
                        lecturers
                ) []
        let practicioners =
            discipline 
            |> Seq.fold (fun acc r -> 
                if r.workType = "Промежуточная аттестация (зач)" then 
                    (teacherOverride r r.teacher) :: acc
                else
                    acc
                ) []
        {lecturers = lecturers; practicioners = practicioners}

    let foldSemester (semester: WorkDistributionRecord seq) =
        let disciplineMap = semester |> Seq.groupBy (fun r -> r.discipline) |> Map.ofSeq
        disciplineMap |> Map.map (fun _ v -> foldDiscipline v)

    let firstSemester, secondSemester = getFirstCourseData ()
    let semester3, semester4, semester5, semester6, semester7, semester8 = getOtherCoursesData ()

    let parsedData =
        Map.empty
            .Add(1, foldSemester firstSemester)
            .Add(2, foldSemester secondSemester)
            .Add(3, foldSemester semester3)
            .Add(4, foldSemester semester4)
            .Add(5, foldSemester semester5)
            .Add(6, foldSemester semester6)
            .Add(7, foldSemester semester7)
            .Add(8, foldSemester semester8)

    interface IWorkDistribution with

        member _.Teachers semester discipline = parsedData.[semester].[discipline]
