module OfficialWorkDistributionParser

open LessPainfulXlsx
open WorkDistributionTypes
open LessPainfulGoogleSheets
open HardcodedData
open CurriculumUtils

type WorkDistributionRecord =
    {
        teacher: string
        discipline: string
        semester: int
        workType: string
    }

type OfficialWorkDistributionParser() =

    let knownTeacherData () =
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

    let fixIncorrectTeacherNames names =
        names |> Seq.map (fun s -> if s = "Нестеров В В" then "Нестеров Владимир Викторович" else s)

    let teacherDataFromFirstCourse () =
        let sheet = openXlsxSheet "../../МатОбес РПП бак маг.xlsx" "17_5006_1 курс"
        let surnameColumn = readColumnByName sheet "Фамилия преподавателя (заполнить)" |> Seq.map (fun s -> s.Trim()) |> Seq.toList
        let nameColumn = readColumnByName sheet "Имя (заполнить)" |> Seq.map (fun s -> s.Trim()) |> Seq.toList
        let middleNameColumn = readColumnByName sheet "Отчество (заполнить)" |> Seq.map (fun s -> s.Trim()) |> Seq.toList

        let names = Seq.zip3 surnameColumn nameColumn middleNameColumn |> Seq.map (fun (s, n, m) -> $"{s} {n} {m}")
        names |> fixIncorrectTeacherNames |> Seq.toList

    let groupBySemester records =
        records |> Seq.groupBy (fun r -> r.semester) |> Map.ofSeq

    let getFirstCourseOf5006'2017Data () =
        let sheet = openXlsxSheet "../../МатОбес РПП бак маг.xlsx" "17_5006_1 курс"
        let disciplineColumn = readColumnByName sheet "Наименование дисц" |> Seq.toList
        let semesterColumn = readColumnByName sheet "Период обучения" |> Seq.toList
        let workTypeColumn = readColumnByName sheet "Вид учебной работы" |> Seq.toList
        let surnameColumn = readColumnByName sheet "Фамилия преподавателя (заполнить)" |> Seq.map (fun s -> s.Trim()) |> Seq.toList
        let nameColumn = readColumnByName sheet "Имя (заполнить)" |> Seq.map (fun s -> s.Trim()) |> Seq.toList
        let middleNameColumn = readColumnByName sheet "Отчество (заполнить)" |> Seq.map (fun s -> s.Trim()) |> Seq.toList

        let names = Seq.zip3 surnameColumn nameColumn middleNameColumn |> Seq.map (fun (s, n, m) -> $"{s} {n} {m}")
        let names = names |> fixIncorrectTeacherNames |> Seq.toList
        let works = Seq.zip3 disciplineColumn semesterColumn workTypeColumn
        let records = 
            Seq.zip names works 
            |> Seq.map (fun (n, (d, s, wt)) -> {
                                                teacher = n; 
                                                discipline = d; 
                                                semester = s.Replace("Семестр ", "") |> int; 
                                                workType = wt
                                               })

        records |> Seq.filter (fun r -> r.semester = 1 || r.semester = 2) |> groupBySemester

    let correctYears =
        match Config.programStartYear with
        | 2017 -> fun (_, (_, semester, _), year) ->
            year = "2020-2021 уч. год" && (semester = 7 || semester = 8)
            || year = "2019-2020 уч. год" && (semester = 5 || semester = 6)
            || year = "2018-2019 уч. год" && (semester = 3 || semester = 4)
        | 2018 -> fun (_, (_, semester, _), year) ->
            year = "2020-2021 уч. год" && (semester = 5 || semester = 6)
            || year = "2019-2020 уч. год" && (semester = 3 || semester = 4)
            || year = "2018-2019 уч. год" && (semester = 1 || semester = 2)
            // Here we are using last year's assignments since 4th course was not implemented yet and we need to guess 
            // work distribution. 2017's work distribution for 4th course is good enough.
            || year = "2020-2021 уч. год" && (semester = 7 || semester = 8)
        | 2019 -> fun (_, (_, semester, _), year) ->
            year = "2020-2021 уч. год" && (semester = 3 || semester = 4)
            || year = "2019-2020 уч. год" && (semester = 1 || semester = 2)
        | _ -> failwith "Unknown program start year!"

    let parseCommonSheet sheetName =
        let sheet = openXlsxSheet "../../МатОбес РПП бак маг.xlsx" sheetName
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
            |> Seq.filter correctYears
            |> Seq.map (fun (t, w, _) -> t, w)
            |> Seq.map (fun (n, (d, s, wt)) -> {
                                                teacher = n
                                                discipline = d
                                                semester = s
                                                workType = wt
                                               })
        records

    let identifyTeachers records =
        let knownTeachers = (knownTeacherData ()) @ (teacherDataFromFirstCourse ()) @ additionalTeacherFullNames |> List.distinct

        records 
        |> Seq.map (fun r ->
            let rawTeacher = r.teacher
            let rawTeacherSurname = rawTeacher.Substring(0, rawTeacher.IndexOf('.'))
            let candidates = knownTeachers |> List.filter (fun t -> t.StartsWith rawTeacherSurname)
            let fullName = 
                if candidates = [] then
                    //failwith $"Unknown teacher: {rawTeacher}"
                    printfn "Unknown teacher: %s" rawTeacher
                    ""
                elif candidates.Length > 1 then
                    failwith $"Ambiguous teacher: {rawTeacher}"
                else 
                    candidates.[0]
            {r with teacher = fullName}
            ) |> Seq.toList

    let get5006'2017OtherCoursesData () =
        let records = 
            parseCommonSheet "5006"
            |> Seq.filter (fun r -> r.semester <> 1)
            |> Seq.filter (fun r -> r.semester <> 2)

        identifyTeachers records |> groupBySemester

    let get5006'2018RawData () =
        let records = parseCommonSheet "5006"
        identifyTeachers records |> groupBySemester

    let getMastersRawData () =
        let records = parseCommonSheet "5665"
        identifyTeachers records |> groupBySemester

    let getTeacherOverride (discipline: WorkDistributionRecord) teacher =
        if teacher = "Пименов Александр Александрович" && (discipline.semester = 7 || discipline.semester = 8) then
            "Смирнов Михаил Николаевич"
        elif teacher = "Пименов Александр Александрович" && (discipline.semester = 5 || discipline.semester = 6) && Config.programStartYear = 2018 then
            "Смирнов Михаил Николаевич"
        elif teacherOverride.ContainsKey discipline.semester && teacherOverride.[discipline.semester].ContainsKey discipline.discipline then
            teacherOverride.[discipline.semester].[discipline.discipline]
        else
            teacher

    let foldDiscipline (discipline: WorkDistributionRecord seq) =
        let mergeLists list1 list2 =
            let groupedList1 = list1 |> List.groupBy id |> List.map (fun (k, v) -> (k, v.Length))
            let groupedList2 = list2 |> List.groupBy id |> List.map (fun (k, v) -> (k, v.Length))
            let result = 
                groupedList1 @ groupedList2 
                |> List.groupBy fst
                |> List.map (fun (k, s) -> (k, s |> Seq.map snd |> Seq.max))
                |> List.map (fun (k, v) -> List.replicate v k)
                |> List.concat
            result

        let getTeachersFor workTypes =
            discipline 
            |> Seq.groupBy (fun d -> d.workType)
            |> Seq.filter (fun (k, _) -> List.contains k workTypes)
            |> Seq.map (fun (_, s) -> s |> Seq.fold (fun acc r ->(getTeacherOverride r r.teacher) :: acc) [])
            |> Seq.fold mergeLists []

        let practicioners = getTeachersFor practiceWorkTypes
        let lecturers = getTeachersFor lectionWorkTypes |> List.distinct

        {lecturers = lecturers; practicioners = practicioners}

    let foldSemester (semester: WorkDistributionRecord seq) =
        let disciplineMap = semester |> Seq.groupBy (fun r -> r.discipline) |> Map.ofSeq
        disciplineMap |> Map.map (fun _ v -> foldDiscipline v)

    let foldMap map =
        map |> Map.map (fun _ v -> foldSemester v)

    let get5006'2017Data () =
        let firstCourse = getFirstCourseOf5006'2017Data ()
        let otherCourses = get5006'2017OtherCoursesData ()
        
        (Map.toList firstCourse) @ (Map.toList otherCourses) 
        |> Map.ofList 
        |> foldMap

    let get5665Data () =
        getMastersRawData() |>foldMap

    let get5006'2018Data () =
        get5006'2018RawData () |> foldMap

    let parsedData = 
        match Config.programStartYear with 
        | 2017 -> get5006'2017Data ()
        | 2018 -> get5006'2018Data ()
        | 2019 -> get5665Data ()
        | _ -> failwith "Unknown program start year!"

    interface IWorkDistribution with

        member _.Teachers semester discipline = parsedData.[semester].[discipline]

        member _.HaveData semester discipline =
            parsedData.ContainsKey semester && parsedData.[semester].ContainsKey discipline
