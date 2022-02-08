module AdcOperational

open AdcDomain
open CurriculumUtils
open HardcodedData
open WorkDistributionTypes
open OfficialWorkDistributionParser
open EnglishTeachers
open RoomData

open System.Text.RegularExpressions
open System.Threading
open CurriculumParser
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging

let physicalTrainingTeacher () =
    match Config.programStartYear with
    | 2017 -> physicalTraining2017Teacher
    | 2018 -> physicalTraining2018Teacher
    | 2020 -> physicalTraining2020Teacher
    | _ -> failwith "No data!"

let physicalTrainingWorkTypes semester trajectory =
    match (Config.programStartYear, semester) with
    | 2017, 1 -> physicalTraining2017Sem1WorkTypes
    | 2017, 2 -> physicalTraining2017Sem2WorkTypes
    | 2017, 3 -> physicalTraining2017Sem3WorkTypes
    | 2017, 4 -> physicalTraining2017Sem4WorkTypes
    | 2018, semester -> physicalTraining2018WorkTypes.[semester].[trajectory]
    | 2020, 1 -> failwith "No data!"
    | 2020, 2 -> physicalTraining2020Sem2WorkTypes
    | 2020, 3 -> physicalTraining2020Sem3WorkTypes
    | 2020, 4 -> physicalTraining2020Sem4WorkTypes
    | _ -> failwith "No data!"

let onlineWorkTypes () =
    match Config.programStartYear with
    | 2017 -> online2017WorkTypes
    | 2018 -> online2018WorkTypes
    | 2019 -> online2019WorkTypes
    | 2020 -> online2020WorkTypes
    | _ -> failwith "No data!"

let englishWorkTypes trajectory semester =
    englishWorkTypes.[Config.programStartYear].[trajectory].[semester]

let makeRoomsMap workTypes rooms =
    workTypes |> Seq.map fst |> Seq.map (fun w -> w, rooms) |> Map.ofSeq

let extendRoomMap (rooms: Map<string, string list>) workTypes =
    let oneOf workTypes (map: Map<string, string list>) =
        let first = workTypes |> Seq.tryFind (fun wt -> map.ContainsKey wt)
        match first with
        | Some wt -> map.[wt]
        | None -> []

    let workTypes = workTypes |> Seq.map fst
    let mutable result: Map<string, string list> = Map.empty

    for workType in workTypes do
        // Trying to get best matching rooms list.
        let roomsForWorkType = 
            if List.contains workType CurriculumUtils.lectionWorkTypes then 
                oneOf (workType :: CurriculumUtils.lectionWorkTypes) rooms
            elif List.contains workType CurriculumUtils.practiceWorkTypes then
                oneOf (workType :: CurriculumUtils.practiceWorkTypes) rooms
            else
                failwith $"Unknown work type {workType}"
        // If failed, trying to get rooms for any work type.
        let roomsForWorkType = 
            if roomsForWorkType = [] then 
                rooms |> Map.toList |> List.head |> snd 
            else 
                roomsForWorkType

        if roomsForWorkType = [] then
            failwith $"No rooms for work type '{workType}' found!"

        result <- result.Add(workType, roomsForWorkType)
    result

let englishRooms trajectory semester =
    let workTypes = englishWorkTypes trajectory semester
    makeRoomsMap workTypes englishRooms.[Config.programStartYear].[trajectory]

let isRelevantExperience teacher semester =
    if irrelevantIndustrialExperience.Contains teacher then 
        false
    elif limitedIndustrialExperience.ContainsKey teacher then 
        let expStartYear = limitedIndustrialExperience.[teacher]
        match semester with
        | 1 | 2 -> Config.programStartYear - expStartYear >= 3
        | 3 | 4 -> Config.programStartYear - expStartYear >= 2
        | 5 | 6 -> Config.programStartYear - expStartYear >= 1
        | 7 | 8 -> Config.programStartYear - expStartYear >= 0
        | _ -> failwith "Unknown semester!"
    else
        true

let addTeachers teachers workTypes semester =
    switchTab Teachers
    for teacher in teachers do
        addTeacher teacher workTypes (isRelevantExperience teacher semester)

let processTypicalRecord teachers workTypes rooms semester =
    wipeOutTeachers ()
    wipeOutRooms ()
    addTeachers teachers workTypes semester
    refresh ()
    removeWorkTypes 0 (workTypes |> Seq.length)
    addRooms rooms

let processLectionWithPractices lecturers lecturerWorkTypes practitioners practicionerWorkTypes rooms semester =
    wipeOutTeachers ()
    wipeOutRooms ()

    addTeachers lecturers lecturerWorkTypes semester
    addTeachers practitioners practicionerWorkTypes semester
    refresh ()
    removeWorkTypes 0 (lecturerWorkTypes @ practicionerWorkTypes |> Seq.length)
    addRooms rooms

let doEnglishMagic semester (englishTeachers: EnglishTeachers) =
    Thread.Sleep 1000
    let trajectoryFullName = getRecordCaption ()
    if trajectoryFullName.Contains("Английский язык по специальности") then
        processTypicalRecord 
            englishForSpecialPurposesTeachers.[semester]
            englishForSpecialPurposesWorkTypes.[semester]
            (makeRoomsMap englishForSpecialPurposesWorkTypes.[semester] englishForSpecialPurposesRooms.[semester])
            semester
    else
        let regexMatch = Regex.Match(trajectoryFullName, @".*(Траектория \d).*")
        let trajectory = if regexMatch.Success then regexMatch.Groups.[1].Value else ""
        match trajectory with 
        | "Траектория 1" -> 
            // Trajectory 1 was never implemented, so there is no teacher pool or rooms for it.
            wipeOutTeachers ()
            wipeOutRooms ()
            markAsDone ()
        | "Траектория 2"
        | "Траектория 3"
        | "Траектория 4" ->
            let teachers = 
                if englishTeachers.HaveData semester trajectory then
                    englishTeachers.Teachers semester trajectory
                else
                    additionalEnglistTeachers Config.programStartYear semester trajectory

            processTypicalRecord 
                teachers
                (englishWorkTypes trajectory semester)
                (englishRooms trajectory semester)
                semester
        | _ -> failwith "Unknown trajectory!"

let doRussianMagic semester =
    Thread.Sleep 1000
    wipeOutTeachers ()
    wipeOutRooms ()
    if Config.rkiWasImplemented then
        let trajectoryFullName = getRecordCaption ()
        let regexMatch = Regex.Match(trajectoryFullName, @".*(Траектория \d).*")
        let trajectory = if regexMatch.Success then regexMatch.Groups.[1].Value else ""
        processTypicalRecord 
            rkiTeacher
            // Russian was not implemented, so we need to add 0 as work hours
            (rkiWorkTypes.[Config.programStartYear].[trajectory].[semester] |> List.map (fun (wt, _) -> wt, 0))
            (makeRoomsMap rkiWorkTypes.[Config.programStartYear].[trajectory].[semester] rkiRoom)
            semester
    else
        markAsDone ()

let doPreGraduationMagic semester =
    Thread.Sleep 1000
    wipeOutTeachers ()
    wipeOutRooms ()
    addTeachers 
        (preGraduationTeachers Config.programStartYear)
        (preGraduationWorkTypes Config.programStartYear)
        semester 

    let vkrAdvisors = VkrAdvisorsParser.parseVkrAdvisors Config.vkrAdvisorsFileName
    vkrAdvisors 
    |> Seq.iter (fun student -> 
        addTeacher 
            student.advisor 
            ["Под руководством преподавателя", 60] 
            (isRelevantExperience student.advisor semester)
            )
    addRooms (makeRoomsMap ["Под руководством преподавателя", 60] computerClasses)

let doMagicForDiscipline disciplineRow (curriculum: Curriculum) (workload: IWorkDistribution) (englishTeachers: EnglishTeachers) (roomInfo: RoomData) =
    let disciplineName = getDisciplineNameFromTable disciplineRow
    let semester = getSemesterFromTable disciplineRow

    openRecord disciplineRow

    if notImplementedCourses.[Config.programStartYear].Contains (disciplineName, semester) then
        wipeOutTeachers ()
        wipeOutRooms ()
        refresh ()
        markAsDone ()
    else 
        match disciplineName with 
        | "Английский язык" 
        | "Английский язык по специальности" ->
            doEnglishMagic semester englishTeachers
        | "Немецкий язык" ->
            processTypicalRecord 
                germanTeachers.[semester]
                germanWorkTypes.[semester]
                (makeRoomsMap germanWorkTypes.[semester] germanRooms.[semester])
                semester
        | "Испанский язык" ->
            processTypicalRecord 
                spanishTeachers.[semester]
                spanishWorkTypes.[semester]
                (makeRoomsMap spanishWorkTypes.[semester] spanishRooms.[semester])
                semester
        | "Французский язык" ->
            processTypicalRecord 
                frenchTeachers.[semester]
                frenchWorkTypes.[semester]
                (makeRoomsMap frenchWorkTypes.[semester] frenchRooms.[semester])
                semester
        | "Английский язык в сфере профессиональной коммуникации" ->
            let rooms = roomInfo.Rooms semester disciplineName
            processTypicalRecord 
                mastersEnglishTeachers.[semester]
                mastersEnglishWorkTypes.[semester]
                (extendRoomMap rooms (curriculum.GetWorkTypes disciplineName semester))
                semester
        | "Немецкий язык в сфере профессиональной коммуникации" ->
            ()  // Do nothing. Was not implemented this year.
        | "Учебная практика 1 (научно-исследовательская работа)" 
        | "Учебная практика 1" 
        | "Учебная практика 2 (научно-исследовательская работа)" 
        | "Учебная практика 2" 
        | "Производственная практика (проектно-технологическая)"
        | "Производственная практика (научно-исследовательская работа)"
        | "Производственная практика" ->
            processTypicalRecord 
                ["Литвинов Юрий Викторович"; "Литвинов Юрий Викторович"; "Литвинов Юрий Викторович"; "Литвинов Юрий Викторович"]
                practiceWorkTypes
                (makeRoomsMap practiceWorkTypes computerClasses)
                semester
        | "Курсовая работа 1" 
        | "Курсовая работа 2" 
        | "Курсовая работа 3" ->
            processTypicalRecord 
                courseWorkTeachers
                courseWorkWorkTypes
                (makeRoomsMap courseWorkWorkTypes computerClasses)
                semester
        | "Подготовка выпускной квалификационной работы"  ->
            processTypicalRecord 
                courseWorkTeachers
                vkrWorkTypes
                (makeRoomsMap vkrWorkTypes computerClasses)
                semester
        | "Производственная практика (преддипломная)"
        | "Преддипломная практика" ->
            doPreGraduationMagic semester
        | "Учебная практика (научно-исследовательская работа)"
        | "Научно-исследовательская (производственная) практика" ->
            processTypicalRecord 
                ["Абрамов Максим Викторович"; "Ловягин Юрий Никитич"]
                mastersPracticeWorkTypes
                (makeRoomsMap mastersPracticeWorkTypes computerClasses)
                semester
        | "Научно-производственная практика" ->
            processTypicalRecord 
                ["Луцив Дмитрий Вадимович"; "Ловягин Юрий Никитич"]
                mastersPracticeWorkTypes
                (makeRoomsMap mastersPracticeWorkTypes computerClasses)
                semester
        | "Педагогическая практика" ->
            processTypicalRecord 
                ["Ловягин Юрий Никитич"; "Тулупьева Татьяна Валентиновна"]
                mastersPedagogicalPracticeWorkTypes
                (makeRoomsMap mastersPedagogicalPracticeWorkTypes ["4511"])
                semester
        | "Физическая культура и спорт" ->
            let trajectoryFullName = getRecordCaption ()
            let trajectory = if trajectoryFullName.Contains "Спортивный" then "Спортивный" else "Основной"
            processTypicalRecord 
                (physicalTrainingTeacher ())
                (physicalTrainingWorkTypes semester trajectory)
                (makeRoomsMap (physicalTrainingWorkTypes semester trajectory) physicalTrainingRooms)
                semester
        | "Русский язык как иностранный" ->
            doRussianMagic semester
        | "Администрирование информационных систем (на английском языке)" ->
            processTypicalRecord
                additionalTeachers.[semester].[disciplineName]
                ["Лекции", 24; "Промежуточная аттестация (зач)", 2]
                (extendRoomMap (roomInfo.Rooms semester disciplineName) ["Лекции", 24; "Промежуточная аттестация (зач)", 2])
                semester
        | _ ->
            if onlineCourses.ContainsKey disciplineName then
                if disciplineName = "Безопасность жизнедеятельности (онлайн-курс)" then
                    processTypicalRecord 
                        onlineCourses.[disciplineName]
                        onlineSafetyWorkTypes
                        (makeRoomsMap onlineSafetyWorkTypes computerClasses)
                        semester
                else
                    processTypicalRecord 
                        onlineCourses.[disciplineName]
                        (onlineWorkTypes ())
                        (makeRoomsMap (onlineWorkTypes ()) computerClasses)
                        semester
            elif facultatives.ContainsKey disciplineName then
                processTypicalRecord 
                    facultatives.[disciplineName]
                    facultativeWorkTypes.[disciplineName]
                    (makeRoomsMap facultativeWorkTypes.[disciplineName] computerClasses)
                    semester
            else
                let teachers = 
                    if commonUniversityCourses.ContainsKey disciplineName then 
                        {lecturers = commonUniversityCourses.[disciplineName]; practicioners = []}
                    elif workload.HaveData semester disciplineName then
                        workload.Teachers semester disciplineName
                    else
                        {lecturers = additionalTeachers.[semester].[disciplineName]; practicioners = []}

                if teachers.practicioners = [] && teachers.lecturers = [] then
                    failwith "No data!"

                let rooms = if knownRooms.ContainsKey disciplineName then knownRooms.[disciplineName] else []
                let rooms = 
                    if rooms = [] then 
                        roomInfo.Rooms semester disciplineName 
                    else 
                        makeRoomsMap (curriculum.GetWorkTypes disciplineName semester) rooms

                if teachers.lecturers <> [] && teachers.practicioners <> [] then
                    processLectionWithPractices
                        teachers.lecturers
                        (curriculum.GetLectionWorkTypes disciplineName semester)
                        teachers.practicioners
                        (curriculum.GetPracticeWorkTypes disciplineName semester)
                        (extendRoomMap rooms (curriculum.GetWorkTypes disciplineName semester))
                        semester
                else
                    processTypicalRecord
                        (teachers.lecturers @ teachers.practicioners)
                        (curriculum.GetWorkTypes disciplineName semester)
                        (extendRoomMap rooms (curriculum.GetWorkTypes disciplineName semester))
                        semester
   
        if knownSoftware.ContainsKey disciplineName then
            addSoftware knownSoftware.[disciplineName] (curriculum.GetWorkTypes disciplineName semester |> Seq.length)

        refresh ()

let logIn () =
    let login, password = 
        System.IO.File.ReadLines(Config.credentialsFileName) 
        |> Seq.toList 
        |> fun list -> (List.head list, list |> List.item 1)

    logIn login password

let doMagic filter =
    let curriculum = Curriculum(Config.curriculumFileName)
    let englishTeachers = EnglishTeachers(Config.workPlan)
    let rooms = RoomData(Config.roomDataSheetId)

    logIn ()
    switchFilter filter

    let workload = OfficialWorkDistributionParser()

    doMagicForDiscipline 1 curriculum workload englishTeachers rooms

let doEverythingRight () =
    let curriculum = Curriculum(Config.curriculumFileName)
    let englishTeachers = EnglishTeachers(Config.workPlan)
    let rooms = RoomData(Config.roomDataSheetId)

    logIn ()
    switchFilter NotStarted

    let workload = OfficialWorkDistributionParser()

    while tableSize () > 0 do
        doMagicForDiscipline 1 curriculum workload englishTeachers rooms

        backToTable ()

let autoAddRooms rooms =
    logIn ()
    let curriculum = Curriculum(Config.curriculumFileName)

    //switchFilter InProgress

    let disciplineName = getDisciplineNameFromTable 1
    let semester = getSemesterFromTable 1

    openRecord 1
    
    if disciplineName = "Адаптация и обучение в Университете (ЭО)" then
        addRooms (makeRoomsMap ["Промежуточная аттестация (зач)", 2] rooms)
    else
        addRooms (makeRoomsMap (curriculum.GetWorkTypes disciplineName semester) rooms)

let autoAddSoftware software =
    logIn ()
    let curriculum = Curriculum(Config.curriculumFileName)

    switchFilter InProgress

    let disciplineName = getDisciplineNameFromTable 1
    let semester = getSemesterFromTable 1

    openRecord 1
    
    addSoftware software (curriculum.GetWorkTypes disciplineName semester |> Seq.length)

let autoCorrectRoomsForFirst () =
    let roomInfo = RoomData(Config.roomDataSheetId)
    let curriculum = Curriculum(Config.curriculumFileName)

    logIn ()

    switchFilter InProgress

    let row = 1
    let disciplineName = getDisciplineNameFromTable row
    let semester = getSemesterFromTable row
    let rooms = roomInfo.Rooms semester disciplineName
    let workTypes = curriculum.GetWorkTypes disciplineName semester

    if rooms <> Map.empty then
        openRecord row
        wipeOutRooms ()
        addRooms (extendRoomMap rooms workTypes)

let autoCorrectRoomsForAll filter offset =
    let roomInfo = RoomData(Config.roomDataSheetId)
    let curriculum = Curriculum(Config.curriculumFileName)

    logIn ()

    switchFilter filter

    let disciplines = tableSize ()
    for row in [(1 + offset)..disciplines] do
        let disciplineName = getDisciplineNameFromTable row
        let semester = getSemesterFromTable row
        let rooms = roomInfo.Rooms semester disciplineName
        let workTypes = curriculum.GetWorkTypes disciplineName semester

        if rooms <> Map.empty then
            openRecord row
            wipeOutRooms ()
            addRooms (extendRoomMap rooms workTypes)
            backToTable ()

(*
let autoAddRoomsForAll () =
    let roomInfo = RoomData(Config.roomDataSheetId)
    let curriculum = Curriculum(Config.curriculumFileName)

    logIn ()

    let disciplines = tableSize ()
    for row in [1..disciplines] do
        let disciplineName = getDisciplineNameFromTable row
        let semester = getSemesterFromTable row

        if notImplementedCourses.[Config.programStartYear].Contains (disciplineName, semester) |> not then
            match disciplineName with 
            | "Английский язык" 
            | "Английский язык по специальности" ->
                openRecord row
                Thread.Sleep 1000
                let trajectoryFullName = getRecordCaption ()
                if trajectoryFullName.Contains("Английский язык по специальности") then
                    addRooms (extendRoomMap englishForSpecialPurposesWorkTypes.[semester] englishForSpecialPurposesRooms.[semester])
                else
                    let regexMatch = Regex.Match(trajectoryFullName, @".*(Траектория \d).*")
                    let trajectory = if regexMatch.Success then regexMatch.Groups.[1].Value else ""
                    match trajectory with 
                    | "Траектория 1" -> 
                        // Trajectory 1 was never implemented, so there is no teacher pool or rooms for it.
                        wipeOutTeachers ()
                        wipeOutRooms ()
                        markAsDone ()
                    | "Траектория 2"
                    | "Траектория 3"
                    | "Траектория 4" ->
                        processTypicalRecord 
                            teachers
                            (englishWorkTypes trajectory semester)
                            (englishRooms trajectory semester)
                            semester
                    | _ -> failwith "Unknown trajectory!"

            | "Немецкий язык" ->
                processTypicalRecord 
                    germanTeachers.[semester]
                    germanWorkTypes.[semester]
                    (makeRoomsMap germanWorkTypes.[semester] germanRooms.[semester])
                    semester
            | "Испанский язык" ->
                processTypicalRecord 
                    spanishTeachers.[semester]
                    spanishWorkTypes.[semester]
                    (makeRoomsMap spanishWorkTypes.[semester] spanishRooms.[semester])
                    semester
            | "Французский язык" ->
                processTypicalRecord 
                    frenchTeachers.[semester]
                    frenchWorkTypes.[semester]
                    (makeRoomsMap frenchWorkTypes.[semester] frenchRooms.[semester])
                    semester
            | "Английский язык в сфере профессиональной коммуникации" ->
                let rooms = roomInfo.Rooms semester disciplineName
                processTypicalRecord 
                    mastersEnglishTeachers.[semester]
                    mastersEnglishWorkTypes.[semester]
                    (extendRoomMap rooms (curriculum.GetWorkTypes disciplineName semester))
                    semester
            | "Немецкий язык в сфере профессиональной коммуникации" ->
                ()  // Do nothing. Was not implemented this year.
            | "Учебная практика 1 (научно-исследовательская работа)" 
            | "Учебная практика 1" 
            | "Учебная практика 2 (научно-исследовательская работа)" 
            | "Учебная практика 2" 
            | "Производственная практика (проектно-технологическая)"
            | "Производственная практика (научно-исследовательская работа)"
            | "Производственная практика" ->
                processTypicalRecord 
                    ["Литвинов Юрий Викторович"; "Литвинов Юрий Викторович"; "Литвинов Юрий Викторович"; "Литвинов Юрий Викторович"]
                    practiceWorkTypes
                    (makeRoomsMap practiceWorkTypes computerClasses)
                    semester
            | "Курсовая работа 1" 
            | "Курсовая работа 2" 
            | "Курсовая работа 3" ->
                processTypicalRecord 
                    courseWorkTeachers
                    courseWorkWorkTypes
                    (makeRoomsMap courseWorkWorkTypes computerClasses)
                    semester
            | "Подготовка выпускной квалификационной работы"  ->
                processTypicalRecord 
                    courseWorkTeachers
                    vkrWorkTypes
                    (makeRoomsMap vkrWorkTypes computerClasses)
                    semester
            | "Производственная практика (преддипломная)"
            | "Преддипломная практика" ->
                doPreGraduationMagic semester
            | "Учебная практика (научно-исследовательская работа)"
            | "Научно-исследовательская (производственная) практика" ->
                processTypicalRecord 
                    ["Абрамов Максим Викторович"; "Ловягин Юрий Никитич"]
                    mastersPracticeWorkTypes
                    (makeRoomsMap mastersPracticeWorkTypes computerClasses)
                    semester
            | "Научно-производственная практика" ->
                processTypicalRecord 
                    ["Луцив Дмитрий Вадимович"; "Ловягин Юрий Никитич"]
                    mastersPracticeWorkTypes
                    (makeRoomsMap mastersPracticeWorkTypes computerClasses)
                    semester
            | "Педагогическая практика" ->
                processTypicalRecord 
                    ["Ловягин Юрий Никитич"; "Тулупьева Татьяна Валентиновна"]
                    mastersPedagogicalPracticeWorkTypes
                    (makeRoomsMap mastersPedagogicalPracticeWorkTypes ["4511"])
                    semester
            | "Физическая культура и спорт" ->
                let trajectoryFullName = getRecordCaption ()
                let trajectory = if trajectoryFullName.Contains "Спортивный" then "Спортивный" else "Основной"
                processTypicalRecord 
                    (physicalTrainingTeacher ())
                    (physicalTrainingWorkTypes semester trajectory)
                    (makeRoomsMap (physicalTrainingWorkTypes semester trajectory) physicalTrainingRooms)
                    semester
            | "Русский язык как иностранный" ->
                doRussianMagic semester
            | "Администрирование информационных систем (на английском языке)" ->
                processTypicalRecord
                    additionalTeachers.[semester].[disciplineName]
                    ["Лекции", 24; "Промежуточная аттестация (зач)", 2]
                    (extendRoomMap (roomInfo.Rooms semester disciplineName) ["Лекции", 24; "Промежуточная аттестация (зач)", 2])
                    semester
            | _ ->
                if onlineCourses.ContainsKey disciplineName then
                    if disciplineName = "Безопасность жизнедеятельности (онлайн-курс)" then
                        processTypicalRecord 
                            onlineCourses.[disciplineName]
                            onlineSafetyWorkTypes
                            (makeRoomsMap onlineSafetyWorkTypes computerClasses)
                            semester
                    else
                        processTypicalRecord 
                            onlineCourses.[disciplineName]
                            (onlineWorkTypes ())
                            (makeRoomsMap (onlineWorkTypes ()) computerClasses)
                            semester
                elif facultatives.ContainsKey disciplineName then
                    processTypicalRecord 
                        facultatives.[disciplineName]
                        facultativeWorkTypes.[disciplineName]
                        (makeRoomsMap facultativeWorkTypes.[disciplineName] computerClasses)
                        semester
                else
                    let teachers = 
                        if commonUniversityCourses.ContainsKey disciplineName then 
                            {lecturers = commonUniversityCourses.[disciplineName]; practicioners = []}
                        elif workload.HaveData semester disciplineName then
                            workload.Teachers semester disciplineName
                        else
                            {lecturers = additionalTeachers.[semester].[disciplineName]; practicioners = []}

                    if teachers.practicioners = [] && teachers.lecturers = [] then
                        failwith "No data!"

                    let rooms = if knownRooms.ContainsKey disciplineName then knownRooms.[disciplineName] else []
                    let rooms = 
                        if rooms = [] then 
                            roomInfo.Rooms semester disciplineName 
                        else 
                            makeRoomsMap (curriculum.GetWorkTypes disciplineName semester) rooms

                    if teachers.lecturers <> [] && teachers.practicioners <> [] then
                        processLectionWithPractices
                            teachers.lecturers
                            (curriculum.GetLectionWorkTypes disciplineName semester)
                            teachers.practicioners
                            (curriculum.GetPracticeWorkTypes disciplineName semester)
                            (extendRoomMap rooms (curriculum.GetWorkTypes disciplineName semester))
                            semester
                    else
                        processTypicalRecord
                            (teachers.lecturers @ teachers.practicioners)
                            (curriculum.GetWorkTypes disciplineName semester)
                            (extendRoomMap rooms (curriculum.GetWorkTypes disciplineName semester))
                            semester




            let rooms = roomInfo.Rooms semester disciplineName
            let workTypes = curriculum.GetWorkTypes disciplineName semester

            if rooms <> Map.empty then
                openRecord row
                addRooms (extendRoomMap rooms workTypes)
                backToTable ()
*)

let printCurriculumDisciplines semester =
    let curriculum = DocxCurriculum(Config.curriculumFileName)
    curriculum.Disciplines 
    |> Seq.map (fun d -> d.Implementations) 
    |> Seq.concat 
    |> Seq.filter (fun i -> i.Semester = semester)
    |> Seq.map (fun i -> 
        let disciplineType = 
            match i.Discipline.Type with
            | DisciplineType.Base -> "Базовая"
            | DisciplineType.Elective -> "Электив"
            | DisciplineType.Facultative -> "Факультатив"
            | _ -> failwith "Unknown discipline type"
        $"{i.Discipline.RussianName} {disciplineType}")
    |> Seq.sort
    |> Seq.iter (printfn "%s")

let checkWorkDistribution semester =
    let curriculum = DocxCurriculum(Config.curriculumFileName)
    let disciplines = 
        curriculum.Disciplines 
        |> Seq.map (fun d -> d.Implementations) 
        |> Seq.concat 
        |> Seq.filter (fun i -> i.Semester = semester)
        |> Seq.map (fun i -> i.Discipline.RussianName)
        |> Seq.sort

    let workload = OfficialWorkDistributionParser() :> IWorkDistribution
    disciplines |> Seq.iter (fun d -> if not (workload.HaveData semester d) then printfn "%s" d)

let checkRpds rpdFolder =
    let curriculum = DocxCurriculum(Config.curriculumFileName)
    let disciplines = curriculum.Disciplines
    let rpds = System.IO.Directory.EnumerateFiles rpdFolder
    for discipline in disciplines do
        if rpds |> Seq.tryFind (fun s -> s.Contains(discipline.Code)) |> Option.isNone then
            printfn "[%s] %s" discipline.Code discipline.RussianName

let createSoftwareReport rpdFolder =
    use result = System.IO.File.Create("SoftwareReport.txt")
    use writer = new System.IO.StreamWriter(result)
    let rpds = System.IO.Directory.EnumerateFiles rpdFolder
    for rpd in rpds do
        let wordDocument = WordprocessingDocument.Open(rpd, false)
        let body = wordDocument.MainDocumentPart.Document.Body
        let text = body.InnerText
        let disciplineCode = System.IO.FileInfo(rpd).Name.Substring(0, 6)
        let disciplineName = System.IO.FileInfo(rpd).Name.Substring(7).Replace(".docx", "")
        let commonSoftHeader = "Характеристики аудиторного оборудования, в том числе неспециализированного компьютерного оборудования и программного обеспечения общего пользования"
        let specialSoftHeader = "Характеристики специализированного программного обеспечения"
        if text.Contains commonSoftHeader
            && text.Contains "3.3.3" 
            && text.Contains specialSoftHeader
            && text.Contains "3.3.5" 
        then
            writer.WriteLine $"[{disciplineCode}] {disciplineName}"
            let commonSoft = text.Substring(text.IndexOf(commonSoftHeader))
            let commonSoft = commonSoft.Replace(commonSoftHeader, "")
            let commonSoft = commonSoft.Substring(0, commonSoft.IndexOf("3.3.3"))

            let specialSoft = text.Substring(text.IndexOf(specialSoftHeader))
            let specialSoft = specialSoft.Replace(specialSoftHeader, "")
            let specialSoft = specialSoft.Substring(0, specialSoft.IndexOf("3.3.5"))

            writer.WriteLine $"{commonSoft}"
            writer.WriteLine ()
            writer.WriteLine $"{specialSoft}"
            writer.WriteLine ()
        else
            printfn "%s" $"[{disciplineCode}] {disciplineName} --- missing required fields"
