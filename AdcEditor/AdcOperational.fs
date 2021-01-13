module AdcOperational

open AdcDomain
open CurriculumUtils
open HardcodedData
open WorkDistributionTypes
open OfficialWorkDistributionParser
open EnglishTeachers

open System.Text.RegularExpressions
open System.Threading

let physicalTrainingTeacher () =
    match Config.programStartYear with
    | 2017 -> physicalTraining2017Teacher
    | 2020 -> physicalTraining2020Teacher
    | _ -> failwith "No data!"

let physicalTrainingWorkTypes semester =
    match (Config.programStartYear, semester) with
    | 2017, 1 -> physicalTraining2017Sem1WorkTypes
    | 2017, 2 -> physicalTraining2017Sem2WorkTypes
    | 2017, 3 -> physicalTraining2017Sem3WorkTypes
    | 2017, 4 -> physicalTraining2017Sem4WorkTypes
    | 2020, 1 -> failwith "No data!"
    | 2020, 2 -> physicalTraining2020Sem2WorkTypes
    | 2020, 3 -> physicalTraining2020Sem3WorkTypes
    | 2020, 4 -> physicalTraining2020Sem4WorkTypes
    | _ -> failwith "No data!"

let onlineWorkTypes () =
    match Config.programStartYear with
    | 2017 -> online2017WorkTypes
    | 2020 -> online2020WorkTypes
    | _ -> failwith "No data!"

let englishWorkTypes trajectory semester =
    englishWorkTypes.[Config.programStartYear].[trajectory].[semester]

let englishRooms trajectory =
    englishRooms.[Config.programStartYear].[trajectory]

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

let addTypicalRecordWithoutOpening teachers workTypes rooms semester =
    addTeachers teachers workTypes semester
    refresh ()
    removeWorkTypes 0 (workTypes |> Seq.length)
    addRooms rooms (workTypes |> Seq.length)

let addTypicalRecord recordNumber teachers workTypes rooms =
    openRecord recordNumber
    addTypicalRecordWithoutOpening teachers workTypes rooms

let processTypicalRecordNoWipe recordNumber teachers workTypes rooms =
    openRecord recordNumber
    addTypicalRecordWithoutOpening teachers workTypes rooms

let processTypicalRecord recordNumber teachers workTypes rooms =
    openRecord recordNumber

    wipeOutTeachers ()
    wipeOutRooms ()

    addTypicalRecordWithoutOpening teachers workTypes rooms

let processLectionWithPracticesNoOpen lecturers lecturerWorkTypes practitioners practicionerWorkTypes rooms semester =
    addTeachers lecturers lecturerWorkTypes semester
    addTeachers practitioners practicionerWorkTypes semester
    refresh ()
    removeWorkTypes 0 (lecturerWorkTypes @ practicionerWorkTypes |> Seq.length)
    addRooms rooms (lecturerWorkTypes @ practicionerWorkTypes |> Seq.length)

let processLectionWithPractices recordNumber lecturers lecturerWorkTypes practitioners practicionerWorkTypes rooms =
    openRecord recordNumber
    processLectionWithPracticesNoOpen lecturers lecturerWorkTypes practitioners practicionerWorkTypes rooms

let doEnglishMagic row semester (englishTeachers: EnglishTeachers) =
    openRecord row
    Thread.Sleep 1000
    wipeOutTeachers ()
    wipeOutRooms ()
    let trajectoryFullName = getRecordCaption ()
    let regexMatch = Regex.Match(trajectoryFullName, @".*(Траектория \d).*")
    let trajectory = if regexMatch.Success then regexMatch.Groups.[1].Value else ""
    match trajectory with 
    | "Траектория 1" -> 
        // Do nothing (yet). Trajectory 1 was never implemented, so there is no teacher pool or rooms for it.
        ()
    | "Траектория 2"
    | "Траектория 3"
    | "Траектория 4" ->
        addTypicalRecordWithoutOpening 
            (englishTeachers.Teachers semester trajectory)
            (englishWorkTypes trajectory semester)
            (englishRooms trajectory)
            semester
    | _ -> failwith "Unknown trajectory!"

let doRussianMagic row semester =
    openRecord row
    Thread.Sleep 1000
    wipeOutTeachers ()
    wipeOutRooms ()
    let trajectoryFullName = getRecordCaption ()
    let regexMatch = Regex.Match(trajectoryFullName, @".*(Траектория \d).*")
    let trajectory = if regexMatch.Success then regexMatch.Groups.[1].Value else ""
    addTypicalRecordWithoutOpening 
        rkiTeacher
        // Russian was not implemented, so we need to add 0 as work hours
        (rkiWorkTypes.[Config.programStartYear].[trajectory].[semester] |> List.map (fun (wt, _) -> wt, 0))
        rkiRoom

let doPreGraduationMagic row semester =
    openRecord row
    Thread.Sleep 1000
    addTeachers 
        preGraduationTeachers
        preGraduationWorkTypes
        semester 

    let vkrAdvisors = VkrAdvisorsParser.parseVkrAdvisors Config.vkrAdvisorsFileName
    vkrAdvisors 
    |> Seq.iter (fun student -> 
        addTeacher 
            student.advisor 
            ["Под руководством преподавателя", 60] 
            (isRelevantExperience student.advisor semester)
            )
    addRooms computerClasses 3

let doMagicForFirstDiscipline (curriculum: Curriculum) (workload: IWorkDistribution) (englishTeachers: EnglishTeachers) =
    let disciplineRow = 1

    let disciplineName = getDisciplineNameFromTable disciplineRow
    let semester = getSemesterFromTable disciplineRow

    match disciplineName with 
    | "Английский язык" ->
        doEnglishMagic disciplineRow semester englishTeachers
    | "Учебная практика 1 (научно-исследовательская работа)" 
    | "Учебная практика 2 (научно-исследовательская работа)" 
    | "Производственная практика (проектно-технологическая)"
    | "Производственная практика (научно-исследовательская работа)" ->
        processTypicalRecordNoWipe 
            1 
            ["Литвинов Юрий Викторович"; "Литвинов Юрий Викторович"; "Литвинов Юрий Викторович"; "Литвинов Юрий Викторович"]
            practiceWorkTypes
            computerClasses
            semester
    | "Курсовая работа 1" 
    | "Курсовая работа 2" 
    | "Курсовая работа 3" ->
        processTypicalRecordNoWipe 
            1 
            courseWorkTeachers
            courseWorkWorkTypes
            computerClasses
            semester
    | "Производственная практика (преддипломная)" ->
        doPreGraduationMagic disciplineRow semester
    | "Физическая культура и спорт" ->
        // Physical training also has trajectories and they actually have different teachers. But we don't have enough
        // information, so we use Timetable data and curriculum, where all trajectories are basically the same.
        processTypicalRecordNoWipe 
            1 
            (physicalTrainingTeacher ())
            (physicalTrainingWorkTypes semester)
            physicalTrainingRooms
            semester
    | "Русский язык как иностранный" ->
        doRussianMagic disciplineRow semester semester
    | _ ->
        if onlineCourses.ContainsKey disciplineName then
            if disciplineName = "Безопасность жизнедеятельности (онлайн-курс)" then
                processTypicalRecordNoWipe 
                    1 
                    onlineCourses.[disciplineName]
                    onlineSafetyWorkTypes
                    computerClasses
                    semester
            else
                processTypicalRecordNoWipe 
                    1 
                    onlineCourses.[disciplineName]
                    (onlineWorkTypes ())
                    computerClasses
                    semester
        elif facultatives.ContainsKey disciplineName then
            processTypicalRecordNoWipe 
                1 
                facultatives.[disciplineName]
                ["Промежуточная аттестация (зач)", 0]
                computerClasses
                semester
        else
            let teachers = workload.Teachers semester disciplineName

            if teachers.practicioners = [] && teachers.lecturers = [] then
                failwith "No data!"

            let rooms = if knownRooms.ContainsKey disciplineName then knownRooms.[disciplineName] else []

            if teachers.lecturers <> [] && teachers.practicioners <> [] then
                processLectionWithPractices
                    disciplineRow
                    teachers.lecturers
                    (curriculum.GetLectionWorkTypes disciplineName semester)
                    teachers.practicioners
                    (curriculum.GetPracticeWorkTypes disciplineName semester)
                    rooms
                    semester
            else
                processTypicalRecordNoWipe
                    disciplineRow
                    (teachers.lecturers @ teachers.practicioners)
                    (curriculum.GetWorkTypes disciplineName semester)
                    rooms
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

    logIn ()
    switchFilter filter

    let workload = OfficialWorkDistributionParser()

    doMagicForFirstDiscipline curriculum workload englishTeachers

let doEverythingRight () =
    let curriculum = Curriculum(Config.curriculumFileName)
    let englishTeachers = EnglishTeachers(Config.workPlan)

    logIn ()
    switchFilter NotStarted

    let workload = OfficialWorkDistributionParser()

    while tableSize () > 0 do
        doMagicForFirstDiscipline curriculum workload englishTeachers

        backToTable ()

let autoAddRooms rooms =
    logIn ()
    let curriculum = Curriculum(Config.curriculumFileName)

    switchFilter InProgress

    let disciplineName = getDisciplineNameFromTable 1
    let semester = getSemesterFromTable 1

    openRecord 1
    
    addRooms rooms (curriculum.GetWorkTypes disciplineName semester |> Seq.length)

let autoAddSoftware software =
    logIn ()
    let curriculum = Curriculum(Config.curriculumFileName)

    switchFilter InProgress

    let disciplineName = getDisciplineNameFromTable 1
    let semester = getSemesterFromTable 1

    openRecord 1
    
    addSoftware software (curriculum.GetWorkTypes disciplineName semester |> Seq.length)
