module AdcOperational

open AdcDomain
open CurriculumUtils
open WorkDistributionData
open HardcodedData

open System.Text.RegularExpressions
open System.Threading

let addTeachers teachers workTypes =
    switchTab Teachers
    for teacher in teachers do
        addTeacher teacher workTypes

let addTypicalRecordWithoutOpening teachers workTypes rooms =
    addTeachers teachers workTypes
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

let processLectionWithPracticesNoOpen lecturer lecturerWorkTypes practitioners practicionerWorkTypes rooms =
    addTeachers [lecturer] lecturerWorkTypes
    addTeachers practitioners practicionerWorkTypes
    refresh ()
    removeWorkTypes 0 (lecturerWorkTypes @ practicionerWorkTypes |> Seq.length)
    addRooms rooms (lecturerWorkTypes @ practicionerWorkTypes |> Seq.length)

let processLectionWithPractices recordNumber lecturer lecturerWorkTypes practitioners practicionerWorkTypes rooms =
    openRecord recordNumber
    processLectionWithPracticesNoOpen lecturer lecturerWorkTypes practitioners practicionerWorkTypes rooms

let doEnglishMagic row semester =
    openRecord row
    Thread.Sleep 1000
    let trajectoryFullName = getRecordCaption ()
    let regexMatch = Regex.Match(trajectoryFullName, @".*(Траектория \d).*")
    let trajectory = if regexMatch.Success then regexMatch.Groups.[1].Value else ""
    match trajectory with 
    | "Траектория 1" -> 
        // Do nothing (yet). Trajectory 1 was never implemented, so there is no teacher pool or rooms for it.
        ()
    | "Траектория 2" -> 
        addTypicalRecordWithoutOpening 
            englishT2Teachers
            (match semester with 
            | 3 -> englishT2S3WorkTypes
            | 4 -> englishT2S4WorkTypes
            | _ -> failwith "Unknown semester"
            )
            englishT2Rooms
    | "Траектория 3" -> 
        addTypicalRecordWithoutOpening 
            englishT3Teachers
            (match semester with 
            | 3 -> englishT3S3WorkTypes
            | 4 -> englishT3S4WorkTypes
            | _ -> failwith "Unknown semester"
            )
            englishT3Rooms
    | "Траектория 4" -> 
        addTypicalRecordWithoutOpening 
            englishT4Teachers
            (match semester with 
            | 3 -> englishT4S3WorkTypes
            | 4 -> englishT4S4WorkTypes
            | _ -> failwith "Unknown semester"
            )
            englishT3Rooms
    | _ -> failwith "Unknown trajectory!"

let doRussianMagic row semester =
    openRecord row
    Thread.Sleep 1000
    let trajectoryFullName = getRecordCaption ()
    let regexMatch = Regex.Match(trajectoryFullName, @".*(Траектория \d).*")
    let trajectory = if regexMatch.Success then regexMatch.Groups.[1].Value else ""
    match trajectory with 
    | "Траектория 1" -> 
        addTypicalRecordWithoutOpening 
            rkiTeacher
            (match semester with 
            | 3 -> rkiT1S3WorkTypes
            | 4 -> rkiT1S4WorkTypes
            | _ -> failwith "Unknown semester"
            )
            rkiRoom
    | "Траектория 2" -> 
        addTypicalRecordWithoutOpening 
            rkiTeacher
            (match semester with 
            | 3 -> rkiT2S3WorkTypes
            | 4 -> rkiT2S4WorkTypes
            | _ -> failwith "Unknown semester"
            )
            rkiRoom
    | _ -> failwith "Unknown trajectory!"

let doMagicForFirstDiscipline (curriculum: Curriculum) (workload: WorkDistribution) =
    let disciplineRow = 1

    let disciplineName = getDisciplineNameFromTable disciplineRow
    let semester = getSemesterFromTable disciplineRow

    match disciplineName with 
    | "Английский язык" ->
        doEnglishMagic disciplineRow semester
    | "Учебная практика 1 (научно-исследовательская работа)" 
    | "Учебная практика 2 (научно-исследовательская работа)" 
    | "Производственная практика (проектно-технологическая)" ->
        processTypicalRecordNoWipe 
            1 
            ["Литвинов Юрий Викторович"; "Литвинов Юрий Викторович"; "Литвинов Юрий Викторович"; "Литвинов Юрий Викторович"]
            practiceWorkTypes
            computerClasses
    | "Физическая культура и спорт" ->
        // Physical training also has trajectories and they actually have different teachers. But we don't have enough
        // information, so we use Timetable data and curriculum, where all trajectories are basically the same.
        processTypicalRecordNoWipe 
            1 
            physicalTrainingTeacher
            physicalTrainingSem4WorkTypes
            physicalTrainingRooms
    | "Русский язык как иностранный" ->
        doRussianMagic disciplineRow semester
    | _ ->
        if onlineCourses.ContainsKey disciplineName then
            processTypicalRecordNoWipe 
                1 
                onlineCourses.[disciplineName]
                onlineWorkTypes
                computerClasses
        else
            let teachers = workload.Teachers disciplineName

            if teachers.practicioners = [] then
                failwith "No data!"

            let rooms = if knownRooms.ContainsKey disciplineName then knownRooms.[disciplineName] else []

            if teachers.lecturer = "" then
                processTypicalRecordNoWipe
                    disciplineRow
                    teachers.practicioners
                    (curriculum.GetWorkTypes disciplineName semester)
                    rooms
            else
                processLectionWithPractices
                    disciplineRow
                    teachers.lecturer
                    (curriculum.GetLectionWorkTypes disciplineName semester)
                    teachers.practicioners
                    (curriculum.GetPracticeWorkTypes disciplineName semester)
                    rooms

    refresh ()

let logIn pathToAdcCredentials =
    let login, password = 
        System.IO.File.ReadLines(pathToAdcCredentials) 
        |> Seq.toList 
        |> fun list -> (List.head list, list |> List.item 1)

    logIn login password

let doMagic curriculimFileName pathToAdcCredentials filter =
    let curriculum = Curriculum(curriculimFileName)

    logIn pathToAdcCredentials
    switchFilter filter

    let semester = getSemesterFromTable 1
    let workload = WorkDistribution(semester)

    doMagicForFirstDiscipline curriculum workload

let doEverythingRight curriculimFileName pathToAdcCredentials =
    let curriculum = Curriculum(curriculimFileName)

    logIn pathToAdcCredentials
    switchFilter NotStarted

    let semester = getSemesterFromTable 1
    let workload = WorkDistribution(semester)

    while tableSize () > 0 do
        doMagicForFirstDiscipline curriculum workload

        backToTable ()

let autoAddRooms pathToAdcCredentials rooms workTypes =
    logIn pathToAdcCredentials

    switchFilter InProgress
    openRecord 1
    
    addRooms rooms workTypes
