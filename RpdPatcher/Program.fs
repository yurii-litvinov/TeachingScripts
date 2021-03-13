open CurriculumParser
open System.IO
open System.Text.RegularExpressions

open LessPainfulGoogleSheets
open ProgramContentChecker
open ProgramPatcher

let plansFolder = "WorkingPlans"

let owners = Config.owners |> Map.ofList

let isProgramNameOk fileName =
    let matches = Regex.Match(fileName, @"^\d{6}_(\w|\s|\+|-|\.|,)+_\d{2}_\d{4}_\d(-\d)?с_\w+.docx$")
    matches.Success

let excluded code =
    Config.exclusions |> List.contains code

let getCode (disciplineFileName: string) =
    disciplineFileName.Substring(0, 6)

let planNameToCode fileName =
    FileInfo(fileName).Name.Substring(3, "9999-2084".Length)

let planCodeToFileName planCode =
    Directory.EnumerateFiles plansFolder
    |> Seq.find (fun f -> planNameToCode f = planCode)

let isKnownPlanCode planCode =
    Directory.EnumerateFiles plansFolder
    |> Seq.exists (fun f -> planNameToCode f = planCode)

let inventorize planCode programsFolder =
    let curriculum = DocxCurriculum(planCodeToFileName planCode)
    let disciplines = curriculum.Disciplines
    let programs = 
        Directory.EnumerateFiles programsFolder
        |> Seq.map (fun p -> FileInfo(p).Name)

    let printMissingDisciplines () =
        let missingDisciplines = 
            disciplines
            |> Seq.filter (fun d -> excluded d.Code |> not)
            |> Seq.filter (fun d -> programs |> Seq.tryFind (fun s -> s.Contains(d.Code)) |> Option.isNone)

        if missingDisciplines |> Seq.isEmpty then
            printfn "Всё в порядке."
        else
            printfn "Отсутствующие рабочие программы:"
            missingDisciplines
            |> Seq.sortBy (fun d -> d.Code)
            |> Seq.iter (fun d -> printfn "[%s] %s" d.Code d.RussianName)

    let printUnneededPrograms () =
        let unneededPrograms = 
            programs
            |> Seq.filter 
                (fun p -> 
                    disciplines 
                    |> Seq.tryFind (fun d -> d.Code = getCode p) 
                    |> Option.isNone)

        if unneededPrograms |> Seq.isEmpty |> not then
            printfn ""
            printfn "РПД, которых нет в учебном плане:"
            unneededPrograms
            |> Seq.sort 
            |> Seq.iter (printfn "%s")

    let printIncorrectlyNamedPrograms () =
        let incorrectlyNamedPrograms =
            programs
            |> Seq.filter (not << excluded << getCode)
            |> Seq.filter (isProgramNameOk >> not)

        if incorrectlyNamedPrograms |> Seq.isEmpty |> not then
            printfn ""
            printfn "Неправильно названные РПД:"
            incorrectlyNamedPrograms
            |> Seq.sort 
            |> Seq.iter (printfn "%s")

    printMissingDisciplines ()
    printUnneededPrograms ()
    printIncorrectlyNamedPrograms ()

let loadDisciplineOwners planCode =
    let service = openGoogleSheet "AssignmentMatcher"
    let sheet, page = owners.[planCode]
    let rawData = readGoogleSheet service sheet page "A" "B" 2
    rawData
    |> Seq.map (fun row -> (row |> Seq.head, if Seq.length row >= 2 then row |> Seq.item 1 else ""))
    |> Seq.map (fun (name, owner) -> (name.Substring(1, 6), owner.Split [|' '|] |> Array.item 0))
    |> Map.ofSeq

let autoRename planCode programsFolder =
    let curriculum = DocxCurriculum(planCodeToFileName planCode)
    let disciplines = curriculum.Disciplines
    let owners = loadDisciplineOwners planCode

    let findDiscipline code = disciplines |> Seq.find (fun d -> d.Code = code)

    let programs = 
        Directory.EnumerateFiles programsFolder
        |> Seq.map FileInfo

    programs
    |> Seq.filter (fun p -> isProgramNameOk p.Name |> not)
    |> Seq.filter (fun p -> excluded (getCode p.Name) |> not)
    |> Seq.iter 
        (fun p ->
            let code = p.Name.Substring(0, 6)
            let discipline = findDiscipline code
            let name = discipline.RussianName
            let year = curriculum.CurriculumCode.Substring(0, 2)
            let curriculumCode = curriculum.CurriculumCode.Substring(3, 4)
            let semesters = discipline.Implementations |> Seq.map (fun i -> i.Semester) |> Seq.sort
            let startSemester = semesters |> Seq.head
            let endSemester = semesters |> Seq.last
            let semesterPart = if startSemester = endSemester then string startSemester else $"{startSemester}-{endSemester}"
            let owner = if owners.ContainsKey code then owners.[code] else ""

            if p.Extension <> ".docx" then
                printfn "Файл '%s' не будет переименован, он не .docx!" p.Name
            elif owner = "" then
                printfn "Файл '%s' не будет переименован, неизвестен 'хозяин' дисциплины (см. соответствующую таблицу в Google Docs)!" p.Name
            else
                let newName = $"{code}_{name}_{year}_{curriculumCode}_{semesterPart}с_{owner}.docx"
                p.MoveTo($"{p.DirectoryName}/{newName}")
                printfn "Файл '%s' переименован в  '%s'" p.Name newName
        )

let deleteExcluded programsFolder =
    Directory.EnumerateFiles programsFolder
    |> Seq.map FileInfo
    |> Seq.filter (fun p -> excluded (getCode p.Name))
    |> Seq.iter (fun p -> 
        printfn "Удаляю '%s'" p.Name
        p.Delete ())

let checkContent programsFolder =
    Directory.EnumerateFiles programsFolder
    |> Seq.map FileInfo
    |> Seq.filter (fun p -> excluded (getCode p.Name) |> not)
    |> Seq.iter (fun p -> checkProgram p.FullName)

let patchPrograms planCode programsFolder =
    Directory.EnumerateFiles programsFolder
    |> Seq.map FileInfo
    |> Seq.filter (fun p -> excluded (getCode p.Name) |> not)
    |> Seq.iter (fun p -> patchProgram (planCodeToFileName planCode) p.FullName)

let patchProgram planCode programFileName =
    ProgramPatcher.patchProgram (planCodeToFileName planCode) programFileName

let printHelp () =
    printfn "Добро пожаловать в RPD Patcher, утилиту для проверки и автоматических исправлений РПД СПбГУ."
    printfn "Использование:"
    printfn "\tdotnet run -- -i <номер рабочего плана и год> <папка с РПД>"
    printfn "\t\t- проверить, что все РПД в наличии и канонично называются."
    printfn "\tdotnet run -- -r <номер рабочего плана и год> <папка с РПД>"
    printfn "\t\t- автоматически переименовать РПД, если их имена не в каноничном формате. Требует авторизации в Google Drive."
    printfn "\tdotnet run -- --delete-excluded <папка с РПД>"
    printfn "\t\t- удалить из папки все РПД, перечисленные в списке 'исключённые' в конфиге. Используйте с осторожностью!"
    printfn "\tdotnet run -- -c <номер рабочего плана и год> <папка с РПД>"
    printfn "\t\t- проверить структуру и содержимое всех РПД в папке. Проверяет следование приказу 539/1 от 17.02.2014 о форме РПД, ищет нарушения приказа 1395/1 и типичные ошибки."
    printfn "\tdotnet run -- -cp <файл РПД>"
    printfn "\t\t- как '-c', только для одной РПД."
    printfn "\tdotnet run -- -p <номер рабочего плана и год> <папка с РПД>"
    printfn "\t\t- автомагически исправляет типичные ошибки во всех РПД в папке и приводит их в соответствие приказу 1395/1. После этого ТРЕБУЕТСЯ ручная проверка!"
    printfn "\tdotnet run -- -pp <файл РПД>"
    printfn "\t\t- как '-p', только для одной РПД."
    printfn ""
    printfn "Имеющиеся рабочие планы:"
    Directory.EnumerateFiles(plansFolder) 
    |> Seq.map (fun p -> FileInfo(p).Name.Substring(3, "9999-2084".Length))
    |> Seq.iter (printf "%s ")

[<EntryPoint>]
let main argv =
    if argv.Length < 2 then
        printHelp ()
    else
        let call f =
            if argv.Length < 3 then
                printHelp ()
            else
                f argv.[1] argv.[2]

        match argv.[0] with
        | "--inventorize" | "-i" -> call inventorize 
        | "--auto-rename" | "-r" -> call autoRename
        | "--delete-excluded" -> deleteExcluded argv.[1]
        | "--check-content" | "-c" -> checkContent argv.[1]
        | "--check-program" | "-cp" -> checkProgram argv.[1]
        | "--patch" | "-p" -> call patchPrograms
        | "--patch-program" | "-pp" -> call patchProgram
        | _ -> printHelp ()
    0
