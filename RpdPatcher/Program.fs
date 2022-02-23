open CurriculumParser
open System.IO
open System.Text.RegularExpressions

open LessPainfulGoogleSheets
open ProgramContentChecker
open ProgramPatcher
open System
open Argu

let plansFolder = "../WorkingPlans"

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
    Directory.EnumerateFiles (AppDomain.CurrentDomain.BaseDirectory + "/../../../" + plansFolder)
    |> Seq.find (fun f -> planNameToCode f = planCode)

let isKnownPlanCode planCode =
    Directory.EnumerateFiles plansFolder
    |> Seq.exists (fun f -> planNameToCode f = planCode)

let inventorize (planCode, programsFolder) =
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
            printfn "Все нужные дисциплины присутствуют."
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

    let printDuplicates () =
        let duplicates =
            programs 
            |> Seq.map (fun p -> (p, getCode p))
            |> Seq.groupBy (fun (_, code) -> code)
            |> Seq.filter (fun (_, s) -> s |> Seq.length > 1)
            |> Seq.map (fun (_, s) -> s)
            |> Seq.concat
            |> Seq.map fst

        if duplicates |> Seq.isEmpty |> not then
            printfn ""
            printfn "Дублирующиеся РПД:"
            duplicates
            |> Seq.sort 
            |> Seq.iter (printfn "%s")

    let printProgramsInOldDocFormat () =
        let programs =
            programs
            |> Seq.filter (not << excluded << getCode)
            |> Seq.filter (fun p -> FileInfo(p).Extension = ".doc")

        if programs |> Seq.isEmpty |> not then
            printfn ""
            printfn "РПД в старом бинарном формате .doc:"
            programs
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
    printDuplicates ()
    printProgramsInOldDocFormat ()
    printIncorrectlyNamedPrograms ()

let loadDisciplineOwners planCode =
    let service = openGoogleSheet "AssignmentMatcher"
    let sheet, page = owners.[planCode]
    let rawData = readGoogleSheet service sheet page "A" "B" 2
    rawData
    |> Seq.map (fun row -> (row |> Seq.head, if Seq.length row >= 2 then row |> Seq.item 1 else ""))
    |> Seq.map (fun (name, owner) -> (name.Substring(1, 6), owner.Split [|' '|] |> Array.item 0))
    |> Map.ofSeq

let autoRename (planCode, programsFolder) force =
    let curriculum = DocxCurriculum(planCodeToFileName planCode)
    let disciplines = curriculum.Disciplines
    let owners = loadDisciplineOwners planCode

    let findDiscipline code = disciplines |> Seq.tryFind (fun d -> d.Code = code)

    let programs = 
        Directory.EnumerateFiles programsFolder
        |> Seq.map FileInfo

    programs
    |> Seq.filter (fun p -> isProgramNameOk p.Name |> not || force)
    |> Seq.filter (fun p -> excluded (getCode p.Name) |> not)
    |> Seq.iter 
        (fun p ->
            let oldName = p.Name
            let code = p.Name.Substring(0, 6)
            let discipline = findDiscipline code
            if discipline.IsSome then
                let discipline = discipline.Value
                let name = discipline.RussianName.Replace(":", "")
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
                    printfn "Файл '%s' переименован в  '%s'" oldName newName
            else
                printfn "Файл '%s' не соответствует ни одной дисциплине из учебного плана" p.Name
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

let patchPrograms (planCode, programsFolder) =
    Directory.EnumerateFiles programsFolder
    |> Seq.map FileInfo
    |> Seq.filter (fun p -> excluded (getCode p.Name) |> not)
    |> Seq.iter (fun p -> patchProgram (planCodeToFileName planCode) p.FullName)

let patchProgram (planCode, programFileName) =
    ProgramPatcher.patchProgram (planCodeToFileName planCode) programFileName

let patchYear (year, programFileName) =
    ProgramPatcher.patchYear year programFileName

let addIndicatorsTable (planCode, programFileName) =
    ProgramPatcher.addIndicatorsTable (planCodeToFileName planCode) programFileName

let patchYears (year, programsFolder) =
    Directory.EnumerateFiles programsFolder
    |> Seq.map FileInfo
    |> Seq.filter (fun p -> excluded (getCode p.Name) |> not)
    |> Seq.iter (fun p -> patchYear (year, p.FullName))

let removeCompetences programFileName =
    ProgramPatcher.removeCompetences programFileName

let removeCompetencesInFolder programsFolder =
    Directory.EnumerateFiles programsFolder
    |> Seq.map FileInfo
    |> Seq.filter (fun p -> excluded (getCode p.Name) |> not)
    |> Seq.iter (fun p -> removeCompetences p.FullName)

let printSoftware programsFolder =
    let printSoftForProgram (out: StreamWriter) (program: FileInfo) =
        let content, _ = parseProgramFile (program.FullName)
        let generalSoft = "3.3.2. Характеристики аудиторного оборудования, в том числе неспециализированного компьютерного оборудования и программного обеспечения общего пользования"
        let specialSoft = "3.3.4. Характеристики специализированного программного обеспечения"
        out.WriteLine $"{program.Name}:"
        out.WriteLine $"{generalSoft}:"
        if content.ContainsKey generalSoft then
            out.WriteLine content.[generalSoft]
        out.WriteLine ""
        out.WriteLine $"{specialSoft}:"
        if content.ContainsKey specialSoft then
            out.WriteLine content.[specialSoft]
        out.WriteLine ""
        out.WriteLine ""
        ()

    use out = new StreamWriter("softwareReport.txt")

    Directory.EnumerateFiles programsFolder
    |> Seq.map FileInfo
    |> Seq.filter (fun p -> excluded (getCode p.Name) |> not)
    |> Seq.iter (printSoftForProgram out)

let printAuthors programsFolder =
    let printAuthorsForProgram (out: StreamWriter) (program: FileInfo) =
        let content, _ = parseProgramFile (program.FullName)
        let authors = "Раздел 4. Разработчики программы"
        out.WriteLine $"{program.Name}:"
        if content.ContainsKey authors then
            out.WriteLine content.[authors]
        out.WriteLine ""
        out.WriteLine ""
        ()

    use out = new StreamWriter("authorsReport.txt")

    Directory.EnumerateFiles programsFolder
    |> Seq.map FileInfo
    |> Seq.filter (fun p -> excluded (getCode p.Name) |> not)
    |> Seq.iter (printAuthorsForProgram out)

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

type Arguments =
    | [<AltCommandLine("-i")>] Inventorize of plan: string * programsFolder: string
    | [<AltCommandLine("-r")>] Auto_Rename of plan: string * programsFolder: string
    | Delete_Excluded of programsFolder: string
    | [<AltCommandLine("-c")>] Check_Content of programsFolder: string
    | [<AltCommandLine("-cp")>] Check_Program of programFileName: string
    | [<AltCommandLine("-p")>] Patch of plan: string * programsFolder: string
    | [<AltCommandLine("-pp")>] Patch_Program of plan: string * programFileName: string
    | Patch_Year of year: string * programFileName: string
    | Patch_Years of year: string * programsFolder: string
    | Remove_Competences of programFileName: string
    | Remove_Competences_All of programsFolder: string
    | Print_Soft of programsFolder: string
    | Print_Authors of programsFolder: string
    | Add_Indicators_Table of plan: string * programFileName: string
    | [<AltCommandLine("-f")>] Force
with
    interface IArgParserTemplate with
        member arg.Usage =
            match arg with
            | Inventorize _ -> "проверить, что все РПД в наличии и канонично называются."
            | Auto_Rename _ -> "автоматически переименовать РПД, если их имена не в каноничном формате. Требует авторизации в Google Drive."
            | Delete_Excluded _ -> "удалить из папки все РПД, перечисленные в списке 'исключённые' в конфиге. Используйте с осторожностью!"
            | Check_Content _ -> "проверить структуру и содержимое всех РПД в папке. Проверяет следование приказу 539/1 от 17.02.2014 о форме РПД, ищет нарушения приказа 1395/1 и типичные ошибки."
            | Check_Program _ -> "как '-c', только для одной РПД."
            | Patch _ -> "автомагически исправляет типичные ошибки во всех РПД в папке и приводит их в соответствие приказу 1395/1. После этого ТРЕБУЕТСЯ ручная проверка!"
            | Patch_Program _ -> "как '-p', только для одной РПД."
            | Patch_Year _ -> "меняет в указанной РПД год на титульнике на заданный."
            | Patch_Years _ -> "меняет во всех РПД в папке год на титульнике на заданный."
            | Remove_Competences _ -> "Удаляет компетенции из контрольно-оценочных материалов для данной РПД."
            | Remove_Competences_All _ -> "Удаляет компетенции из контрольно-оценочных материалов для всех РПД в папке."
            | Print_Soft _ -> "Печатает требуемое программное обеспечение, специализированное и общего назначения, для каждой РПД в папке."
            | Print_Authors _ -> "Печатает автора/авторов (как указано в разделе 4) для всех РПД в папке."
            | Add_Indicators_Table _ -> "Добавляет шаблон таблицы с индикаторами, с компетенциями из учебного плана."
            | Force -> "Заставляет команду работать агрессивнее (применяется только для -r, переименует файлы, даже если они соответствуют стилю именования)."

[<EntryPoint>]
let main argv =
    let errorHandler = ProcessExiter(colorizer = function ErrorCode.HelpText -> None | _ -> Some ConsoleColor.Red)
    let parser = ArgumentParser.Create<Arguments>(programName = "RpdPatcher", errorHandler = errorHandler)
    let args = parser.ParseCommandLine argv

    if (args.Contains Inventorize) then
        inventorize (args.GetResult Inventorize)

    if (args.Contains Auto_Rename) then
        autoRename (args.GetResult Auto_Rename) (args.Contains Force)

    if (args.Contains Delete_Excluded) then
        deleteExcluded (args.GetResult Delete_Excluded)

    if (args.Contains Check_Content) then
        checkContent (args.GetResult Check_Content)

    if (args.Contains Check_Program) then
        checkProgram (args.GetResult Check_Program)

    if (args.Contains Patch) then
        patchPrograms (args.GetResult Patch)

    if (args.Contains Patch_Program) then
        patchProgram (args.GetResult Patch_Program)

    if (args.Contains Patch_Year) then
        patchYear (args.GetResult Patch_Year)

    if (args.Contains Patch_Years) then
        patchYears (args.GetResult Patch_Years)

    if (args.Contains Remove_Competences) then
        removeCompetences (args.GetResult Remove_Competences)

    if (args.Contains Remove_Competences_All) then
        removeCompetencesInFolder (args.GetResult Remove_Competences_All)

    if (args.Contains Print_Soft) then
        printSoftware (args.GetResult Print_Soft)

    if (args.Contains Print_Authors) then
        printAuthors (args.GetResult Print_Authors)

    if (args.Contains Add_Indicators_Table) then
        addIndicatorsTable (args.GetResult Add_Indicators_Table)
    
    0