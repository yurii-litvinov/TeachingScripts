open CurriculumParser
open System.IO
open System.Text.RegularExpressions

open LessPainfulGoogleSheets
open ProgramContentChecker

let plansFolder = "../.."
let programsFolder = "../.."

let plans = 
    [ "5162-2020", "СВ.5162-2020.docx"
      "5006-2019", "СВ.5006-2019.docx"
      "5006-2018", "СВ.5006-2018.docx"
      "5006-2017", "СВ.5006-2017.docx"
      "5006-2020", "ВМ.5665-2020.docx"
      "5006-2019", "ВМ.5665-2019.docx" ]
    |> List.map (fun x -> (fst x, $"{plansFolder}/{snd x}"))
    |> Map.ofList

let owners =
    [ "5162-2020", ("1aP3biZf9vVMLG26Ra3acreMW2iZLCoXz93WPqK2ZV-U", "2020")
    ]
    |> Map.ofList

let exclusions =
    ["057495"; "057596"; "058037"; "058039"; "058041"; "059998"; "060000"; "060008"; "060009"; "060010"; "060059"; 
        "062762"; "700000"; "800000"; "900000"]  // Общеуниверситетские дисциплины
    @ ["064764"; "064792"; "064793"; "064795"; "064797"]  // РПП, у них совсем другой формат

let isProgramNameOk fileName =
    let matches = Regex.Match(fileName, @"^\d{6}_(\w|\s|\+|-|\.|,)+_\d{2}_\d{4}_\d(-\d)?с_\w+.docx$")
    matches.Success

let excluded code =
    exclusions |> List.contains code

let getCode (disciplineFileName: string) =
    disciplineFileName.Substring(0, 6)

let inventorize planName =
    let curriculum = DocxCurriculum(plans.[planName])
    let disciplines = curriculum.Disciplines
    let programs = 
        Directory.EnumerateFiles $"{programsFolder}/{planName}"
        |> Seq.map (fun p -> FileInfo(p).Name)

    let printMissingDisciplines () =
        let missingDisciplines = 
            disciplines
            |> Seq.filter (fun d -> excluded d.Code |> not)
            |> Seq.filter (fun d -> programs |> Seq.tryFind (fun s -> s.Contains(d.Code)) |> Option.isNone)

        if missingDisciplines |> Seq.isEmpty then
            printfn "Everything present!"
        else
            printfn "Missing working programs:"
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
            printfn "There are working programs not from a working plan:"
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
            printfn "Incorrectly named working programs:"
            incorrectlyNamedPrograms
            |> Seq.sort 
            |> Seq.iter (printfn "%s")

    printMissingDisciplines ()
    printUnneededPrograms ()
    printIncorrectlyNamedPrograms ()

let loadDisciplineOwners planName =
    let service = openGoogleSheet "AssignmentMatcher"
    let sheet, page = owners.[planName]
    let rawData = readGoogleSheet service sheet page "A" "B" 2
    rawData
    |> Seq.map (fun row -> (row |> Seq.head, if Seq.length row >= 2 then row |> Seq.item 1 else ""))
    |> Seq.map (fun (name, owner) -> (name.Substring(1, 6), owner.Split [|' '|] |> Array.item 0))
    |> Map.ofSeq

let autoRename planName =
    let curriculum = DocxCurriculum(plans.[planName])
    let disciplines = curriculum.Disciplines
    let owners = loadDisciplineOwners planName

    let findDiscipline code = disciplines |> Seq.find (fun d -> d.Code = code)

    let programs = 
        Directory.EnumerateFiles $"{programsFolder}/{planName}"
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
                printfn "Not going to rename %s, it is in wrong format!" p.Name
            elif owner = "" then
                printfn "Not going to rename %s, unknown owner!" p.Name
            else
                let newName = $"{code}_{name}_{year}_{curriculumCode}_{semesterPart}с_{owner}.docx"
                printfn "Going to rename %s to %s" p.Name newName
                p.MoveTo($"{p.DirectoryName}/{newName}")
        )

let removeExcluded planName =
    Directory.EnumerateFiles $"{programsFolder}/{planName}"
    |> Seq.map FileInfo
    |> Seq.filter (fun p -> excluded (getCode p.Name))
    |> Seq.iter (fun p -> 
        printfn "Removing %s" p.Name
        p.Delete ())

let checkContent planName =
    Directory.EnumerateFiles $"{programsFolder}/{planName}"
    |> Seq.iter checkProgram

let printHelp () =
    printfn "Welcome to RPD Patcher, an utility to fix working programs of SPbSU!"
    printfn "Usage:"
    printfn "\tdotnet run -- -i <working plan name>"
    printfn "\t\t- check that all working programs are present and named correctly."
    printfn "\tdotnet run -- -r <working plan name>"
    printfn "\t\t- try to automatically rename working programs if they are in incorrect format."
    printfn ""
    printfn "Supported working plans:"
    plans |> Map.iter (fun key _ -> printf "%s " key)

[<EntryPoint>]
let main argv =
    if argv.Length < 2 then
        printHelp ()
    else
        match argv.[0] with
        | "--inventorize" | "-i" -> inventorize argv.[1]
        | "--auto-rename" | "-r" -> autoRename argv.[1]
        | "--delete-excluded" -> removeExcluded argv.[1]
        | "--check-content" | "-c" -> checkContent argv.[1]
        | _ -> printHelp ()
    0
