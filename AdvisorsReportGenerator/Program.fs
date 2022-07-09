open DocUtils

/// Configuration parameter --- id of a Google Spreadsheet where to put generated report. Shall be initially empty.
let targetTable = "1kvx5dVd8PTpxDDnhGnRizsQzDtZt6OD2XaoYJAV__2k"

/// Configuration parameter --- names of chairs of interest.
let chairsOfInterest = Set.ofList ["ИАС"; "Информатики"; "СП"; "ПА" ]

/// Record describing from where to get info for a given chair.
type ChairMetadata =
    { Chair: string
      AdvisorColumn: string
      ResultColumn: string }

/// Record describing data sourec for practices for a given semester.
type Metadata = 
    { Semester: int
      Spreadsheet: string
      Chairs: ChairMetadata list }
    with 
        static member New semester spreadsheet =
            { Semester = semester; Spreadsheet = spreadsheet; Chairs = [] }

/// "Builder" pattern implementation for easier configuring.
let withChair chair advisorRow resultRow metadata =
    { metadata with Chairs = { Chair = chair; AdvisorColumn = advisorRow; ResultColumn = resultRow } :: metadata.Chairs }

/// Configuration parameter --- a table that describes data sources. Contains semesters, corresponding spreadsheet ids and
/// chair infos --- chair name, column with advisor name for a given chair, and column with practice result.
let metadata = 
    [
        Metadata.New 3 "1YXLpkUSWunT5m6_XANpHKAU3IU0NoBZE_aFQwwX_TjA"
            |> withChair "ИАС" "C" "J"
            |> withChair "Информатики" "C" "K"
            |> withChair "СП" "C" "P"
            |> withChair "ПА" "C" "J"
        Metadata.New 4 "1OyDYJ8N_fvdphXYdY2h_O1Cr1hB6noPcfv4XXFUkEYc"
            |> withChair "ИАС" "C" "K"
            |> withChair "Информатики" "C" "K"
            |> withChair "СП" "C" "P"
            |> withChair "ПА" "C" "K"
        Metadata.New 5 "1MCVf88nLnYuRdPKURYbX8dNcMLweTWUUrqOCmfcJmvI"
            |> withChair "ИАС" "D" "K"
            |> withChair "Информатики" "D" "K"
            |> withChair "СП" "D" "P"
            |> withChair "ПА" "D" "K"
        Metadata.New 6 "1MCVf88nLnYuRdPKURYbX8dNcMLweTWUUrqOCmfcJmvI"
            |> withChair "ИАС" "D" "O"
            |> withChair "Информатики" "D" "O"
            |> withChair "СП" "D" "Y"
            |> withChair "ПА" "D" "O"
        Metadata.New 7 "103E-S9SxzRAPiTt_Yr6428243hzIKzjy1WaLR_W7HDM"
            |> withChair "ИАС" "D" "N"
            |> withChair "Информатики" "D" "N"
            |> withChair "СП" "E" "P"
            |> withChair "ПА" "D" "N"
        Metadata.New 8 "103E-S9SxzRAPiTt_Yr6428243hzIKzjy1WaLR_W7HDM"
            |> withChair "ИАС" "D" "V"
            |> withChair "Информатики" "D" "V"
            |> withChair "СП" "E" "T"
            |> withChair "ПА" "D" "V"
        // Virtual semester for qualification work defence.
        Metadata.New 9 "103E-S9SxzRAPiTt_Yr6428243hzIKzjy1WaLR_W7HDM"
            |> withChair "ИАС" "D" "AC"
            |> withChair "Информатики" "D" "AC"
            |> withChair "СП" "E" "AA"
            |> withChair "ПА" "D" "AC" 
    ]

/// Collected data.
let mutable data: Map<string, Map<int, int>> = Map.empty

/// Parses a single sheet containing info about practices of a single semester of a single chair. Adds to collected data.
let getDataFromSheet (service: GoogleSheetService) sheetMetadata =
    let successfulResults = Set.ofList ["A"; "B"; "C"; "D"; "E"; "да"; "зачёт"]
    let semester, spreadsheetId, chairInfo = sheetMetadata
    let sheet = service.Sheet(spreadsheetId, chairInfo.Chair)
    let advisors = sheet.ReadColumn (chairInfo.AdvisorColumn, 1)
    let results = sheet.ReadColumn (chairInfo.ResultColumn, 1)
    let merged = Seq.zip advisors results |> Seq.map (fun (a, r) -> (a, semester, r))
    let filtered = merged |> Seq.filter (fun (_, _, r) -> successfulResults.Contains r)
    filtered
    |> Seq.iter (
        fun (a, s, _) -> 
            if data.ContainsKey a |> not then
                data <- data.Add (a, Map.empty)
            let semesters = if data[a].ContainsKey s then data[a] else data[a].Add (s, 0)
            data <- data.Add (a, semesters.Add (s, semesters[s] + 1))
       )
    ()

/// Parses a single semester.
let getDataFromSemester service metadata =
    metadata.Chairs
    |> List.filter (fun c -> chairsOfInterest.Contains c.Chair)
    |> List.map (fun c -> (metadata.Semester, metadata.Spreadsheet, c))
    |> List.iter (getDataFromSheet service)

/// Reads data from data sources described in metadata.
let getData service =
    metadata |> List.iter (getDataFromSemester service)

/// Prints report into targetTable using collected data.
let generateReport (service: GoogleSheetService) =
    let columnHeaders = 
        Map.empty
            .Add(3, "Учебных практик, 2 курс, осень")
            .Add(4, "Учебных практик, 2 курс, весна")
            .Add(5, "Учебных практик, 3 курс, осень")
            .Add(6, "Учебных практик, 3 курс, весна")
            .Add(7, "Производственных практик")
            .Add(8, "Преддипломных практик")
            .Add(9, "ВКР")

    let sheet = service.Sheet (targetTable, "Sheet1")
    let advisorColumn = "ФИО" :: (List.ofSeq data.Keys)

    let worksBySemester semester = 
        data.Keys |> Seq.map (fun a -> if data[a].ContainsKey semester then data[a][semester] else 0)

    let semesterColumns =
        [3..9]
        |> List.map (fun i -> columnHeaders[i] :: (List.ofSeq <| (worksBySemester i |> Seq.map string)))
        |> List.map (Seq.map (fun cell -> if cell = "0" then "" else cell))

    let semesterToColumn semester =
        (int 'A') + semester - 2 |> char |> string

    sheet.WriteColumn ("A", advisorColumn)
    semesterColumns
    |> List.iteri (fun i s -> sheet.WriteColumn (semesterToColumn (i + 3), s))

/// Entry point of the application.
[<EntryPoint>]
let main _ =
    let service = new GoogleSheetService("credentials.json", "AdvisorsReportGenerator")
    getData service
    generateReport service
    
    0
