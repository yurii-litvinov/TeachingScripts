/// URL of a folder with table on Yandex Disk.
let yandexSheetFolderUrl =
    @"https://disk.yandex.ru/client/disk/Административная%20деятельность%20СПбГУ/Комиссия%20по%20переводам%20и%20восстановлениям/2024%2C%20осень"

/// File name (without extension) of a table in Yandex Disk.
let yandexSheetFileName = @"Оценки"

/// File name (with extension) of an input .xlsx table with applicants.
let dataFileName = "applicants.xlsx"

/// Tab name within applicants table.
let tabName = "Список полный"

/// False if we actually need to modify Yandex Sheet table, true if we only want to calculate statistics.
let dryRun = true

#r "nuget: Google.Apis.Sheets.v4"
#r "nuget: DocumentFormat.OpenXml, 2.20.0"
#r "nuget: FSharp.Json"

#load "DocUtils/DocUtils/XlsxUtils.fs"
#load "DocUtils/DocUtils/YandexSheetsUtils.fs"

open DocUtils.YandexSheets
open DocUtils.Xlsx

type Level =
    | Bachelor
    | Master
    | Specialist

type Statement =
    { Programme: string
      Course: string
      Level: Level
      StatementType: string
      Source: string }

type Student =
    { Name: string
      Statements: Statement list }

let collectData () =
    let sheet = Spreadsheet.FromFile(dataFileName).Sheet(tabName)

    let readAndClean (column: string) =
        sheet.Column column |> Seq.skip 3 |> Seq.filter ((<>) "")

    let formStatements (statements: (string * string * string * Level * string * string) seq) =
        statements
        |> Seq.map (fun (_, p, c, l, st, s) ->
            { Programme = p
              Course = c
              Level = l
              StatementType = st.ToLower()
              Source = s })
        |> Seq.toList

    let level =
        function
        | "бакалавриат" -> Bachelor
        | "магистр" | "магистратура" -> Master
        | "специалитет" -> Specialist
        | s -> failwithf "Неизвестный уровень образования '%s'" s

    let names = readAndClean "B"
    let programmes = readAndClean "C"
    let levels = readAndClean "D"
    let courses = readAndClean "E"
    let statementTypes = readAndClean "F"
    let sources = readAndClean "H"

    let merged =
        Seq.zip3 names programmes courses
        |> Seq.zip3 levels statementTypes
        |> Seq.zip sources

    let merged =
        merged
        |> Seq.map (fun (s, (l, st, (n, p, c))) -> (n.Trim(), p.Replace('\n', ' ').Trim(), c, level l, st, s))

    let combined = merged |> Seq.groupBy (fun (n, _, _, _, _, _) -> n)

    let result =
        combined
        |> Seq.map (fun (name, statements) ->
            { Name = name
              Statements = formStatements statements })

    result

let data = collectData ()

if dryRun |> not then
    let namesColumn = data |> Seq.map (fun student -> student.Name)

    let smartFold selector lineBreak =
        let smartReduce data =
            if data |> Seq.distinct |> Seq.length = 1 then
                data |> Seq.head
            else
                data
                |> Seq.reduce (fun acc el -> $"{acc},{if lineBreak then '\n' else ' '}{el}")

        data
        |> Seq.map (fun student -> student.Statements |> Seq.map selector |> smartReduce)

    let programmesColumn = smartFold _.Programme true

    let coursesColumn =
        smartFold (fun st -> st.Course + (if st.Level = Master then " (магистр)" else "")) false

    let statementTypeColumn = smartFold _.StatementType true

    let service = YandexService.FromClientSecretsFile()

    let yandexSpreadsheet =
        service
            .GetSpreadsheetByFolderAndFileNameAsync(yandexSheetFolderUrl, yandexSheetFileName)
            .Result

    let sheet = yandexSpreadsheet.Sheet "Сортированные по алфавиту"
    sheet.WriteColumn "A" 1 namesColumn
    sheet.WriteColumn "C" 1 programmesColumn
    sheet.WriteColumn "D" 1 coursesColumn
    sheet.WriteColumn "E" 1 statementTypeColumn

    try
        yandexSpreadsheet.SaveAsync() |> Async.AwaitTask |> Async.RunSynchronously
    with ServerCommunicationException(response) ->
        printf "Сервер вернул ошибку %s" response
else
    printfn "dryRun выставлен в true, таблица на Яндекс.Диске не была модифицирована!"
    printfn ""

/// Статистика поданных заявлений по образовательным программам, курсам и бюджет/договор
type ReportByProgrammes =
    { Programme: string
      Course: string
      BudgetStatements: int
      PaidStatements: int }

/// Формирование отчётов по количеству поданных заявлений по образовательным программам
let reportByProgrammes () =
    let getReport (student: Student) =
        student.Statements
        |> List.fold
            (fun reports statement ->
                { Programme = statement.Programme //.Substring(0, "02.03.03".Length)
                  Course = statement.Course
                  BudgetStatements = if statement.StatementType = "бюджет" then 1 else 0
                  PaidStatements = if statement.StatementType = "договор" then 1 else 0 }
                :: reports)
            []
        |> List.groupBy (fun r -> r.Programme + r.Course)
        |> List.map (fun (_, statementsList) ->
            match statementsList with
            | head :: [] -> head
            | first :: second :: [] ->
                { Programme = first.Programme
                  Course = first.Course
                  BudgetStatements = first.BudgetStatements + second.BudgetStatements
                  PaidStatements = first.PaidStatements + second.PaidStatements }
            | _ -> failwithf "Слишком много разных видов заявлений на один курс и программу, так вообще не должно быть (Студент %s)" student.Name)

    let addReports (reports1: ReportByProgrammes list) (reports2: ReportByProgrammes list) =
        let mergedReports =
            reports1
            |> List.map (fun report ->
                let matchingReport =
                    reports2
                    |> List.tryFind (fun r -> r.Programme = report.Programme && r.Course = report.Course)

                if matchingReport.IsSome then
                    { Programme = report.Programme
                      Course = report.Course
                      BudgetStatements = report.BudgetStatements + matchingReport.Value.BudgetStatements
                      PaidStatements = report.PaidStatements + matchingReport.Value.PaidStatements }
                else
                    report)

        let newReports =
            reports2
            |> List.filter (fun report ->
                reports1
                |> List.tryFind (fun r -> r.Programme = report.Programme && r.Course = report.Course)
                |> Option.isNone)

        newReports @ mergedReports

    let report =
        data
        |> Seq.fold (fun programmes student -> getReport student |> addReports programmes) []

    report

let reports = reportByProgrammes ()

/// печать отчётов
let printReport (programmeName: string, programmeReport: ReportByProgrammes list) =
    printfn "%s: " programmeName

    let programmeReport = programmeReport |> List.sortBy _.Course

    programmeReport
    |> List.iter (fun r ->
        printfn "Курс: %s — бюджетных: %d, договорных: %d" r.Course r.BudgetStatements r.PaidStatements)

    printfn
        "Всего, бюджет: %d, договор: %d"
        (programmeReport |> List.sumBy _.BudgetStatements)
        (programmeReport |> List.sumBy _.PaidStatements)
    
    printfn ""

reports
|> List.sortBy _.Programme
|> List.groupBy _.Programme
|> List.iter printReport

printfn ""

printfn
    "Всего, бюджет: %d, договор: %d"
    (reports |> List.sumBy _.BudgetStatements)
    (reports |> List.sumBy _.PaidStatements)

/// Отчётность по тому, откуда заявление — перевод, восстановление, другой вуз и т.п.
let totalSources =
    data
    |> Seq.map _.Statements
    |> Seq.concat
    |> Seq.countBy _.Source

printfn "Всего кто откуда: "

totalSources |> Seq.sortBy fst |> Seq.iter (fun (key, value) -> printf "%s: %d, " key value)

data
|> Seq.map _.Statements
|> Seq.concat
|> Seq.groupBy _.Programme
|> Seq.sortBy fst
|> Seq.map (fun (p, s) -> (p, s |> Seq.countBy (fun statement -> statement.Source)))
|> Seq.iter (fun (p, statistics) ->
    printf "\n%s: " p
    statistics |> Seq.sortBy fst |> Seq.iter (fun (key, value) -> printf "%s: %d, " key value))

let uniqueStudentsFromOtherUniversities =
    data |> Seq.filter (fun s -> s.Statements.Head.Source = "др ВУЗ") |> Seq.length

printfn ""
printfn ""
printfn "Всего студентов: %d" (data |> Seq.length)
printfn "Всего студентов из других вузов: %d" uniqueStudentsFromOtherUniversities
