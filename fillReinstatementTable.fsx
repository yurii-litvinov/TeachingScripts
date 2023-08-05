let yandexSheetFolderUrl =
    @"https://disk.yandex.ru/client/disk/Административная%20деятельность%20СПбГУ/Комиссия%20по%20переводам%20и%20восстановлениям/2023%2C%20лето"

let yandexSheetFileName = @"Оценки"

let dataFileName = "pretendents.xlsx"
let tabName = "Список полный"

#r "nuget: Google.Apis.Sheets.v4"
#r "nuget: DocumentFormat.OpenXml"

#load "DocUtils/DocUtils/XlsxUtils.fs"
#load "DocUtils/DocUtils/YandexSheetsUtils.fs"

open DocUtils.YandexSheets
open DocUtils.Xlsx
open System

type Statement =
    { Programme: string
      Course: string
      StatementType: string }

type Student =
    { Name: string
      Statements: Statement list }

let collectData () =
    let sheet = Spreadsheet.FromFile(dataFileName).Sheet(tabName)

    let readAndClean (columnIndex: int) =
        sheet.Column columnIndex |> Seq.skip 2 |> Seq.filter ((<>) "")

    let formStatements (statements: (string * string * string * string) seq) =
        statements
        |> Seq.map (fun (_, p, c, st) ->
            { Programme = p
              Course = c
              StatementType = st.ToLower() })
        |> Seq.toList

    let names = readAndClean 1
    let programmes = readAndClean 2
    let courses = readAndClean 4
    let statementTypes = readAndClean 5

    let merged =
        Seq.zip3 names programmes courses
        |> Seq.zip statementTypes
        |> Seq.map (fun (st, (n, p, c)) -> (n.Trim(), p.Replace('\n', ' ').Trim(), c, st))

    let combined = merged |> Seq.groupBy (fun (n, _, _, _) -> n)

    let result =
        combined
        |> Seq.map (fun (name, statements) ->
            { Name = name
              Statements = formStatements statements })

    result

let data = collectData ()

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

let programmesColumn = smartFold (fun st -> st.Programme) true
let coursesColumn = smartFold (fun st -> st.Course) false
let statementTypeColumn = smartFold (fun st -> st.StatementType) true

let service = YandexService.FromClientSecretsFile()

let yandexSpreadsheet =
    service
        .GetSpreadsheetByFolderAndFileNameAsync(yandexSheetFolderUrl, yandexSheetFileName)
        .Result

let sheet = yandexSpreadsheet.Sheet "Сортированные по алфавиту"
sheet.WriteColumn 0 2 namesColumn
sheet.WriteColumn 2 2 programmesColumn
sheet.WriteColumn 3 2 coursesColumn
sheet.WriteColumn 4 2 statementTypeColumn

yandexSpreadsheet.SaveAsync() |> Async.AwaitTask |> Async.RunSynchronously
