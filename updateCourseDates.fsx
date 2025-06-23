let holidays = ["08.03.2025"; "01.05.2025"; "08.05.2025"; "09.05.2025"; "12.06.2025"; "04.11.2025"; "31.12.2025"]
let startDate = "17.02.2025"
let finishDate = "30.05.2025"
let spreadsheetPath = "https://disk.yandex.ru/client/disk/Курсы/Проектирование%20ПО%2C%20ВШЭ%20и%20ИТМО"
let spreadsheetName = "Проектирование ПО, ВШЭ, программа"
let datePos = {| row = 2; column = "B" |}
let sheetName = "Sheet1"

#r "nuget:DocumentFormat.OpenXml, 2.20.0"
#r "nuget:FSharp.Json"
#load "DocUtils/DocUtils/XlsxUtils.fs"
#load "DocUtils/DocUtils/YandexSheetsUtils.fs"

open System
open DocUtils.YandexSheets

let convertedStartDate = DateTime.Parse startDate

let courseDates = 
    Seq.initInfinite (fun i -> convertedStartDate + (float (i * 7) |> TimeSpan.FromDays))
    |> Seq.map (fun date -> date.ToShortDateString())
    |> Seq.except holidays
    |> Seq.takeWhile (fun date -> (DateTime.Parse date) <= (DateTime.Parse finishDate))

let yandexService = YandexService.FromClientSecretsFile()

let run task =
    task |> Async.AwaitTask |> Async.RunSynchronously

let spreadsheet =
    run <| yandexService.GetSpreadsheetByFolderAndFileNameAsync(spreadsheetPath, spreadsheetName)

let sheet = spreadsheet.Sheet sheetName

sheet.WriteColumn datePos.column datePos.row courseDates

run <| spreadsheet.SaveAsync ()