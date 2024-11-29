let holidays = ["23.02.2024"; "08.03.2024"; "01.05.2024"; "09.05.2024"; "12.06.2024"; "04.11.2024"]
let startDate = "13.02.2024"
let finishDate = "30.05.2024"
let spreadsheetPath = "https://disk.yandex.ru/client/disk/Курсы/Проектирование%20и%20архитектура%20ПО"
let spreadsheetName = "Проектирование и архитектура ПО, программа"
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
    Seq.initInfinite (fun i -> convertedStartDate + (TimeSpan.FromDays <| float (i * 7)))
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