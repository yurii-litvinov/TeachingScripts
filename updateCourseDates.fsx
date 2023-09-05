let holidays = ["23.02.2023"; "24.02.2023"; "08.03.2023"; "01.05.2023"; "08.05.2023"; "09.05.2023"; "12.06.2023"; "04.11.2023"]
let startDate = "17.02.2023"
let spreadsheedUrl = @"https://disk.yandex.ru/edit/disk/disk%2F%D0%9A%D1%83%D1%80%D1%81%D1%8B%2F%D0%9E%D0%9E%D0%9F%20%D0%BD%D0%B0%20C%23%2F%D0%9F%D1%80%D0%BE%D0%B3%D1%80%D0%B0%D0%BC%D0%BC%D0%B0%20%D0%BA%D1%83%D1%80%D1%81%D0%B0%20%28%D0%9E%D0%9E%D0%9F%20%D0%BD%D0%B0%20C%23%29.xlsx?sk=yc8fe08c1e06bffc98b4a58fe72195b8b"
let sheetName = "Sheet1"
let datePos = {| row = 19; column = "B" |}
let finishDate = "30.05.2023"

#r "nuget:DocumentFormat.OpenXml"
#r "nuget:FSharp.Json"
#load "DocUtils/DocUtils/XlsxUtils.fs"
#load "DocUtils/DocUtils/YandexSheetsUtils.fs"

open System
open DocUtils.YandexSheets

let convertedStartDate = DateTime.Parse startDate
let decodedUrl = Uri.UnescapeDataString spreadsheedUrl
let spreadsheedPathWithSuffix = decodedUrl.Remove(0, "https://disk.yandex.ru/edit/disk/disk/".Length)
let spreadsheedPath = spreadsheedPathWithSuffix.Remove(spreadsheedPathWithSuffix.LastIndexOf("?sk="))

let courseDates = 
    Seq.initInfinite (fun i -> convertedStartDate + (TimeSpan.FromDays <| float (i * 7)))
    |> Seq.map (fun date -> date.ToShortDateString())
    |> Seq.except holidays
    |> Seq.takeWhile (fun date -> (DateTime.Parse date) <= (DateTime.Parse finishDate))

let yandexService = YandexService.FromClientSecretsFile()

let run task =
    task |> Async.AwaitTask |> Async.RunSynchronously

let spreadsheet =
    run <| yandexService.GetSpreadsheetAsync spreadsheedPath

let sheet = spreadsheet.Sheet sheetName

sheet.WriteColumn (int datePos.column[0] - (int 'A')) (datePos.row - 1) courseDates

run <| spreadsheet.SaveAsync ()