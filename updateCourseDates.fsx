let holidays = ["23.02.2023"; "24.02.2023"; "08.03.2023"; "01.05.2023"; "08.05.2023"; "09.05.2023"; "12.06.2023"; "04.11.2023"]
let startDate = "14.02.2023"
let spreadsheedPath = "Курсы/ТРПО/ТРПО, программа курса-копия.xlsx"
let datePos = (1, 1)
let finishDate = "30.05.2023"

#r "nuget:DocumentFormat.OpenXml"
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
    run <| yandexService.GetSpreadsheetAsync "Курсы/ТРПО/ТРПО, программа курса-копия.xlsx"

let sheet = spreadsheet.Sheet "Лист1"

sheet.WriteColumn (snd datePos) (fst datePos) courseDates

run <| spreadsheet.Save ()