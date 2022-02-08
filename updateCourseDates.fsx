let holidays = ["23.02.2022"; "08.03.2022"; "01.05.2022"; "02.05.2022"; "09.05.2022"; "10.05.2022"; "12.06.2022"; "13.06.2022"]
let startDate = "17.02.2022"
let spreadSheetId = "1NMOh0yxsQZqzz0E3TuG_uYyS5VxgfqbBvM2ko1wSEjU"
let datePos = ("B", 2)
let finishDate = "30.05.2022"

#r "nuget: Google.Apis.Sheets.v4"
#load "ReviewGenerator/LessPainfulGoogleSheets.fs"

open System
open LessPainfulGoogleSheets

let convertedStartDate = DateTime.Parse startDate

let courseDates = 
    Seq.initInfinite (fun i -> convertedStartDate + (TimeSpan.FromDays <| float (i * 7)))
    |> Seq.map (fun date -> date.ToShortDateString())
    |> Seq.except holidays
    |> Seq.takeWhile (fun date -> (DateTime.Parse date) <= (DateTime.Parse finishDate))

let service = openGoogleSheet "AssignmentMatcher" 

writeGoogleSheetColumn service spreadSheetId "Sheet1" (fst datePos) (snd datePos) courseDates

service.Dispose()