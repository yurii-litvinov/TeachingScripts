module LessPainfulGoogleSheets

open System
open Google.Apis.Sheets.v4
open System.IO
open Google.Apis.Auth.OAuth2
open Google.Apis.Services
open System.Threading
open Google.Apis.Util.Store
open Google.Apis.Sheets.v4.Data

let openGoogleSheet applicationName =
    let scopes = [ SheetsService.Scope.Spreadsheets ]

    use credentialsStream = new FileStream("../../../../../credentials.json", FileMode.Open, FileAccess.Read)
    let tokenPath = "token.json"

    let credential = 
        GoogleWebAuthorizationBroker.AuthorizeAsync(
            GoogleClientSecrets.Load(credentialsStream).Secrets,
            scopes,
            "user",
            CancellationToken.None,
            new FileDataStore(tokenPath, true)).Result

    let service = 
        new SheetsService(
            BaseClientService.Initializer(
                HttpClientInitializer = credential,
                ApplicationName = applicationName
            )
        )

    service

let writeGoogleSheetColumn (service: SheetsService) spreadsheetId sheet column offset (data: #seq<string>) =
    let range = sheet + "!" + column + (string offset) + ":" + column + (string ((data |> Seq.length) + offset))
    let valueRange = ValueRange(Values = [| data |> Seq.cast<obj> |> Seq.toArray |])
    valueRange.MajorDimension <- "COLUMNS"
    let request = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range)
    request.ValueInputOption <- Nullable(SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW)

    request.Execute() |> ignore

/// Returns sequence of rows of specified google sheet.
let readGoogleSheet (service: SheetsService) spreadsheetId sheet startColumn endColumn (offset: int) =
    let range = sheet + "!" + startColumn + (string offset) + ":" + endColumn + "1000"

    let request = service.Spreadsheets.Values.Get(spreadsheetId, range)

    let values = request.Execute().Values 
    if values <> null then
        values |> Seq.map (Seq.map string)
    else
        Seq.empty

let readGoogleSheetColumn (service: SheetsService) spreadsheetId sheet column offset =
    let column = readGoogleSheet service spreadsheetId sheet column column offset
    column |> Seq.concat
