let gradingListId = "1IkoCocAKy5hb6KfMHEOlIZG5RAdZwFnm77RaoUhDMAc"

let secondCourseSpreadsheetId = "17TxV9OgKSN37RaFiGRgRe_T45E20YtODrnn5U3q3P9k"
let thirdCourseSpreadsheetId ="1zUMT5TkjDuiwjABzu-83n972tlx0Jg1b_wwZcYFzLQs"
let fourthCourseSpreadsheetId = "12ypnsGlkpD1hJa_n2J9QWOxwbIClXf67bALYH0mr8d0"

let registrationListId = "13eC38aia4WwysOUCYcDKeH2Sfi2hwIBARNjmeoN8x_Q"
let registrationRangeStart = 63
let registrationRangeEnd = 72

#r "nuget: Google.Apis.Sheets.v4"

#load "DocUtils/DocUtils/GoogleSheetsUtils.fs"

open DocUtils

type Student =
    { 
        Surname: string
        Course: string 
        FullName: string
        Theme: string
    }

let collectStudents (service: GoogleSheetService) = 
    let sheet = service.Sheet(registrationListId, "Sheet1")
    let rawData = sheet.ReadColumn ("B", registrationRangeStart, registrationRangeEnd - registrationRangeStart + 1)
    rawData
        |> Seq.filter ((<>) "")
        |> Seq.map (fun s -> 
            let splittedRow = s.Split(' ')
            { 
                Surname = splittedRow[0]
                Course = splittedRow[2]
                FullName = ""
                Theme = ""
            }
        )

let updateStudentFullNamesAndThemes (service: GoogleSheetService) (data: seq<Student>) =
    let getInfo id = service.Sheet(id, "СП").ReadByHeaders(["ФИО"; "Тема"]) |> Seq.map (fun map -> (map["ФИО"], map["Тема"]))
    let secondCourseInfo = getInfo secondCourseSpreadsheetId
    let thirdCourseInfo = getInfo thirdCourseSpreadsheetId
    let fourthCourseInfo = getInfo fourthCourseSpreadsheetId
    let infos = Seq.concat [secondCourseInfo; thirdCourseInfo; fourthCourseInfo]

    let findInfo surname = 
        let fullNameCandidates = infos |> Seq.filter (fun (n, _) -> n.Split(' ')[0] = surname) |> Seq.toList
        match fullNameCandidates with 
        | [] -> printfn "No info for %s" surname; ("", "")
        | [n] -> n
        | n :: _ :: _ -> printfn "Multiple full names for %s" surname; n

    data |> Seq.map (fun student -> 
        let fullName, theme = findInfo student.Surname
        { student with FullName = fullName; Theme = theme})

let generateGradingList (service: GoogleSheetService) (data: seq<Student>) =
    let sheetIds = service.Spreadsheet(gradingListId).Sheets() |> Seq.filter ((<>) "Итого")
    let mainSheet = service.Spreadsheet(gradingListId).Sheet("Итого")

    let write column mapper =
        let info = Seq.concat [data |> Seq.map mapper; Seq.replicate 10 ""]
        mainSheet.WriteColumn (column, info, 7)

    write "A" (fun s -> s.FullName)
    write "B" (fun s -> s.Theme)
    write "C" (fun s -> s.Course)

    for id in sheetIds do
        let write column mapper =
            let sheet = service.Spreadsheet(gradingListId).Sheet(id)
            let info = Seq.concat [data |> Seq.map mapper; Seq.replicate 10 ""]
            sheet.WriteColumn (column, info, 4)
        write "A" (fun s -> s.FullName)

let main() =
    use service = new GoogleSheetService("../credentials.json", "TeachingScripts")
    let data = collectStudents service
    let data = updateStudentFullNamesAndThemes service data
    generateGradingList service data
    ()

main()


