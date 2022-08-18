let spreadsheetId = "1sRclQ1HRUN8_hnky0XOeVMehipZJue6CRvY6vW646gk"
let dataFileName = "reinstatementPretendents.xlsx"
let tabName = "Список полный"


#r "nuget: Google.Apis.Sheets.v4"
#r "nuget: DocumentFormat.OpenXml"

#load "ReviewGenerator/LessPainfulGoogleSheets.fs"
#load "AdcEditor/LessPainfulXlsx.fs"

open LessPainfulGoogleSheets
open LessPainfulXlsx

type Statement =
    { Programme: string
      Course: string 
      StatementType: string
      }

type Student =
    { Name: string
      Statements: Statement list
      }

let collectData () = 
    let sheet = openXlsxSheet dataFileName tabName
    
    let readAndClean columnIndex =
        readColumn sheet columnIndex |> Seq.skip 2 |> Seq.filter ((<>) "")

    let formStatements (statements: (string * string * string * string) seq) =
        statements 
        |> Seq.map (fun (_, p, c, st) -> { Programme = p; Course = c; StatementType = st.ToLower () }) 
        |> Seq.toList

    let names = readAndClean 1
    let programmes = readAndClean 2
    let courses = readAndClean 4
    let statementTypes = readAndClean 5

    let merged = 
        Seq.zip3 names programmes courses 
        |> Seq.zip statementTypes
        |> Seq.map (fun (st, (n, p, c)) -> (n, p, c, st))

    let combined = merged |> Seq.groupBy (fun (n, _, _, _) -> n)
    let result = combined |> Seq.map (fun (name, statements) -> { Name = name; Statements = formStatements statements})
    result

let data = collectData ()

let service = openGoogleSheet "TeachingScripts" 

let namesColumn = data |> Seq.map (fun student -> student.Name)

let smartFold selector =
    let smartReduce data =
        if data |> Seq.distinct |> Seq.length = 1 then
            data |> Seq.head
        else
            data |> Seq.reduce (fun acc el -> acc + ", " + el)

    data |> Seq.map (fun student -> student.Statements |> Seq.map selector |> smartReduce)

let programmesColumn = smartFold (fun st -> st.Programme)
let coursesColumn = smartFold (fun st -> st.Course)
let statementTypeColumn = smartFold (fun st -> st.StatementType)

writeGoogleSheetColumn service spreadsheetId "Сортированные по алфавиту" "A" 2 namesColumn
writeGoogleSheetColumn service spreadsheetId "Сортированные по алфавиту" "C" 2 programmesColumn
writeGoogleSheetColumn service spreadsheetId "Сортированные по алфавиту" "D" 2 coursesColumn
writeGoogleSheetColumn service spreadsheetId "Сортированные по алфавиту" "E" 2 statementTypeColumn

service.Dispose()