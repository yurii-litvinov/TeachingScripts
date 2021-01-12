module VkrAdvisorsParser

open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing

type VkrRecord = 
    {
        studentName: string
        theme: string
        englishTheme: string
        advisor: string
        advisorJobTitle: string
        chair: string
    }

let parseRow (row: TableRow) =
    let value rowNum =
        row.Elements<TableCell>() |> Seq.item rowNum |> fun x -> x.InnerText

    let themeCell = row.Elements<TableCell>() |> Seq.item 2
    let runs = themeCell.Descendants<Run>()

    if runs |> Seq.length < 3 then
        { 
            studentName = value 1
            theme = ""
            englishTheme = ""
            advisor = ""
            advisorJobTitle = ""
            chair = ""
        }
    else
        let mutable theme = ""
        let mutable englishTheme = ""
        let mutable seenBreak = false
        for run in runs do
            for element in run.ChildElements do
                match element with
                | :? Text when not seenBreak -> theme <- theme + element.InnerText
                | :? Text when seenBreak -> englishTheme <- englishTheme + element.InnerText
                | :? Break -> seenBreak <- true
                | _ -> ()

        { 
            studentName = value 1
            theme = theme
            englishTheme = englishTheme
            advisor = value 3
            advisorJobTitle = value 4
            chair = value 5
        }

let parseVkrAdvisors (fileName: string) :VkrRecord seq =
    use wordDocument = WordprocessingDocument.Open(fileName, false)
    
    let body = wordDocument.MainDocumentPart.Document.Body
    let mainTable = body.Elements<Table>() |> Seq.skip 1 |> Seq.head

    mainTable.Elements<TableRow>() 
    |> Seq.skip 2
    |> Seq.map parseRow
