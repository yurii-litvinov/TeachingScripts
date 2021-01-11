module WorkDistributionData

open LessPainfulGoogleSheets
open System.Text.RegularExpressions

type WorkloadInfo =
    {
        lecturer: string
        practicioners: string list
    }

type WorkDistribution(semester) =
    let service = openGoogleSheet "AssignmentMatcher"
    let page = 
        match semester with
        | 1 | 2 -> "ТП, 1-2 семестры"
        | 3 | 4 -> "Матобес, 3-4 семестры"
        | 5 -> "Матобес, 5 семестр"
        | 6 -> "Матобес, 6 семестр"
        | 7 -> "Матобес, 7 семестр"
        | 8 -> "Матобес, 8 семестр"
        | _ -> failwith "Unknown semester"
    let rawData = readGoogleSheet service "1aLSgrsDnWAiGD7nMbAts6SxHA9PybMaNlWOPiNgsk7w" page "A" "C" 2
    
    let parseDisciplineName disciplineString =
        let regexMatch = Regex.Match(disciplineString, @"\[(\d+)\]\s+(.+)")
        if regexMatch.Success then 
            let russianName = regexMatch.Groups.[2].Value
            russianName.Trim ()
        else
            ""

    let parseWorkloadInfo (workloadString: string) = 
        let strings = workloadString.Split [|'\n'|]
        if strings.[0].StartsWith("Лекции") then
            let lecturer = strings.[0].Replace("Лекции: ", "").Trim()
            let practicioners = strings |> List.ofArray |> List.skip 2 |> List.map (fun s -> s.Trim())
            { lecturer = lecturer; practicioners = practicioners }
        else
            let practicioners = strings |> List.ofArray |> List.map (fun s -> s.Trim())
            { lecturer = ""; practicioners = practicioners }

    let parsedData = 
        rawData 
        |> Seq.map (fun row -> (row |> Seq.head, if Seq.length row > 2 then row |> Seq.skip 2 |> Seq.head else ""))
        |> Seq.map (fun (name, teachers) -> (parseDisciplineName name, parseWorkloadInfo teachers))
        |> Map.ofSeq

    member _.Teachers discipline =
        if parsedData.ContainsKey discipline then 
            parsedData.[discipline]
        else
            { lecturer = ""; practicioners = [] }
