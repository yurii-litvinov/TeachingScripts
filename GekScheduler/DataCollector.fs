module DataCollector

open LessPainfulGoogleSheets
open DataModel
open Utils

let service = openGoogleSheet "AssignmentMatcher"

let parseDates (str: string) =
    if str <> "" then
        let splitDates = str.Split(", ")
        splitDates
        |> Seq.map (fun d -> d.Split(" – ").[0])
        |> Seq.map Date
        |> Seq.toList
    else 
        []

let getSlots () =
    let rawData = readGoogleSheet service Config.gekSchedule "Sheet1" "B" "C" 2
    let slots =
        rawData
        |> Seq.filter (fun s -> s |> Seq.isEmpty |> not)
        |> Seq.filter (fun s -> s |> Seq.head |> (fun str -> System.Char.IsDigit(str.ToCharArray().[0])))
        |> Seq.pairwise
        |> Seq.filter (fun (f, _) -> f |> Seq.length = 1)
        |> Seq.map (fun (f, s) -> (Seq.head f, Seq.head s, Seq.item 1 s))
        |> Seq.filter (fun (_, _, p) -> p.ToString().StartsWith("СП"))
        |> Seq.map (fun (d, t, p) -> 
            let date = Date d
            let time = 
                match t with
                | "10:00" -> Morning
                | "14:00" -> Day
                | _ -> failwith "Unknown time slot"

            let program = 
                let programPart = p.Split(", ").[1]
                match programPart with
                | "бакалавры ПИ" -> PiBachelor
                | "бакалавры МОиАИС" -> MatobesBachelor
                | "магистры ПИ" -> PiMaster
                | "магистры МОиАИС" -> MatobesMaster
                | _ -> failwith "Unknown program"

            {date = date; time = time; program = program}
            )
    slots |> Seq.toList

let getWorks () =
    let worksData = readGoogleSheet service Config.worksData "Sheet1" "A" "E" 2
    let worksData = 
        worksData
        |> Seq.map Seq.toList
        |> Seq.map (fun l ->
            match l with
            | [name; program; theme; advisor; reviewer] ->
                let parseProgram = function
                    | "бакалавр ПИ" -> PiBachelor
                    | "бакалавр матобеса" -> MatobesBachelor
                    | "магистр ПИ" -> PiMaster
                    | "магистр матобеса" -> MatobesMaster
                    | _ -> failwith "Invalid format"
                {
                    theme = theme
                    author = name
                    program = parseProgram program
                    advisor = advisor
                    reviewer = reviewer
                    convenientDates = []
                    inconvenientDates = []
                }
            | _ -> failwith "Invalid format"
            )

    let studentPreferences = readGoogleSheet service Config.studentPreferences "Form Responses 1" "B" "D" 2

    let studentPreferences =
        studentPreferences
        |> Seq.map (fun s -> if s |> Seq.length = 2 then Seq.append s (seq {""}) else s)
        |> Seq.map Seq.toList
        |> Seq.map (fun l ->
            match l with
            | [name; preferred; preferredNot] ->
                (getParticipantId name, (parseDates preferred, parseDates preferredNot))
            | _ -> failwith "Invalid format"
            )
        |> Map.ofSeq

    worksData
    |> Seq.map (fun r -> 
        let splitAuthor = r.author.Split(" ")
        let studentId = $"{splitAuthor.[0]} {splitAuthor.[1].ToCharArray().[0]}"
        if studentPreferences.ContainsKey studentId then
            { r with
                convenientDates = fst studentPreferences.[studentId]
                inconvenientDates = snd studentPreferences.[studentId] }
        else
            r
    )
    |> Seq.toList

let getAdvisorPreferences () =
    let advisorPreferences = readGoogleSheet service Config.advisorPreferences "Form Responses 1" "B" "D" 2

    advisorPreferences
    |> Seq.map (fun s -> if s |> Seq.length = 2 then Seq.append s (seq {""}) else s)
    |> Seq.map Seq.toList
    |> Seq.map (fun l ->
        match l with
        | [name; preferred; preferredNot] ->
            (getParticipantId name, (parseDates preferred, parseDates preferredNot))
        | _ -> failwith "Invalid format"
        )
    |> Map.ofSeq

let getLinkedWorks () =
    let worksData = readGoogleSheet service Config.worksData "Связанные работы" "A" "A" 1
    worksData |> Seq.concat |> Seq.map (fun s -> s.Split(", ") |> Seq.map getParticipantId |> Seq.toList)

let getCommissionMembers () =
    let rawData = readGoogleSheet service Config.gekMembers "Form Responses 1" "B" "N" 2

    let parseInterestingThemes (str: string) =
        str.Split(", ") |> List.ofArray

    let parseInvolvementPreference = function
    | "Чем больше, тем лучше" -> Max
    | "Поменьше" -> Min
    | _ -> Min

    let parseChair = function
    | "да" -> true
    | _ -> false

    let parsePossibleDays = function
    | "" -> 1
    | x -> int x

    let gekMembers =
        rawData
        |> Seq.map (fun s -> Seq.append s (Seq.init (13 - Seq.length s) (fun _ -> "")))
        |> Seq.map Seq.toList
        |> Seq.map (fun l ->
            match l with
            | [ name
                preferred
                preferredNot
                possibleDays
                involvementPreference
                _
                interestingThemes1
                interestingThemes2
                interestingThemes3
                interestingThemes4
                _
                isChair 
                company ] ->
                { name = name
                  company = company
                  convenientDates = parseDates preferred
                  inconvenientDates = parseDates preferredNot
                  maxDays = parsePossibleDays possibleDays
                  involvementPreference = parseInvolvementPreference involvementPreference
                  interestingThemes = (parseInterestingThemes interestingThemes1) 
                      @ (parseInterestingThemes interestingThemes2)
                      @ (parseInterestingThemes interestingThemes3)
                      @ (parseInterestingThemes interestingThemes4) 
                  isChair = parseChair isChair }
            | _ -> failwith "Invalid format"
            )
        |> Seq.toList

    gekMembers
