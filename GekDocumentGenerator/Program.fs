open LessPainfulGoogleSheets
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Wordprocessing
open System.IO

let fillSeq size seq =
    let length = Seq.length seq
    if size > length then
        Seq.append seq (Seq.replicate (size - length) "")
    else
        seq

type SittingTime = 
    Morning | Day | Evening
    override this.ToString() =
        match this with
        | Morning -> "10:00"
        | Day -> "14:00"
        | Evening -> "17:00"
    static member fromString str =
        match str with
        | "10:00" -> Morning
        | "14:00" -> Day
        | "17:00" -> Evening
        | _ -> failwith "Incorrect format."

type Program = 
    PIBachelors | MOBachelors | PIMasters | MOMasters | Aspirants
    override this.ToString() =
        match this with
        | PIBachelors -> "09.03.04 Программная инженерия (бакалавры)"
        | MOBachelors -> "02.03.03 Математическое обеспечение и администрирование информационных систем (бакалавры)"
        | PIMasters -> "09.04.04 Программная инженерия (магистры)"
        | MOMasters -> "02.04.03 Математическое обеспечение и администрирование информационных систем (магистры)"
        | Aspirants -> "09.06.01 Информатика (аспиранты)"
    static member fromString str =
        match str with
        | "бакалавры ПИ" -> PIBachelors
        | "бакалавры МОиАИС" -> MOBachelors
        | "магистры ПИ" -> PIMasters
        | "магистры МОиАИС" -> MOMasters
        | "Аспиранты" -> Aspirants
        | _ -> failwith "Unknown program"

type Sitting = 
    { Program: Program
      CommissionCode: string 
      Time: SittingTime }

type CommissionMember =
    { Name: string
      Company: string 
      Position: string 
      Mail: string
      Degree: string }
    static member empty = { Name = ""; Company = ""; Position = ""; Mail = ""; Degree = "" }
    static member create name = { CommissionMember.empty with Name = name }
    override this.ToString() =
        let degree = if this.Degree = "нет" then "" else this.Degree
        let degree = if degree = "" then "" else degree + ", "
        $"{this.Name}, {degree}{this.Position}, {this.Company}, {this.Mail}"

type Day =
    { Date: string
      Chair: CommissionMember
      Commission: CommissionMember list
      Sittings: Sitting list }
    static member empty =
        { Date = ""; Chair = CommissionMember.empty; Commission = []; Sittings = [] }

type ParserState =
    { Accumulator: Day list
      Current: Day }
    static member empty = { Accumulator = []; Current = Day.empty }
    member this.NextState currentDay = { this with Current = currentDay }

let parseSchedule service =
    let parser (state: ParserState) line =
        match line with
        | [ date; ""; ""; ""; chair ] when chair.StartsWith "Председатель: " ->
            let current = 
                { Day.empty with 
                    Date = date 
                    Chair = CommissionMember.create (chair.Substring ("Председатель: ".Length)) }
            if state.Current = Day.empty then
                state.NextState current
            else
                { state with Accumulator = state.Current :: state.Accumulator; Current = current }
        | [ date; ""; ""; ""; "" ] when date <> "" && System.Char.IsDigit(date.Chars 0) -> 
            if date <> state.Current.Date then failwith "Invalid schedule"
            state
        | [ time; commission; ""; ""; _ ] when time = "10:00" || time = "14:00" || time = "17:00" ->
            let time = SittingTime.fromString time
            let splitCommission = commission.Split ", "
            let program, code =
                if splitCommission.Length = 3 then
                    splitCommission.[1], splitCommission.[2]
                else
                    splitCommission.[0], splitCommission.[1]
            let sitting = { Program = Program.fromString program; CommissionCode = code; Time = time }
            state.NextState { state.Current with Sittings =  sitting :: state.Current.Sittings }
        | [ _; _; _; ""; commissionMember ] when commissionMember <> "" ->
            let commissionMember = CommissionMember.create commissionMember
            state.NextState { state.Current with Commission = commissionMember :: state.Current.Commission }
        | [ _; _; _; ""; "" ] -> state
        | _ -> failwith "Wrong format"

    let scheduleRawData = readGoogleSheet service "15KWQOg8i9-3BkbZOOPfbAqVp1EYPeqKFpjH7pKO8JIs" "Sheet1" "B" "F" 2
    
    let state =
        scheduleRawData 
        |> Seq.map (fillSeq 5)
        |> Seq.map Seq.toList
        |> Seq.fold parser ParserState.empty

    state.Current :: state.Accumulator |> List.rev

let addMemberInfo service (schedule: Day list) =
    let memberData = 
        readGoogleSheet service "1IEFRLzyTqM0bG2HzKxQNuSh_rmpHvKS2qE8ywhQAEPg" "Члены ГЭК" "A" "I" 3
        |> Seq.map (fillSeq 9)
        |> Seq.map Seq.toList
        |> Seq.map (fun l -> match l with h :: t -> (h, t) | _ -> failwith "Invalid format")
        |> Map.ofSeq

    let addInfo (commissionMember: CommissionMember) =
        let info = memberData.[commissionMember.Name]
        { commissionMember with
            Company = info.[0]
            Position = info.[1]
            Mail = info.[2]
            Degree = info.[7] }

    let schedule = schedule |> List.map (fun day -> { day with Commission = day.Commission |> List.map addInfo })

    let chairData = 
        readGoogleSheet service "1IEFRLzyTqM0bG2HzKxQNuSh_rmpHvKS2qE8ywhQAEPg" "Председатели" "A" "G" 3
        |> Seq.map (fillSeq 7)
        |> Seq.map Seq.toList
        |> Seq.map (fun l -> match l with h :: t -> (h, t) | _ -> failwith "Invalid format")
        |> Map.ofSeq

    let addChairInfo (chair: CommissionMember) =
        let info = chairData.[chair.Name]
        { chair with
            Company = info.[1]
            Position = info.[2]
            Mail = info.[3]
            Degree = info.[5] }

    schedule
    |> List.map (fun day -> { day with Chair = addChairInfo day.Chair })

let generateScheduleDoc (schedule: Day list) =
    let print (body: Body) justification text =
        let paragraph = LessPainfulDocx.createParagraph text justification false false false
        body.AppendChild paragraph |> ignore

    use wordDocument = WordprocessingDocument.Create("Result.docx", WordprocessingDocumentType.Document)
    let mainPart = wordDocument.AddMainDocumentPart()
    mainPart.Document <- new Document()
    let body = mainPart.Document.AppendChild(new Body())

    let generateSitting (day: Day) (sitting: Sitting) =
        let (!<--) str = print body LessPainfulDocx.Left str
        let (!<->) str = print body LessPainfulDocx.Center str

        !<-> $"ГЭК для {sitting.Program.ToString ()}"
        !<-> $"Дата заседания: {day.Date}"
        !<-> $"Место заседания: Университетский пр., д. 28, ауд. 405"
        !<-> $"Время заседания: с {sitting.Time}"
        !<-- $"Состав:"

        let numberingId = LessPainfulDocx.createNumbering wordDocument

        let createNumberingPr (numberingId: Int32Value) =
            let ilvl = NumberingLevelReference(Val = Int32Value(0))
            let numId = NumberingId(Val = numberingId)
            let numPr = NumberingProperties(NumberingId = numId, NumberingLevelReference = ilvl)
            numPr

        let (!-) str =
            let paragraph = LessPainfulDocx.createParagraph str LessPainfulDocx.Stretch false false false
            let numPr = createNumberingPr numberingId
            paragraph.ParagraphProperties.AppendChild numPr |> ignore
            body.AppendChild paragraph |> ignore

        !- $"{day.Chair} — председатель;"
        
        day.Commission 
        |> List.sortBy (fun cm -> cm.Name)
        |> List.iteri (fun i m -> if i = day.Commission.Length - 1 then !- $"{m}." else !- $"{m};")
        
        !<-- ""

    let generateDay (day: Day) =
        day.Sittings |> List.rev |> List.iter (generateSitting day)

    schedule |> List.iter generateDay

let generateMails (schedule: Day list) =
    let generateMail (writer: StreamWriter) (commissionMember: CommissionMember, list: (string * Program list) list) =
        let program = function
            | PIBachelors -> "бакалавров программной инженерии"
            | MOBachelors -> "бакалавров математического обеспечения"
            | PIMasters -> "магистров программной инженерии"
            | MOMasters -> "магистров математического обеспечения"
            | Aspirants -> "аспирантов"

        let programs (p: Program list) =
            p |> List.fold (fun acc cur -> (if acc <> "" then acc + ", " else "") + (program cur)) ""

        writer.WriteLine $"to: {commissionMember.Mail}"
        writer.WriteLine $""
        let splitName = commissionMember.Name.Split(" ")
        writer.WriteLine $"Здравствуйте, {splitName.[1]} {splitName.[2]}!"
        writer.WriteLine $""

        if list.Length = 1 then
            writer.WriteLine $"Расписание защит готово, хотим Вас занять на один день, {fst list.Head}, на защиты {programs (snd list.Head)}"
        else
            writer.WriteLine $"Расписание защит готово, хотим Вас занять на следующие дни:"
            list |> List.iter (fun day -> writer.WriteLine $"- {fst day}, на защиты {programs (snd day)}")

        writer.WriteLine $""
        writer.WriteLine $"Если никаких изменений не будет, где-то за неделю до защит я разошлю тексты дипломов для предварительного ознакомления."
        writer.WriteLine $""
        writer.WriteLine $"С уважением,"
        writer.WriteLine $"Юрий Литвинов."
        writer.WriteLine $""
        writer.WriteLine $""

    let commissionMembers = 
        schedule
        |> List.map (fun day -> (day.Date, day.Sittings |> List.map (fun s -> s.Program), day.Chair :: day.Commission))
        |> List.map (fun (date, sittings, commission) -> commission |> List.map (fun m -> (m, date, sittings)))
        |> List.concat
        |> List.groupBy (fun (m, _, _) -> m)
        |> List.map (fun (k, v) -> (k, v |> List.map (fun (_, d, s) -> (d, s))))
        |> List.map (fun (k, v) -> (k, v |> List.groupBy (fun (date, _) -> date)))
        |> List.map (fun (k, v) -> 
            (k, v |> List.map (fun (date, list) -> 
                (date, list |> List.map (fun (_, list) -> list) |> List.concat |> List.distinct))))

    use file = File.Open("mails.txt", FileMode.Create)
    use writer = new StreamWriter(file)
    commissionMembers |> List.iter (generateMail writer)

[<EntryPoint>]
let main _ =
    let service = openGoogleSheet "AssignmentMatcher"
    let schedule = parseSchedule service
    let schedule = addMemberInfo service schedule

    generateScheduleDoc schedule

    generateMails schedule

    0
