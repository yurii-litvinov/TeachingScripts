open System.IO
open LessPainfulGoogleSheets

type Work = 
    {
        name: string
        o1Score: string
        o1Comment: string
        o2Score: string
        o2Comment: string
        t1Score: string
        t1Comment: string
        t2Score: string
        t2Comment: string
        p1Score: string
        p1Comment: string
        p2Score: string
        p2Comment: string
        general: string
    }

[<EntryPoint>]
let main argv =
    printfn "Connecting to Google Sheets..."

    let spreadsheetId = argv.[0]
    let offset = argv.[1]
    let page = "Form Responses 1"
    use service = openGoogleSheet "AssignmentMatcher" 

    let sheet = readGoogleSheet service spreadsheetId page "A" "P" offset

    let works = sheet 
                |> Seq.map(fun line -> 
                    let line = line |> Seq.toArray
                    { 
                        name = line.[1]
                        o1Score = line.[3]
                        o1Comment = line.[4]
                        o2Score = line.[5]
                        o2Comment = line.[6]
                        t1Score = line.[7]
                        t1Comment = line.[8]
                        t2Score = line.[9]
                        t2Comment = line.[10]
                        p1Score = line.[11]
                        p1Comment = line.[12]
                        p2Score = line.[13]
                        p2Comment = if line.Length > 14 then line.[14] else ""
                        general = if line.Length > 15 then line.[15] else ""
                    }
                    )

    works 
    |> Seq.iter(fun work ->
        use file = new FileStream(work.name.Replace('\"', ' ').Replace('/', '-').Replace('\\', '-') + ".txt", FileMode.Create)
        use writer = new StreamWriter(file)
        writer.WriteLine("Здравствуйте.")
        writer.WriteLine()
        writer.WriteLine("Вот рецензия, по критериям. Каждый критерий оценивается по пятибалльной шкале.")
        writer.WriteLine()

        writer.WriteLine("О1. Соответствие содержания и оформления предъявленным требованиям:")
        writer.WriteLine(work.o1Score)
        writer.WriteLine(work.o1Comment)
        writer.WriteLine()

        writer.WriteLine("О2. Умение работать с информацией, опубликованной в научных и иных источниках:")
        writer.WriteLine(work.o2Score)
        writer.WriteLine(work.o2Comment)
        writer.WriteLine()

        writer.WriteLine("Т1. Обоснование принятых решений/Теоретический анализ:")
        writer.WriteLine(work.t1Score)
        writer.WriteLine(work.t1Comment)
        writer.WriteLine()

        writer.WriteLine("Т2. Сравнение с аналогами:")
        writer.WriteLine(work.t2Score)
        writer.WriteLine(work.t2Comment)
        writer.WriteLine()

        writer.WriteLine("П1. Качество практической части:")
        writer.WriteLine(work.p1Score)
        if work.p1Comment <> "" then
            writer.WriteLine(work.p1Comment)
        writer.WriteLine()

        writer.WriteLine("П2. Качество проводимых измерений и постановки экспериментов:")
        writer.WriteLine(work.p2Score)
        if work.p2Comment <> "" then
            writer.WriteLine(work.p2Comment)
        writer.WriteLine()

        if work.general <> "" then
            writer.WriteLine("Общий комментарий к работе:")
            writer.WriteLine(work.general)
            writer.WriteLine()

        writer.WriteLine("В общем, надо поправить, выложить в папку и сообщить мне, что работа готова к повторному рецензированию")

        ()
        )

    printfn "Done!"

    0
