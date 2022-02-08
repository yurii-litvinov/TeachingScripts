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
        reviewLink: string
        status: string
    }

[<EntryPoint>]
let main argv =
    printfn "Connecting to Google Sheets..."

    let spreadsheetId = argv.[0]
    let offset = argv.[1]
    let page = "Form Responses 1"
    use service = openGoogleSheet "AssignmentMatcher" 

    let sheet = readGoogleSheet service spreadsheetId page "A" "R" offset

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
                        reviewLink = if line.Length > 16 then line.[16] else ""
                        status = if line.Length > 17 then line.[17] else ""
                    }
                    )

    works 
    |> Seq.iter(fun work ->
        use file = new FileStream(work.name.Replace('\"', ' ').Replace('/', '-').Replace('\\', '-').Replace(':', ' ') + ".txt", FileMode.Create)
        use writer = new StreamWriter(file)

        let writeSection (title: string) (score: string) (comment: string) =
            writer.WriteLine(title)
            writer.WriteLine(score)
            if comment <> "" then
                writer.WriteLine(comment)
            writer.WriteLine()

        writer.WriteLine("Здравствуйте.")
        writer.WriteLine()
        writer.WriteLine("Вот рецензия на текст практики. Каждый критерий оценивается по пятибалльной шкале.")
        writer.WriteLine()

        writeSection "О1. Соответствие содержания и оформления предъявленным требованиям:" work.o1Score work.o1Comment
        writeSection "О2. Умение работать с информацией, опубликованной в научных и иных источниках:" work.o2Score work.o2Comment

        writeSection "Т1. Обоснование принятых решений/Теоретический анализ:" work.t1Score work.t1Comment
        writeSection "Т2. Сравнение с аналогами:" work.t2Score work.t2Comment

        writeSection "П1. Качество практической части:" work.p1Score work.p1Comment
        writeSection "П2. Качество проводимых измерений и постановки экспериментов:" work.p2Score work.p2Comment

        if work.general <> "" then
            writer.WriteLine("Общий комментарий к работе:")
            writer.WriteLine(work.general)
            writer.WriteLine()

        if work.reviewLink <> "" then
            writer.WriteLine("Ссылка на подробное ревью: " + work.reviewLink)
            writer.WriteLine()

        if work.status <> "Можно зачесть" then
            writer.WriteLine("В общем, надо поправить, выложить в папку и сообщить мне, что работа готова к повторному рецензированию.")
        else
            writer.WriteLine("В общем, текст зачтён.")

        ()
        )

    printfn "Done!"

    0
