﻿module ProgramPatcher

open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing
open System.IO
open DocumentFormat.OpenXml
open ProgramContentChecker
open CurriculumUtils

type Justification =
    | Center
    | Stretch

let createParagraph text (justification: Justification) (bold: bool) (italic: bool) =
    let text = Text(text)
    let run = Run()

    let runProperties = RunProperties()
    let fonts = 
        RunFonts
            (Ascii = StringValue("Times New Roman"), 
             HighAnsi = StringValue("Times New Roman"), 
             ComplexScript = StringValue("Times New Roman"))

    let size = FontSize(Val = StringValue("24"))
    let sizeCs = FontSizeComplexScript(Val = StringValue("24"))
    
    runProperties.AppendChild(fonts) |> ignore
    runProperties.AppendChild(size) |> ignore
    runProperties.AppendChild(sizeCs) |> ignore

    if bold then 
        runProperties.AppendChild(Bold()) |> ignore
    if italic then 
        runProperties.AppendChild(Italic()) |> ignore

    run.AppendChild(runProperties) |> ignore
    run.AppendChild(text) |> ignore

    let paragraphProperties = ParagraphProperties()
    let indent = Indentation(FirstLine = StringValue("720"))
    let justificationValue = if justification = Center then JustificationValues.Center else JustificationValues.Both
    if justification <> Justification.Center then
        // Do not indent paragraphs with centering.
        paragraphProperties.Indentation <- indent
    let justification = Justification(Val = EnumValue<_>(justificationValue))
    paragraphProperties.Justification <- justification
    paragraphProperties.SpacingBetweenLines <- SpacingBetweenLines(After = StringValue("120"))

    let paragraph = Paragraph()
    paragraph.AppendChild paragraphProperties |> ignore
    paragraph.AppendChild run |> ignore
    paragraph

let findParagraph (body: Body) (text: string) =
    let paragraphs = body.Descendants<Paragraph>()
    paragraphs |> Seq.tryFind (fun p -> p.InnerText.Contains(text))

let patchSection (body: Body) (sectionName: string) (replaceTo: string list) =
    let paragraphs = body.Descendants<Paragraph>()

    let index = paragraphs |> Seq.tryFindIndex (fun p -> p.InnerText.Contains(sectionName))
    if index.IsSome then
        let index = index.Value
        let sectionHeader = paragraphs |> Seq.item index
        let target = paragraphs |> Seq.item (index + 1)

        if not (target.InnerText.Contains(replaceTo.[0])) then
            printfn "Заменяю содержимое секции '%s':\n\
                %s\n\
                на стандартную фразу\n" 
                sectionName target.InnerText

            target.Remove()
            for replacePart in replaceTo |> List.rev do
                let paragraph = createParagraph replacePart Stretch false false
                sectionHeader.InsertAfterSelf paragraph |> ignore

let addGovernmentIfNeeded (body: Body) =
    let text = body.InnerText.TrimStart()
    if not (text.StartsWith "Правительство Российской Федерации") then
        printfn "Добавляю 'Правительство Российской Федерации' в начало"
        let university = 
            body.Descendants<Paragraph>() 
            |> Seq.find (fun p -> p.InnerText.Contains("Санкт-Петербургский государственный университет"))
        let government = createParagraph "Правительство Российской Федерации" Center true false
        let props = government.Descendants<RunProperties>() |> Seq.exactlyOne
        let spacing = Spacing(Val = Int32Value(20))
        props.AppendChild spacing |> ignore
        university.InsertBeforeSelf government |> ignore

let fixTextInRun (body: Body) (from: string) (replaceTo: string) =
    let texts = body.Descendants<Text>()
    let text = texts |> Seq.tryFind (fun r -> r.InnerText = from)
    if text.IsSome then
        printfn "Меняю '%s' на '%s'" from replaceTo
        text.Value.Text <- replaceTo

let addLiterature (body: Body) =
    let sectionHeader = findParagraph body "Перечень иных информационных источников"
    if sectionHeader.IsSome then
        let sectionHeader = sectionHeader.Value

        let literature = 
            [ "Сайт Научной библиотеки им. М. Горького СПбГУ: http://www.library.spbu.ru/"
              "Электронный каталог Научной библиотеки им. М. Горького СПбГУ: http://www.library.spbu.ru/cgi-bin/irbis64r/cgiirbis_64.exe?C21COM=F&I21DBN=IBIS&P21DBN=IBIS" 
              "Перечень электронных ресурсов, находящихся в доступе СПбГУ: http://cufts.library.spbu.ru/CRDB/SPBGU/" 
              "Перечень ЭБС, на платформах которых представлены российские учебники, находящиеся в доступе СПбГУ: http://cufts.library.spbu.ru/CRDB/SPBGU/browse?name=rures&resource_type=8" ]
            |> List.rev

        printfn "Добавляю в перечень иных информационных источников стандартный список литературы (без форматирования!)" 

        literature 
        |> List.iter (fun s ->
            let p = createParagraph s Stretch false false
            sectionHeader.InsertAfterSelf p |> ignore)

let addCompetencesToAttestationMaterials (content: ProgramContent) (curriculum: Curriculum) (body: Body) (disciplineCode: string) =
    if content |> controlMaterialsShallReferenceCompetences |> List.isEmpty |> not then
        let addCompetences (competences: Competence seq) (insertAfter: Paragraph) =
            if competences |> Seq.isEmpty then
                let p = createParagraph "Нет." Stretch false false
                insertAfter.InsertAfterSelf p
            else 
                competences 
                |> Seq.fold (fun (lastParagraph: Paragraph) competence -> 
                        let description = $"{competence.Description.[0]}".ToLower() + competence.Description.Substring(1)
                        let p = createParagraph ($"{competence.Code} — {description}.") Stretch false false
                        lastParagraph.InsertAfterSelf p
                    ) insertAfter

        let addHeader (header: string) (insertAfter: Paragraph) =
            createParagraph header Stretch true false
            |> insertAfter.InsertAfterSelf

        let sectionHeader = findParagraph body "Методические материалы для проведения текущего контроля успеваемости и промежуточной аттестации (контрольно-измерительные материалы, оценочные средства)"
        if sectionHeader.IsSome then
            printfn "Вставляю компетенции и шкалу оценивания в раздел 3.1.4."
            let sectionHeader = sectionHeader.Value
            let discipline = curriculum.Disciplines |> Seq.find (fun d -> d.Code = disciplineCode)
            let formedCompetences = discipline.FormedCompetences
            let improvedCompetences = discipline.ImprovedCompetences
            let fullyFormedCompetences = discipline.FullyFormedCompetences
        
            let competencesParagraph =
                addHeader "Компетенции, впервые формируемые дисциплиной:" sectionHeader
                |> addCompetences formedCompetences
                |> addHeader "Компетенции, развиваемые дисциплиной:" 
                |> addCompetences improvedCompetences
                |> addHeader "Компетенции, полностью сформированные по результатам освоения дисциплины:" 
                |> addCompetences fullyFormedCompetences
        
            let p = createParagraph "Для каждой компетенции применяется линейная шкала оценивания, определяемая долей успешно выполненных заданий, проверяющих данную компетенцию" Stretch false false
            competencesParagraph.InsertAfterSelf p |> ignore

            let sectionHeader = findParagraph body "Методические материалы для оценки обучающимися содержания и качества учебного процесса"
            if sectionHeader.IsSome then
                let sectionHeader = sectionHeader.Value
                let competences = 
                    [formedCompetences; improvedCompetences; fullyFormedCompetences] 
                    |> Seq.concat
                    |> Seq.map (fun c -> c.Code)
                    |> Seq.reduce (fun acc s -> acc + ", " + s)
                let competences = "Проверяемые компетенции: " + competences
                let p = createParagraph competences Stretch false true
                sectionHeader.InsertBeforeSelf p |> ignore

                let p = createParagraph "Сформированность компетенций считается пропорционально доле успешных ответов на вопросы и выполненности заданий." Stretch false true
                competencesParagraph.InsertAfterSelf p |> ignore


let fixStudent (content: ProgramContent) (body: Body) =
    if content |> noStudent |> List.isEmpty |> not then
        printfn "Пытаюсь заменить запрещённое слово 'студент' на 'обучающийся'"
        fixTextInRun body "студента" "обучающегося"
        fixTextInRun body "студенту" "обучающемуся"
        fixTextInRun body "студентом" "обучающимся"
        fixTextInRun body "студенте" "обучающемся"
        fixTextInRun body "студент" "обучающийся"

        fixTextInRun body "Студента" "Обучающегося"
        fixTextInRun body "Студенту" "Обучающемуся"
        fixTextInRun body "Студентом" "Обучающимся"
        fixTextInRun body "Студенте" "Обучающемся"
        fixTextInRun body "Студент" "Обучающийся"

let patchProgram (curriculumFile: string) (programFileName: string) =
    try
        use wordDocument = WordprocessingDocument.Open(programFileName, true)
        let content, _ = parseProgram wordDocument
        let curriculum = Curriculum curriculumFile
        let disciplineCode = FileInfo(programFileName).Name.Substring(0, 6)
        let body = wordDocument.MainDocumentPart.Document.Body

        printfn "Программа %s:\n" (FileInfo(programFileName).Name)

        addGovernmentIfNeeded body

        fixTextInRun 
            body
            "Требования подготовленности обучающегося к освоению содержания учебных занятий ("
            "Требования к подготовленности обучающегося к освоению содержания учебных занятий ("

        addCompetencesToAttestationMaterials content curriculum body disciplineCode

        patchSection 
            body 
            "Методические материалы для оценки обучающимися содержания и качества учебного процесса"
            [ "Для оценки обучающимися содержания и качества учебного процесса применяется анкетирование в соответствии с методикой и графиком, утвержденными в установленном порядке." ]

        patchSection 
            body 
            "Характеристики аудиторий (помещений, мест) для проведения занятий"
            [ "Учебные аудитории для проведения учебных занятий, оснащенные стандартным оборудованием, используемым для обучения в СПбГУ в соответствии с требованиями материально-технического обеспечения." ]

        patchSection 
            body 
            "Характеристики аудиторного оборудования, в том числе неспециализированного компьютерного оборудования и программного обеспечения общего пользования"
            [ "Стандартное оборудование, используемое для обучения в СПбГУ."
              "MS Windows, MS Office, Mozilla FireFox, Google Chrome, Acrobat Reader DC, WinZip, Антивирус Касперского." ]

        if content |> libraryLinksShallPresent |> List.isEmpty |> not then
            addLiterature body

        fixStudent content body

        wordDocument.Save()
    with
    | :? OpenXmlPackageException -> 
        printfn "%s настолько коряв, что даже не читается, пропущен" programFileName
