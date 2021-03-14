﻿module ProgramPatcher

open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing
open System.IO
open DocumentFormat.OpenXml
open ProgramContentChecker
open CurriculumUtils
open System

type Justification =
    | Center
    | Stretch
    | Left

let createParagraph text (justification: Justification) (bold: bool) (italic: bool) (indent: bool) =
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
    
    runProperties.AppendChild fonts |> ignore
    runProperties.AppendChild size |> ignore
    runProperties.AppendChild sizeCs |> ignore

    if bold then 
        runProperties.AppendChild(Bold()) |> ignore
    if italic then 
        runProperties.AppendChild(Italic()) |> ignore

    run.AppendChild(runProperties) |> ignore
    run.AppendChild(text) |> ignore

    let paragraphProperties = ParagraphProperties()
    let justificationValue = 
        match justification with
        | Center -> JustificationValues.Center 
        | Stretch -> JustificationValues.Both
        | Left -> JustificationValues.Left

    if justification <> Justification.Center && indent then
        // Do not indent paragraphs with centering.
        let indent = Indentation(FirstLine = StringValue("720"))
        paragraphProperties.Indentation <- indent
    let justification = Justification(Val = EnumValue<_>(justificationValue))
    paragraphProperties.Justification <- justification
    paragraphProperties.SpacingBetweenLines <- SpacingBetweenLines(After = StringValue("120"))

    let paragraph = Paragraph()
    paragraph.AppendChild paragraphProperties |> ignore
    paragraph.AppendChild run |> ignore
    paragraph

let createNumbering (document: WordprocessingDocument) =
    let numberingDefinitionsPart = document.MainDocumentPart.NumberingDefinitionsPart
    if numberingDefinitionsPart = null then
        document.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>() |> ignore
        document.MainDocumentPart.NumberingDefinitionsPart.Numbering <- Numbering()
    let numberingPart = document.MainDocumentPart.NumberingDefinitionsPart.Numbering
    let lastAbstractNumberingId = 
        if numberingPart.Descendants<AbstractNum>() |> Seq.isEmpty then
            0
        else
            numberingPart.Descendants<AbstractNum>() 
            |> Seq.map (fun an -> an.AbstractNumberId.Value) 
            |> Seq.sortDescending 
            |> Seq.head

    let abstractNum = AbstractNum()
    abstractNum.AbstractNumberId <- Int32Value(lastAbstractNumberingId + 1)
    let lvl0 = Level()
    lvl0.LevelIndex <- Int32Value 0
    lvl0.StartNumberingValue <- StartNumberingValue(Val = Int32Value 1)
    lvl0.NumberingFormat <- NumberingFormat(Val = EnumValue<NumberFormatValues>(NumberFormatValues.Decimal))
    lvl0.LevelText <- LevelText(Val = StringValue "%1.")
    lvl0.LevelJustification <- LevelJustification(Val = EnumValue<LevelJustificationValues>(LevelJustificationValues.Left))

    let lvl0PPr = ParagraphProperties()
    lvl0PPr.Indentation <- Indentation(Left = StringValue("360"), FirstLine = StringValue("0"), Hanging = StringValue("360"))
    lvl0.AppendChild lvl0PPr |> ignore

    let lvl0RunPr = RunProperties()
    let fonts = 
        RunFonts
            (Ascii = StringValue("Times New Roman"), 
             HighAnsi = StringValue("Times New Roman"), 
             ComplexScript = StringValue("Times New Roman"))

    let size = FontSize(Val = StringValue("24"))
    let sizeCs = FontSizeComplexScript(Val = StringValue("24"))
    
    lvl0RunPr.AppendChild fonts |> ignore
    lvl0RunPr.AppendChild size |> ignore
    lvl0RunPr.AppendChild sizeCs |> ignore
    lvl0.AppendChild lvl0RunPr |> ignore

    abstractNum.AppendChild lvl0 |> ignore

    numberingPart.AppendChild abstractNum |> ignore

    let num = NumberingInstance()
    let lastNumberingId = 
        if numberingPart.Descendants<NumberingInstance>() |> Seq.isEmpty then
            0
        else
            numberingPart.Descendants<NumberingInstance>() 
            |> Seq.map (fun n -> n.NumberID.Value) 
            |> Seq.sortDescending 
            |> Seq.head

    num.NumberID <- Int32Value(lastNumberingId + 1)

    let abstractNumberingRef = AbstractNumId(Val = Int32Value(lastAbstractNumberingId + 1))
    num.AppendChild abstractNumberingRef |> ignore

    numberingPart.AppendChild num |> ignore

    num.NumberID

let createHyperlink link (document: WordprocessingDocument) =
    let text = Text(link)
    let run = Run()

    let runProperties = RunProperties()
    let fonts = 
        RunFonts
            (Ascii = StringValue("Times New Roman"), 
             HighAnsi = StringValue("Times New Roman"), 
             ComplexScript = StringValue("Times New Roman"))

    let size = FontSize(Val = StringValue("24"))
    let sizeCs = FontSizeComplexScript(Val = StringValue("24"))
    let underline = Underline(Val = EnumValue<_>(UnderlineValues.Single))
    let color = Color(ThemeColor = EnumValue<_>(ThemeColorValues.Hyperlink))

    runProperties.AppendChild fonts |> ignore
    runProperties.AppendChild size |> ignore
    runProperties.AppendChild sizeCs |> ignore
    runProperties.AppendChild underline|> ignore
    runProperties.AppendChild color|> ignore

    run.AppendChild(runProperties) |> ignore
    run.AppendChild(text) |> ignore

    let paragraphProperties = ParagraphProperties()

    let indent = Indentation(Left = StringValue("360"))
    paragraphProperties.Indentation <- indent
    let justification = Justification(Val = EnumValue<_>(JustificationValues.Left))
    paragraphProperties.Justification <- justification
    paragraphProperties.SpacingBetweenLines <- SpacingBetweenLines(After = StringValue("120"))

    let linkParagraph = Paragraph()
    linkParagraph.AppendChild paragraphProperties |> ignore

    let rel = document.MainDocumentPart.AddHyperlinkRelationship(Uri(link), true)

    let hyperlink = Hyperlink()
    hyperlink.AppendChild run |> ignore
    hyperlink.Id <- StringValue(rel.Id)
    hyperlink.History <- OnOffValue.FromBoolean(true)            

    linkParagraph.AppendChild hyperlink |> ignore
    linkParagraph

let findParagraph (body: Body) (text: string) =
    let paragraphs = body.Descendants<Paragraph>()
    paragraphs |> Seq.tryFind (fun p -> p.InnerText.Contains(text))

let findPreviousParagraph (body: Body) (paragraph: Paragraph) =
    body.Descendants<Paragraph>() 
    |> Seq.pairwise 
    |> Seq.find (fun (_, next) -> next = paragraph) 
    |> fst

let findNextParagraph (body: Body) (paragraph: Paragraph) =
    body.Descendants<Paragraph>() 
    |> Seq.pairwise 
    |> Seq.find (fun (prev, _) -> prev = paragraph) 
    |> snd

let rec removeNextParagraphIfEmpty (body: Body) (paragraph: Paragraph) (recursive: bool) =
    let nextParagraph = findNextParagraph body paragraph

    if nextParagraph.InnerText = "" then
        nextParagraph.Remove()
        if recursive then
            removeNextParagraphIfEmpty body paragraph recursive

let rec removePreviousParagraphIfEmpty (body: Body) (paragraph: Paragraph) (recursive: bool) =
    let previousParagraph = findPreviousParagraph body paragraph

    if previousParagraph.InnerText = "" then
        previousParagraph.Remove()
        if recursive then
            removePreviousParagraphIfEmpty body paragraph recursive


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
                let paragraph = createParagraph replacePart Stretch false false true
                sectionHeader.InsertAfterSelf paragraph |> ignore
    else
        printfn "Секция '%s' не найдена, замена не произведена!" sectionName

let addGovernment (body: Body) =
    let text = body.InnerText.TrimStart()
    if not (text.StartsWith "Правительство Российской Федерации") then
        printfn "Добавляю 'Правительство Российской Федерации' в начало"
        let university = 
            body.Descendants<Paragraph>() 
            |> Seq.find (fun p -> p.InnerText.Contains("Санкт-Петербургский государственный университет"))
        let government = createParagraph "Правительство Российской Федерации" Center true false true
        let props = government.Descendants<RunProperties>() |> Seq.exactlyOne
        let spacing = Spacing(Val = Int32Value(20))
        props.AppendChild spacing |> ignore
        university.InsertBeforeSelf government |> ignore

let addStPetersburg (errors: string list) (body: Body) =
    if errors |> List.exists (fun s -> s.StartsWith "Не найден 'Санкт-Петербург'") then
        let nextSectionHeader = findParagraph body "Раздел 1"
        if nextSectionHeader.IsSome then
            let nextSectionHeader = nextSectionHeader.Value
            removePreviousParagraphIfEmpty body nextSectionHeader false
            printfn "Добавляю 'Санкт-Петербург 2020' примерно в титульник"
            let stPetersburg = createParagraph "Санкт-Петербург" Center false false true
            nextSectionHeader.InsertBeforeSelf stPetersburg |> ignore
            let year = createParagraph "2020" Center false false true
            nextSectionHeader.InsertBeforeSelf year |> ignore

            let breakElement = Break(Type = EnumValue<BreakValues>(BreakValues.Page))
            let breakRun = Run()
            breakRun.AppendChild breakElement |> ignore
            let pageBreak = Paragraph()
            pageBreak.AppendChild breakRun |> ignore
            nextSectionHeader.InsertBeforeSelf pageBreak |> ignore
        else
            printfn "Раздел 1 не найден, не могу добавить 'Санкт-Петербург'"

let fixTextInRun (body: Body) (from: string) (replaceTo: string) =
    let texts = body.Descendants<Text>()
    let text = texts |> Seq.tryFind (fun r -> r.InnerText = from)
    if text.IsSome then
        printfn "Меняю '%s' на '%s'" from replaceTo
        text.Value.Text <- replaceTo

let fixTextEverywhere (body: Body) (from: string) (replaceTo: string) =
    body.Descendants<Text>()
    |> Seq.iter (fun t -> t.Text <- t.Text.Replace(from, replaceTo))

let addCompetencesToLearningOutcomes (content: ProgramContent) (body: Body) =
    if content |> shallContainCompetences |> List.isEmpty |> not then
        let nextSectionHeader = findParagraph body "Перечень и объём активных и интерактивных форм учебных занятий"
        if nextSectionHeader.IsSome then
            printfn "Вставляю стандартную фразу про компетенции в раздел 1.3."
            let p = createParagraph "Дисциплина участвует в формировании компетенций обучающихся по образовательной программе, установленных учебным планом для данной дисциплины." Stretch false false true
            removePreviousParagraphIfEmpty body nextSectionHeader.Value true
            nextSectionHeader.Value.InsertBeforeSelf p |> ignore
        else
            printfn "Раздел 1.4 не найден, пропускаю добавление фразы про компетенции в раздел 1.3."

let addCompetencesToAttestationMaterials (content: ProgramContent) (curriculum: Curriculum) (body: Body) (disciplineCode: string) =
    if content |> controlMaterialsShallReferenceCompetences |> List.isEmpty |> not then
        let addCompetences (competences: Competence seq) (insertAfter: Paragraph) =
            if competences |> Seq.isEmpty then
                let p = createParagraph "Нет." Stretch false false true
                insertAfter.InsertAfterSelf p
            else 
                competences 
                |> Seq.fold (fun (lastParagraph: Paragraph) competence -> 
                        let description = $"{competence.Description.[0]}".ToLower() + competence.Description.Substring(1)
                        let p = createParagraph ($"{competence.Code} — {description}.") Stretch false false true
                        lastParagraph.InsertAfterSelf p
                    ) insertAfter

        let addHeader (header: string) (insertAfter: Paragraph) =
            createParagraph header Stretch true false true
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
        
            let p = createParagraph "Для каждой компетенции применяется линейная шкала оценивания, определяемая долей успешно выполненных заданий, проверяющих данную компетенцию" Stretch false false true
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
                let p = createParagraph competences Stretch false true true
                sectionHeader.InsertBeforeSelf p |> ignore

                let p = createParagraph "Сформированность компетенций считается пропорционально доле успешных ответов на вопросы и доле выполненных заданий." Stretch false true true
                sectionHeader.InsertBeforeSelf p |> ignore
            else
                printfn "'Методические материалы для оценки обучающимися содержания и качества учебного процесса' не удалось найти, список проверяемых ФОС компетенций не сгенерирован!"
        else
            printfn "'Методические материалы для проведения текущего контроля успеваемости и промежуточной аттестации' не удалось найти, компетенции в ФОС и шкалы оценивания не сгенерированы!"

let fillRequiredLiterature (content: ProgramContent) (body: Body) =
    if content.["3.4.1. Список обязательной литературы"].Trim() = "" then
        let sectionHeader = findParagraph body "Список обязательной литературы"
        if sectionHeader.IsSome then
            let sectionHeader = sectionHeader.Value
            printfn "Список обязательной литературы пуст, добавляю 'Не требуется'." 
            removeNextParagraphIfEmpty body sectionHeader true
            let p = createParagraph "Не требуется" Stretch false false true
            sectionHeader.InsertAfterSelf p |> ignore
        else
            printfn "'Список обязательной литературы' не найден!" 

let fillOptionalLiterature (content: ProgramContent) (body: Body) =
    if content.["3.4.2. Список дополнительной литературы"].Trim() = "" then
        let sectionHeader = findParagraph body "Список дополнительной литературы"
        if sectionHeader.IsSome then
            let sectionHeader = sectionHeader.Value
            removeNextParagraphIfEmpty body sectionHeader true
            printfn "Список дополнительной литературы пуст, добавляю 'Не требуется'." 
            let p = createParagraph "Не требуется" Stretch false false true
            sectionHeader.InsertAfterSelf p |> ignore
        else
            printfn "'Список дополнительной литературы' не найден!" 

let addLiterature (document: WordprocessingDocument) (body: Body) =
    let sectionHeader = findParagraph body "Перечень иных информационных источников"
    if sectionHeader.IsSome then
        let sectionHeader = sectionHeader.Value

        let createNumberingPr (numberingId: Int32Value) =
            let ilvl = NumberingLevelReference(Val = Int32Value(0))
            let numId = NumberingId(Val = numberingId)
            let numPr = NumberingProperties(NumberingId = numId, NumberingLevelReference = ilvl)
            numPr

        let numberingId = createNumbering document

        let literature = 
            [ "Сайт Научной библиотеки им. М. Горького СПбГУ:", "http://www.library.spbu.ru/"
              "Электронный каталог Научной библиотеки им. М. Горького СПбГУ:", "http://www.library.spbu.ru/cgi-bin/irbis64r/cgiirbis_64.exe?C21COM=F&I21DBN=IBIS&P21DBN=IBIS" 
              "Перечень электронных ресурсов, находящихся в доступе СПбГУ:", "http://cufts.library.spbu.ru/CRDB/SPBGU/" 
              "Перечень ЭБС, на платформах которых представлены российские учебники, находящиеся в доступе СПбГУ:", "http://cufts.library.spbu.ru/CRDB/SPBGU/browse?name=rures&resource_type=8" ]
            |> List.rev

        printfn "Добавляю в перечень иных информационных источников стандартный список литературы." 

        literature 
        |> List.iter (fun (name, link) ->
            let nameParagraph = createParagraph name Left false false false
            let numPr = createNumberingPr numberingId
            nameParagraph.ParagraphProperties.AppendChild numPr |> ignore
            sectionHeader.InsertAfterSelf nameParagraph |> ignore
            let linkParagraph = createHyperlink link document
            nameParagraph.InsertAfterSelf linkParagraph |> ignore
            )
    else
        printfn "'Перечень иных информационных источников' не найден, литература не сгенерирована!" 

let fixStudent (content: ProgramContent) (body: Body) =
    if content |> noStudent |> List.isEmpty |> not then
        printfn "Пытаюсь заменить запрещённое слово 'студент' на 'обучающийся'"

        fixTextEverywhere body "студентами" "обучающимися"
        fixTextEverywhere body "студентам" "обучающимся"
        fixTextEverywhere body "студентах" "обучающихся"
        fixTextEverywhere body "студента" "обучающегося"
        fixTextEverywhere body "студенте" "обучающемся"
        fixTextEverywhere body "студентов" "обучающихся"
        fixTextEverywhere body "студентом" "обучающимся"
        fixTextEverywhere body "студенту" "обучающемуся"
        fixTextEverywhere body "студенты" "обучающиеся"
        fixTextEverywhere body "студент" "обучающийся"

        fixTextEverywhere body "Студентами" "Обучающимися"
        fixTextEverywhere body "Студентам" "Обучающимся"
        fixTextEverywhere body "Студентах" "Обучающихся"
        fixTextEverywhere body "Студента" "Обучающегося"
        fixTextEverywhere body "Студенте" "Обучающемся"
        fixTextEverywhere body "Студентов" "Обучающихся"
        fixTextEverywhere body "Студентом" "Обучающимся"
        fixTextEverywhere body "Студенту" "Обучающемуся"
        fixTextEverywhere body "Студенты" "Обучающиеся"
        fixTextEverywhere body "Студент" "Обучающийся"

let patchProgram (curriculumFile: string) (programFileName: string) =
    try
        use wordDocument = WordprocessingDocument.Open(programFileName, true)
        let content, errors = parseProgram wordDocument
        let curriculum = Curriculum curriculumFile
        let disciplineCode = FileInfo(programFileName).Name.Substring(0, 6)
        let body = wordDocument.MainDocumentPart.Document.Body

        printfn "Программа %s:\n" (FileInfo(programFileName).Name)

        addGovernment body
        
        addStPetersburg errors body

        fixTextInRun 
            body
            "Требования подготовленности обучающегося к освоению содержания учебных занятий ("
            "Требования к подготовленности обучающегося к освоению содержания учебных занятий ("

        fixTextInRun 
            body
            "Требования подготовленности обучающегося к освоению содержания учебных занятий (пререквизиты)"
            "Требования к подготовленности обучающегося к освоению содержания учебных занятий (пререквизиты)"

        addCompetencesToLearningOutcomes content body

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

        fillRequiredLiterature content body
        fillOptionalLiterature content body

        if content |> libraryLinksShallPresent |> List.isEmpty |> not then
            addLiterature wordDocument body

        fixStudent content body

        wordDocument.Save()

        printfn "\n"
    with
    | :? OpenXmlPackageException
    | :? InvalidDataException ->
        printfn "%s настолько коряв, что даже не читается, пропущен" programFileName
