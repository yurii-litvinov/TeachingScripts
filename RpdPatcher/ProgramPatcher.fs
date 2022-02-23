module ProgramPatcher

open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing
open System.IO
open DocumentFormat.OpenXml
open ProgramContentChecker
open CurriculumUtils
open LessPainfulDocx

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
                        let description = $"{competence.Description.[0]}".ToLower() + competence.Description.Substring(1).TrimEnd('.')
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
        
            let p = createParagraph "Для каждой компетенции применяется линейная шкала оценивания, определяемая долей успешно выполненных заданий, проверяющих данную компетенцию." Stretch false false true
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
    if content.ContainsKey "3.4.1. Список обязательной литературы" 
        && content.["3.4.1. Список обязательной литературы"].Trim() = "" then
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
    if content.ContainsKey "3.4.2. Список дополнительной литературы"
        && content.["3.4.2. Список дополнительной литературы"].Trim() = "" then
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
            nameParagraph.ParagraphProperties.SpacingBetweenLines <- SpacingBetweenLines(After = StringValue("0"))
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

let fixDoubleSpaces (body: Body) =
    let rec fixDoubleSpacesInText (text: string) (result: bool)=
        let newText = text.Replace("  ", " ")
        if newText = text then
            (text, result)
        else
            fixDoubleSpacesInText newText true

    body.Descendants<Text>()
    |> Seq.toList
    |> List.filter (fun t -> t.Text.Contains "Р А Б О Ч А Я" |> not)
    |> List.map (fun t -> 
        let text, wasReplaced = fixDoubleSpacesInText t.Text false
        if wasReplaced then
            t.Text <- text
        wasReplaced)
    |> List.contains true
    |> (fun result -> if result then printfn "Избавляюсь от двойных пробелов.")

let fixRunFonts (body: Body) =
    body.Descendants<Run>()
    |> Seq.toList
    |> List.map (fun r -> 
        let fonts = 
            RunFonts
                (Ascii = StringValue("Times New Roman"), 
                    HighAnsi = StringValue("Times New Roman"), 
                    ComplexScript = StringValue("Times New Roman"))

        if r.RunProperties = null then
            r.RunProperties <- RunProperties()

        if r.RunProperties.RunFonts = null || r.RunProperties.FontSize = null then
            r.RunProperties.RunFonts <- fonts
            r.RunProperties.FontSize <- FontSize(Val = StringValue("24"))
            true
        else
            false)
    |> List.contains true
    |> (fun result -> if result then printfn "Выставляю шрифт в Times New Roman.")

let patchProgram (curriculumFile: string) (programFileName: string) =
    try
        use wordDocument = WordprocessingDocument.Open(programFileName, true)
        let content, errors = parseProgram wordDocument
        let curriculum = Curriculum curriculumFile
        let disciplineCode = FileInfo(programFileName).Name.Substring(0, 6)
        let body = wordDocument.MainDocumentPart.Document.Body

        printfn "Программа %s:\n" (FileInfo(programFileName).Name)

        addGovernment body

        fixDoubleSpaces body
        
        fixTextInRun 
            body
            "Р А Б О Ч А Я П Р О Г Р А М "
            "Р А Б О Ч А Я   П Р О Г Р А М "

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

        fixRunFonts body

        wordDocument.Save()

        printfn "\n"
    with
    | :? OpenXmlPackageException
    | :? InvalidDataException ->
        printfn "%s настолько коряв, что даже не читается, пропущен" programFileName

let patchYear (year: string) (programFileName: string) =
    try
        use wordDocument = WordprocessingDocument.Open(programFileName, true)
        let texts = wordDocument.MainDocumentPart.Document.Body.Descendants<Text>()
        let pair = texts |> Seq.pairwise |> Seq.tryFind (fun (r1, _) -> r1.Text = "Санкт-Петербург")
        match pair with
        | Some (_, yearText) -> 
            if yearText.Text.Length = 4 then
                printfn "Меняю год в '%s' на %s" programFileName year
                yearText.Text <- year
                wordDocument.Save()
            elif yearText.Text.Length = 2 then
                let nextRun = yearText.Parent.ElementsAfter() |> Seq.tryHead
                if nextRun.IsSome && nextRun.Value.InnerText.Length = 2 then
                    printfn "Меняю год в '%s' на %s" programFileName year
                    nextRun.Value.Remove()
                    yearText.Text <- year
                    wordDocument.Save()
                else
                    printfn "Неверный формат титульника, год в '%s' не изменён!" programFileName 
            else 
                printfn "Неверный формат титульника, год в '%s' не изменён!" programFileName 
        | _ -> 
            printfn "Неверный формат титульника, год в '%s' не изменён!" programFileName 

        printfn "\n"
    with
    | :? OpenXmlPackageException
    | :? InvalidDataException ->
        printfn "%s настолько коряв, что даже не читается, пропущен" programFileName

let removeCompetences (programFileName: string) =
    try
        use wordDocument = WordprocessingDocument.Open(programFileName, true)
        let paragraphs = wordDocument.MainDocumentPart.Document.Body.Descendants<Paragraph>()
        let firstCompetenceParagraph = paragraphs |> Seq.tryFind (fun p -> p.InnerText = "Компетенции, впервые формируемые дисциплиной:")

        match firstCompetenceParagraph with
        | Some firstCompetenceParagraph ->
            let lastCompetenceParagraph = 
                paragraphs 
                |> Seq.find (fun p -> 
                    p.InnerText.StartsWith "Для каждой компетенции применяется линейная шкала оценивания, определяемая долей успешно выполненных заданий, проверяющих данную компетенцию")

            let finishingParagraph = 
                paragraphs 
                |> Seq.find (fun p -> 
                    p.InnerText.StartsWith "Проверяемые компетенции:")

        
            let next = finishingParagraph.ElementsAfter() |> Seq.head 
            if next.InnerText.StartsWith "Сформированность компетенций считается пропорционально доле успешных ответов на вопросы и доле выполненных заданий" |> not then
                printfn "Файл '%s' не преобразован, неправильная структура ФОС!" programFileName
            else
                printfn "Удаляю компетенции из '%s'." programFileName

                finishingParagraph.Remove()
                next.Remove()

                let paragraphs =  firstCompetenceParagraph.ElementsAfter () |> Seq.filter (fun e -> e :? Paragraph) |> Seq.cast<Paragraph> |> Seq.toList
                firstCompetenceParagraph.Remove()

                let rec removeParagraph paragraphs =
                    match paragraphs with
                    | h :: _ when h = lastCompetenceParagraph ->
                        h.Remove ()
                    | h :: t -> 
                        h.Remove ()
                        removeParagraph t
                    | _ -> ()
            
                removeParagraph paragraphs

            wordDocument.Save()
        | None -> printfn "Файл '%s' не преобразован, там и так нет компетенций (или не могу найти)." programFileName

        printfn "\n"
    with
    | :? OpenXmlPackageException
    | :? InvalidDataException ->
        printfn "%s настолько коряв, что даже не читается, пропущен" programFileName

let addIndicatorsTable (curriculumFile: string) (programFileName: string) =
    try
        use wordDocument = WordprocessingDocument.Open(programFileName, true)
        let body = wordDocument.MainDocumentPart.Document.Body
        let learningOutcomesParagraph = findParagraph body "Перечень результатов обучения (learning outcomes)"
        if learningOutcomesParagraph.IsSome then

            let curriculum = Curriculum curriculumFile
            let disciplineCode = FileInfo(programFileName).Name.Substring(0, 6)
            let discipline = curriculum.Disciplines |> Seq.find (fun d -> d.Code = disciplineCode)

            let competenceCategory (code: string) =
                // Competence categories below are specific to 02.03.03 Federal Standart
                if curriculumFile.Contains "5162" |> not then ""
                elif code.StartsWith "ПКА" then "Профессиональные компетенции (академические)"
                elif code.StartsWith "ПКП" then "Профессиональные компетенции (практические)"
                elif code = "ОПК-1" || code = "ОПК-2" then "Теоретические и практические основы профессиональной деятельности"
                elif code = "ОПК-3" || code = "ОПК-4" || code = "ОПК-5" || code = "ОПК-6" then "Информационно-коммуникационные технологии для профессиональной деятельности"
                elif code = "УК-1" then "Системное и критическое мышление"
                elif code = "УК-2" then "Разработка и реализация проектов"
                elif code = "УК-3" then "Командная работа и лидерство"
                elif code = "УК-4" then "Коммуникация"
                elif code = "УК-5" then "Межкультурное взаимодействие"
                elif code = "УК-6" || code = "УК-7" then "Самоорганизация и саморазвитие (в том числе здоровьесбережение)"
                elif code = "УК-8" then "Безопасность жизнедеятельности"
                elif code = "УК-9" then "Экономическая культура, в том числе финансовая грамотность"
                elif code = "УК-10" then "Гражданская позиция"
                elif code = "УК-11" then ""
                else ""

            let competenceIndicators (code: string) =
                if curriculumFile.Contains "5162" |> not then ""
                elif code = "УК-1" then "УК 1.1. Анализирует задачу, выделяя ее базовые составляющие; 
                    УК 1.2. Определяет информацию, необходимую для решения поставленной задачи; 
                    УК 1.3. Осуществляет по различным запросам поиск информации, необходимой для решения поставленной задачи; 
                    УК 1.4. Оценивает достоинства, недостатки и последствия вариантов решения поставленных задач; 
                    УК 1.5. Грамотно, логично, аргументированно формирует собственные суждения, решения и оценки."
                elif code = "УК-2" then "УК-2.1. Определяет круг задач в рамках поставленной цели; 
                    УК-2.2. Предлагает способы решения поставленных задач; 
                    УК-2.3. Оценивает соответствие способов решения цели проекта; 
                    УК-2.4. Планирует реализацию задач в зоне своей ответственности с учетом имеющихся ресурсов и ограничений, действующих правовых норм; 
                    УК-2.5. Выполняет задачи в зоне своей ответственности в соответствии с запланированными результатами и точками контроля."
                elif code = "УК-3" then "УК-3.1. Определяет свою роль в социальном взаимодействии и командной работе исходя из стратегии сотрудничества для достижения поставленной цели;  
                    УК-3.2. При реализации своей роли в социальном взаимодействии и командной работе учитывает особенности поведения и интересы других участников;  
                    УК-3.3. Строит продуктивное взаимодействие с учетом возможных последствий личных действий в социальном взаимодействии и командной работе;  
                    УК-3.4. Осуществляет обмен информацией, знаниями и опытом с членами команды; 
                    УК-3.5. Соблюдает нормы и установленные правила командной работы."
                elif code = "УК-4" then "УК-4.1. Выбирает стиль общения на русском языке в зависимости от цели и условий партнерства;  
                    УК-4.2 Адаптирует речь, стиль общения и язык жестов к ситуациям взаимодействия; 
                    УК-4.3. Ведет деловую переписку на русском языке с учетом особенностей стилистики официальных и неофициальных писем;  
                    УК-4.4. Ведет деловую переписку на иностранном языке с учетом особенностей стилистики официальных писем и социокультурных различий; 
                    УК-4.5. Выполняет для личных целей перевод официальных и профессиональных текстов с иностранного языка на русский, с русского языка на иностранный."
                elif code = "УК-5" then "УК-5.1. Знает философские, этические, исторические, религиозные предпосылки культурного разнообразия. 
                    УК-5.2. Владеет навыками философского, исторического, религиоведческого анализа явлений культуры. 
                    УК-5.3. Формулирует собственную этическую позицию в обстоятельствах межкультурного взаимодействия."
                elif code = "УК-6" then "УК-6.1. Применяет приемы управления своим временем; 
                    УК-6.2. Применяет приемы целеполагания и планирования для выстраивания траектории саморазвития; 
                    УК-6.3. Выстраивает траекторию саморазвития на основе принципов образования."
                elif code = "УК-7" then "УК-7.1. Выбирает здоровьесберегающие технологии для поддержания здорового образа жизни с учетом физиологических особенностей организма; 
                    УК-7.2. Планирует свое рабочее и свободное время для оптимального сочетания физической и умственной нагрузки и обеспечения работоспособности; 
                    УК-7.3. Соблюдает и пропагандирует нормы здорового образа жизни в различных жизненных ситуациях и в профессиональной деятельности."
                elif code = "УК-8" then "УК-8.1. Анализирует факторы вредного влияния на жизнедеятельность элементов среды обитания (технических средств, технологических процессов, материалов, зданий и сооружений, природных и социальных явлений);  
                    УК-8.2. Идентифицирует опасные и вредные факторы в рамках осуществляемой деятельности;  
                    УК-8.3. Выявляет проблемы, связанные с нарушениями техники безопасности на рабочем месте;  
                    УК-8.4. Предлагает мероприятия по предотвращению чрезвычайных ситуаций;  
                    УК-8.5. Разъясняет правила поведения при возникновении чрезвычайных ситуаций природного и техногенного происхождения;  
                    УК-8.6. Оказывает первую помощь, описывает способы участия в восстановительных мероприятиях."
                elif code = "УК-9" then "УК-9.1. Знает базовые принципы функционирования экономики и экономического развития, цели и формы участия государства в экономике; 
                    УК-9.2. Применяет методы личного экономического и финансового планирования для достижения текущих и долгосрочных финансовых целей; 
                    УК-9.3. Использует финансовые инструменты для управления личными финансами (личным бюджетом); 
                    УК-9.4. Контролирует собственные экономические и финансовые риски."
                elif code = "УК-10" then "УК-10.1. Понимает значение основных правовых категорий, сущность коррупционного поведения, формы его проявления в различных сферах общественной жизни; 
                    УК-10.2. Демонстрирует знание российского законодательства; 
                    УК-10.3. Демонстрирует знание антикоррупционных стандартов поведения,  
                    УК-10.4. Демонстрирует уважение к праву и закону; 
                    УК-10.5. Идентифицирует и оценивает коррупционные риски; 
                    УК-10.6. Проявляет нетерпимое отношение к коррупционному поведению; 
                    УК-10.7. Умеет правильно анализировать, толковать и применять нормы права в различных сферах социальной деятельности, а также в сфере противодействия коррупции; 
                    УК-10.8. Осуществляет социальную и  профессиональную деятельность на основе развитого правосознания и сформированной правовой культуры."
                elif code = "УКБ-1" then "1.1. Определяет круг задач в рамках поставленной цели; 
                    1.2. Предлагает способы решения поставленных задач; 
                    1.3. Оценивает соответствие способов решения цели проекта; 
                    1.4. Планирует реализацию задач в зоне своей ответственности с учетом имеющихся ресурсов и ограничений, действующих правовых норм; 
                    1.5. Выполняет задачи в зоне своей ответственности в соответствии с запланированными результатами и точками контроля; 
                    1.6. Представляет результаты проекта; 
                    1.7. Предлагает возможности использования результатов проекта и/или совершенствования."
                elif code = "УКБ-2" then "2.1. Определяет свою роль в социальном взаимодействии и командной работе, исходя из стратегии сотрудничества для достижения поставленной цели;  
                    2.2. При реализации своей роли в социальном взаимодействии и командной работе учитывает особенности поведения и интересы других участников;  
                    2.3. Осуществляет обмен информацией, знаниями и опытом с членами команды; 
                    2.4. Оценивает идеи других членов команды для достижения поставленной цели;  
                    2.5. Соблюдает нормы и установленные правила командной работы."
                elif code = "УКБ-3" then "3.1. Находит и использует различные источники информации.
                    3.2. Точно определяет тип и форму необходимой информации.
                    3.3. Получает информацию и сохраняет ее в удобном для работы формате.
                    3.4. Проверяет достоверность собранной информации."
                elif code = "УКБ-4" then "4.1. Выбирает стиль общения на русском языке в зависимости от цели и условий партнерства;  
                    4.2 Адаптирует речь, стиль общения и язык жестов к ситуациям взаимодействия; 
                    4.3. Ведет деловую переписку на русском языке с учетом особенностей стилистики официальных и неофициальных писем;  
                    4.4. Публично выступает на русском языке, строит свое выступление с учетом аудитории и цели общения; 
                    4.5. Устно представляет результаты своей деятельности на иностранном языке, может поддержать разговор в ходе их обсуждения."
                else ""

            let rows = 
                discipline.FormedCompetences @ discipline.FullyFormedCompetences @ discipline.ImprovedCompetences 
                |> Seq.sortBy (fun c -> c.Code)
                |> Seq.indexed
                |> Seq.map (fun (i, c) -> (i + 1, competenceCategory c.Code, sprintf "%s. %s" c.Code c.Description, competenceIndicators c.Code))
                |> Seq.map (fun (i, category, description, indicators) -> [string i; category; description; ""; indicators])
                |> Seq.toList

            let indicatorsTable = 
                createTable 
                    ([
                        [ "№" 
                          "Наименование категории (группы) компетенций" 
                          "Код и наименование компетенции" 
                          "Планируемые результаты обучения, обеспечивающие формирование компетенции"
                          "Код индикатора и индикатор достижения универсальной компетенции"
                        ]
                        [""; "1"; "2"; "3"; "4"]
                    ]
                    @ rows)

            learningOutcomesParagraph.Value.InsertAfterSelf indicatorsTable |> ignore
            wordDocument.Save()
        else
            printfn "Раздел 1.3 не найден, не могу добавить таблицу с индикаторами"

    with
    | :? OpenXmlPackageException
    | :? InvalidDataException ->
        printfn "%s настолько коряв, что даже не читается, пропущен" programFileName


