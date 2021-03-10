module ProgramContentChecker

open DocumentFormat.OpenXml.Packaging
open System.IO
open System.Text.RegularExpressions

type ProgramContent = Map<string, string>

type ParserState = 
    { Input: string
      Output: ProgramContent
      Errors: string list }

type ParserStep =
    { Prefix: string
      Action: string -> ProgramContent -> ProgramContent }

let nop = fun _ -> id

let accept str = { Prefix = str; Action = nop }

let parseAfter str action =
    { Prefix = str; Action = action str }

let skip = { Prefix = ""; Action = nop }

let parse (steps: ParserStep list) inputString =
    let rec relaxSectionTitleParsing (steps: ParserStep list) (acc: ParserStep list) =
        match steps with
        | [] -> List.rev acc
        | h :: t when h.Prefix.StartsWith("Раздел") -> 
            let partPrefix = h.Prefix.Substring(0, "Раздел 9".Length)
            let partName = h.Prefix.Split('.') |> Seq.last |> fun s -> s.TrimStart ()
            relaxSectionTitleParsing t (([accept partPrefix; skip; { Prefix = partName; Action = h.Action } ] |> List.rev) @ acc)
        | h :: t when Regex.Match(h.Prefix, @"^(\d\.)+\d?").Success ->
            let matches = Regex.Match(h.Prefix, @"^((\d\.)+\d?)((\w|\s|\(|\)|,|-)+)$")
            if matches.Success then
                let digitPart = matches.Groups.[1].Value.TrimEnd('.').Trim()
                let namePart = matches.Groups.[3].Value.Trim()
                relaxSectionTitleParsing t (([accept digitPart; skip; { Prefix = namePart; Action = h.Action } ] |> List.rev) @ acc)
            else 
                relaxSectionTitleParsing t (h :: acc)
        | h :: t -> relaxSectionTitleParsing t (h :: acc)

    let rec parseRec (steps: ParserStep list) (state: ParserState) =
        match steps with
        | [] -> state
        | h :: n :: t ->
            if state.Input.StartsWith h.Prefix then
                let newInput = state.Input.Substring(h.Prefix.Length).TrimStart()
                let suffixStart = newInput.IndexOf n.Prefix
                if suffixStart = -1 then
                    // Panic silently, next step will report error.
                    parseRec (n :: t) { state with Input = newInput }
                else
                    // Everything fine, cutting off parsed part.
                    let content = newInput.Substring(0, suffixStart).Trim()
                    let newInput = newInput.Substring(suffixStart).TrimStart()
                    parseRec (n :: t) { state with Input = newInput; Output = h.Action content state.Output }
            else
                // Trying to skip until next prefix.
                let recover = state.Input.IndexOf n.Prefix
                if recover = -1 then
                    // Failed. Skipping step, reporting our error and hoping that next step will be more successful in recovery.
                    parseRec (n :: t) { state with Errors = $"Not found {h.Prefix}, seen '{state.Input.Substring(0, 20)}...'" :: state.Errors }
                else
                    // Succeeded. Skipping input until next prefix.
                    let newInput = state.Input.Substring(recover)
                    parseRec (n :: t) { state with Input = newInput; Errors = $"Not found {h.Prefix}, seen '{state.Input.Substring(0, 30)}..." :: state.Errors }
        | h :: [] ->
            // Last step, parsed string is everything until end.
            if state.Input.StartsWith h.Prefix then
                let content = state.Input.Substring(h.Prefix.Length).Trim()
                { state with Input = ""; Output = h.Action content state.Output }
            else
                { state with Input = ""; Errors = $"Not found {h.Prefix}" :: state.Errors }

    let steps = relaxSectionTitleParsing steps []
    let state = parseRec steps {Input = inputString; Output = Map.empty; Errors = []}
    (state.Output, state.Errors |> List.rev)

let parseProgram (programFileName: string) : (ProgramContent * string list) =
    let wordDocument = WordprocessingDocument.Open(programFileName, false)
    let body = wordDocument.MainDocumentPart.Document.Body
    let text = body.InnerText.Trim ()

    let note field = fun s (c: Map<string, string>) -> c.Add(field, s)

    let parser = 
        [ accept "Правительство Российской Федерации" 
          accept "Санкт-Петербургский государственный университет" 
          parseAfter "Р А Б О Ч А Я   П Р О Г Р А М М АУЧЕБНОЙ ДИСЦИПЛИНЫ" note
          accept "Язык(и) обучения"
          skip
          parseAfter "Трудоемкость в зачетных единицах: " note
          parseAfter "Регистрационный номер рабочей программы: " note
          parseAfter "Санкт-Петербург" note
          accept "Раздел 1. Характеристики учебных занятий"
          parseAfter "1.1. Цели и задачи учебных занятий" note
          parseAfter "1.2. Требования к подготовленности обучающегося к освоению содержания учебных занятий (пререквизиты)" note
          parseAfter "1.3. Перечень результатов обучения (learning outcomes)" note
          parseAfter "1.4. Перечень и объём активных и интерактивных форм учебных занятий" note
          accept "Раздел 2. Организация, структура и содержание учебных занятий"
          accept "2.1. Организация учебных занятий" 
          accept "2.1.1. Основной курс" 
          skip
          parseAfter "2.2. Структура и содержание учебных занятий" note
          accept "Раздел 3. Обеспечение учебных занятий"
          accept "3.1. Методическое обеспечение"
          parseAfter "3.1.1. Методические указания по освоению дисциплины" note
          parseAfter "3.1.2. Методическое обеспечение самостоятельной работы" note 
          parseAfter "3.1.3. Методика проведения текущего контроля успеваемости и промежуточной аттестации и критерии оценивания" note
          parseAfter "3.1.4. Методические материалы для проведения текущего контроля успеваемости и промежуточной аттестации (контрольно-измерительные материалы, оценочные средства)" note
          parseAfter "3.1.5. Методические материалы для оценки обучающимися содержания и качества учебного процесса" note
          accept "3.2. Кадровое обеспечение"
          parseAfter "3.2.1. Образование и (или) квалификация штатных преподавателей и иных лиц, допущенных к проведению учебных занятий" note
          parseAfter "3.2.2. Обеспечение учебно-вспомогательным и (или) иным персоналом" note
          accept "3.3. Материально-техническое обеспечение"
          parseAfter "3.3.1. Характеристики аудиторий (помещений, мест) для проведения занятий" note
          parseAfter "3.3.2. Характеристики аудиторного оборудования, в том числе неспециализированного компьютерного оборудования и программного обеспечения общего пользования" note
          parseAfter "3.3.3. Характеристики специализированного оборудования" note
          parseAfter "3.3.4. Характеристики специализированного программного обеспечения" note
          parseAfter "3.3.5. Перечень и объёмы требуемых расходных материалов" note
          accept "3.4. Информационное обеспечение"
          parseAfter "3.4.1. Список обязательной литературы" note
          parseAfter "3.4.2. Список дополнительной литературы" note
          parseAfter "3.4.3. Перечень иных информационных источников" note
          parseAfter "Раздел 4. Разработчики программы" note ]

    parse parser text

let checkEveryFieldFilled (content: ProgramContent) =
    content
    |> Map.toSeq
    |> Seq.iter (fun (k, v) -> if v = "" then printfn "\tСекция '%s' пуста." k)

let noStudent (content: ProgramContent) =
    content
    |> Map.toSeq
    |> Seq.iter (fun (k, v) -> if v.Contains "студен" then printfn "\tСекция '%s' содержит запрещённое слово 'студент'." k)

let shallContainCompetences (content: ProgramContent) =
    if not (content.["1.3. Перечень результатов обучения (learning outcomes)"].Contains "компетенц") then
        printfn "\tПеречень результатов обучения не содержит указания на компетенции."

let attestationMethodologyShallBeBigEnough (content: ProgramContent) =
    let attestationMethodology = content.["3.1.3. Методика проведения текущего контроля успеваемости и промежуточной аттестации и критерии оценивания"].Trim()
    if attestationMethodology.Length < 300 then
        printfn "\tМетодика проведения аттестации подозрительно короткая: \n%s" attestationMethodology

let controlMaterialsShallBeBigEnough (content: ProgramContent) =
    let controlMaterials = content.["3.1.4. Методические материалы для проведения текущего контроля успеваемости и промежуточной аттестации (контрольно-измерительные материалы, оценочные средства)"].Trim()
    if controlMaterials.Length < 300 then
        printfn "\tКонтрольно-измерительные материалы подозрительно короткие: \n%s" controlMaterials

let controlMaterialsShallReferenceCompetences (content: ProgramContent) =
    let controlMaterials = content.["3.1.4. Методические материалы для проведения текущего контроля успеваемости и промежуточной аттестации (контрольно-измерительные материалы, оценочные средства)"].Trim()
    if not (controlMaterials.Contains "омпетенц") then
        printfn "\tКонтрольно-измерительные материалы не содержат указаний по оцениванию компетенций: \n%s" controlMaterials

let feedbackMethodShallBeStandard (content: ProgramContent) =
    let roomRequirements = content.["3.1.5. Методические материалы для оценки обучающимися содержания и качества учебного процесса"].Trim()
    if roomRequirements <> "Для оценки обучающимися содержания и качества учебного процесса применяется анкетирование в соответствии \
        с методикой и графиком, утвержденными в установленном порядке."
    then
        printfn "\tРаздел 3.1.5 не содержит стандартную фразу."

let roomRequirementsShallBeStandard (content: ProgramContent) =
    let roomRequirements = content.["3.3.1. Характеристики аудиторий (помещений, мест) для проведения занятий"].Trim()
    if roomRequirements <> "Учебные аудитории для проведения учебных занятий, \
        оснащенные стандартным оборудованием, используемым для обучения в СПбГУ \
        в соответствии с требованиями материально-технического обеспечения."
        && roomRequirements <> "Помещение, оснащенное специализированным оборудованием."
    then
        printfn "\tРаздел 3.3.1 не содержит стандартную фразу."

let softRequirementsShallBeStandard (content: ProgramContent) =
    let roomRequirements = content.["3.3.2. Характеристики аудиторного оборудования, в том числе неспециализированного компьютерного оборудования и программного обеспечения общего пользования"].Trim()
    if roomRequirements <> "Стандартное оборудование, используемое для обучения в СПбГУ.\
        MS Windows, MS Office, Mozilla FireFox, Google Chrome, Acrobat Reader DC, \
        WinZip, Антивирус Касперского."
    then
        printfn "\tРаздел 3.3.2 не содержит стандартную фразу."

let libraryLinksShallPresent (content: ProgramContent) =
    let otherSources = content.["3.4.3. Перечень иных информационных источников"].Trim()
    if not (otherSources.Contains "Сайт Научной библиотеки им. М. Горького СПбГУ: http://www.library.spbu.ru/"
        && otherSources.Contains "Электронный каталог Научной библиотеки им. М. Горького СПбГУ"
        && otherSources.Contains "http://www.library.spbu.ru/cgi-bin/irbis64r/cgiirbis_64.exe?C21COM=F&I21DBN=IBIS&P21DBN=IBIS"
        && otherSources.Contains "Перечень электронных ресурсов, находящихся в доступе СПбГУ:"
        && otherSources.Contains "http://cufts.library.spbu.ru/CRDB/SPBGU/"
        && otherSources.Contains "Перечень ЭБС, на платформах которых представлены российские учебники, находящиеся в доступе СПбГУ:"
        && otherSources.Contains "http://cufts.library.spbu.ru/CRDB/SPBGU/browse?name=rures&resource_type=8")
    then
        printfn "\tРаздел 3.4.3 не содержит ссылки на электронную библиотеку СПбГУ."

let checkProgram (programFileName: string) =
    printfn "%s:" (FileInfo(programFileName).Name)
    let content, errors = parseProgram programFileName
    if errors <> [] then
        printfn "\tОшибки разбора:"
        for error in errors do
            printfn "\t\t%s" error

    [ checkEveryFieldFilled
      noStudent
      shallContainCompetences
      attestationMethodologyShallBeBigEnough
      controlMaterialsShallBeBigEnough
      controlMaterialsShallReferenceCompetences
      feedbackMethodShallBeStandard
      roomRequirementsShallBeStandard
      softRequirementsShallBeStandard 
      libraryLinksShallPresent ] 
    |> List.iter (fun f -> f content)

    printfn ""
    printfn ""
