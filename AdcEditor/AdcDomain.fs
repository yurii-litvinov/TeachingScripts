module AdcDomain

open canopy.classic
open canopy.types

open HardcodedData

type Filter =
    | NotStarted
    | InProgress
    | Done

type Tab =
    | Teachers
    | Rooms

let (!) path = 
    click (xpath path)

let waitForElementVisibleBySelector selector =
    waitFor (fun () -> (element selector).Displayed)

let waitForElementVisible (element: OpenQA.Selenium.IWebElement) =
    waitFor (fun () -> element.Displayed)

let tryUntilDone f =
    let mutable isDone = false
    while not isDone do
        try
            f ()
            isDone <- true
        with
        | _ -> System.Threading.Thread.Sleep 10

let logIn login password =
    start chrome
    !^ "https://adc.spbu.ru/"
    pin FullScreen
    (xpath "/html/body/div[1]/form/div[4]/div/div[3]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div/div/table/tbody/tr/td/div[2]/table[2]/tbody/tr[2]/td/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/input") << login
    (xpath "/html/body/div[1]/form/div[4]/div/div[3]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div/div/table/tbody/tr/td/div[2]/table[2]/tbody/tr[2]/td/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/input") << password
    click "#Logon_PopupActions_Menu_DXI0_T"

let getFilterStateFromTable row =
    sleep 1
    if (someElement "td[onclick*='GVScheduleCommand'] + td + td + td + td + td + td + td + td").IsNone then 
        InProgress
    else 
        let filterState = (elements "td[onclick*='GVScheduleCommand'] + td + td + td + td + td + td + td + td" |> Seq.skip (row - 1) |> Seq.head).Text
        match filterState with 
        | "не обработана" -> NotStarted
        | "в работе" -> InProgress
        | "обработана" -> Done
        | _ -> failwith (sprintf "Unknown filter state: %s" filterState)

let switchFilter filter =
    let filterState = getFilterStateFromTable 1
    if filterState <> filter then
        sleep 1
        ! "/html/body/form/div[4]/div[5]/div[3]/div[2]/div[2]/div/div[1]/table/tbody/tr/td/table[1]/tbody/tr[1]/td[9]/table/tbody/tr/td[2]/img[1]"
        sleep 1
        let filtersTable = element (xpath "/html/body/form/div[4]/div[5]/div[3]/div[2]/div[2]/div/div[1]/table/tbody/tr/td/div[3]/div/div/div[1]/div/table[2]/tbody/tr/td/div[2]/div/table[2]")
        let filterString = 
            match filter with
            | NotStarted -> "не обработана"
            | InProgress -> "в работе"
            | Done -> "обработана"
        filtersTable |> elementWithin (text filterString) |> click
        sleep 1
        // Click on "Модуль" to close filter window if it was not closed yet
        ! "/html/body/form/div[4]/div[5]/div[3]/div[1]/table/tbody/tr/td[3]/table/tbody/tr/td[2]/div/span/span[2]"

let switchTab tab =
    match tab with
    | Teachers -> click (text "Преподаватели")
    | Rooms -> click (text "Помещения для вида работы")

let openRecord num =
    ! $"/html/body/form/div[4]/div[5]/div[3]/div[2]/div[2]/div/div[1]/table/tbody/tr/td/table[1]/tbody/tr[{num + 1}]/td[9]"

let addRooms rooms workTypes =
    switchTab Rooms
    let rooms = rooms |> Seq.map string

    for room in rooms do
        for i in [2..(workTypes + 1)] do
            sleep 2
            ! "/html/body/form/div[4]/div[5]/div[3]/div[2]/div[2]/div/div/div/div/div/div[3]/div/div/div/div/div[1]/table/tbody/tr/td[2]/div/div/div/div/div[1]/ul/li[1]/a/img"
            sleep 1
            ! "/html/body/form/div[4]/div[2]/div[2]/div[2]/div/div/table[1]/tbody/tr[2]/td/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div/table[1]/tbody/tr/td[2]/table/tbody/tr/td[3]/img"
            sleep 1
            ! $"/html/body/form/div[4]/div[2]/div[2]/div[2]/div/div/table[1]/tbody/tr[2]/td/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div/table[1]/tbody/tr/td[2]/div/div/div/div/table/tbody/tr/td/div/div/table[2]/tr[{i}]/td"
            sleep 1
            ! "/html/body/form/div[4]/div[2]/div[2]/div[2]/div/div/table[1]/tbody/tr[2]/td/div[3]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input"
            (xpath "/html/body/form/div[4]/div[2]/div[2]/div[2]/div/div/table[1]/tbody/tr[2]/td/div[3]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input") << room
            sleep 2
            if room = "405" then
                // There are two rooms with number 405. One in Smolny, second one is correct
                click "table[id*='Room_Edit_DDD_gv_DXMainTable'] > tbody > tr + tr + tr > td"
            else
                click "table[id*='Room_Edit_DDD_gv_DXMainTable'] > tbody > tr + tr > td"
            sleep 2
            // Click "OK"
            click "li > a[id*='Dialog_SAC_Menu_DXI0']"

let isChecked checkboxName =
    (element $"span input[name*='{checkboxName}']").GetAttribute("value") = "C"

let check checkboxName =
    click $"table[id*='{checkboxName}'] > tbody > tr > td > span"

let addTeacher teacher workTypes =
    for workType in workTypes do
        click "#viewSite li[title^='Добавить'] img"
        sleep 2
        let mutable retries = 0
        while retries < 3 do
            try
                // Выпадающий список "Вид работ в модуле"
                click "input[onchange*='StudyModuleWorkKind_Edit_dropdown']"
                sleep 1
                // Получившаяся по клику на выпадающий список таблица
                let workTypeDropDown = element (xpath "/html/body/form/div[4]/div[2]/div[2]/div[2]/div/div/table[1]/tbody/tr[2]/td/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div/table[1]/tbody/tr/td[2]/div/div/div/div/table/tbody/tr/td/div/div/table[2]")
                waitForElementVisible workTypeDropDown
                workTypeDropDown |> elementWithin (text (fst workType)) |> click
                retries <- 3
            with 
            | _ -> 
                retries <- retries + 1
        try
            click "#viewSite input[spellcheck=\"false\"]"
            "#viewSite input[spellcheck=\"false\"]" << teacher

            let surname = (teacher.Split [|' '|]).[0]
            let fathersName = (teacher.Split [|' '|]).[2]
            waitForElementVisibleBySelector (xpath $"//*[text() = '{surname}']")
            // Default Лебедева Анастасия Владимировна is librarian, we need mathematician
            if teacher <> "Лебедева Анастасия Владимировна" then
                click (xpath $"//*[text() = '{fathersName}']")
            else
                click (xpath $"(//*[text() = 'Владимировна'])[2]")
        
            if not (irrelevantEducation.Contains teacher) then
                check "IsEducationLevelMatch_Edit"

            sleep 2
            if isChecked "HasWorkplaceInquiry_Edit" && not (isChecked "PracticalExperience_Edit") && not (irrelevantIndustrialExperience.Contains teacher) then
                check "HasPracticalExperience_Edit"

            (xpath "/html/body/form/div[4]/div[2]/div[2]/div[2]/div/div/table[1]/tbody/tr[2]/td/table[2]/tbody/tr/td/div[1]/div/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[1]/input") << string (snd workType)
        with
        | _ -> ()

        press enter

let wipeOutTeachers () =
    switchTab Teachers

    let length () = 
        sleep 1
        elements "img[title='Отменить нагрузку']" |> Seq.length

    let mutable quickPassCounter = (length ()) - 3

    while quickPassCounter > 0 do
        tryUntilDone (fun () -> 
            let eraser = elements "img[title='Отменить нагрузку']" |> Seq.skip (quickPassCounter / 2) |> Seq.head
            click eraser
        )
        quickPassCounter <- quickPassCounter - 1

    let mutable lastLength = length ()
    let mutable isDone = false
    let mutable offset = 0
    while not isDone do
        while ((length ()) = lastLength && offset < lastLength) do
            tryUntilDone (fun () -> 
                let eraser = elements "img[title='Отменить нагрузку']" |> Seq.skip offset |> Seq.head
                click eraser
            )
            if length () = lastLength then 
                offset <- offset + 1
        lastLength <- length ()
        isDone <- offset >= lastLength

let wipeOutRooms () =
    switchTab Rooms
    sleep 1
    if (elements "span[title='Состояние выбранных строк на всех страницах']" |> Seq.length) > 1 then
        sleep 1
        elements "span[title='Состояние выбранных строк на всех страницах']" |> Seq.skip 1 |> Seq.head |> click
        sleep 1
        element "img[src*='Action_Delete']" |> click
        sleep 1

let removeWorkTypes offset count =
    switchTab Teachers
    sleep 1
    for _ in [0..count - 1] do
        let eraser = elements "img[title='Отменить нагрузку']" |> Seq.skip offset |> Seq.head
        click eraser
        sleep 2

let refresh () =
    sleep 1
    click "img[src*='Action_Refresh']"
    sleep 2

let correctTeachersData (workTypes: Map<string, int>) =
    switchTab Teachers
    let mutable offset = 0
    let length = elements "img[title='Отменить нагрузку']" |> Seq.length
    while (offset < length) do
        tryUntilDone (fun () -> 
            let edit = elements "img[title='Правка']" |> Seq.skip offset |> Seq.head
            click edit
        )
        
        let workKind = (element "input[onchange*='StudyModuleWorkKind_Edit_dropdown']").GetAttribute("value")
        tryUntilDone (fun () -> "input[onfocus*='HoursPlan_Edit']" << string workTypes.[workKind])

        if not (isChecked "IsEducationLevelMatch_Edit") then
            check "IsEducationLevelMatch_Edit"

        if isChecked "HasWorkplaceInquiry_Edit" && not (isChecked "PracticalExperience_Edit") then
            check "HasPracticalExperience_Edit"

        press enter
        offset <- offset + 1

let getDisciplineNameFromTable row =
    (elements "td[onclick*='GVScheduleCommand'] + td + td + td + td + td + td" |> Seq.skip (row - 1) |> Seq.head).Text

let getSemesterFromTable row =
    let semester = (elements "td[onclick*='GVScheduleCommand'] + td + td + td" |> Seq.skip (row - 1) |> Seq.head).Text
    semester.Replace("Семестр ", "") |> int

let getRecordCaption () =
    (element "span .MainMenuTruncateCaption + span").GetAttribute "title"

let backToTable () =
    click "a[href*='StudyModule_ListView']"
    sleep 3

let tableSize () =
    ((elements "table[id*='MainTable'] > tbody > tr") |> Seq.length) - 1
