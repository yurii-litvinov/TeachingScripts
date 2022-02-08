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
    | Software
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
    tryUntilDone (fun () -> 
        match tab with
        | Teachers -> click (text "Преподаватели")
        | Software -> click (text "Программное обеспечение по видам работ")
        | Rooms -> click (text "Study Module Work Kind Rooms")
    )

let openRecord num =
    ! $"/html/body/form/div[4]/div[5]/div[3]/div[2]/div[2]/div/div[1]/table/tbody/tr/td/table[1]/tbody/tr[{num + 1}]/td[9]"
    sleep 1

let addRooms (rooms: Map<string, string list>) =
    switchTab Rooms
    
    rooms |> Map.iter (
        fun workType rooms ->
            for room in rooms do
                // Large green plus button
                click "img[id*='StudyModuleWorkKindRooms_ToolBar_Menu_DXI0_Img']"

                // Sometimes room selection doesn't work, will try several times
                while (element "input[id*='StudyModuleWorkKind_Edit_dropdown_DD_I']").GetAttribute("value") = "Н/Д" do
                    // Sometimes click on dropdown does not work for some reason, try several times
                    let mutable isDone = false
                    while not isDone do
                        // Work type selection dropdown
                        click "img[id*='StudyModuleWorkKind_Edit_dropdown_DD_B-1']"
                        sleep 1
                        isDone <- (someElement (xpath $"//tr[contains(@class, 'List')]/td[text() = '{workType}']")).IsSome
                    
                    // Work type
                    click (xpath $"//tr[contains(@class, 'List')]/td[text() = '{workType}']")
                    // Giving some time
                    sleep 1
                // Room
                click "input[id*='dviRoom_Edit_I']"
                // Giving some time
                sleep 1
                "input[id*='dviRoom_Edit_I']" << room
                // Wait for it to load room data
                sleep 2

                // Large auditories are ambiguous as hell so we need some involved special processing
                if room.Length = 1 then
                    // Click on address filter
                    click "td[id*='Room_Edit_DDD_gv_col0'] img[class*='Filter']"
                    // Let it load filter data
                    sleep 1
                    // Click on input field
                    click "input[id*='HFListBox0_LBFE']"
                    // Enter part of an address
                    "input[id*='HFListBox0_LBFE']" << "28"
                    // Let filter do its work
                    sleep 1
                    // Select address.
                    click "table[id*='Room_Edit_DDD_gv_HFListBox0_LBT'] > tr > td"
                    // Waiting for a filter to do its thing
                    sleep 1
                    // Now sifting through pages of filter output
                    while (someElement (xpath $"//tr[contains(@id,'Room_Edit_DDD_gv_DXDataRow')]/td[text() = '{room}']")).IsNone 
                        && (someElement "a[id*='Room_Edit_DDD_gv_DXPagerBottom'] > img[alt='Следующая']").IsSome do
                        // Click next page
                        click "a[id*='Room_Edit_DDD_gv_DXPagerBottom'] > img[alt='Следующая']"
                        // Let selection window load next page
                        sleep 1
                    // Hopefully we found what we were looking for. If not, script will stall here and user will be 
                    // aware that something is wrong.
                    click (xpath $"//tr[contains(@id,'Room_Edit_DDD_gv_DXDataRow')]/td[text() = '{room}']")
                
                elif room = "405" then
                    // There are many rooms with number 405, second one is correct (for now)
                    click "table[id*='Room_Edit_DDD_gv_DXMainTable'] > tbody > tr + tr + tr > td"
                elif room = "актовый зал правая" then
                    click "table[id*='Room_Edit_DDD_gv_DXMainTable'] > tbody > tr + tr + tr + tr + tr + tr > td"

                else
                    click "table[id*='Room_Edit_DDD_gv_DXMainTable'] > tbody > tr + tr > td"

                // Wait for ADC to load room data
                sleep 2
                // Click "OK"
                click "li > a[id*='Dialog_SAC_Menu_DXI0']"
    )

let addSoftware software workTypes =
    switchTab Software
    for soft in software do
        for i in [2..(workTypes + 1)] do
            sleep 2
            click "img[id*='WorkKindSoftwareProducts_ToolBar_Menu_DXI0_Img']"
            sleep 1
            click "img[id*='StudyModuleWorkKind_Edit_dropdown_DD_B-1Img']"
            sleep 1
            ! $"/html/body/form/div[4]/div[2]/div[2]/div[2]/div/div/table[1]/tbody/tr[2]/td/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div/table[1]/tbody/tr/td[2]/div/div/div/div/table/tbody/tr/td/div/div/table[2]/tr[{i}]/td"
            sleep 1
            click "input[id*='SoftwareProduct_Edit_I']"
            "input[id*='SoftwareProduct_Edit_I']" << soft
            sleep 2
            click (xpath $"//*[text() = '{soft}']")
            sleep 2
            // Click "OK"
            click "li > a[id*='Dialog_SAC_Menu_DXI0']"

let isChecked checkboxName =
    (element $"span input[name*='{checkboxName}']").GetAttribute("value") = "C"

let check checkboxName =
    while not (isChecked checkboxName) do
        click $"table[id*='{checkboxName}'] > tbody > tr > td > span"
        System.Threading.Thread.Sleep 50

let addTeacher teacher workTypes isRelevantExperience =
    for workType in workTypes do
        try
            click "#viewSite li[title^='Добавить'] img"
        with
        | _ ->
            click "li > a[id*='Dialog_SAC_Menu_DXI0']"
            sleep 2
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
            // Default Лебедева Анастасия Владимировна is librarian, we need mathematician. 
            // Same with Смирнов Михаил Николаевич, we need second.
            if teacher <> "Лебедева Анастасия Владимировна" && teacher <> "Смирнов Михаил Николаевич" then
                click (xpath $"//*[text() = '{fathersName}']")
            else
                click (xpath $"(//*[text() = '{fathersName}'])[2]")

            sleep 2
        
            if not (irrelevantEducation.Contains teacher) then
                check "IsEducationLevelMatch_Edit"

            sleep 2
            if isChecked "HasWorkplaceInquiry_Edit" && not (isChecked "PracticalExperience_Edit") && isRelevantExperience then
                check "HasPracticalExperience_Edit"

            sleep 1
            (xpath "/html/body/form/div[4]/div[2]/div[2]/div[2]/div/div/table[1]/tbody/tr[2]/td/table[2]/tbody/tr/td/div[1]/div/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[1]/input") << string (snd workType)
            sleep 1
        with
        | _ -> ()

        press enter
        sleep 2

let wipeOutTeachers () =
    switchTab Teachers
    click "tr[id*='HeadersRow'] span[id*='WorkKindTeachers']"
    sleep 2
    click "img[src*='Action_Clear']"
    sleep 1

let wipeOutRooms () =
    switchTab Rooms
    sleep 1
    let removalElements = elements "span[title='Состояние выбранных строк на всех страницах']"
    if (element "tr[id*='StudyModuleWorkKindRooms'] + tr").Text <> "Нет данных для отображения" then
        if (removalElements |> Seq.length) > 1 then
            removalElements |> Seq.skip 1 |> Seq.head |> click
        if (removalElements |> Seq.length) = 1 then
            click "span[title='Состояние выбранных строк на всех страницах']"
        sleep 1
        click "img[src*='Action_Delete']"
        sleep 1

let removeWorkTypes offset count =
    switchTab Teachers
    sleep 1
    let checkboxes = 
        (elements "table[id*='WorkKindTeachers'] span[id*='WorkKindTeachers']") 
        |> Seq.skip (1 + offset) |> Seq.take count |> Seq.toList
    
    checkboxes |> Seq.iter (fun e -> click e; sleep 1)
    click "img[src*='Action_Clear']"
    sleep 1

let refresh () =
    tryUntilDone (fun () -> click "img[src*='Action_Refresh']")
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
    sleep 1
    let semester = (elements "td[onclick*='GVScheduleCommand'] + td + td + td" |> Seq.skip (row - 1) |> Seq.head).Text
    semester.Replace("Семестр ", "") |> int

let getRecordCaption () =
    (element "span .MainMenuTruncateCaption + span").GetAttribute "title"

let backToTable () =
    click "a[href*='StudyModule_ListView']"
    sleep 4

let markAsDone () =
    tryUntilDone (fun () -> click "a[title*='Установить статус записи в \"Обработана\"']")
    sleep 1

let tableSize () =
    ((elements "table[id*='MainTable'] > tbody > tr") |> Seq.length) - 1
