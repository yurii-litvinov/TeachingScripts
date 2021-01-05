module AdcFunctions

open canopy.classic
open canopy.types

type Filter =
    | NotStarted
    | InProgress
    | Done

type Tab =
    | Teachers
    | Rooms

let wait path = waitForElement (xpath path)
let (!) path = 
    wait path
    click (xpath path)

let logIn login password =
    start chrome
    !^ "https://adc.spbu.ru/"
    pin FullScreen
    (xpath "/html/body/div[1]/form/div[4]/div/div[3]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div/div/table/tbody/tr/td/div[2]/table[2]/tbody/tr[2]/td/div[1]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/input") << login
    (xpath "/html/body/div[1]/form/div[4]/div/div[3]/table/tbody/tr[2]/td/table/tbody/tr/td/div[1]/div/div/table/tbody/tr/td/div[2]/table[2]/tbody/tr[2]/td/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/input") << password
    click "#Logon_PopupActions_Menu_DXI0_T"

let switchFilter filter =
    ! "/html/body/form/div[4]/div[5]/div[3]/div[2]/div[2]/div/div[1]/table/tbody/tr/td/table[1]/tbody/tr[1]/td[9]/table/tbody/tr/td[2]/img[1]"
    let filtersTable = element (xpath "/html/body/form/div[4]/div[5]/div[3]/div[2]/div[2]/div/div[1]/table/tbody/tr/td/div[3]/div/div/div[1]/div/table[2]/tbody/tr/td/div[2]/div/table[2]")
    sleep 1
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
            ! "/html/body/form/div[4]/div[2]/div[2]/div[2]/div/div/table[1]/tbody/tr[2]/td/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div/table[1]/tbody/tr/td[2]/table/tbody/tr/td[3]/img"
            sleep 1
            ! $"/html/body/form/div[4]/div[2]/div[2]/div[2]/div/div/table[1]/tbody/tr[2]/td/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div/table[1]/tbody/tr/td[2]/div/div/div/div/table/tbody/tr/td/div/div/table[2]/tr[{i}]/td"
            sleep 1
            ! "/html/body/form/div[4]/div[2]/div[2]/div[2]/div/div/table[1]/tbody/tr[2]/td/div[3]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input"
            (xpath "/html/body/form/div[4]/div[2]/div[2]/div[2]/div/div/table[1]/tbody/tr[2]/td/div[3]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input") << room
            sleep 5
            press enter

let addTeacher teacher workTypes hours =
    let workTypes = List.zip workTypes hours

    for workType in workTypes do
        click "#viewSite li[title^='Добавить'] img"
        // Выпадающий список "Вид работ в модуле"
        ! "/html/body/form/div[4]/div[2]/div[2]/div[2]/div/div/table[1]/tbody/tr[2]/td/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div/table[1]/tbody/tr/td[2]/table/tbody/tr/td[2]/input"
        sleep 1
        // Получившаяся по клику на выпадающий список таблица
        let workTypeDropDown = element (xpath "/html/body/form/div[4]/div[2]/div[2]/div[2]/div/div/table[1]/tbody/tr[2]/td/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div/table[1]/tbody/tr/td[2]/div/div/div/div/table/tbody/tr/td/div/div/table[2]")
        workTypeDropDown |> elementWithin (text (fst workType)) |> click
        click "#viewSite input[spellcheck=\"false\"]"
        (element "#viewSite input[spellcheck=\"false\"]") << teacher
        sleep 3
        let surname = (teacher.Split [|' '|]).[0]
        click (xpath $"//*[text() = '{surname}']")
        click "#viewSite .WebEditorCell[title^='\"Образование\"'] > table > tbody > tr> td > span"
        (xpath "/html/body/form/div[4]/div[2]/div[2]/div[2]/div/div/table[1]/tbody/tr[2]/td/table[2]/tbody/tr/td/div[1]/div/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[1]/input") << string (snd workType)
        press enter

let wipeOutTeachers () =
    switchTab Teachers
    let length () = elements "img[title='Отменить нагрузку']" |> Seq.length
    let mutable lastLength = length ()
    let mutable isDone = false
    let mutable offset = 0
    while not isDone do
        while ((length ()) = lastLength && offset < lastLength) do
            let eraser = elements "img[title='Отменить нагрузку']" |> Seq.skip offset |> Seq.head
            click eraser
            sleep 2
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

let addTeachers teachers workTypes hours =
    switchTab Teachers
    for teacher in teachers do
        addTeacher teacher workTypes hours

let addTypicalRecordWithoutOpening teachers workTypes hours rooms =
    addTeachers 
        teachers
        workTypes
        hours

    sleep 1

    click "img[src*='Action_Refresh']"

    sleep 2

    removeWorkTypes 0 (workTypes |> Seq.length)

    addRooms rooms (workTypes |> Seq.length)

let addTypicalRecord recordNumber teachers workTypes hours rooms =
    openRecord recordNumber
    addTypicalRecordWithoutOpening teachers workTypes hours rooms

let processTypicalRecord recordNumber teachers workTypes hours rooms =
    openRecord recordNumber

    wipeOutTeachers ()
    wipeOutRooms ()

    addTypicalRecordWithoutOpening teachers workTypes hours rooms
