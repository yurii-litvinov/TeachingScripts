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
    | Teachers -> ! "/html/body/form/div[4]/div[5]/div[3]/div[2]/div[2]/div/div/div/div/ul/li[2]/div/ul/li[2]/a/span"
    | Rooms -> ! "/html/body/form/div[4]/div[5]/div[3]/div[2]/div[2]/div/div/div/div/ul/li[2]/div/ul/li[7]/a/span"

let openRecord num =
    ! $"/html/body/form/div[4]/div[5]/div[3]/div[2]/div[2]/div/div[1]/table/tbody/tr/td/table[1]/tbody/tr[{num + 1}]/td[9]"

let addRooms rooms =
    switchTab Rooms
    let rooms = rooms |> Seq.map string

    for room in rooms do
        for i in [2..4] do
            ! "/html/body/form/div[4]/div[5]/div[3]/div[2]/div[2]/div/div/div/div/div/div[3]/div/div/div/div/div[1]/table/tbody/tr/td[2]/div/div/div/div/div[1]/ul/li[1]/a/img"
            ! "/html/body/form/div[4]/div[2]/div[2]/div[2]/div/div/table[1]/tbody/tr[2]/td/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div/table[1]/tbody/tr/td[2]/table/tbody/tr/td[3]/img"
            sleep 1
            ! $"/html/body/form/div[4]/div[2]/div[2]/div[2]/div/div/table[1]/tbody/tr[2]/td/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div/table[1]/tbody/tr/td[2]/div/div/div/div/table/tbody/tr/td/div/div/table[2]/tr[{i}]/td"
            sleep 1
            ! "/html/body/form/div[4]/div[2]/div[2]/div[2]/div/div/table[1]/tbody/tr[2]/td/div[3]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input"
            (xpath "/html/body/form/div[4]/div[2]/div[2]/div[2]/div/div/table[1]/tbody/tr[2]/td/div[3]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input") << room
            sleep 3
            press enter

let addTeacher teacher hours =
    let workTypes = ["Под руководством преподавателя"; "Практические занятия"; "Промежуточная аттестация (зач)"]
    let workTypes = List.zip workTypes hours

    for workType in workTypes do
        click "#viewSite li[title^='Добавить'] img"
        ! "/html/body/form/div[4]/div[2]/div[2]/div[2]/div/div/table[1]/tbody/tr[2]/td/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td/div/div/table[1]/tbody/tr/td[2]/table/tbody/tr/td[2]/input"
        sleep 1
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

let wipeOutTeachers offset =
    switchTab Teachers
    while elements "img[title='Отменить нагрузку']" |> Seq.length > 3 do
        let eraser = elements "img[title='Отменить нагрузку']" |> Seq.skip offset |> Seq.head
        click eraser
        sleep 2

let addTeachers teachers hours =
    switchTab Teachers
    for teacher in teachers do
        addTeacher teacher hours
