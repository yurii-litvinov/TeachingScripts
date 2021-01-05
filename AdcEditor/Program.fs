module Main

open AdcFunctions

let login = ""
let password = ""

let englishRooms = [3503; 2509; 3507; 3512; 2503; 2510] 
let englishWorkTypes = ["Под руководством преподавателя"; "Практические занятия"; "Промежуточная аттестация (зач)"]
let rkiWorkTypes = ["Под руководством преподавателя"; "Практические занятия"; "В присутствии преподавателя"; "Промежуточная аттестация (зач)"]
let onlineWorkTypes = ["Консультации"; "Промежуточная аттестация (зач)"]
let rkiRoom = [3512]
let computerClasses = ["2406"; "2408"; "2412"; "2414"; "2444-1"; "2444-2"; "2446"]
let physicalTrainingWorkTypes = ["Практические занятия"; "Текущий контроль (ауд)"]

logIn login password
switchFilter InProgress

processTypicalRecord
    1
    ["Наливайко Роман Алексеевич"] 
    onlineWorkTypes
    [10; 2] 
    ("1" :: computerClasses)
