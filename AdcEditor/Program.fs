module Main

open AdcDomain
open AdcOperational
open CurriculumUtils
open WorkDistributionData
open HardcodedData
open VkrAdvisorsParser
open OfficialWorkDistributionParser
open WorkDistributionTypes
open EnglishTeachers
open Config
open RoomData
open CurriculumParser

[<EntryPoint>]
let main _ =

    //logIn ()

    //checkRpds "../../5006"
    //createSoftwareReport "../../5006"

    //doEverythingRight ()

    //doMagic InProgress

    //checkWorkDistribution 8

    //let teachers = EnglishTeachers(Config.workPlan)
    //printfn "%A" (teachers.Teachers 1 "Траектория 3")

    //let roomData = RoomData(Config.roomDataSheetId)
    //printfn "%A" (roomData.Rooms 2 "Программирование")

    // autoCorrectRoomsForFirst ()

    autoAddRooms ["3509"; "2505"]
    
    //autoAddSoftware ["Office "; "Denwer"; "Notepad++"; "TeXstudio"]

    //addTypicalRecord
    //    1
    //    ["Смирнов Михаил Николаевич"]
    //    ["Промежуточная аттестация (зач)", 2; "Лекции", 24]
    //    []

    //checkWorkDistribution 7

    //let test = OfficialWorkDistributionParser() :> IWorkDistribution
    //printfn "%A" <| test.Teachers 3 "Алгебра и теория чисел"

    printfn "%s" "Done!"
    0
