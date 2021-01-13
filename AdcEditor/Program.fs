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

[<EntryPoint>]
let main _ =

    //logIn ()

    //let teachers = EnglishTeachers(Config.workPlan)
    
    //printfn "%A" (teachers.Teachers 1 "Траектория 3")

    doEverythingRight ()

    //doMagic InProgress

    //autoAddRooms ["405"; "2414"; "2448"]
    
    //autoAddSoftware ["Office "; "Denwer"; "Notepad++"; "TeXstudio"]

    //addTypicalRecord
    //    1
    //    ["Смирнов Михаил Николаевич"]
    //    ["Промежуточная аттестация (зач)", 2; "Лекции", 24]
    //    []

    printfn "%s" "Done!"
    0