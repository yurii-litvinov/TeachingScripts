module Main

open AdcDomain
open AdcOperational
open CurriculumUtils
open WorkDistributionData
open HardcodedData
open VkrAdvisorsParser

[<EntryPoint>]
let main _ =

    //logIn "../../adcCredentials.txt"

    //doEverythingRight 
    //    "20_5162_1.docx" 
    //    "../../adcCredentials.txt"

    //doMagic
    //    "20_5162_1.docx" 
    //    "../../adcCredentials.txt"
    //    InProgress

    autoAddRooms ["2414"]
    
    //autoAddSoftware["ANTLR"; "Bison"]

    //addTypicalRecord
    //    1
    //    ["Демьянович Юрий Казимирович"]
    //    ["Лабораторные работы", 16; "Консультации", 2; "Промежуточная аттестация (экз)", 2; "Лекции", 48]
    //    ["405"; "2448"]

    let students = parseVkrAdvisors "../../Приложение к приказу ВКР и НР5006 23.11.2020 20.05.01.docx"

    students 
    |> Seq.iter (fun s -> printfn "%s" s.advisor)

    printfn "%s" "Done!"
    0