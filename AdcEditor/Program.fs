module Main

open AdcDomain
open AdcOperational
open CurriculumUtils
open WorkDistributionData

[<EntryPoint>]
let main _ =

    doEverythingRight 
        "20_5162_1.docx" 
        "../../adcCredentials.txt"

    //doMagic
    //    "20_5162_1.docx" 
    //    "../../adcCredentials.txt"
    //    InProgress

    printfn "%s" "Done!"
    0