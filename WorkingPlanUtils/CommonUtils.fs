module CommonUtils

open System.IO

let plansFolder = "../WorkingPlans"

let planNameToCode fileName =
    FileInfo(fileName).Name.Substring(3, "9999-2084".Length)

let planCodeToFileName planCode =
    Directory.EnumerateFiles (System.AppDomain.CurrentDomain.BaseDirectory + "/../../../" + plansFolder)
    |> Seq.find (fun f -> planNameToCode f = planCode)