open CompetencesToDisciplines
open ComparePlans
open PrintBySemester
open CommonUtils

open System.IO
open Argu

type Arguments = 
    | [<CliPrefix(CliPrefix.None); Unique; First>] Print_By_Semester of plan: string * semester: int
    | [<CliPrefix(CliPrefix.None); Unique; First>] Compare of plan1: string * plan2: string
    | [<CliPrefix(CliPrefix.None); Unique; First>] Compare_To_Semester of plan1: string * plan2: string * semester: int
    | [<CliPrefix(CliPrefix.None); Unique; First>] Competences_To_Disciplines of plan: string
    interface IArgParserTemplate with
        member this.Usage: string = 
            match this with
            | Print_By_Semester _ -> "specify a plan code and a semester."
            | Compare _ -> "specify two working plans"
            | Compare_To_Semester _ -> "specify two working plans and a semester up to which they shall be compared"
            | Competences_To_Disciplines _ -> "specify a working plan for which competence to disciplines map is to be printed"

[<EntryPoint>]
let main argv =
    let parser = ArgumentParser.Create<Arguments>(programName = "WorkingPlanUtils.exe")
    
    let result = parser.Parse (inputs=argv, raiseOnUsage=false)

    if result.Contains Print_By_Semester then
        let plan, semester = result.GetResult Print_By_Semester
        printBySemester plan semester
    elif result.Contains Compare then
        let plan1, plan2 = result.GetResult Compare
        comparePlans plan1 plan2
    elif result.Contains Compare_To_Semester then
        let plan1, plan2, semester = result.GetResult Compare_To_Semester
        compareToSemester plan1 plan2 semester
    elif result.Contains Competences_To_Disciplines then
        let plan = result.GetResult Competences_To_Disciplines
        generateCompetencesToDisciplinesMap plan "competencesToDisciplines.json"
    else
        printfn "%s" <| parser.PrintUsage()

        let existingPlans =
            Directory.EnumerateFiles(plansFolder) 
            |> Seq.map (fun p -> FileInfo(p).Name.Substring(3, "9999-2084".Length) + " ")
            |> Seq.reduce (+)

        printfn "\n\nExisting plans:\n %s" existingPlans

    0
