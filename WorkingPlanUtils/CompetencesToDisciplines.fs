module CompetencesToDisciplines

open CommonUtils

open CurriculumParser
open FSharp.Json
open System.IO

type DisciplineDto =
    { Code: string
      Name: string }

type CompetenceToDisciplinesDto =
    { CompetenceCode: string
      CompetenceName: string
      Disciplines: DisciplineDto list }

let generateCompetencesToDisciplinesMap (plan: string) (outFileName: string) =
    let curriculum = DocxCurriculum(planCodeToFileName plan)
    
    let competenceToDisciplines =
        curriculum.Disciplines
        |> Seq.collect _.Implementations
        |> Seq.collect (fun (i: DisciplineImplementation) -> i.Competences |> Seq.map (fun c -> (c, i.Discipline)))
        |> Seq.distinctBy (fun (c: Competence, d: Discipline) -> c.Code + " " + d.Code)
        |> Seq.groupBy fst
        |> Seq.map (fun (c, d) -> (c, d |> Seq.map snd))
        |> Seq.sortBy (fun (c: Competence, _) -> c.Code)
        |> Seq.map (fun (c, ds) -> c, ds |> Seq.sortBy _.Code)

    let convertDiscipline (d: Discipline) =
        { Code = d.Code
          Name = d.RussianName }

    let dto = 
        competenceToDisciplines 
        |> Seq.map (fun (c, ds) -> 
            { CompetenceCode = c.Code
              CompetenceName = c.Description
              Disciplines = ds |> Seq.map convertDiscipline |> Seq.toList })
        |> Seq.toArray

    let result = Json.serialize dto

    File.WriteAllText(outFileName, result)
