module CurriculumUtils

open System.Text.RegularExpressions
open System.IO

open CurriculumParser

/// Information about a competence.
type Competence =
    {
        /// Competence code, for example, "ОПК-2".
        Code: string

        /// Competence description, for example, "способен применять современный математический аппарат..."
        Description: string
    }

/// Information about workload item for a given semester.
type WorkloadItem =
    {
        /// Name of an item, for example, "Лекции". Full list of possible workload item names 
        /// is provided in CurriculumUtils.fs.
        Name: string

        /// Hours of work, for example, "30".
        Hours: int
    }

/// Information about one semester of a discipline. Most disciplines are single semester, but some span multiple 
// semesters, like Programming or Mathematical Analysis.
type DisciplineSemester =
    {
        /// Number of semester in which the discipline is implemented, from 1 to 8.
        Semester: int

        /// Labor intensity of a semester, in ECTS credits.
        LaborIntensity: int

        /// Type of an attestation, can be "зачёт" or "экзамен".
        AttestationType: string

        /// A list of workload items with non-zero hours of work.
        Workload: WorkloadItem list
    }

/// Information about discipline.
type Discipline =
    {
        /// Code of a working program, for example, "002212".
        Code: string

        /// Russian name of a discipline.
        Name: string

        /// A list of competences that this discipline starts to form (there are no disciplines with these competences 
        /// that precede this discipline in a working plan).
        FormedCompetences: Competence list

        /// A list of competences improved by this discipline (so we can have them in prerequisites).
        ImprovedCompetences: Competence list

        /// A list of competences fully formed by studying this discipline (so we can test students after 
        /// this discipline, no disciplines later in working plan improve these competences).
        FullyFormedCompetences: Competence list

        /// A list of semesters in which the discipline is implemented, with semester-specific info.
        Semesters: DisciplineSemester list
    }

/// Wrapper around CurriculumParser, provides additional info about working plan.
type Curriculum(fileName: string) =
    let curriculum = DocxCurriculum(fileName)

    let programme = FileInfo(fileName).Name.Substring(0, "ВМ.5665".Length)
    let year = FileInfo(fileName).Name.Substring("ВМ.5665-".Length, 4)

    let workTypes = 
        [
            "Лекции"
            "Семинары"
            "Консультации"
            "Практические занятия"
            "Лабораторные работы"
            "Контрольные работы"
            "Коллоквиумы"
            "Текущий контроль"
            "Промежуточная аттестация"
            "Под руководством преподавателя (сам. раб.)"
            "В присутствии преподавателя (сам. раб.)"
            "В т.ч. с использованием учебно-методич. материалов (сам. раб.)"
            "Текущий контроль (сам. раб.)"
            "Промежуточная аттестация (сам. раб.)"
            "Объём занятий в активных и интерактивных формах, часов"
        ] 

    let getWorkTypes (impl: DisciplineImplementation) =
        let workHours = impl.WorkHours.Split [|' '|] |> Seq.map int

        Seq.zip workTypes workHours 
        |> Seq.filter ((<=) 0 << snd) 
        |> Seq.map (fun (t, h) -> {Name = t; Hours = h})
        |> Seq.toList

    let parseImplementations (impl: DisciplineImplementation) =
        {
            Semester = impl.Semester
            LaborIntensity = impl.LaborIntensity
            AttestationType = impl.MonitoringTypes
            Workload = getWorkTypes impl
        }

    let semesters (discipline: CurriculumParser.Discipline) =
        let sortedSemesters = discipline.Implementations |> Seq.map (fun i -> i.Semester) |> Seq.sort
        (Seq.head sortedSemesters, Seq.last sortedSemesters)

    let isCompetenceFormedInSemester (semester: int) (curriculum: DocxCurriculum) (competence: Competence) =
        if semester = 1 then
            true
        else
            curriculum.Disciplines
            |> Seq.map (fun d -> 
                let startSemester, _ = semesters d
                (startSemester, d.Implementations.[0].Competences)
                )
            |> Seq.filter (fun (s, _) -> s < semester)
            |> Seq.filter (fun (_, cs) -> cs |> Seq.exists (fun c -> c.Code = competence.Code))
            |> Seq.isEmpty

    let isCompetenceFinishedFormingInSemester (semester: int) (curriculum: DocxCurriculum) (competence: Competence) =
        if semester = 8 then
            true
        else
            curriculum.Disciplines
            |> Seq.map (fun d -> 
                let _, endSemester = semesters d
                (endSemester, d.Implementations.[0].Competences)
                )
            |> Seq.filter (fun (e, _) -> e > semester)
            |> Seq.filter (fun (_, cs) -> cs |> Seq.exists (fun c -> c.Code = competence.Code))
            |> Seq.isEmpty

    let competences (discipline: CurriculumParser.Discipline) =
        discipline.Implementations.[0].Competences
        |> Seq.map (fun c -> {Code = c.Code; Description = c.Description})
        |> Seq.toList

    let parseDiscipline (curriculum: DocxCurriculum) (discipline: CurriculumParser.Discipline) =
        let semestersList = 
            discipline.Implementations
            |> Seq.map parseImplementations
            |> Seq.toList

        let competencesList = competences discipline
        let startSemester, endSemester = semesters discipline
        let formedCompetences = 
            competencesList 
            |> Seq.filter (isCompetenceFormedInSemester startSemester curriculum)
            |> Seq.toList

        let fullyFormedCompetences = 
            competencesList 
            |> Seq.filter (isCompetenceFinishedFormingInSemester endSemester curriculum)
            |> Seq.toList

        let improvedCompetences = 
            competencesList 
            |> Seq.filter ((isCompetenceFinishedFormingInSemester startSemester curriculum) >> not)
            |> Seq.filter ((isCompetenceFormedInSemester startSemester curriculum) >> not)
            |> Seq.toList

        { 
            Code = discipline.Code
            Name = discipline.RussianName
            FormedCompetences = formedCompetences
            ImprovedCompetences = improvedCompetences
            FullyFormedCompetences = fullyFormedCompetences
            Semesters = semestersList
        }

    /// List of all disciplines in this plan.
    member _.Disciplines =
        curriculum.Disciplines
        |> Seq.map (parseDiscipline curriculum)

    /// Programme name, for example, "СВ.5006".
    member _.Programme = programme

    /// Year of admission of a plan, for example, "2019".
    member _.Year = year
