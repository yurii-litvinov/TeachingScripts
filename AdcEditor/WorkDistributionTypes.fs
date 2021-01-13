module WorkDistributionTypes

type WorkloadInfo =
    {
        lecturers: string list
        practicioners: string list
    }

type IWorkDistribution =
    abstract Teachers: semester: int -> discipline: string -> WorkloadInfo
