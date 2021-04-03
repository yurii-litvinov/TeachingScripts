module DataModel

open System

type Program = 
    | MatobesBachelor
    | MatobesMaster
    | PiBachelor
    | PiMaster

type Date = Date of string

type TimeSlot = Morning | Day

type SittingSlot =
    { date: Date
      time: TimeSlot
      program: Program }

type Diploma =
    { theme: string
      author: string
      program: Program
      advisor: string
      reviewer: string
      convenientDates: Date list
      inconvenientDates: Date list }

type InvolvementPreference = Min | Max

[<CustomEquality>]
[<CustomComparison>]
type CommissionMember =
    {  name: string
       company: string
       convenientDates: Date list
       inconvenientDates: Date list
       maxDays: int
       involvementPreference: InvolvementPreference
       interestingThemes: string list
       isChair: bool }
    with
        override this.Equals (o: obj) = this.name = (o :?> CommissionMember).name
        override this.GetHashCode() = this.name.GetHashCode()
        interface System.IComparable with
            override this.CompareTo (o: obj) = (this.name :> IComparable).CompareTo((o :?> CommissionMember).name)

type Sitting =
    { slot: SittingSlot
      works: Diploma list }

type Day =
    { date: Date
      chair: CommissionMember
      commission: CommissionMember list
      firstSitting: option<Sitting>
      secondSitting: option<Sitting> }

type Schedule = 
    { schedule: Day list
      totalDays: Collections.Generic.Dictionary<CommissionMember, int>
      advisorTotalDays: Collections.Generic.Dictionary<string, int> }

type Fitness =
    { commissionDates: int
      chairDates: int
      studentDates: int
      advisorDates: int 
      themes: int
      totalDaysCommissionPenalty: int
      totalDaysAdvisorPenalty: int
      commissionBalance: int }
