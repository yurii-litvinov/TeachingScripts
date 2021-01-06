module AdcOperational

open AdcDomain

let addTeachers teachers workTypes =
    switchTab Teachers
    for teacher in teachers do
        addTeacher teacher workTypes

let addTypicalRecordWithoutOpening teachers workTypes rooms =
    addTeachers teachers workTypes
    refresh ()
    removeWorkTypes 0 (workTypes |> Seq.length)
    addRooms rooms (workTypes |> Seq.length)

let addTypicalRecord recordNumber teachers workTypes rooms =
    openRecord recordNumber
    addTypicalRecordWithoutOpening teachers workTypes rooms

let processTypicalRecord recordNumber teachers workTypes rooms =
    openRecord recordNumber

    wipeOutTeachers ()
    wipeOutRooms ()

    addTypicalRecordWithoutOpening teachers workTypes rooms

let processLectionWithPractices recordNumber lecturer lecturerWorkTypes practitioners practicionerWorkTypes =
    openRecord recordNumber
    wipeOutTeachers ()

    addTypicalRecordWithoutOpening [lecturer] lecturerWorkTypes []

    addTeachers practitioners practicionerWorkTypes
