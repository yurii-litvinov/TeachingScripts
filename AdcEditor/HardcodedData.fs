﻿module HardcodedData

let englishT2Rooms = [2507; 2509]
let englishT2Teachers = ["Журавлева Юлия Владимировна"; "Юлова Екатерина Сергеевна"]
let englishT2S3WorkTypes = ["Под руководством преподавателя", 48; "Практические занятия", 58; "Промежуточная аттестация (зач)", 2]
let englishT2S4WorkTypes = ["Под руководством преподавателя", 48; "Практические занятия", 50; "Промежуточная аттестация (зач)", 10]

let englishT3Rooms = [3503; 2509; 3507; 3512; 2503; 2510]
let englishT3Teachers = ["Парфенов Андрей Сергеевич"; "Шашукова Анна Сергеевна"; "Кононов Борис Викторович"]
let englishT3S3WorkTypes = ["Под руководством преподавателя", 48; "Практические занятия", 58; "Промежуточная аттестация (зач)", 2]
let englishT3S4WorkTypes = ["Под руководством преподавателя", 48; "Практические занятия", 50; "Промежуточная аттестация (зач)", 10]

let englishT4Rooms = [2505]
let englishT4Teachers = ["Гукалина Александра Владимировна"; "Аношина Елена Владимировна"] 
let englishT4S3WorkTypes = ["Под руководством преподавателя", 76; "Практические занятия", 30; "Промежуточная аттестация (зач)", 2]
let englishT4S4WorkTypes = ["Под руководством преподавателя", 70; "Практические занятия", 28; "Промежуточная аттестация (зач)", 10]

let onlineWorkTypes = ["Консультации", 10; "Промежуточная аттестация (зач)", 2]

let rkiTeacher = ["Быстрых Марина Викторовна"]
let rkiRoom = [3512]
let rkiT1S3WorkTypes = ["Под руководством преподавателя", 18; "В присутствии преподавателя", 30; "Практические занятия", 58; "Промежуточная аттестация (зач)", 2]
let rkiT2S3WorkTypes = ["Под руководством преподавателя", 48; "Практические занятия", 58; "Промежуточная аттестация (зач)", 2]
let rkiT1S4WorkTypes = ["Под руководством преподавателя", 10; "В присутствии преподавателя", 30; "Практические занятия", 56; "Промежуточная аттестация (экз)", 10; "Консультации", 2]
let rkiT2S4WorkTypes = ["Под руководством преподавателя", 46; "Практические занятия", 50; "Промежуточная аттестация (экз)", 10; "Консультации", 2]

let computerClasses = ["2406"; "2408"; "2412"; "2414"; "2444-1"; "2444-2"; "2446"]

let physicalTrainingWorkTypes = ["Практические занятия", 34; "Текущий контроль (ауд)", 2]
let physicalTrainingSem3WorkTypes = ["В присутствии преподавателя", 62]
let physicalTrainingSem4WorkTypes = ["В присутствии преподавателя", 62]
let physicalTrainingTeacher = ["Поципун Анатолий Антонович"]
let physicalTrainingRooms = ["по спортивным объектам СПбГУ"]

let algebraWorkTypes = ["Практические занятия", 26; "Лекции", 30; "Консультации", 2; "Промежуточная аттестация (экз)", 2; "Контрольные работы", 4; "Промежуточная аттестация (зач)", 2]

let practiceS4WorkTypes = ["Под руководством преподавателя", 15; "Промежуточная аттестация (зач)", 2]

let onlineCourses =
    Map.empty
        // Have no idea, used first author of the working program.
        .Add("Основы противодействия коррупции и экстремизму (онлайн-курс)", ["Дмитрикова Екатерина Александровна"])
        // Same here.
        .Add("Основы финансовой грамотности (онлайн-курс)", ["Белозеров Сергей Анатольевич"])

let irrelevantIndustrialExperience = 
    Set.empty
        .Add("Михайлова Елена Георгиевна")
        .Add("Абрамов Максим Викторович")
