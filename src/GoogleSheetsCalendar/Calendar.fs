module Calendar

open System

let DaysPerWeek = Enum.GetValues<DayOfWeek>().Length

type Week =
    {
        StartDate: DateOnly
        EndDate: DateOnly
        DaysActive: bool array
    }

type Month = Month of Week list

type Calendar = Calendar of Month list

module Int =
    let between (min, max) x = x >= min && x <= max

module DayOfWeek =
    let diff (x: DayOfWeek, y: DayOfWeek) =
        (int x - int y + DaysPerWeek) % DaysPerWeek

    let addDays days (dayOfWeek: DayOfWeek) =
        (int dayOfWeek + int days) % DaysPerWeek
        |> LanguagePrimitives.EnumOfValue

[<RequireQualifiedAccess>]
module Calendar =
    let private calculateMonth firstDayOfWeeek year month =
        let dayCount = DateTime.DaysInMonth(year, month)
        let startDate = DateOnly(year, month, 1)
        let monthFistDayOfWeek = DayOfWeek.diff (startDate.DayOfWeek, firstDayOfWeeek)
        [ 1..dayCount ]
        |> List.groupBy (fun day -> (day - 1 + monthFistDayOfWeek) / DaysPerWeek)
        |> List.map (fun (weekNumber, _) ->
            let startDay = weekNumber * DaysPerWeek - monthFistDayOfWeek
            {
                StartDate = startDate.AddDays(startDay)
                EndDate = startDate.AddDays(startDay + DaysPerWeek - 1)
                DaysActive =
                    Array.init DaysPerWeek (fun dayOfWeek ->
                        startDay + dayOfWeek
                        |> Int.between (0, dayCount - 1))
            })
        |> Month

    let calculate firstDayOfWeeek year =
        let monthCount = DateOnly(year + 1, 1, 1).AddDays(-1).Month
        [ 1..monthCount ]
        |> List.map (calculateMonth firstDayOfWeeek year)
        |> Calendar

    let getMonths (Calendar months) = months

    let getWeeks (Calendar months) =
        months
        |> List.collect (fun (Month weeks) -> weeks)

    let getWeekNumberRanges (Calendar months) =
        months
        |> List.scan
            (fun (_, nextWeekStartNumber) (Month weeks) ->
                (weeks, nextWeekStartNumber + weeks.Length))
            ([], 0)
        |> List.tail
        |> List.map (fun (weeks, nextWeekStartNumber) ->
            (nextWeekStartNumber - weeks.Length, weeks.Length))

    let getFirstDayOfWeek calendar =
        let firstFullWeek = calendar |> getWeeks |> List.find (fun week -> week.DaysActive[0])
        firstFullWeek.StartDate.DayOfWeek

