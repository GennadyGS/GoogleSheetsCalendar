﻿module internal CalendarRenderer

open System
open System.Globalization
open Google.Apis.Sheets.v4.Data
open GoogleSheets
open Calendar

let private clearFormatting (spreadsheet: Spreadsheet) sheetId =
    let request = SheetsRequests.createClearFormattingOfSheetRequest sheetId
    Spreadsheet.update spreadsheet request
    
let renderCalendar (spreadsheet: Spreadsheet) sheetId calendar =

    clearFormatting spreadsheet sheetId

    let weeks = Calendar.getWeeks calendar

    let headerRowRange = Range.single 0
    let weeksRowRange =
        headerRowRange
        |> Range.nextRangeWithCount weeks.Length
    let totalRowRange = weeksRowRange |> Range.nextSingleRange
    let dataRowRange = Range.union (weeksRowRange, totalRowRange)

    let headerColumnRange = Range.withStartAndCount (0, 2)
    let daysOfWeekColumnRange =
        headerColumnRange
        |> Range.nextRangeWithCount DaysPerWeek
    let weekTotalColumnRange = daysOfWeekColumnRange |> Range.nextSingleRange
    let monthTotalColumnRange = weekTotalColumnRange |> Range.nextSingleRange
    let dataColumnRange =
        Range.unionAll
            [
                daysOfWeekColumnRange
                weekTotalColumnRange
                monthTotalColumnRange
            ]

    let sheetProperties =
        { SheetProperties.defaultValue with
            FrozenRowCount = Some(Range.getCount headerRowRange)
            FrozenColumnCount = Some(Range.getCount headerColumnRange)
        }
    let setSheetPropertiesRequest =
        SheetsRequests.createSetSheetPropertiesRequest sheetProperties

    let columnCount = Range.getEndIndexValue dataColumnRange + 1
    let rowCount = Range.getEndIndexValue dataRowRange + 1
    let setDimensionLengthRequests =
        [
            yield! SheetsRequests.createSetDimensionLengthRequests (sheetId, "COLUMNS", columnCount)
            yield! SheetsRequests.createSetDimensionLengthRequests (sheetId, "ROWS", rowCount)
        ]

    let unmergeAllRequest =
        GridRange.unbounded (Some sheetId)
        |> SheetsRequests.createUnmergeCellsRequest

    let mergeCellRequests =
        calendar
        |> Calendar.getWeekNumberRanges
        |> List.map (fun (startWeekNumber, weekCount) ->
            {
                SheetId = Some sheetId
                Rows =
                    weeksRowRange
                    |> Range.subrangeWithStartAndCount (startWeekNumber, weekCount)
                Columns = monthTotalColumnRange
            })
        |> List.map SheetsRequests.createMergeCellsRequest
        |> List.toArray

    let solidBorder = new Border(Style = "SOLID")
    let outerBorderRequest =
        let range = GridRange.unbounded (Some sheetId)
        SheetsRequests.createUpdateBorderRequest (range, Borders.outer solidBorder)

    let monthBorderRequests =
        calendar
        |> Calendar.getWeekNumberRanges
        |> List.map (fun (startWeekNumber, weekCount) ->
            let range =
                {
                    SheetId = Some sheetId
                    Rows =
                        weeksRowRange
                        |> Range.subrangeWithStartAndCount (startWeekNumber, weekCount)
                    Columns = Range.unbounded
                }
            SheetsRequests.createUpdateBorderRequest (range, Borders.outer solidBorder))

    let dayOfWeeksBorderRequest =
        let range =
            {
                SheetId = Some sheetId
                Rows = Range.unbounded
                Columns = daysOfWeekColumnRange
            }
        SheetsRequests.createUpdateBorderRequest (range, Borders.outer solidBorder)

    let weekTotalBorderRequest =
        let range =
            {
                SheetId = Some sheetId
                Rows = Range.unbounded
                Columns = weekTotalColumnRange
            }
        SheetsRequests.createUpdateBorderRequest (range, Borders.outer solidBorder)

    let setBordersRequests =
        [
            outerBorderRequest
            yield! monthBorderRequests
            dayOfWeeksBorderRequest
            weekTotalBorderRequest
        ]

    let setCellBackgroundColorRequests =
        let inactiveDayColor = Color(Red = 0.75f, Green = 0.75f, Blue = 0.75f)
        [|
            for (weekNumber, week) in List.indexed weeks do
                for dayOfWeekNumber in [ 0 .. DaysPerWeek - 1 ] do
                    if not week.DaysActive[dayOfWeekNumber] then
                        let range =
                            {
                                SheetId = Some sheetId
                                Rows = Range.subrangeSingle weekNumber weeksRowRange
                                Columns = Range.subrangeSingle dayOfWeekNumber daysOfWeekColumnRange
                            }
                        SheetsRequests.createSetBackgroundColorRequest range inactiveDayColor
        |]

    let requests =
        [
            setSheetPropertiesRequest
            yield! setDimensionLengthRequests
            unmergeAllRequest
            yield! mergeCellRequests
            yield! setBordersRequests
            yield! setCellBackgroundColorRequests
        ]
    Spreadsheet.batchUpdate spreadsheet requests

    let updateValues () =
        let range =
            {
                SheetId = Some sheetId
                Rows = headerRowRange
                Columns = Range.unbounded
            }
        let firstDayOfWeek = Calendar.getFirstDayOfWeek calendar
        let dayOfWeekNames =
            Enum.GetValues<DayOfWeek>()
            |> Array.map (DayOfWeek.addDays (int firstDayOfWeek))
            |> Array.map CultureInfo.InvariantCulture.DateTimeFormat.GetDayName
            |> Array.toList

        let values =
            [
                "Start Date"
                "End Date"
                yield! dayOfWeekNames
                "Week Total"
                "Month Total"
            ]
            |> List.singleton
        Spreadsheet.updateValuesInRange spreadsheet (range, values)

        let range =
            {
                SheetId = Some sheetId
                Rows = weeksRowRange
                Columns = headerColumnRange
            }
        let dateValues =
            [
                for week in weeks -> [ week.StartDate; week.EndDate ]
            ]
        Spreadsheet.updateValuesInRange spreadsheet (range, dateValues)

        let weekSumFormulaValues =
            weeksRowRange
            |> Range.getIndexValues
            |> List.map (fun row ->
                {
                    SheetId = Some sheetId
                    Rows = Range.single row
                    Columns = daysOfWeekColumnRange
                }
                |> SheetFormula.sumofRange
                |> List.singleton)
        let range =
            {
                SheetId = Some sheetId
                Rows = weeksRowRange
                Columns = weekTotalColumnRange
            }
        Spreadsheet.updateValuesInRange spreadsheet (range, weekSumFormulaValues)

        let monthSumFormulaValues =
            calendar
            |> Calendar.getWeekNumberRanges
            |> List.collect (fun (startWeekNumber, weekCount) ->
                {
                    SheetId = Some sheetId
                    Rows =
                        weeksRowRange
                        |> Range.subrangeWithStartAndCount (startWeekNumber, weekCount)
                    Columns = daysOfWeekColumnRange
                }
                |> SheetFormula.sumofRange
                |> List.singleton
                |> List.replicate weekCount)

        let range =
            {
                SheetId = Some sheetId
                Rows = weeksRowRange
                Columns = monthTotalColumnRange
            }
        Spreadsheet.updateValuesInRange spreadsheet (range, monthSumFormulaValues)

        let dayOfWeekSumFormulaValues =
            dataColumnRange
            |> Range.getIndexValues
            |> List.map (fun column ->
                {
                    SheetId = Some sheetId
                    Rows = weeksRowRange
                    Columns = Range.single column
                }
                |> SheetFormula.sumofRange)
            |> List.singleton
        let range =
            {
                SheetId = Some sheetId
                Rows = totalRowRange
                Columns = dataColumnRange
            }
        Spreadsheet.updateValuesInRange spreadsheet (range, dayOfWeekSumFormulaValues)

    updateValues ()

