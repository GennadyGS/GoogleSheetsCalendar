module internal CalendarRenderer

open System
open System.Globalization
open Google.Apis.Sheets.v4
open Google.Apis.Sheets.v4.Data
open GoogleSheets
open Calendar

let renderCalendar (sheetsService: SheetsService) (spreadsheetId, sheetId) calendar =

    let clearFormatting () =
        let range = TwoDimensionRange.unbounded (Some sheetId)
        let clearFormattingRequest = SheetsRequests.createClearFormattingRequest range

        SheetsService.batchUpdate sheetsService spreadsheetId [ clearFormattingRequest ]

    clearFormatting ()

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
        SheetsService.updateValuesInRange sheetsService spreadsheetId (range, values)

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
        SheetsService.updateValuesInRange sheetsService spreadsheetId (range, dateValues)

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
        SheetsService.updateValuesInRange sheetsService spreadsheetId (range, weekSumFormulaValues)

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
        SheetsService.updateValuesInRange sheetsService spreadsheetId (range, monthSumFormulaValues)

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
        SheetsService.updateValuesInRange
            sheetsService
            spreadsheetId
            (range, dayOfWeekSumFormulaValues)

    updateValues ()

    let setSheetPropertiesRequest =
        SheetsRequests.createSetSheetPropertiesRequest (Some 1, Some 2)

    let columnCount = Range.getEndIndexValue dataColumnRange + 1
    let rowCount = Range.getEndIndexValue dataRowRange + 1
    let setDimensionLengthRequests =
        [
            yield! SheetsRequests.createSetDimensionLengthRequests (sheetId, "COLUMNS", columnCount)
            yield! SheetsRequests.createSetDimensionLengthRequests (sheetId, "ROWS", rowCount)
        ]

    let unmergeAllRequest =
        TwoDimensionRange.unbounded (Some sheetId)
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
        let range = TwoDimensionRange.unbounded (Some sheetId)
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
    SheetsService.batchUpdate sheetsService spreadsheetId requests
