module internal CalendarRenderer

open System
open System.Globalization
open Google.Apis.Sheets.v4.Data
open GoogleSheets
open Calendar

type private RowRanges =
    {
        Header: Range
        Weeks: Range
        Total: Range
    }
    member this.Data = Range.union (this.Weeks, this.Total)

type private ColumnRanges =
    {
        Header: Range
        DaysOfWeek: Range
        WeekTotal: Range
        MonthTotal: Range
    }
    member this.Data =
        Range.unionAll
            [
                this.DaysOfWeek
                this.WeekTotal
                this.MonthTotal
            ]

let private getRowRanges calendar =
    let weeks = Calendar.getWeeks calendar
    let header = Range.single 0
    let weeks = header |> Range.nextRangeWithCount weeks.Length
    let total = weeks |> Range.nextSingleRange
    {
        RowRanges.Header = header
        Weeks = weeks
        Total = total
    }

let private columnRanges =
    let header = Range.withStartAndCount (0, 2)
    let daysOfWeek = header |> Range.nextRangeWithCount DaysPerWeek
    let weekTotal = daysOfWeek |> Range.nextSingleRange
    let monthTotal = weekTotal |> Range.nextSingleRange
    {
        ColumnRanges.Header = header
        DaysOfWeek = daysOfWeek
        WeekTotal = weekTotal
        MonthTotal = monthTotal
    }

let renderCalendar (spreadsheet: Spreadsheet) sheetId calendar =

    let weeks = Calendar.getWeeks calendar

    let rowRanges = getRowRanges calendar

    let sheetProperties =
        { SheetProperties.defaultValue with
            FrozenRowCount = Some(Range.getCount rowRanges.Header)
            FrozenColumnCount = Some(Range.getCount columnRanges.Header)
        }
    let setSheetPropertiesRequest =
        SheetsRequests.createSetSheetPropertiesRequest sheetProperties

    let clearFormattingRequest =
        SheetsRequests.createClearFormattingOfSheetRequest sheetId

    let columnCount = Range.getEndIndexValue columnRanges.Data + 1
    let rowCount = Range.getEndIndexValue rowRanges.Data + 1
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
                    rowRanges.Weeks
                    |> Range.subrangeWithStartAndCount (startWeekNumber, weekCount)
                Columns = columnRanges.MonthTotal
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
                        rowRanges.Weeks
                        |> Range.subrangeWithStartAndCount (startWeekNumber, weekCount)
                    Columns = Range.unbounded
                }
            SheetsRequests.createUpdateBorderRequest (range, Borders.outer solidBorder))

    let dayOfWeeksBorderRequest =
        let range =
            {
                SheetId = Some sheetId
                Rows = Range.unbounded
                Columns = columnRanges.DaysOfWeek
            }
        SheetsRequests.createUpdateBorderRequest (range, Borders.outer solidBorder)

    let weekTotalBorderRequest =
        let range =
            {
                SheetId = Some sheetId
                Rows = Range.unbounded
                Columns = columnRanges.WeekTotal
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
                                Rows = Range.subrangeSingle weekNumber rowRanges.Weeks
                                Columns =
                                    Range.subrangeSingle dayOfWeekNumber columnRanges.DaysOfWeek
                            }
                        SheetsRequests.createSetBackgroundColorRequest range inactiveDayColor
        |]

    let requests =
        [
            yield! setDimensionLengthRequests
            setSheetPropertiesRequest
            clearFormattingRequest
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
                Rows = rowRanges.Header
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
                Rows = rowRanges.Weeks
                Columns = columnRanges.Header
            }
        let dateValues =
            [
                for week in weeks -> [ week.StartDate; week.EndDate ]
            ]
        Spreadsheet.updateValuesInRange spreadsheet (range, dateValues)

        let weekSumFormulaValues =
            rowRanges.Weeks
            |> Range.getIndexValues
            |> List.map (fun row ->
                {
                    SheetId = Some sheetId
                    Rows = Range.single row
                    Columns = columnRanges.DaysOfWeek
                }
                |> SheetFormula.sumofRange
                |> List.singleton)
        let range =
            {
                SheetId = Some sheetId
                Rows = rowRanges.Weeks
                Columns = columnRanges.WeekTotal
            }
        Spreadsheet.updateValuesInRange spreadsheet (range, weekSumFormulaValues)

        let monthSumFormulaValues =
            calendar
            |> Calendar.getWeekNumberRanges
            |> List.collect (fun (startWeekNumber, weekCount) ->
                {
                    SheetId = Some sheetId
                    Rows =
                        rowRanges.Weeks
                        |> Range.subrangeWithStartAndCount (startWeekNumber, weekCount)
                    Columns = columnRanges.DaysOfWeek
                }
                |> SheetFormula.sumofRange
                |> List.singleton
                |> List.replicate weekCount)

        let range =
            {
                SheetId = Some sheetId
                Rows = rowRanges.Weeks
                Columns = columnRanges.MonthTotal
            }
        Spreadsheet.updateValuesInRange spreadsheet (range, monthSumFormulaValues)

        let dayOfWeekSumFormulaValues =
            columnRanges.Data
            |> Range.getIndexValues
            |> List.map (fun column ->
                {
                    SheetId = Some sheetId
                    Rows = rowRanges.Weeks
                    Columns = Range.single column
                }
                |> SheetFormula.sumofRange)
            |> List.singleton
        let range =
            {
                SheetId = Some sheetId
                Rows = rowRanges.Total
                Columns = columnRanges.Data
            }
        Spreadsheet.updateValuesInRange spreadsheet (range, dayOfWeekSumFormulaValues)

    updateValues ()
