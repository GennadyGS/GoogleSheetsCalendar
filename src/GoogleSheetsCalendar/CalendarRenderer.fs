module internal CalendarRenderer

open System
open System.Globalization
open Google.Apis.Sheets.v4.Data
open GoogleSheets
open Calendar

[<Literal>]
let private InactiveCellColorIntencity = 0.75f

type private RowRanges =
    {
        Header: Range
        Weeks: Range
        Totals: Range
    }
    member this.Data = Range.union (this.Weeks, this.Totals)

type private ColumnRanges =
    {
        Header: Range
        DaysOfWeek: Range
        WeekTotals: Range
        MonthTotals: Range
    }
    member this.Data =
        Range.unionAll
            [
                this.DaysOfWeek
                this.WeekTotals
                this.MonthTotals
            ]

let private getRowRanges calendar =
    let weeks = Calendar.getWeeks calendar
    let header = Range.single 0
    let weeks = header |> Range.nextRangeWithCount weeks.Length
    let total = weeks |> Range.nextSingleRange
    {
        RowRanges.Header = header
        Weeks = weeks
        Totals = total
    }

let private columnRanges =
    let header = Range.withStartAndCount (0, 2)
    let daysOfWeek = header |> Range.nextRangeWithCount DaysPerWeek
    let weekTotal = daysOfWeek |> Range.nextSingleRange
    let monthTotal = weekTotal |> Range.nextSingleRange
    {
        ColumnRanges.Header = header
        DaysOfWeek = daysOfWeek
        WeekTotals = weekTotal
        MonthTotals = monthTotal
    }

let private getUpdateSheetRequests sheetId calendar =
    let createGrigRange = GridRange.create (Some sheetId)

    let rowRanges = getRowRanges calendar

    let setDimensionLengthRequests =
        let columnCount = Range.getEndIndexValue columnRanges.Data + 1
        let rowCount = Range.getEndIndexValue rowRanges.Data + 1
        let createSetDimensionLengthRequests (dimension, length) =
            SheetsRequests.createSetDimensionLengthRequests (Some sheetId, dimension, length)
        [
            yield! createSetDimensionLengthRequests (Dimension.Columns, columnCount)
            yield! createSetDimensionLengthRequests (Dimension.Rows, rowCount)
        ]

    let setSheetPropertiesRequest =
        { SheetProperties.createDefault (Some sheetId) with
            FrozenRowCount = Some(Range.getCount rowRanges.Header)
            FrozenColumnCount = Some(Range.getCount columnRanges.Header)
        }
        |> SheetsRequests.createSetSheetPropertiesRequest

    let clearFormattingRequest =
        createGrigRange (Range.unbounded, Range.unbounded)
        |> SheetsRequests.createClearFormattingRequest

    let unmergeAllRequest =
        createGrigRange (Range.unbounded, Range.unbounded)
        |> SheetsRequests.createUnmergeCellsRequest

    let mergeCellRequests =
        calendar
        |> Calendar.getWeekNumberRanges
        |> List.map (fun (startWeekNumber, weekCount) ->
            let rows =
                rowRanges.Weeks
                |> Range.subrangeWithStartAndCount (startWeekNumber, weekCount)
            createGrigRange (rows, columnRanges.MonthTotals))
        |> List.map SheetsRequests.createMergeCellsRequest
        |> List.toArray

    let setBordersRequests =
        let solidBorder = new Border(Style = "SOLID")

        let outerBorderRequest =
            let range = createGrigRange (Range.unbounded, Range.unbounded)
            SheetsRequests.createUpdateBorderRequest (range, Borders.outer solidBorder)

        let monthBorderRequests =
            calendar
            |> Calendar.getWeekNumberRanges
            |> List.map (fun (startWeekNumber, weekCount) ->
                let rows =
                    rowRanges.Weeks
                    |> Range.subrangeWithStartAndCount (startWeekNumber, weekCount)
                let range = createGrigRange (rows, Range.unbounded)
                SheetsRequests.createUpdateBorderRequest (range, Borders.outer solidBorder))

        let dayOfWeeksBorderRequest =
            let range = createGrigRange (Range.unbounded, columnRanges.DaysOfWeek)
            SheetsRequests.createUpdateBorderRequest (range, Borders.outer solidBorder)

        let weekTotalBorderRequest =
            let range = createGrigRange (Range.unbounded, columnRanges.WeekTotals)
            SheetsRequests.createUpdateBorderRequest (range, Borders.outer solidBorder)

        [
            outerBorderRequest
            yield! monthBorderRequests
            dayOfWeeksBorderRequest
            weekTotalBorderRequest
        ]

    let setInactiveCellBackgroundColorRequests =
        let inactiveCellColor = Color.grey InactiveCellColorIntencity
        [|
            let weeks = Calendar.getWeeks calendar
            for (weekNumber, week) in List.indexed weeks do
                for dayOfWeekNumber in [ 0 .. DaysPerWeek - 1 ] do
                    if not week.DaysActive[dayOfWeekNumber] then
                        let rows = Range.subrangeSingle weekNumber rowRanges.Weeks
                        let columns = Range.subrangeSingle dayOfWeekNumber columnRanges.DaysOfWeek
                        let range = createGrigRange (rows, columns)
                        SheetsRequests.createSetBackgroundColorRequest range inactiveCellColor
        |]

    [
        yield! setDimensionLengthRequests
        setSheetPropertiesRequest
        clearFormattingRequest
        unmergeAllRequest
        yield! mergeCellRequests
        yield! setBordersRequests
        yield! setInactiveCellBackgroundColorRequests
    ]

let private getUpdateValuesRequests sheetId calendar =
    let createGrigRange = GridRange.create (Some sheetId)

    let rowRanges = getRowRanges calendar

    let titlesRowValueRange =
        let range = createGrigRange (rowRanges.Header, Range.unbounded)
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
        (range, values)

    let weekDatesValueRange =
        let range = createGrigRange (rowRanges.Weeks, columnRanges.Header)
        let values =
            [
                for week in Calendar.getWeeks calendar -> [ week.StartDate; week.EndDate ]
            ]
        (range, values)

    let weekTotalsValueRange =
        let formulaValues =
            createGrigRange (rowRanges.Weeks, columnRanges.DaysOfWeek)
            |> SheetFormulaValues.rowWiseAggregation AggregationFunction.Sum
        let range = createGrigRange (rowRanges.Weeks, columnRanges.WeekTotals)
        (range, formulaValues)

    let monthTotalsValueRange =
        let formulaValues =
            calendar
            |> Calendar.getWeekNumberRanges
            |> List.collect (fun (startWeekNumber, weekCount) ->
                let rows =
                    rowRanges.Weeks
                    |> Range.subrangeWithStartAndCount (startWeekNumber, weekCount)
                createGrigRange (rows, columnRanges.DaysOfWeek)
                |> SheetFormulaValue.aggregate AggregationFunction.Sum
                |> List.singleton
                |> List.replicate weekCount)
        let range = createGrigRange (rowRanges.Weeks, columnRanges.MonthTotals)
        (range, formulaValues)

    let totalsRowValueRange =
        let formulaValues =
            createGrigRange (rowRanges.Weeks, columnRanges.Data)
            |> SheetFormulaValues.columnWiseAggregation AggregationFunction.Sum
        let range = createGrigRange (rowRanges.Totals, columnRanges.Data)
        (range, formulaValues)

    [
        titlesRowValueRange |> ValueRange.box
        weekDatesValueRange |> ValueRange.box
        weekTotalsValueRange |> ValueRange.box
        monthTotalsValueRange |> ValueRange.box
        totalsRowValueRange |> ValueRange.box
    ]

let renderCalendar (spreadsheet: Spreadsheet) sheetId calendar =
    calendar
    |> getUpdateSheetRequests sheetId
    |> Spreadsheet.batchUpdate spreadsheet

    calendar
    |> getUpdateValuesRequests sheetId
    |> Spreadsheet.batchUpdateValuesInRange spreadsheet
