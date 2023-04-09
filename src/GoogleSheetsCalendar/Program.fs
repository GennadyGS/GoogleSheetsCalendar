open System
open System.IO
open System.Collections.Generic
open System.Globalization
open System.Reflection
open Microsoft.Extensions.Configuration
open Google.Apis.Sheets.v4
open Google.Apis.Sheets.v4.Data
open Calendar
open GoogleSheets

[<CLIMutable>]
type Configuration =
    {
        SpreadsheetId: string
        SheetId: int
        Year: int
        FirstDayOfWeek: DayOfWeek
    }

[<Literal>]
let CredentialFileName = "credential.json"

module NullCoalesce =
    let coalesce (b: 'a Lazy) (a: 'a) =
        if obj.ReferenceEquals(a, null) then
            b.Value
        else
            a

let getRootDirectoryPath () =
    let executableFilePath = Assembly.GetExecutingAssembly().Location
    Path.GetDirectoryName(executableFilePath)

let createConfiguration rootDirectoryPath =
    let nativeConfiguration =
        (new ConfigurationBuilder())
            .SetBasePath(rootDirectoryPath)
            .AddJsonFile("appsettings.json", true, true)
            .AddUserSecrets<Configuration>()
            .Build()

    nativeConfiguration.Get<Configuration>()
    |> NullCoalesce.coalesce (lazy raise (InvalidOperationException("Configuration is missing")))

let renderCalendar (sheetsService: SheetsService) configuration calendar =
    let weeks = Calendar.getWeeks calendar
    let spreadsheetId = configuration.SpreadsheetId

    let clearFormatting () =
        let range = TwoDimensionRange.unbounded (Some configuration.SheetId)
        let clearFormattingRequest = SheetsRequests.createClearFormattingRequest range

        SheetsService.batchUpdate sheetsService spreadsheetId [ clearFormattingRequest ]

    clearFormatting ()

    let headerRowRange = Range.single 0
    let weeksRowRange =
        headerRowRange
        |> Range.nextRangeWithCount weeks.Length
    let totalRowRange = weeksRowRange |> Range.nextSingleRange

    let headerColumnRange = Range.fromStartAndCount (0, 2)
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
                SheetId = Some configuration.SheetId
                Rows = headerRowRange
                Columns = Range.unbounded
            }
        let dayOfWeekNames =
            Enum.GetValues<DayOfWeek>()
            |> Array.map (DayOfWeek.addDays (int configuration.FirstDayOfWeek))
            |> Array.map CultureInfo.InvariantCulture.DateTimeFormat.GetDayName
            |> Array.toList

        [
            "Start Date"
            "End Date"
            yield! dayOfWeekNames
            "Week Total"
            "Month Total"
        ]
        |> List.singleton
        |> SheetsService.updateValuesInRange sheetsService spreadsheetId range

        let range =
            {
                SheetId = Some configuration.SheetId
                Rows = weeksRowRange
                Columns = headerColumnRange
            }
        let dateValues =
            [
                for week in weeks -> [ week.StartDate; week.EndDate ]
            ]
        SheetsService.updateValuesInRange sheetsService spreadsheetId range dateValues

        let weekSumFormulaValues =
            [ 0 .. weeks.Length - 1 ]
            |> List.map (fun weekNumber ->
                {
                    SheetId = Some configuration.SheetId
                    Rows = weeksRowRange |> Range.singleSubrange weekNumber
                    Columns = daysOfWeekColumnRange
                }
                |> SheetFormula.sumofRange
                |> List.singleton)
        let range =
            {
                SheetId = Some configuration.SheetId
                Rows = weeksRowRange
                Columns = weekTotalColumnRange
            }
        SheetsService.updateValuesInRange sheetsService spreadsheetId range weekSumFormulaValues

        let monthSumFormulaValues =
            calendar
            |> Calendar.getWeekNumberRanges
            |> List.collect (fun (startWeekNumber, weekCount) ->
                {
                    SheetId = Some configuration.SheetId
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
                SheetId = Some configuration.SheetId
                Rows = weeksRowRange
                Columns = monthTotalColumnRange
            }
        SheetsService.updateValuesInRange sheetsService spreadsheetId range monthSumFormulaValues

        let dayOfWeekSumFormulaValues =
            [ 2 .. DaysPerWeek + 3 ]
            |> List.map (fun column ->
                {
                    SheetId = Some configuration.SheetId
                    Rows = weeksRowRange
                    Columns = Range.single column
                }
                |> SheetFormula.sumofRange)
            |> List.singleton
        let range =
            {
                SheetId = Some configuration.SheetId
                Rows = totalRowRange
                Columns = dataColumnRange
            }
        SheetsService.updateValuesInRange
            sheetsService
            spreadsheetId
            range
            dayOfWeekSumFormulaValues

    updateValues ()

    let createSingleCellRange (rowIndex, columnIndex) =
        GridRange(
            StartColumnIndex = Nullable columnIndex,
            EndColumnIndex = Nullable(columnIndex + 1),
            StartRowIndex = Nullable rowIndex,
            EndRowIndex = Nullable(rowIndex + 1),
            SheetId = configuration.SheetId
        )

    let setSheetPropertiesRequest =
        SheetsRequests.createSetSheetPropertiesRequest (Some 1, Some 2)

    let createSetDimensionLengthRequests (sheetId, dimension, length) =
        [
            SheetsRequests.createAppendDimensionRequest (sheetId, dimension, length)
            let deleteDimensionRange =
                DimensionRange(Dimension = dimension, StartIndex = length)
            SheetsRequests.createDeleteDimensionRequest deleteDimensionRange
        ]

    let setDimensionLengthRequests =
        [
            yield! createSetDimensionLengthRequests (configuration.SheetId, "COLUMNS", 11)
            yield!
                createSetDimensionLengthRequests (configuration.SheetId, "ROWS", 2 + weeks.Length)
        ]

    let unmergeAllRequest =
        TwoDimensionRange.unbounded (Some configuration.SheetId)
        |> SheetsRequests.createUnmergeCellsRequest

    let mergeCellRequests =
        calendar
        |> Calendar.getWeekNumberRanges
        |> List.map (fun (startWeekNumber, weekCount) ->
            {
                SheetId = Some configuration.SheetId
                Rows =
                    weeksRowRange
                    |> Range.subrangeWithStartAndCount (startWeekNumber, weekCount)
                Columns = monthTotalColumnRange
            })
        |> List.map SheetsRequests.createMergeCellsRequest
        |> List.toArray

    let solidBorder = new Border(Style = "SOLID")
    let outerBorderRequest =
        let range = TwoDimensionRange.unbounded (Some configuration.SheetId)
        SheetsRequests.createUpdateBorderRequest (range, Borders.outer solidBorder)

    let monthBorderRequests =
        calendar
        |> Calendar.getWeekNumberRanges
        |> List.map (fun (startWeekNumber, weekCount) ->
            let range =
                {
                    SheetId = Some configuration.SheetId
                    Rows =
                        weeksRowRange
                        |> Range.subrangeWithStartAndCount (startWeekNumber, weekCount)
                    Columns = Range.unbounded
                }
            SheetsRequests.createUpdateBorderRequest (range, Borders.outer solidBorder))

    let dayOfWeeksBorderRequest =
        let range =
            {
                SheetId = Some configuration.SheetId
                Rows = Range.unbounded
                Columns = daysOfWeekColumnRange
            }
        SheetsRequests.createUpdateBorderRequest (range, Borders.outer solidBorder)

    let weekTotalBorderRequest =
        let range =
            {
                SheetId = Some configuration.SheetId
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
                        let range = createSingleCellRange (weekNumber + 1, dayOfWeekNumber + 2)
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

let sheetsService =
    let credentialFileName = Path.Combine(getRootDirectoryPath (), CredentialFileName)
    SheetsService.create credentialFileName

let configuration = createConfiguration (getRootDirectoryPath ())

configuration.Year
|> Calendar.calculate configuration.FirstDayOfWeek 
|> renderCalendar sheetsService configuration
