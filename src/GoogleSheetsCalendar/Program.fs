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
        let createClearFormattingRequest sheetId =
            let request = Request()
            request.UpdateCells <-
                UpdateCellsRequest(
                    Range = GridRange(SheetId = (sheetId |> Option.toNullable)),
                    Fields = nameof (Unchecked.defaultof<CellData>.UserEnteredFormat)
                )
            request

        let clearFormattingRequest =
            Some configuration.SheetId
            |> createClearFormattingRequest

        let updateRequestBody =
            BatchUpdateSpreadsheetRequest(Requests = [| clearFormattingRequest |])

        sheetsService
            .Spreadsheets
            .BatchUpdate(updateRequestBody, spreadsheetId)
            .Execute()
        |> ignore

    clearFormatting ()

    let updateValuesInRange range values =
        let valueArray =
            values
            |> List.toArray
            |> Array.map (
                List.toArray
                >> Array.map box
                >> (fun array -> array :> IList<obj>)
            )
        let updateData =
            [|
                ValueRange(Range = TwoDimensionRange.toString range, Values = valueArray)
            |]

        let valueUpdateRequestBody =
            BatchUpdateValuesRequest(ValueInputOption = "USER_ENTERED", Data = updateData)

        sheetsService
            .Spreadsheets
            .Values
            .BatchUpdate(valueUpdateRequestBody, spreadsheetId)
            .Execute()
        |> ignore

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
                Rows = headerRowRange
                Columns = Range.all
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
        |> updateValuesInRange range

        let range =
            {
                Rows = weeksRowRange
                Columns = headerColumnRange
            }
        let dateValues =
            [
                for week in weeks -> [ week.StartDate; week.EndDate ]
            ]
        updateValuesInRange range dateValues

        let weekSumFormulaValues =
            [ 0 .. weeks.Length - 1 ]
            |> List.map (fun weekNumber ->
                {
                    Rows = weeksRowRange |> Range.singleSubrange weekNumber
                    Columns = daysOfWeekColumnRange
                }
                |> SheetFormula.sumofRange
                |> List.singleton)
        let range =
            {
                Rows = weeksRowRange
                Columns = weekTotalColumnRange
            }
        updateValuesInRange range weekSumFormulaValues

        let monthSumFormulaValues =
            calendar
            |> Calendar.getWeekNumberRanges
            |> List.collect (fun (startWeekNumber, weekCount) ->
                {
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
                Rows = weeksRowRange
                Columns = monthTotalColumnRange
            }
        updateValuesInRange range monthSumFormulaValues

        let dayOfWeekSumFormulaValues =
            [ 2 .. DaysPerWeek + 3 ]
            |> List.map (fun column ->
                {
                    Rows = weeksRowRange
                    Columns = Range.single column
                }
                |> SheetFormula.sumofRange)
            |> List.singleton
        let range =
            {
                Rows = totalRowRange
                Columns = dataColumnRange
            }
        updateValuesInRange range dayOfWeekSumFormulaValues

    updateValues ()

    let createSingleCellRange (rowIndex, columnIndex) =
        GridRange(
            StartColumnIndex = Nullable columnIndex,
            EndColumnIndex = Nullable(columnIndex + 1),
            StartRowIndex = Nullable rowIndex,
            EndRowIndex = Nullable(rowIndex + 1),
            SheetId = configuration.SheetId
        )
    let greyColor = Color(Red = 0.75f, Green = 0.75f, Blue = 0.75f)

    let createSetSheetPropertiesRequest (frozenRowCount, frozenColumnCount) =
        let request = new Request()
        let sheetProperties =
            SheetProperties(
                GridProperties =
                    GridProperties(
                        FrozenRowCount = Option.toNullable frozenRowCount,
                        FrozenColumnCount = Option.toNullable frozenColumnCount
                    )
            )

        request.UpdateSheetProperties <-
            UpdateSheetPropertiesRequest(
                Properties = sheetProperties,
                Fields =
                    ([
                        $"{nameof (sheetProperties.GridProperties)}.{nameof (sheetProperties.GridProperties.FrozenRowCount)}"
                        $"{nameof (sheetProperties.GridProperties)}.{nameof (sheetProperties.GridProperties.FrozenColumnCount)}"
                     ]
                     |> String.concat ",")
            )
        request

    let setSheetPropertiesRequest = createSetSheetPropertiesRequest (Some 1, Some 2)

    let createDeleteDimensionRequest dimensionRange =
        let result = new Request()
        result.DeleteDimension <- DeleteDimensionRequest(Range = dimensionRange)
        result

    let createAppendDimensionRequest (sheetId: int, dimension, length: int) =
        let result = new Request()
        result.AppendDimension <-
            AppendDimensionRequest(SheetId = sheetId, Dimension = dimension, Length = length)
        result

    let createSetDimensionLengthRequests (sheetId, dimension, length) =
        [
            createAppendDimensionRequest (sheetId, dimension, length)
            let deleteDimensionRange =
                DimensionRange(Dimension = dimension, StartIndex = length)
            createDeleteDimensionRequest deleteDimensionRange
        ]

    let setDimensionLengthRequests =
        [
            yield! createSetDimensionLengthRequests (configuration.SheetId, "COLUMNS", 11)
            yield!
                createSetDimensionLengthRequests (configuration.SheetId, "ROWS", 2 + weeks.Length)
        ]

    let createUnmergeCellsRequest gridRange =
        let result = new Request()
        result.UnmergeCells <- UnmergeCellsRequest(Range = gridRange)
        result

    let unmergeAllRequest =
        GridRange(SheetId = configuration.SheetId)
        |> createUnmergeCellsRequest

    let createMergeCellsRequest gridRange =
        let result = new Request()
        result.MergeCells <- MergeCellsRequest(MergeType = "MERGE_ALL", Range = gridRange)
        result

    let mergeCellRequests =
        calendar
        |> Calendar.getWeekNumberRanges
        |> List.map (fun (startWeekNumber, weekCount) ->
            GridRange(
                SheetId = configuration.SheetId,
                StartRowIndex = startWeekNumber + 1,
                EndRowIndex = startWeekNumber + 1 + weekCount,
                StartColumnIndex = 10,
                EndColumnIndex = 11
            )
            |> createMergeCellsRequest)
        |> List.toArray

    let createUpdateBorderRequest (range, borders) =
        let updateBordersRequest = new Request()
        updateBordersRequest.UpdateBorders <-
            UpdateBordersRequest(
                Range = range,
                Left = Option.defaultValue null borders.Left,
                Right = Option.defaultValue null borders.Right,
                Top = Option.defaultValue null borders.Top,
                Bottom = Option.defaultValue null borders.Bottom
            )
        updateBordersRequest

    let solidBorder = new Border(Style = "SOLID")
    let outerBorderRequest =
        let range =
            GridRange(
                StartColumnIndex = Nullable 0,
                EndColumnIndex = Nullable 11,
                StartRowIndex = 0,
                EndRowIndex = weeks.Length + 2,
                SheetId = configuration.SheetId
            )
        createUpdateBorderRequest (range, Borders.outer solidBorder)

    let monthBorderRequests =
        calendar
        |> Calendar.getWeekNumberRanges
        |> List.map (fun (startWeekNumber, weekCount) ->
            let range =
                GridRange(
                    StartColumnIndex = Nullable 0,
                    EndColumnIndex = Nullable 11,
                    StartRowIndex = startWeekNumber + 1,
                    EndRowIndex = startWeekNumber + weekCount + 1,
                    SheetId = configuration.SheetId
                )
            createUpdateBorderRequest (range, Borders.outer solidBorder))

    let dayOfWeeksBorderRequest =
        let range =
            GridRange(
                StartColumnIndex = Nullable 2,
                EndColumnIndex = Nullable(2 + DaysPerWeek),
                StartRowIndex = 0,
                EndRowIndex = weeks.Length + 2,
                SheetId = configuration.SheetId
            )
        createUpdateBorderRequest (range, Borders.outer solidBorder)

    let monthTotalBorderRequest =
        let range =
            GridRange(
                StartColumnIndex = Nullable(2 + DaysPerWeek),
                EndColumnIndex = Nullable(3 + DaysPerWeek),
                StartRowIndex = 0,
                EndRowIndex = weeks.Length + 2,
                SheetId = configuration.SheetId
            )
        createUpdateBorderRequest (range, Borders.outer solidBorder)

    let setBordersRequests =
        [
            outerBorderRequest
            yield! monthBorderRequests
            dayOfWeeksBorderRequest
            monthTotalBorderRequest
        ]

    let createSetBackgroundColorRequest gridRange color =
        let updateCellFormatRequest = Request()
        updateCellFormatRequest.RepeatCell <-
            let cellFormat = CellFormat(BackgroundColor = color)
            let cellData = CellData(UserEnteredFormat = cellFormat)
            RepeatCellRequest(
                Range = gridRange,
                Cell = cellData,
                Fields =
                    $"{nameof (cellData.UserEnteredFormat)}.{nameof (cellFormat.BackgroundColor)}"
            )
        updateCellFormatRequest

    let setCellBackgroundColorRequests =
        [|
            for (weekNumber, week) in List.indexed weeks do
                for dayOfWeekNumber in [ 0 .. DaysPerWeek - 1 ] do
                    if not week.DaysActive[dayOfWeekNumber] then
                        let range = createSingleCellRange (weekNumber + 1, dayOfWeekNumber + 2)
                        createSetBackgroundColorRequest range greyColor
        |]

    let updateRequestBody =
        BatchUpdateSpreadsheetRequest(
            Requests =
                [|
                    setSheetPropertiesRequest
                    yield! setDimensionLengthRequests
                    unmergeAllRequest
                    yield! mergeCellRequests
                    yield! setBordersRequests
                    yield! setCellBackgroundColorRequests
                |]
        )

    sheetsService
        .Spreadsheets
        .BatchUpdate(updateRequestBody, spreadsheetId)
        .Execute()
    |> ignore

let rootDirectoryPath = getRootDirectoryPath ()

let credentialFileName = Path.Combine(rootDirectoryPath, CredentialFileName)

let sheetsService = SheetsService.create credentialFileName

let configuration = createConfiguration rootDirectoryPath

let calendar = Calendar.calculate configuration.FirstDayOfWeek configuration.Year

renderCalendar sheetsService configuration calendar
