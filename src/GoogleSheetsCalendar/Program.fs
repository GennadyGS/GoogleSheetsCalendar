open System
open System.IO
open System.Collections.Generic
open System.Reflection
open Microsoft.Extensions.Configuration
open Google.Apis.Auth.OAuth2
open Google.Apis.Services
open Google.Apis.Sheets.v4
open Google.Apis.Sheets.v4.Data

type Week =
    {
        StartDate: DateOnly
        EndDate: DateOnly
        DaysActive: bool array
    }

type Month = Month of Week list

type Calendar = Calendar of Month list

[<CLIMutable>]
type Configuration =
    {
        SpreadsheetId: string
        SheetId: int
        Year: int
        FirstDayOfWeek: DayOfWeek
    }

[<Literal>]
let credentialsFileName = "credentials.json"

let DaysPerWeek = Enum.GetValues<DayOfWeek>().Length

module NullCoalesce =
    let coalesce (b: 'a Lazy) (a: 'a) =
        if obj.ReferenceEquals(a, null) then
            b.Value
        else
            a

module Int =
    let between (min, max) x = x >= min && x <= max

module DayOfWeek =
    let diff (x: DayOfWeek, y: DayOfWeek) =
        (int x - int y + DaysPerWeek) % DaysPerWeek

module Calendar =
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
        |> List.map (fun (weeks, nextWeekStartNumber) ->
            (nextWeekStartNumber - weeks.Length, weeks.Length))

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

let createSheetsService rootDirectoryPath =
    let credentialFileName = Path.Combine(rootDirectoryPath, credentialsFileName)

    if not (File.Exists(credentialFileName)) then
        raise (InvalidOperationException($"File {credentialsFileName} is missing"))

    let credential =
        GoogleCredential
            .FromFile(credentialFileName)
            .CreateScoped([| SheetsService.Scope.Spreadsheets |])

    let initializer =
        new BaseClientService.Initializer(HttpClientInitializer = credential)

    new SheetsService(initializer)

let calculateCalendar firstDayOfWeeek year =
    let calculateMonth month =
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

    let monthCount = DateOnly(year + 1, 1, 1).AddDays(-1).Month
    [ 1..monthCount ]
    |> List.map calculateMonth
    |> Calendar

let renderCalendar (sheetsService: SheetsService) configuration calendar =
    let weeks = Calendar.getWeeks calendar
    let spreadsheetId = configuration.SpreadsheetId

    let clearFormattingRequest = Request()
    clearFormattingRequest.UpdateCells <-
        UpdateCellsRequest(
            Range = GridRange(SheetId = configuration.SheetId),
            Fields = nameof (Unchecked.defaultof<CellData>.UserEnteredFormat)
        )

    let updateRequestBody =
        BatchUpdateSpreadsheetRequest(Requests = [| clearFormattingRequest |])

    sheetsService
        .Spreadsheets
        .BatchUpdate(updateRequestBody, spreadsheetId)
        .Execute()
    |> ignore

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
                ValueRange(Range = range, Values = valueArray)
            |]

        let valueUpdateRequestBody =
            BatchUpdateValuesRequest(ValueInputOption = "USER_ENTERED", Data = updateData)

        sheetsService
            .Spreadsheets
            .Values
            .BatchUpdate(valueUpdateRequestBody, spreadsheetId)
            .Execute()
        |> ignore

    let dateValues =
        [
            for week in weeks -> [ week.StartDate; week.EndDate ]
        ]
    updateValuesInRange "R2C1:C2" dateValues

    let weekSumFormula = @"=SUM(INDIRECT(""R[0]C[-7]:R[0]C[-1]"", FALSE))"
    let weekSumFormulaValues = List.replicate weeks.Length [ weekSumFormula ]
    updateValuesInRange "R2C10:C10" weekSumFormulaValues

    let monthSumFormulaValues =
        calendar
        |> Calendar.getWeekNumberRanges
        |> List.collect (fun (startWeekNumber, weekCount) ->
            @$"=SUM(INDIRECT(""R{startWeekNumber + 2}C[-1]:R{startWeekNumber + 2 + weekCount - 1}C[-1]"", FALSE))"
            |> List.singleton
            |> List.replicate weekCount)

    updateValuesInRange "R2C11:C11" monthSumFormulaValues

    let dayOfweekSumFormula =
        @$"=SUM(INDIRECT(""R2C[0]:R{weeks.Length + 1}C[0]"", FALSE))"
    let dayOfWeekSumFormulaValues =
        [
            List.replicate (DaysPerWeek + 2) dayOfweekSumFormula
        ]
    updateValuesInRange $"R{weeks.Length + 2}C3:C12" dayOfWeekSumFormulaValues

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

    let createSingleCellRange (rowIndex, columnIndex) =
        GridRange(
            StartColumnIndex = Nullable columnIndex,
            EndColumnIndex = Nullable(columnIndex + 1),
            StartRowIndex = Nullable rowIndex,
            EndRowIndex = Nullable(rowIndex + 1),
            SheetId = configuration.SheetId
        )
    let greyColor = Color(Red = 0.75f, Green = 0.75f, Blue = 0.75f)

    let updateCellFormatRequests =
        [|
            for (weekNumber, week) in List.indexed weeks do
                for dayOfWeekNumber in [ 0 .. DaysPerWeek - 1 ] do
                    if not week.DaysActive[dayOfWeekNumber] then
                        let range = createSingleCellRange (weekNumber + 1, dayOfWeekNumber + 2)
                        createSetBackgroundColorRequest range greyColor
        |]

    let createMergeCellsRequest gridRange =
        let result = new Request()
        result.MergeCells <- MergeCellsRequest(MergeType = "MERGE_ALL", Range = gridRange)
        result

    let mergeRequests =
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

    let updateRequestBody =
        BatchUpdateSpreadsheetRequest(
            Requests = Array.append updateCellFormatRequests mergeRequests
        )

    sheetsService
        .Spreadsheets
        .BatchUpdate(updateRequestBody, spreadsheetId)
        .Execute()
    |> ignore

let rootDirectoryPath = getRootDirectoryPath ()

let configuration = createConfiguration rootDirectoryPath

let sheetsService = createSheetsService rootDirectoryPath

let calendar = calculateCalendar configuration.FirstDayOfWeek configuration.Year

renderCalendar sheetsService configuration calendar
