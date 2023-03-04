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

let renderCalendar (sheetsService: SheetsService) configuration (Calendar months) =
    let spreadsheetId = configuration.SpreadsheetId

    sheetsService
        .Spreadsheets
        .Values
        .Clear(ClearValuesRequest(), spreadsheetId, "A1:ZZ")
        .Execute()
    |> ignore

    let clearFormattingRequest = Request()
    clearFormattingRequest.UpdateCells <-
        UpdateCellsRequest(
            Range = GridRange(SheetId = configuration.SheetId),
            Fields = nameof (Unchecked.defaultof<CellData>.UserEnteredFormat)
        )

    let updateRequestBody =
        BatchUpdateSpreadsheetRequest(
            Requests =
                [|
                    clearFormattingRequest
                |]
        )

    sheetsService
        .Spreadsheets
        .BatchUpdate(updateRequestBody, spreadsheetId)
        .Execute()
    |> ignore

    let values =
        months
        |> List.collect (fun (Month weeks) -> weeks)
        |> List.map (fun week -> [ week.StartDate; week.EndDate ])
        |> List.map (fun list -> list |> List.toArray |> Array.map box :> IList<obj>)
        |> List.toArray

    let updateData =
        [|
            ValueRange(Range = "R2C1:C2", Values = values)
        |]

    let valueUpdateRequestBody =
        BatchUpdateValuesRequest(ValueInputOption = "USER_ENTERED", Data = updateData)

    sheetsService
        .Spreadsheets
        .Values
        .BatchUpdate(valueUpdateRequestBody, spreadsheetId)
        .Execute()
    |> ignore

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

    let range =
        GridRange(
            StartColumnIndex = 2,
            EndColumnIndex = 3,
            StartRowIndex = 1,
            EndRowIndex = 2,
            SheetId = configuration.SheetId
        )
    let greyColor = Color(Red = 0.75f, Green = 0.75f, Blue = 0.75f)
    let updateCellFormatRequest = createSetBackgroundColorRequest range greyColor

    let updateRequestBody =
        BatchUpdateSpreadsheetRequest(
            Requests =
                [|
                    updateCellFormatRequest
                |]
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
