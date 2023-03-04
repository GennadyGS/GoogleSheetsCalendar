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

    let values =
        months
        |> List.collect (fun (Month weeks) -> weeks)
        |> List.map (fun week -> [ week.StartDate; week.EndDate ])
        |> List.map (fun list -> list |> List.toArray |> Array.map box :> IList<obj>)
        |> List.toArray

    let updateData =
        [|
            ValueRange(Range = "A2:B", Values = values)
        |]

    let valueUpdateRequestBody =
        new BatchUpdateValuesRequest(ValueInputOption = "USER_ENTERED", Data = updateData)

    sheetsService
        .Spreadsheets
        .Values
        .BatchUpdate(valueUpdateRequestBody, spreadsheetId)
        .Execute()
    |> ignore

let rootDirectoryPath = getRootDirectoryPath ()

let configuration = createConfiguration rootDirectoryPath

let sheetsService = createSheetsService rootDirectoryPath

let calendar = calculateCalendar configuration.FirstDayOfWeek configuration.Year

renderCalendar sheetsService configuration calendar
