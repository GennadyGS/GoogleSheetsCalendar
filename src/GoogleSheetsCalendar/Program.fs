open System
open System.IO
open System.Collections.Generic
open System.Reflection
open Microsoft.Extensions.Configuration
open Google.Apis.Auth.OAuth2
open Google.Apis.Services
open Google.Apis.Sheets.v4
open Google.Apis.Sheets.v4.Data

module NullCoalesce =
    let coalesce (b: 'a Lazy) (a: 'a) =
        if obj.ReferenceEquals(a, null) then
            b.Value
        else
            a

[<CLIMutable>]
type Configuration = { SpreadsheetId: string; Year: int }

[<Literal>]
let credentialsFileName = "credentials.json"

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

type Week =
    {
        StartDate: DateOnly
        EndDate: DateOnly
        DaysActive: bool array
    }

type Month = Month of Week list

type Calendar = Calendar of Month list

let calculateCalendar year =
    let daysPerWeek = Enum.GetValues<DayOfWeek>().Length
    let calculateMonth month =
        let dayCount = DateTime.DaysInMonth(year, month)
        let startDate = DateOnly(year, month, 1)
        let startDayOfWeek = startDate.DayOfWeek |> int
        [ 1..dayCount ]
        |> List.groupBy (fun day -> (day - 1 + startDayOfWeek) / daysPerWeek)
        |> List.map (fun (weekNumber, _) ->
            let startDay = weekNumber * daysPerWeek - startDayOfWeek
            {
                StartDate = startDate.AddDays(startDay)
                EndDate = startDate.AddDays(startDay + daysPerWeek - 1)
                DaysActive =
                    Array.init daysPerWeek (fun dayOfWeek ->
                        startDay + dayOfWeek > 0
                        && startDay + dayOfWeek < dayCount)
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

let calendar = calculateCalendar configuration.Year

renderCalendar sheetsService configuration calendar
