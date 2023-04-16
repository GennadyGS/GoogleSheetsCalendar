open System
open System.IO
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

let sheetsService =
    let credentialFileName = Path.Combine(getRootDirectoryPath (), CredentialFileName)
    SheetsService.create credentialFileName

let configuration = createConfiguration (getRootDirectoryPath ())

configuration.Year
|> Calendar.calculate configuration.FirstDayOfWeek
|> CalendarRenderer.renderCalendar
    sheetsService
    (configuration.SpreadsheetId, configuration.SheetId)
