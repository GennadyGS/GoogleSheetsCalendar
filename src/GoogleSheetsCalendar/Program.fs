open System
open System.IO
open System.Reflection
open Microsoft.Extensions.Configuration
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

[<RequireQualifiedAccess>]
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

printfn "Loading configuration..."
let configuration = createConfiguration (getRootDirectoryPath ())

printfn "Connecting to Google Sheets..."
let speadsheet =
    let credentialFileName = Path.Combine(getRootDirectoryPath (), CredentialFileName)
    Spreadsheet.createFromCredentialFileName credentialFileName configuration.SpreadsheetId

printfn $"Calculating calendar for year {configuration.Year}..."
let calendar = Calendar.calculate configuration.FirstDayOfWeek configuration.Year

printfn "Updating spreadsheet..."
CalendarRenderer.renderCalendar speadsheet configuration.SheetId calendar

printfn "Done."
