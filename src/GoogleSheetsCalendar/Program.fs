open System
open System.IO
open System.Reflection
open Microsoft.Extensions.Configuration
open Google.Apis.Auth.OAuth2
open Google.Apis.Services
open Google.Apis.Sheets.v4
open Google.Apis.Sheets.v4.Data

module NullCoalesce =
    let coalesce (b: 'a Lazy) (a: 'a) =
        if obj.ReferenceEquals(a, null) then b.Value else a

[<CLIMutable>]
type Configuration = { SpreadsheetId: string }

[<Literal>]
let credentialsFileName = "credentials.json"

let executableFilePath = Assembly.GetExecutingAssembly().Location
let executableDirectoryPath = Path.GetDirectoryName(executableFilePath)

let configuration =
    let nativeConfiguration =
        (new ConfigurationBuilder())
            .SetBasePath(executableDirectoryPath)
            .AddJsonFile("appsettings.json", true, true)
            .AddUserSecrets<Configuration>()
            .Build()

    nativeConfiguration.Get<Configuration>()
    |> NullCoalesce.coalesce (lazy raise (InvalidOperationException("Configuration is missing")))

let credentialFileName = Path.Combine(executableDirectoryPath, credentialsFileName)

if not (File.Exists(credentialFileName)) then
    raise (InvalidOperationException($"File {credentialsFileName} is missing"))

let credential =
    GoogleCredential
        .FromFile(credentialFileName)
        .CreateScoped([| SheetsService.Scope.Spreadsheets |])

let initializer =
    new BaseClientService.Initializer(HttpClientInitializer = credential)

let sheetsService = new SheetsService(initializer)

let spreadsheetId = configuration.SpreadsheetId
let updateData =
    [|
        ValueRange(
            Range = "Sheet1",
            Values = [|[| 2; 3 |]; [| 4; 5 |];|]
        )
    |]

let valueUpdateRequestBody = new BatchUpdateValuesRequest(
    ValueInputOption = "USER_ENTERED",
    Data = updateData
)

sheetsService.Spreadsheets.Values.BatchUpdate(valueUpdateRequestBody, spreadsheetId).Execute()
|> ignore
