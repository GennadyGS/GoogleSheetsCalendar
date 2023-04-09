module GoogleSheets

open System
open System.IO
open Google.Apis.Auth.OAuth2
open Google.Apis.Services
open Google.Apis.Sheets.v4
open Google.Apis.Sheets.v4.Data
open System.Collections.Generic

type Range =
    {
        StartIndex: int option
        EndIndex: int option
    }

type TwoDimensionRange =
    {
        Columns: Range
        Rows: Range
        SheetId: int option
    }

type Borders =
    {
        Left: Border option
        Right: Border option
        Top: Border option
        Bottom: Border option
    }

module Borders =
    let none =
        {
            Left = None
            Right = None
            Top = None
            Bottom = None
        }

    let outer border =
        {
            Left = Some border
            Right = Some border
            Top = Some border
            Bottom = Some border
        }

[<RequireQualifiedAccess>]
module Range =
    let getStartIndexValue range =
        range.StartIndex |> Option.defaultValue 0

    let getEndIndexValue range =
        range.EndIndex
        |> Option.defaultWith (fun _ -> failwith "End index is not defined.")

    let unbounded = { StartIndex = None; EndIndex = None }

    let startingFrom index =
        {
            StartIndex = Some index
            EndIndex = None
        }

    let endingWith index =
        {
            StartIndex = None
            EndIndex = Some index
        }

    let fromBounds (startIndex, endIndex) =
        {
            StartIndex = Some startIndex
            EndIndex = Some endIndex
        }

    let single index = fromBounds (index, index)

    let fromStartAndCount (startIndex, count) =
        fromBounds (startIndex, startIndex + count - 1)

    let nextRangeWithCount count range =
        let endIndexValue = getEndIndexValue range
        fromBounds (endIndexValue + 1, endIndexValue + count)

    let nextSingleRange range = nextRangeWithCount 1 range

    let subrangeWithBounds (startIndex, endIndex) range =
        let startIndexValue = getStartIndexValue range
        fromBounds (startIndexValue + startIndex, startIndexValue + endIndex)

    let subrangeWithStartAndCount (startIndex, count) range =
        subrangeWithBounds (startIndex, startIndex + count - 1) range

    let singleSubrange index range = subrangeWithBounds (index, index) range

    let union (range1, range2) =
        let range1EndIndexValue = getEndIndexValue range1
        let range2StartIndexValue = getStartIndexValue range2
        if (range2StartIndexValue - range1EndIndexValue <> 1) then
            failwith "Ranges should be adjacent."
        {
            StartIndex = range1.StartIndex
            EndIndex = range2.EndIndex
        }

    let unionAll ranges =
        ranges
        |> List.reduce (fun range1 range2 -> union (range1, range2))

[<RequireQualifiedAccess>]
module TwoDimensionRange =
    let unbounded sheetId =
        {
            Rows = Range.unbounded
            Columns = Range.unbounded
            SheetId = sheetId
        }

    let toGridRange (range: TwoDimensionRange) =
        let mapStartIndex index = index |> Option.toNullable
        let mapEndIndex index =
            index |> Option.map ((+) 1) |> Option.toNullable
        GridRange(
            StartColumnIndex = (mapStartIndex range.Columns.EndIndex),
            EndColumnIndex = (mapEndIndex range.Columns.EndIndex),
            StartRowIndex = (mapStartIndex range.Rows.StartIndex),
            EndRowIndex = (mapEndIndex range.Columns.EndIndex),
            SheetId = (range.SheetId |> Option.toNullable)
        )

    let toString (range: TwoDimensionRange) =
        let indexToString dimensionTag index =
            match index with
            | Some value -> dimensionTag + string (value + 1)
            | None -> String.Empty
        let referenceToString (columnIndex, rowIndex) =
            let columnIndexString = columnIndex |> indexToString "C"
            let rowIndexString = rowIndex |> indexToString "R"
            rowIndexString + columnIndexString
        let startReference =
            let makeExplicit index = index |> Option.defaultValue 0 |> Some
            referenceToString (
                range.Columns.StartIndex |> makeExplicit,
                range.Rows.StartIndex |> makeExplicit
            )
        let endReference = referenceToString (range.Columns.EndIndex, range.Rows.EndIndex)
        $"{startReference}:{endReference}"

[<RequireQualifiedAccess>]
module SheetExpression =
    let rangeReference range =
        let rangeString = TwoDimensionRange.toString range
        $"INDIRECT(\"{rangeString}\", FALSE)"

    let sum (expression: string) = $"SUM({expression})"

[<RequireQualifiedAccess>]
module SheetFormula =
    let fromExpression (expression: string) = $"={expression}"

    let sumofRange range =
        range
        |> SheetExpression.rangeReference
        |> SheetExpression.sum
        |> fromExpression

[<RequireQualifiedAccess>]
module SheetsService =
    let create credentialFileName =
        if not (File.Exists(credentialFileName)) then
            raise (InvalidOperationException($"File {credentialFileName} is missing"))

        let credential =
            GoogleCredential
                .FromFile(credentialFileName)
                .CreateScoped([| SheetsService.Scope.Spreadsheets |])

        let initializer =
            new BaseClientService.Initializer(HttpClientInitializer = credential)

        new SheetsService(initializer)

    let batchUpdate (sheetsService: SheetsService) spreadsheetId requests =
        let requestBody =
            BatchUpdateSpreadsheetRequest(Requests = (requests |> List.toArray))

        sheetsService
            .Spreadsheets
            .BatchUpdate(requestBody, spreadsheetId)
            .Execute()
        |> ignore

    let batchUpdateValuesInRange (sheetsService: SheetsService) spreadsheetId rangesAndValues =
        let convertValues values =
            values
            |> List.toArray
            |> Array.map (
                List.toArray
                >> Array.map box
                >> (fun array -> array :> IList<obj>)
            )

        let valueRanges =
            rangesAndValues
            |> List.map (fun (range, values) ->
                ValueRange(Range = TwoDimensionRange.toString range, Values = convertValues values))
            |> List.toArray

        let valueUpdateRequestBody =
            BatchUpdateValuesRequest(ValueInputOption = "USER_ENTERED", Data = valueRanges)

        sheetsService
            .Spreadsheets
            .Values
            .BatchUpdate(valueUpdateRequestBody, spreadsheetId)
            .Execute()
        |> ignore

    let updateValuesInRange (sheetsService: SheetsService) spreadsheetId range values =
        batchUpdateValuesInRange sheetsService spreadsheetId [ (range, values) ]
