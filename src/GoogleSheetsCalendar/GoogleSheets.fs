﻿module GoogleSheets

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

type GridRange =
    {
        Columns: Range
        Rows: Range
        SheetId: int option
    }

type Dimension =
    | Rows
    | Columns
    override this.ToString() =
        match this with
        | Rows -> "ROWS"
        | Columns -> "COLUMNS"

type DimensionRange =
    {
        Dimension: Dimension
        SheetId: int option
        Range: Range
    }

type AggregationFunction =
    | Sum
    | Avg
    | Min
    | Max
    override this.ToString() =
        match this with
        | Sum -> "SUM"
        | Avg -> "AVG"
        | Min -> "MIN"
        | Max -> "MAX"

type SheetProperties =
    {
        SheetId: int option
        FrozenRowCount: int option
        FrozenColumnCount: int option
    }

type Borders =
    {
        Left: Border option
        Right: Border option
        Top: Border option
        Bottom: Border option
    }

type Spreadsheet =
    internal
        {
            SheetsService: SheetsService
            SpreadsheetId: string
        }

[<RequireQualifiedAccess>]
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

    let getCount range =
        getEndIndexValue range - getStartIndexValue range
        + 1

    let getIndexValues range =
        [
            getStartIndexValue range .. getEndIndexValue range
        ]

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

    let withBounds (startIndex, endIndex) =
        {
            StartIndex = Some startIndex
            EndIndex = Some endIndex
        }

    let single index = withBounds (index, index)

    let withStartAndCount (startIndex, count) =
        withBounds (startIndex, startIndex + count - 1)

    let nextRangeWithCount count range =
        let endIndexValue = getEndIndexValue range
        withBounds (endIndexValue + 1, endIndexValue + count)

    let nextSingleRange range = nextRangeWithCount 1 range

    let subrangeWithBounds (startIndex, endIndex) range =
        let startIndexValue = getStartIndexValue range
        withBounds (startIndexValue + startIndex, startIndexValue + endIndex)

    let subrangeWithStartAndCount (startIndex, count) range =
        subrangeWithBounds (startIndex, startIndex + count - 1) range

    let subrangeSingle index range = subrangeWithBounds (index, index) range

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
module GridRange =
    let create sheetId (rows, columns) =
        {
            SheetId = sheetId
            Rows = rows
            Columns = columns
        }

    let unbounded sheetId =
        {
            SheetId = sheetId
            Rows = Range.unbounded
            Columns = Range.unbounded
        }

    let toApiGridRange (range: GridRange) =
        let mapStartIndex index = index |> Option.toNullable
        let mapEndIndex index =
            index |> Option.map ((+) 1) |> Option.toNullable
        GridRange(
            StartColumnIndex = (mapStartIndex range.Columns.StartIndex),
            EndColumnIndex = (mapEndIndex range.Columns.EndIndex),
            StartRowIndex = (mapStartIndex range.Rows.StartIndex),
            EndRowIndex = (mapEndIndex range.Rows.EndIndex),
            SheetId = (range.SheetId |> Option.toNullable)
        )

    let toString (range: GridRange) =
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
module DimensionRange =
    let create sheetId (dimension, range) =
        {
            DimensionRange.SheetId = sheetId
            Dimension = dimension
            Range = range
        }

    let toApiDimensionRange (dimensionRange: DimensionRange) =
        DimensionRange(
            SheetId = Option.toNullable dimensionRange.SheetId,
            Dimension = string dimensionRange.Dimension,
            StartIndex = Option.toNullable dimensionRange.Range.StartIndex,
            EndIndex = Option.toNullable dimensionRange.Range.EndIndex
        )

[<RequireQualifiedAccess>]
module ValueRange =
    let box (range, values) =
        (range, values |> List.map (List.map box))

[<RequireQualifiedAccess>]
module SheetProperties =
    let createDefault sheetId =
        {
            SheetProperties.SheetId = sheetId
            FrozenRowCount = None
            FrozenColumnCount = None
        }

    let toApiSheetProperties sheetProperties =
        SheetProperties(
            SheetId = Option.toNullable sheetProperties.SheetId,
            GridProperties =
                GridProperties(
                    FrozenRowCount = Option.toNullable sheetProperties.FrozenRowCount,
                    FrozenColumnCount = Option.toNullable sheetProperties.FrozenColumnCount
                )
        )

    let getFieldNames () =
        let defaultApiValue = Unchecked.defaultof<Google.Apis.Sheets.v4.Data.SheetProperties>
        let gridPropertiesName = nameof(defaultApiValue.GridProperties)
        let defaultApiGridPropertiesValue = Unchecked.defaultof<GridProperties>
        [
            $"{gridPropertiesName}.{nameof (defaultApiGridPropertiesValue.FrozenRowCount)}"
            $"{gridPropertiesName}.{nameof (defaultApiGridPropertiesValue.FrozenColumnCount)}"
        ]

[<RequireQualifiedAccess>]
module SheetExpression =
    let rangeReference range =
        let rangeString = GridRange.toString range
        $"INDIRECT(\"{rangeString}\", FALSE)"

    let aggregate aggregationFunction range =
        let functionIdentifier = string aggregationFunction
        $"{functionIdentifier}({rangeReference range})"

[<RequireQualifiedAccess>]
module SheetFormulaValue =
    let fromExpression (expression: string) = $"={expression}"

    let aggregate aggregationFunction range =
        range
        |> SheetExpression.aggregate aggregationFunction
        |> fromExpression

[<RequireQualifiedAccess>]
module SheetFormulaValues =
    let rowWiseAggregation aggregationFunction (gridRange: GridRange) =
        gridRange.Rows
        |> Range.getIndexValues
        |> List.map (fun rowIndex ->
            { gridRange with
                Rows = Range.single rowIndex
            }
            |> SheetFormulaValue.aggregate AggregationFunction.Sum
            |> List.singleton)

    let columnWiseAggregation aggregationFunction (gridRange: GridRange) =
        gridRange.Columns
        |> Range.getIndexValues
        |> List.map (fun columnIndex ->
            { gridRange with
                Columns = Range.single columnIndex
            }
            |> SheetFormulaValue.aggregate AggregationFunction.Sum)
        |> List.singleton

[<RequireQualifiedAccess>]
module Color =
    let grey intencity =
        Color(Red = Nullable intencity, Green = Nullable intencity, Blue = Nullable intencity)

[<RequireQualifiedAccess>]
module SheetsRequests =
    let createClearFormattingRequest range =
        let request = Request()
        request.UpdateCells <-
            UpdateCellsRequest(
                Range = (GridRange.toApiGridRange range),
                Fields = nameof (Unchecked.defaultof<CellData>.UserEnteredFormat)
            )
        request

    let createSetSheetPropertiesRequest (sheetProperties: SheetProperties) =
        let request = new Request()
        let apiSheetProperties = SheetProperties.toApiSheetProperties sheetProperties
        request.UpdateSheetProperties <-
            UpdateSheetPropertiesRequest(
                Properties = apiSheetProperties,
                Fields = (SheetProperties.getFieldNames () |> String.concat ","))
        request

    let createDeleteDimensionRequest dimensionRange =
        let result = new Request()
        result.DeleteDimension <-
            DeleteDimensionRequest(Range = (DimensionRange.toApiDimensionRange dimensionRange))
        result

    let createAppendDimensionRequest (sheetId, dimension: Dimension, length: int) =
        let result = new Request()
        result.AppendDimension <-
            AppendDimensionRequest(
                SheetId = (Option.toNullable sheetId),
                Dimension = string dimension,
                Length = length
            )
        result

    let createSetDimensionLengthRequests (sheetId, dimension: Dimension, length) =
        let deleteDimensionRange =
            DimensionRange.create sheetId (dimension, Range.startingFrom length)
        [
            createAppendDimensionRequest (sheetId, dimension, length)
            createDeleteDimensionRequest deleteDimensionRange
        ]

    let createUnmergeCellsRequest range =
        let result = new Request()
        result.UnmergeCells <- UnmergeCellsRequest(Range = (GridRange.toApiGridRange range))
        result

    let createMergeCellsRequest range =
        let result = new Request()
        result.MergeCells <-
            MergeCellsRequest(MergeType = "MERGE_ALL", Range = (GridRange.toApiGridRange range))
        result

    let createUpdateBorderRequest (range, borders) =
        let updateBordersRequest = new Request()
        updateBordersRequest.UpdateBorders <-
            UpdateBordersRequest(
                Range = (range |> GridRange.toApiGridRange),
                Left = Option.defaultValue null borders.Left,
                Right = Option.defaultValue null borders.Right,
                Top = Option.defaultValue null borders.Top,
                Bottom = Option.defaultValue null borders.Bottom
            )
        updateBordersRequest

    let createSetBackgroundColorRequest range color =
        let updateCellFormatRequest = Request()
        updateCellFormatRequest.RepeatCell <-
            let cellFormat = CellFormat(BackgroundColor = color)
            let cellData = CellData(UserEnteredFormat = cellFormat)
            RepeatCellRequest(
                Range = (range |> GridRange.toApiGridRange),
                Cell = cellData,
                Fields =
                    $"{nameof (cellData.UserEnteredFormat)}.{nameof (cellFormat.BackgroundColor)}"
            )
        updateCellFormatRequest

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

    let update (sheetsService: SheetsService) spreadsheetId request =
        batchUpdate sheetsService spreadsheetId [ request ]

    let batchUpdateValuesInRange (sheetsService: SheetsService) spreadsheetId rangesAndValues =
        let convertValues values =
            values
            |> List.toArray
            |> Array.map (fun list -> list |> List.toArray |> Array.map box :> IList<_>)

        let valueRanges =
            rangesAndValues
            |> List.map (fun (range, values) ->
                DataFilterValueRange(
                    DataFilter = DataFilter(GridRange = (range |> GridRange.toApiGridRange)),
                    Values = convertValues values,
                    MajorDimension = (string Dimension.Rows)
                ))
            |> List.toArray

        let valueUpdateRequestBody =
            BatchUpdateValuesByDataFilterRequest(
                ValueInputOption = "USER_ENTERED",
                Data = valueRanges
            )

        sheetsService
            .Spreadsheets
            .Values
            .BatchUpdateByDataFilter(valueUpdateRequestBody, spreadsheetId)
            .Execute()
        |> ignore

    let updateValuesInRange (sheetsService: SheetsService) spreadsheetId (range, values) =
        batchUpdateValuesInRange sheetsService spreadsheetId [ (range, values) ]

[<RequireQualifiedAccess>]
module Spreadsheet =
    let create sheetsService spreadSheetId =
        {
            Spreadsheet.SheetsService = sheetsService
            Spreadsheet.SpreadsheetId = spreadSheetId
        }

    let createFromCredentialFileName credentialFileName spreadSheetId =
        let sheetsService = SheetsService.create credentialFileName
        create sheetsService spreadSheetId

    let batchUpdate (sheetsService: Spreadsheet) requests =
        SheetsService.batchUpdate sheetsService.SheetsService sheetsService.SpreadsheetId requests

    let update (sheetsService: Spreadsheet) request =
        SheetsService.update sheetsService.SheetsService sheetsService.SpreadsheetId request

    let batchUpdateValuesInRange (sheetsService: Spreadsheet) rangesAndValues =
        SheetsService.batchUpdateValuesInRange
            sheetsService.SheetsService
            sheetsService.SpreadsheetId
            rangesAndValues

    let updateValuesInRange (sheetsService: Spreadsheet) (range, values) =
        SheetsService.updateValuesInRange
            sheetsService.SheetsService
            sheetsService.SpreadsheetId
            (range, values)
