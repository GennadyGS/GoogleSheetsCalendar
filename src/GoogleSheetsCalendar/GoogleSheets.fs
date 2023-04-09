﻿module GoogleSheets

open System
open Google.Apis.Sheets.v4.Data

type Range =
    {
        StartIndex: int option
        EndIndex: int option
    }

type TwoDimensionRange = { Columns: Range; Rows: Range }

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

    let all = { StartIndex = None; EndIndex = None }

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
    let toGridRange sheetId (range: TwoDimensionRange) =
        let mapStartIndex index = index |> Option.toNullable
        let mapEndIndex index =
            index |> Option.map ((+) 1) |> Option.toNullable
        GridRange(
            StartColumnIndex = (mapStartIndex range.Columns.EndIndex),
            EndColumnIndex = (mapEndIndex range.Columns.EndIndex),
            StartRowIndex = (mapStartIndex range.Rows.StartIndex),
            EndRowIndex = (mapEndIndex range.Columns.EndIndex),
            SheetId = sheetId
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
