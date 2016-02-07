module NPOI.FSharp.Tests.Xls

open System.IO
open NPOI.FSharp
open NUnit.Framework
open Swensen.Unquote

// Only these declarations should vary between XLS and XLSX tests
let targetFile = "CellTypes.xls"
let fromStream = Workbook.readXls
let fromBytes = Workbook.fromXlsBytes

[<Test>]
let ``Can read from stream`` () =
  use stream = File.OpenRead(targetFile)
  let workbook = stream |> fromStream
  workbook.NumberOfSheets =! 1

[<Test>]
let ``Can read from bytes`` () =
  let workbook = File.ReadAllBytes(targetFile) |> fromBytes
  workbook.NumberOfSheets =! 1

let bytes = lazy (File.ReadAllBytes(targetFile))

let workbook() = bytes.Value |> fromBytes

[<Test>]
let ``Can get active sheet`` () =
  let sheet = workbook() |> Workbook.activeSheet
  sheet.SheetName =! "Sheet1"

[<Test>]
let ``Can get sheet by index`` () =
  let sheet = workbook() |> Workbook.sheetAt 0
  sheet.SheetName =! "Sheet1"

[<Test>]
let ``Can get sheet by name`` () =
  let sheet = workbook() |> Workbook.sheetNamed "Sheet1"
  sheet.SheetName =! "Sheet1"

let sheet() = workbook() |> Workbook.activeSheet

[<Test>]
let ``Can get sheet rows`` () =
  sheet() |> Sheet.rowsWithContent |> Seq.length =! 2

[<Test>]
let ``Can get specific row``() =
  let row = sheet() |> Sheet.row 0
  row.RowNum =! 0

[<Test>]
let ``Can get blank row``() =
  let row = sheet() |> Sheet.row 99
  row.RowNum =! 99

[<Test>]
let ``Can get row cells`` () =
  sheet() |> Sheet.row 0 |> Row.cellsWithContent |> Seq.length =! 13

[<Test>]
let ``Can get blank row cell`` () =
  sheet() |> Sheet.cell(99, 99) |> Cell.value =! CellValue.Blank

[<Test>]
let ``Can read cell values``() =
  let sheet = sheet()
  let expectedColumnValues =
    [ 0,  CellValue.Number 123.456
      1,  CellValue.Number 789.
      2,  CellValue.String "Abc"
      3,  CellValue.Blank
      4,  CellValue.Boolean true
      7,  CellValue.Formula("A2+B2", CellValue.Number 912.456)
      8,  CellValue.Formula("CONCATENATE(C2,\"def\")", CellValue.String "Abcdef")
      // It actually comes out as a number ¯\_(ツ)_/¯
      9,  CellValue.Formula("D2", CellValue.Number 0.)
      10, CellValue.Formula("NOT(E2)", CellValue.Boolean false)
      11, CellValue.Formula("Foo", CellValue.Error 29uy) ]
  let incorrectValues =
    expectedColumnValues
    |> List.choose (fun (columnIndex, expected) ->
      let actual = sheet |> Sheet.cell (1, columnIndex) |> Cell.value
      if actual <> expected then Some(expected, actual) else None)
  incorrectValues =! []
