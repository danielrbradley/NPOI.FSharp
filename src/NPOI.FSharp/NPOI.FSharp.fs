namespace NPOI.FSharp

open System.IO
open NPOI.SS.UserModel
open NPOI.HSSF.UserModel
open NPOI.XSSF.UserModel

module Workbook =

  let readXlsx (stream : Stream) =
    new XSSFWorkbook(stream)

  let readXls (stream : Stream) =
    new HSSFWorkbook(stream)

  let fromXlsxBytes (bytes : byte[]) =
    new XSSFWorkbook(new MemoryStream(bytes))

  let fromXlsBytes (bytes : byte[]) =
    new HSSFWorkbook(new MemoryStream(bytes))

  let activeSheet (workbook : IWorkbook) =
    workbook.GetSheetAt workbook.ActiveSheetIndex

  let sheetAt (sheetIndex : int) (workbook : IWorkbook) =
    workbook.GetSheetAt sheetIndex

  let sheetNamed (sheetName : string) (workbook : IWorkbook) =
    sheetName
    |> workbook.GetSheetIndex
    |> workbook.GetSheetAt

module Sheet =

  let rowsWithContent (sheet : ISheet) =
    seq {
      let enumerator = sheet.GetRowEnumerator()
      while enumerator.MoveNext() do
        yield enumerator.Current :?> IRow
    }

  let row (rowIndex : int)  (sheet : ISheet) =
    match sheet.GetRow(rowIndex) with
    | null -> sheet.CreateRow(rowIndex)
    | row -> row

  let cell (rowIndex : int, columnIndex : int) (sheet : ISheet) =
    match sheet.GetRow(rowIndex) with
    | null ->
      sheet.CreateRow(rowIndex).CreateCell(columnIndex)
    | row ->
      match row.GetCell(columnIndex) with
      | null -> row.CreateCell(columnIndex)
      | cell -> cell

module Row =

  let cellsWithContent (row : IRow) =
    row.Cells

[<RequireQualifiedAccess>]
type CellValue =
  | Number of float
  | String of string
  | Formula of formula : string * cachedResult : CellValue
  | Blank
  | Boolean of bool
  | Error of code : byte
  | Unknown of CellType

module Cell =

  let value (cell : ICell) =
    match cell.CellType with
    | CellType.Numeric -> CellValue.Number cell.NumericCellValue
    | CellType.String -> CellValue.String cell.StringCellValue
    | CellType.Formula ->
      let cachedResult =
        match cell.CachedFormulaResultType with
        | CellType.Numeric -> CellValue.Number cell.NumericCellValue
        | CellType.String -> CellValue.String cell.StringCellValue
        | CellType.Blank -> CellValue.Blank
        | CellType.Boolean -> CellValue.Boolean cell.BooleanCellValue
        | CellType.Error -> CellValue.Error cell.ErrorCellValue
        | _ -> CellValue.Unknown cell.CellType
      CellValue.Formula(cell.CellFormula, cachedResult)
    | CellType.Blank -> CellValue.Blank
    | CellType.Boolean -> CellValue.Boolean cell.BooleanCellValue
    | CellType.Error -> CellValue.Error cell.ErrorCellValue
    | _ -> CellValue.Unknown cell.CellType
