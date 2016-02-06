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

  let fromXslxBytes (bytes : byte[]) =
    new XSSFWorkbook(new MemoryStream(bytes))

  let fromXlsBytes (bytes : byte[]) =
    new HSSFWorkbook(new MemoryStream(bytes))

  let activeSheet (workbook : IWorkbook) =
    workbook.GetSheetAt workbook.ActiveSheetIndex

  let sheetAt (workbook : IWorkbook) (sheetIndex : int) =
    workbook.GetSheetAt sheetIndex

  let sheetNamed (workbook : IWorkbook) (sheetName : string) =
    sheetName
    |> workbook.GetSheetIndex
    |> workbook.GetSheetAt

module Sheet =

  let rows (sheet : ISheet) =
    seq {
      let enumerator = sheet.GetRowEnumerator()
      while enumerator.MoveNext() do
        yield enumerator.Current :?> IRow
    }

module Row =

  let cells (row : IRow) =
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
