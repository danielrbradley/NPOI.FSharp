#load "Scripts/load-project-debug.fsx"

open NPOI.FSharp
open System.Globalization

let rec formatCellValueForCsv = function
  | CellValue.Number n -> n.ToString(CultureInfo.InvariantCulture)
  | CellValue.String s -> sprintf "\"%s\"" (s.Replace("\"","\"\""))
  | CellValue.Formula(formula, cachedResult) ->
    match cachedResult with
    | CellValue.Error _
    | CellValue.Unknown _
    | CellValue.Formula _ -> formula
    | _ -> formatCellValueForCsv cachedResult
  | CellValue.Blank -> ""
  | CellValue.Boolean b -> b.ToString(CultureInfo.InvariantCulture)
  | CellValue.Error e -> sprintf "Error: %i" e
  | CellValue.Unknown t -> "Error: Unknown cell type"

let formatCellForCsv cell =
  cell |> Cell.value |> formatCellValueForCsv

let formatRowForCsv row =
  row
  |> Row.cells
  |> Seq.map formatCellForCsv
  |> String.concat ","

let toCsv workbook =
  workbook
  |> Workbook.activeSheet
  |> Sheet.rows
  |> Seq.map formatRowForCsv
  |> String.concat "\n"

let xlsxToCsv filePath =
  filePath
  |> System.IO.File.ReadAllBytes
  |> Workbook.fromXslxBytes
  |> toCsv

let xlsToCsv filePath =
  filePath
  |> System.IO.File.ReadAllBytes
  |> Workbook.fromXlsBytes
  |> toCsv
