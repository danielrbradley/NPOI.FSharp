(*** hide ***)
// This block of code is omitted in the generated HTML documentation. Use 
// it to define helpers that you do not want to show in the documentation.
#I "../../bin/NPOI.FSharp"

(**
NPOI.FSharp
======================

Documentation

<div class="row">
  <div class="span1"></div>
  <div class="span6">
    <div class="well well-small" id="nuget">
      The NPOI.FSharp library can be <a href="https://nuget.org/packages/NPOI.FSharp">installed from NuGet</a>:
      <pre>PM> Install-Package NPOI.FSharp</pre>
    </div>
  </div>
  <div class="span1"></div>
</div>

Example
-------

This (bad) example demonstrates using the library to convert an exel spreadsheet to CSV.

... don't use this - it has bugs - it's only an example of what the API looks like ;)

*)
#r "NPOI"
#r "NPOI.OOXML"
#r "NPOI.FSharp"

open NPOI.FSharp
open System.Globalization

let rec formatCellValueForCsv = function
  | CellValue.Number n -> n.ToString(CultureInfo.InvariantCulture)
  | CellValue.String s -> sprintf "\"%s\"" (s.Replace("\"","\"\""))
  | CellValue.Formula(formula, cachedResult) -> formula
  | CellValue.Blank -> ""
  | CellValue.Boolean b -> b.ToString(CultureInfo.InvariantCulture)
  | CellValue.Error e -> sprintf "Error: %i" e
  | CellValue.Unknown t -> "Error: Unknown cell type"

let formatCellForCsv cell =
  cell |> Cell.value |> formatCellValueForCsv

let formatRowForCsv row =
  row
  |> Row.cellsWithContent
  |> Seq.map formatCellForCsv
  |> String.concat ","

let toCsv workbook =
  workbook
  |> Workbook.activeSheet
  |> Sheet.rowsWithContent
  |> Seq.map formatRowForCsv
  |> String.concat "\n"

let xlsxToCsv filePath =
  filePath
  |> System.IO.File.ReadAllBytes
  |> Workbook.fromXlsxBytes
  |> toCsv

let xlsToCsv filePath =
  filePath
  |> System.IO.File.ReadAllBytes
  |> Workbook.fromXlsBytes
  |> toCsv

(**
Some more info

Samples & documentation
-----------------------

The library comes with comprehensible documentation. 
It can include tutorials automatically generated from `*.fsx` files in [the content folder][content]. 
The API reference is automatically generated from Markdown comments in the library implementation.

 * [Tutorial](tutorial.html) contains a further explanation of this sample library.

 * [API Reference](reference/index.html) contains automatically generated documentation for all types, modules
   and functions in the library. This includes additional brief samples on using most of the
   functions.
 
Contributing and copyright
--------------------------

The project is hosted on [GitHub][gh] where you can [report issues][issues], fork 
the project and submit pull requests. If you're adding a new public API, please also 
consider adding [samples][content] that can be turned into a documentation. You might
also want to read the [library design notes][readme] to understand how it works.

The library is available under Public Domain license, which allows modification and 
redistribution for both commercial and non-commercial purposes. For more information see the 
[License file][license] in the GitHub repository. 

  [content]: https://github.com/fsprojects/NPOI.FSharp/tree/master/docs/content
  [gh]: https://github.com/fsprojects/NPOI.FSharp
  [issues]: https://github.com/fsprojects/NPOI.FSharp/issues
  [readme]: https://github.com/fsprojects/NPOI.FSharp/blob/master/README.md
  [license]: https://github.com/fsprojects/NPOI.FSharp/blob/master/LICENSE.txt
*)
