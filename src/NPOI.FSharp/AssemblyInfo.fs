namespace System
open System.Reflection

[<assembly: AssemblyTitleAttribute("NPOI.FSharp")>]
[<assembly: AssemblyProductAttribute("NPOI.FSharp")>]
[<assembly: AssemblyDescriptionAttribute("F# wrappers for NPOI")>]
[<assembly: AssemblyVersionAttribute("0.0.1")>]
[<assembly: AssemblyFileVersionAttribute("0.0.1")>]
do ()

module internal AssemblyVersionInformation =
    let [<Literal>] Version = "0.0.1"
