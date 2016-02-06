namespace System
open System.Reflection

[<assembly: AssemblyTitleAttribute("NPOI.FSharp")>]
[<assembly: AssemblyProductAttribute("NPOI.FSharp")>]
[<assembly: AssemblyDescriptionAttribute("F# wrappers for NPOI")>]
[<assembly: AssemblyVersionAttribute("1.0")>]
[<assembly: AssemblyFileVersionAttribute("1.0")>]
do ()

module internal AssemblyVersionInformation =
    let [<Literal>] Version = "1.0"
