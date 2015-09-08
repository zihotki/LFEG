namespace System
open System.Reflection

[<assembly: AssemblyTitleAttribute("LFEG")>]
[<assembly: AssemblyProductAttribute("LFEG")>]
[<assembly: AssemblyDescriptionAttribute("LFEG - Lighting Fast Excel Generator, framework for efficient generation of excel files")>]
[<assembly: AssemblyVersionAttribute("1.0")>]
[<assembly: AssemblyFileVersionAttribute("1.0")>]
do ()

module internal AssemblyVersionInformation =
    let [<Literal>] Version = "1.0"
