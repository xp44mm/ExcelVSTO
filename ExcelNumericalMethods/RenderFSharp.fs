module ExcelNumericalMethods.RenderFSharp

open Microsoft.Office.Interop.Excel
open FSharp.Idioms

open ExcelCompiler
let getFsharp (cell:Range) =
    let addr = cell.get_Address().Replace("$","")

    let fs =
        if unbox<bool> cell.HasFormula then
            cell.Formula
            |> unbox<string>
            |> ExcelExprCompiler.compile
            |> RenderFSharp.norm
        else
            Literal.stringify (cell.get_Value2())

    addr, fs
