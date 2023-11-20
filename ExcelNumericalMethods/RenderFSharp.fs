module ExcelNumericalMethods.RenderFSharp

open Microsoft.Office.Interop.Excel
open ExcelCompiler
open FSharp.Idioms

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
