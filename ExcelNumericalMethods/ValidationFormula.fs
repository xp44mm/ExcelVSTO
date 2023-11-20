module ExcelNumericalMethods.ValidationFormula

open ExcelCompiler
open System
open Microsoft.Office.Interop.Excel

let validate(wb:Workbook) =
    //有公式的单元格
    let cells = 
        Traversal.getCells wb
        |> Seq.filter(fun cell -> unbox<bool> cell.HasFormula)
        |> Array.ofSeq

    let nosupports = 
        cells
        |> Array.Parallel.choose(fun cell -> 
            if unbox<bool> cell.HasArray then
                Some(cell.Worksheet.Name, (cell.CurrentArray.get_Address()), "公式数组")
            else
                let f = unbox<string> cell.Formula
                let tokens = 
                    f
                    |> ExcelTokenUtils.tokenize 0
                    |> Seq.map(fun {value=tok}-> tok)
                    |> Seq.toList

                let msg = Validation.message tokens
                if String.IsNullOrEmpty(msg) then
                    None
                else
                    Some(cell.Worksheet.Name, (cell.get_Address()), msg)
                )

    nosupports 
    |> Array.distinct //公式数组会重复项

