///获取指定的Excel对象的所有内容
module ExcelNumericalMethods.Traversal

open Microsoft.Office.Interop.Excel

/// 获取工作表的名称
let getNames(wb:Workbook) = 
    wb.Names 
    |> Seq.cast<Name> 
    |> Seq.filter(fun nm -> nm.Visible)

/// 获取工作表
let getWorksheets(wb: Workbook) =
    wb.Worksheets
    |> Seq.cast<Worksheet>

/// 范围的每个单元格
let getCellsOfRange(rg: Range) =
    let rows = rg.Rows.Count
    let cols = rg.Columns.Count
    seq {
        for r in 1 .. rows do
            for c in 1 .. cols do
                let cell = rg.Cells.[r, c] :?> Range
                yield cell
    }

/// 获取工作表的所有单元格
let getCellsOfWorksheet(ws: Worksheet) = 
    ws.UsedRange 
    |> getCellsOfRange

/// 获取工作薄所有单元格
let getCells(wb: Workbook) =
    getWorksheets wb
    |> Seq.collect(fun ws -> getCellsOfWorksheet ws)


