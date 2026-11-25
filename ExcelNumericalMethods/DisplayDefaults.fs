///「恢复默认显示」：把当前工作表的显示选项恢复为 Excel 默认状态，
/// 并把行高为 0 的行、列宽为 0 的列恢复为工作表标准尺寸。
module ExcelNumericalMethods.DisplayDefaults

open Microsoft.Office.Interop.Excel
open System.Collections.Generic
open System.Runtime.InteropServices

/// COM 的行高/列宽属性返回 obj（尺寸不一致时为 null），需显式拆箱后比较。
let private isZero (value: obj) =
    match value with
    | :? double as d -> d = 0.0
    | _ -> false

/// 把行高为 0 的行恢复为标准行高。
/// 整表行高一致时一次判断即可；不一致时以可见行做差集，只检查隐藏行，避免逐行遍历整表。
let private restoreZeroRowHeights (ws: Worksheet) =
    let rowCount = int ws.Rows.Count
    match ws.Rows.RowHeight with
    | :? double as h when h <> 0.0 -> 0
    | :? double ->
        ws.Rows.RowHeight <- box ws.StandardHeight
        rowCount
    | _ ->
        let hidden = HashSet<int>(seq { 1 .. rowCount })
        try
            let visible = ws.Rows.SpecialCells(XlCellType.xlCellTypeVisible)
            for a in 1 .. visible.Areas.Count do
                let area = visible.Areas.[a]
                let first = int area.Row
                let count = int area.Rows.Count
                for i in 0 .. count - 1 do
                    hidden.Remove(first + i) |> ignore
        with :? COMException -> ()
        let standard = ws.StandardHeight
        let mutable restored = 0
        for r in hidden do
            let row = ws.Rows.[r] :?> Range
            if isZero row.RowHeight then
                row.RowHeight <- box standard
                restored <- restored + 1
        restored

/// 把列宽为 0 的列恢复为标准列宽。策略同 restoreZeroRowHeights。
let private restoreZeroColumnWidths (ws: Worksheet) =
    let colCount = int ws.Columns.Count
    match ws.Columns.ColumnWidth with
    | :? double as w when w <> 0.0 -> 0
    | :? double ->
        ws.Columns.ColumnWidth <- box ws.StandardWidth
        colCount
    | _ ->
        let hidden = HashSet<int>(seq { 1 .. colCount })
        try
            let visible = ws.Columns.SpecialCells(XlCellType.xlCellTypeVisible)
            for a in 1 .. visible.Areas.Count do
                let area = visible.Areas.[a]
                let first = int area.Column
                let count = int area.Columns.Count
                for i in 0 .. count - 1 do
                    hidden.Remove(first + i) |> ignore
        with :? COMException -> ()
        let standard = ws.StandardWidth
        let mutable restored = 0
        for c in hidden do
            let col = ws.Columns.[c] :?> Range
            if isZero col.ColumnWidth then
                col.ColumnWidth <- box standard
                restored <- restored + 1
        restored

/// 设置显示选项并恢复零高/零宽行列；ScreenUpdating 在内部临时关闭并在结束时恢复。
/// 返回 (恢复行数, 恢复列数)。
let apply (app: Application) (ws: Worksheet) =
    let oldScreenUpdating = app.ScreenUpdating
    try
        app.ScreenUpdating <- false
        let win = app.ActiveWindow
        win.DisplayHeadings <- true
        win.DisplayFormulas <- false
        win.DisplayRightToLeft <- false
        ws.DisplayPageBreaks <- false
        win.DisplayZeros <- true
        win.DisplayOutline <- true
        win.DisplayGridlines <- true
        win.GridlineColorIndex <- XlColorIndex.xlColorIndexAutomatic
        let rows = restoreZeroRowHeights ws
        let cols = restoreZeroColumnWidths ws
        rows, cols
    finally
        app.ScreenUpdating <- oldScreenUpdating
