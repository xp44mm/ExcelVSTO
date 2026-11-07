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

/// 工作表中全部含公式的单元格（SpecialCells 优化；无公式时为空序列）。
/// SpecialCells 的结果可能是多个不连续区域，必须按区域逐个遍历，否则取不到全部单元格。
/// 新式溢出（spill）公式只有锚点单元格会出现在结果中，溢出区其余单元格不含公式；
/// 旧式 CSE 数组公式则区域内每个单元格都会出现。
let getFormulaCells(ws: Worksheet) : seq<Range> =
    try
        let used = ws.UsedRange
        let formulas = used.SpecialCells(XlCellType.xlCellTypeFormulas)
        seq {
            for a in 1 .. formulas.Areas.Count do
                let area = formulas.Areas.[a]
                yield! getCellsOfRange area
        }
    with _ -> Seq.empty


