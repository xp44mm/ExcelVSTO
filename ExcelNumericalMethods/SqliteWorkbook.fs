///工作簿与 SQLite 数据库之间的相互转换
///数据库结构见 ExcelWorkbookDb 项目的 create_excel_db.sql，共三张表：
///  Workbook  工作簿的名称
///  Worksheet 工作表的顺序和名称
///  Cell      单元格：所在工作表、行地址、列地址、公式、数字格式
///本模块只负责 Excel 互操作；数据库读写全部委托给 ExcelWorkbookDb 包装库。
module ExcelNumericalMethods.SqliteWorkbook

open System
open Microsoft.Office.Interop.Excel
open ExcelWorkbookDb

/// 在已用名称集合中生成不重复的名称
let private uniqueName (used: Collections.Generic.HashSet<string>) (baseName: string) =
    let mutable name = baseName
    let mutable i = 2
    while used.Contains name do
        name <- sprintf "%s_%d" baseName i
        i <- i + 1
    used.Add name |> ignore
    name

/// 字符串转换为合法的 Excel 工作表名称（不超过 31 个字符，去掉非法字符）
let private toExcelSheetName (used: Collections.Generic.HashSet<string>) (baseName: string) =
    let invalid = [| ':'; '\\'; '/'; '?'; '*'; '['; ']' |]
    let sb = System.Text.StringBuilder()
    for ch in baseName do
        if Array.contains ch invalid then
            sb.Append '_' |> ignore
        else
            sb.Append ch |> ignore
    let s = sb.ToString()
    let s = if String.IsNullOrEmpty s then "Sheet" else s
    let s = if s.Length > 31 then s.Substring(0, 31) else s
    uniqueName used s

/// 读取整个区域的公式与数字格式，返回基于 1 的下标的 (公式, 数字格式) string[,]
/// 公式为 null 或空表示空单元格（不写入）；数字格式为 null 表示读取失败（写入时用默认值 General）
let private readCells (rg: Range) (rows: int) (cols: int) : (string * string)[,] =
    let arr = Array2D.create (rows + 1) (cols + 1) (null, null)
    if rows = 1 && cols = 1 then
        let f = try (rg.Formula :?> string) with _ -> null
        let nf = try (rg.NumberFormat :?> string) with _ -> null
        arr.[1, 1] <- (f, nf)
    else
        let rawF =
            try
                rg.Formula :?> obj[,]
            with _ -> null
        let rawN =
            try
                rg.NumberFormat :?> obj[,]
            with _ -> null
        if isNull rawF || isNull rawN then
            // 批量读取失败时逐单元格读取公式与数字格式
            for r in 1..rows do
                for c in 1..cols do
                    let cell = rg.Cells.[r, c] :?> Range
                    let f = try (cell.Formula :?> string) with _ -> null
                    let nf = try (cell.NumberFormat :?> string) with _ -> null
                    arr.[r, c] <- (f, nf)
        else
            for r in 1..rows do
                for c in 1..cols do
                    let f =
                        match rawF.[r, c] with
                        | :? string as s -> s
                        | _ -> null
                    let nf =
                        match rawN.[r, c] with
                        | :? string as s -> s
                        | _ -> null
                    arr.[r, c] <- (f, nf)
    arr

/// 将当前工作簿另存为 SQLite 数据库（三张表，直接覆盖目标文件，不利用原有数据）
let saveWorkbookAs (path: string) (wb: Workbook) =
    let worksheets =
        wb.Worksheets
        |> Seq.cast<Worksheet>
        |> Seq.mapi (fun i ws -> { Position = i + 1; Name = ws.Name })
        |> Array.ofSeq
    let cells =
        wb.Worksheets
        |> Seq.cast<Worksheet>
        |> Seq.collect (fun ws ->
            let used = ws.UsedRange
            let rows = used.Rows.Count
            let cols = used.Columns.Count
            if rows > 0 && cols > 0 then
                let cellData = readCells used rows cols
                seq {
                    for r in 1..rows do
                        for c in 1..cols do
                            let f, nf = cellData.[r, c]
                            if not (String.IsNullOrEmpty f) then
                                yield
                                    { Worksheet = ws.Name
                                      Row = r
                                      Col = c
                                      Formula = f
                                      NumberFormat = if isNull nf then WorkbookDb.DefaultNumberFormat else nf }
                }
            else
                Seq.empty)
        |> Array.ofSeq
    WorkbookDb.save
        path
        { Name = Some wb.Name
          Worksheets = worksheets
          Cells = cells }

/// 从 SQLite 数据库创建新的 Excel 工作簿（按三张表重建）
let createWorkbookFrom (app: Application) (path: string) : Workbook =
    // 先读取数据库内容
    let data = WorkbookDb.load path

    if data.Worksheets.Length = 0 then
        invalidOp "数据库中没有工作表记录！"

    let cellsBySheet = data.Cells |> Array.groupBy (fun c -> c.Worksheet) |> Map.ofArray

    // 创建新的工作簿
    let wb = app.Workbooks.Add(Type.Missing)
    let sheets = wb.Worksheets
    let usedSheetNames = Collections.Generic.HashSet<string>()
    data.Worksheets
    |> Array.iteri (fun i ws ->
        let sheet =
            if i = 0 then
                sheets.[1] :?> Worksheet
            else
                sheets.Add(Type.Missing, sheets.[sheets.Count], Type.Missing, Type.Missing)
                :?> Worksheet
        sheet.Name <- toExcelSheetName usedSheetNames ws.Name
        match Map.tryFind ws.Name cellsBySheet with
        | None -> ()
        | Some rows ->
            for cell in rows do
                let excelCell = sheet.Cells.[cell.Row, cell.Col] :?> Range
                // 先设数字格式再写公式：格式为文本(@)时值按文本保存
                try excelCell.NumberFormat <- cell.NumberFormat with _ -> ()
                if not (String.IsNullOrEmpty cell.Formula) then
                    excelCell.Formula <- cell.Formula)
    // 删除多余的空白工作表
    let oldCount = sheets.Count
    if oldCount > data.Worksheets.Length then
        let opt = app.DisplayAlerts
        app.DisplayAlerts <- false
        try
            for i in oldCount .. -1 .. (data.Worksheets.Length + 1) do
                (sheets.[i] :?> Worksheet).Delete()
        finally
            app.DisplayAlerts <- opt
    wb
