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

/// 读取整个区域的公式或常量文本，返回基于 1 的下标的 string[,]（空单元格为 null）
/// 常量单元格的 Formula 返回其值文本；公式单元格返回公式串
let private readFormulas (rg: Range) (rows: int) (cols: int) : string[,] =
    let arr = Array2D.create (rows + 1) (cols + 1) null
    if rows = 1 && cols = 1 then
        let f = try (rg.Formula :?> string) with _ -> null
        if not (String.IsNullOrEmpty f) then arr.[1, 1] <- f
    else
        let raw =
            try
                rg.Formula :?> obj[,]
            with _ -> null
        if isNull raw then
            // 逐单元格读取作为后备
            for r in 1..rows do
                for c in 1..cols do
                    let f = try ((rg.Cells.[r, c] :?> Range).Formula :?> string) with _ -> null
                    if not (String.IsNullOrEmpty f) then arr.[r, c] <- f
        else
            for r in 1..rows do
                for c in 1..cols do
                    match raw.[r, c] with
                    | :? string as s when not (String.IsNullOrEmpty s) -> arr.[r, c] <- s
                    | _ -> ()
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
                let formulas = readFormulas used rows cols
                seq {
                    for r in 1..rows do
                        for c in 1..cols do
                            let f = formulas.[r, c]
                            if not (String.IsNullOrEmpty f) then
                                let fmt =
                                    try
                                        ((used.Cells.[r, c] :?> Range).NumberFormat :?> string)
                                    with _ -> null
                                yield
                                    { Worksheet = ws.Name
                                      Row = r
                                      Col = c
                                      Formula = Some f
                                      NumberFormat = if isNull fmt then None else Some fmt }
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
                match cell.NumberFormat with
                | Some numberFormat ->
                    try excelCell.NumberFormat <- numberFormat with _ -> ()
                | None -> ()
                match cell.Formula with
                | Some formula -> excelCell.Formula <- formula
                | None -> ())
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
