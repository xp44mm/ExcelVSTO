///工作簿与 SQLite 数据库之间的相互转换
///数据库结构见解决方案根目录的 create_excel_db.sql，共三张表：
///  Workbook  工作簿的名称
///  Worksheet 工作表的顺序和名称
///  Cell      单元格：所在工作表、行地址、列地址、公式、格式
module ExcelNumericalMethods.SqliteWorkbook

open System
open System.Data.SQLite
open Microsoft.Office.Interop.Excel

/// 建表 SQL：唯一事实来源为解决方案根目录的 create_excel_db.sql
/// （编译期内嵌为程序集资源，见 ExcelNumericalMethods.fsproj），此处从资源读取
let createSchemaSql =
    let resourceName = "ExcelNumericalMethods.create_excel_db.sql"
    let asm = System.Reflection.Assembly.GetExecutingAssembly()
    use stream = asm.GetManifestResourceStream(resourceName)
    if isNull stream then
        failwithf "未找到内嵌资源 %s：请确认 ExcelNumericalMethods.fsproj 已包含 create_excel_db.sql 的 EmbeddedResource" resourceName
    use reader = new System.IO.StreamReader(stream)
    reader.ReadToEnd()

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
    // 直接覆盖：不使用原有数据库的数据
    if System.IO.File.Exists path then
        System.IO.File.Delete path
    use conn = new SQLiteConnection("Data Source=" + path + ";Version=3;")
    conn.Open()
    use tran = conn.BeginTransaction()
    use cmd = new SQLiteCommand("", conn, tran)
    cmd.CommandText <- createSchemaSql
    cmd.ExecuteNonQuery() |> ignore
    // 工作簿
    use insWb = new SQLiteCommand("INSERT INTO Workbook (name) VALUES (@name);", conn, tran)
    insWb.Parameters.AddWithValue("@name", "") |> ignore
    // 工作表
    use insWs =
        new SQLiteCommand(
            "INSERT INTO Worksheet (position, name) VALUES (@position, @name);",
            conn,
            tran)
    insWs.Parameters.AddWithValue("@position", 0) |> ignore
    insWs.Parameters.AddWithValue("@name", "") |> ignore
    // 单元格：公式（常量时为其值文本） + 数字格式
    use insCell =
        new SQLiteCommand(
            "INSERT INTO Cell (worksheet, row, col, formula, format) VALUES (@worksheet, @row, @col, @formula, @format);",
            conn,
            tran)
    insCell.Parameters.AddWithValue("@worksheet", "") |> ignore
    insCell.Parameters.AddWithValue("@row", 0) |> ignore
    insCell.Parameters.AddWithValue("@col", 0) |> ignore
    insCell.Parameters.AddWithValue("@formula", "") |> ignore
    insCell.Parameters.AddWithValue("@format", "") |> ignore

    insWb.Parameters.["@name"].Value <- wb.Name
    insWb.ExecuteNonQuery() |> ignore

    wb.Worksheets
    |> Seq.cast<Worksheet>
    |> Seq.iteri (fun i ws ->
        let position = i + 1
        insWs.Parameters.["@position"].Value <- position
        insWs.Parameters.["@name"].Value <- ws.Name
        insWs.ExecuteNonQuery() |> ignore

        let used = ws.UsedRange
        let rows = used.Rows.Count
        let cols = used.Columns.Count
        if rows > 0 && cols > 0 then
            let formulas = readFormulas used rows cols
            for r in 1..rows do
                for c in 1..cols do
                    let f = formulas.[r, c]
                    if not (String.IsNullOrEmpty f) then
                        let fmt =
                            try
                                ((used.Cells.[r, c] :?> Range).NumberFormat :?> string)
                            with _ -> null
                        insCell.Parameters.["@worksheet"].Value <- ws.Name
                        insCell.Parameters.["@row"].Value <- r
                        insCell.Parameters.["@col"].Value <- c
                        insCell.Parameters.["@formula"].Value <- box f
                        insCell.Parameters.["@format"].Value <-
                            if isNull fmt then box DBNull.Value else box fmt
                        insCell.ExecuteNonQuery() |> ignore)
    tran.Commit()

/// 从 SQLite 数据库创建新的 Excel 工作簿（按三张表重建）
let createWorkbookFrom (app: Application) (path: string) : Workbook =
    // 先读取数据库内容
    let sheetNames, cells =
        use conn = new SQLiteConnection("Data Source=" + path + ";Version=3;")
        conn.Open()
        let sheetNames =
            use cmd = new SQLiteCommand("SELECT name FROM Worksheet ORDER BY position;", conn)
            use r = cmd.ExecuteReader()
            [| while r.Read() do
                   yield r.GetString 0 |]
        let cells =
            use cmd =
                new SQLiteCommand(
                    "SELECT worksheet, row, col, formula, format FROM Cell ORDER BY worksheet, row, col;",
                    conn)
            use r = cmd.ExecuteReader()
            [|
                while r.Read() do
                    yield
                        r.GetString 0,
                        r.GetInt32 1,
                        r.GetInt32 2,
                        (if r.IsDBNull 3 then null else r.GetString 3),
                        (if r.IsDBNull 4 then null else r.GetString 4)
            |]
        sheetNames, cells

    if sheetNames.Length = 0 then
        invalidOp "数据库中没有工作表记录！"

    let cellsBySheet = cells |> Array.groupBy (fun (ws, _, _, _, _) -> ws) |> Map.ofArray

    // 创建新的工作簿
    let wb = app.Workbooks.Add(Type.Missing)
    let sheets = wb.Worksheets
    let usedSheetNames = Collections.Generic.HashSet<string>()
    sheetNames
    |> Array.iteri (fun i name ->
        let ws =
            if i = 0 then
                sheets.[1] :?> Worksheet
            else
                sheets.Add(Type.Missing, sheets.[sheets.Count], Type.Missing, Type.Missing)
                :?> Worksheet
        ws.Name <- toExcelSheetName usedSheetNames name
        match Map.tryFind name cellsBySheet with
        | None -> ()
        | Some rows ->
            for (_, row, col, formula, format) in rows do
                let cell = ws.Cells.[row, col] :?> Range
                // 先设数字格式再写公式：格式为文本(@)时值按文本保存
                if not (isNull format) then
                    try cell.NumberFormat <- format with _ -> ()
                if not (isNull formula) then
                    cell.Formula <- formula)
    // 删除多余的空白工作表
    let oldCount = sheets.Count
    if oldCount > sheetNames.Length then
        let opt = app.DisplayAlerts
        app.DisplayAlerts <- false
        try
            for i in oldCount .. -1 .. (sheetNames.Length + 1) do
                (sheets.[i] :?> Worksheet).Delete()
        finally
            app.DisplayAlerts <- opt
    wb
