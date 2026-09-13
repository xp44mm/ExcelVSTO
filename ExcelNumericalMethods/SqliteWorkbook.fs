///工作簿与 SQLite 数据库之间的相互转换
///数据库结构见解决方案根目录的 create_excel_db.sql，共三张表：
///  Workbook  工作簿的名称
///  Worksheet 工作表的顺序和名称
///  Cell      单元格：所在工作表、行地址、列地址、值、公式
module ExcelNumericalMethods.SqliteWorkbook

open System
open System.Data.SQLite
open System.Globalization
open Microsoft.Office.Interop.Excel

/// 建表 SQL（与解决方案根目录 create_excel_db.sql 保持一致）
let createSchemaSql =
    """CREATE TABLE IF NOT EXISTS Workbook (
    name TEXT PRIMARY KEY
);
CREATE TABLE IF NOT EXISTS Worksheet (
    position INTEGER NOT NULL,
    name     TEXT NOT NULL UNIQUE,
    PRIMARY KEY (position)
);
CREATE TABLE IF NOT EXISTS Cell (
    worksheet TEXT NOT NULL REFERENCES Worksheet(name),
    row       INTEGER NOT NULL,
    col       INTEGER NOT NULL,
    value     TEXT,
    formula   TEXT,
    PRIMARY KEY (worksheet, row, col)
);"""

/// 单元格值转换为数据库文本
let private toText (v: obj) : string =
    match v with
    | null -> null
    | :? DateTime as d -> d.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
    | :? bool as b -> if b then "1" else "0"
    | :? double as f -> f.ToString("G17", CultureInfo.InvariantCulture)
    | :? float32 as f -> (float f).ToString("G17", CultureInfo.InvariantCulture)
    | :? int as i -> string i
    | :? int64 as i -> string i
    | :? decimal as m -> m.ToString(CultureInfo.InvariantCulture)
    | :? string as s -> s
    | other -> other.ToString()

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

/// 读取单个单元格的值（保留日期，错误值返回 null）
let private readCellValue (cell: Range) : obj =
    try
        let v = cell.get_Value(XlRangeValueDataType.xlRangeValueDefault)
        match v with
        | :? double as d when Double.IsNaN d || Double.IsInfinity d -> null
        | :? int -> null // Excel 错误值
        | _ -> v
    with _ -> null

/// 读取整个区域的值，返回基于 1 的下标的 obj[,]
let private readValues (rg: Range) (rows: int) (cols: int) : obj[,] =
    let arr = Array2D.create (rows + 1) (cols + 1) null
    if rows = 1 && cols = 1 then
        arr.[1, 1] <- readCellValue rg
    else
        let raw =
            try
                rg.get_Value(XlRangeValueDataType.xlRangeValueDefault) :?> obj[,]
            with _ -> null
        if isNull raw then
            // 逐单元格读取作为后备
            for r in 1..rows do
                for c in 1..cols do
                    arr.[r, c] <- readCellValue (rg.Cells.[r, c] :?> Range)
        else
            for r in 1..rows do
                for c in 1..cols do
                    arr.[r, c] <- raw.[r, c]
        // 错误值规整为 null
        for r in 1..rows do
            for c in 1..cols do
                match arr.[r, c] with
                | :? double as d when Double.IsNaN d || Double.IsInfinity d -> arr.[r, c] <- null
                | :? int -> arr.[r, c] <- null
                | _ -> ()
    arr

/// 读取整个区域的公式，返回基于 1 的下标的 string[,]（无公式的单元格为 null）
let private readFormulas (rg: Range) (rows: int) (cols: int) : string[,] =
    let arr = Array2D.create (rows + 1) (cols + 1) null
    if rows = 1 && cols = 1 then
        let f = try (rg.Formula :?> string) with _ -> null
        if not (String.IsNullOrEmpty f) && f.StartsWith "=" then arr.[1, 1] <- f
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
                    if not (String.IsNullOrEmpty f) && f.StartsWith "=" then arr.[r, c] <- f
        else
            for r in 1..rows do
                for c in 1..cols do
                    match raw.[r, c] with
                    | :? string as s when s.StartsWith "=" -> arr.[r, c] <- s
                    | _ -> ()
    arr

/// 读取整个区域的数字格式，返回基于 1 的下标的 string[,]
let private readFormats (rg: Range) (rows: int) (cols: int) : string[,] =
    let arr = Array2D.create (rows + 1) (cols + 1) null
    if rows = 1 && cols = 1 then
        arr.[1, 1] <- try (rg.NumberFormat :?> string) with _ -> null
    else
        let raw =
            try
                rg.NumberFormat :?> obj[,]
            with _ -> null
        if isNull raw then
            // 逐单元格读取作为后备
            for r in 1..rows do
                for c in 1..cols do
                    arr.[r, c] <- try ((rg.Cells.[r, c] :?> Range).NumberFormat :?> string) with _ -> null
        else
            for r in 1..rows do
                for c in 1..cols do
                    arr.[r, c] <- match raw.[r, c] with :? string as s -> s | _ -> null
    arr

/// 数字格式是否表示日期/时间（含 y/m/d/h/s）
let private isDateFormat (fmt: string) =
    not (String.IsNullOrEmpty fmt)
    && fmt.IndexOfAny([| 'y'; 'm'; 'd'; 'h'; 's' |]) >= 0
/// 解析与导出格式一致的 ISO 日期
let private tryParseIsoDate (s: string) =
    let mutable d = DateTime.MinValue
    let ok =
        DateTime.TryParseExact(
            s,
            [| "yyyy-MM-dd HH:mm:ss"; "yyyy-MM-dd" |],
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            &d)
    if ok then Some d else None

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
    // 单元格
    use insCell =
        new SQLiteCommand(
            "INSERT INTO Cell (worksheet, row, col, value, formula) VALUES (@worksheet, @row, @col, @value, @formula);",
            conn,
            tran)
    insCell.Parameters.AddWithValue("@worksheet", "") |> ignore
    insCell.Parameters.AddWithValue("@row", 0) |> ignore
    insCell.Parameters.AddWithValue("@col", 0) |> ignore
    insCell.Parameters.AddWithValue("@value", "") |> ignore
    insCell.Parameters.AddWithValue("@formula", "") |> ignore

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
            let values = readValues used rows cols
            let formulas = readFormulas used rows cols
            let formats = readFormats used rows cols
            for r in 1..rows do
                for c in 1..cols do
                    let v = values.[r, c]
                    let f = formulas.[r, c]
                    if not (isNull v) || not (isNull f) then
                        insCell.Parameters.["@worksheet"].Value <- ws.Name
                        insCell.Parameters.["@row"].Value <- r
                        insCell.Parameters.["@col"].Value <- c
                        insCell.Parameters.["@value"].Value <-
                            if isNull v then
                                box DBNull.Value
                            else
                                match v with
                                | :? double as d when isDateFormat formats.[r, c] ->
                                    box (DateTime.FromOADate(d).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture))
                                | _ -> box (toText v)
                        insCell.Parameters.["@formula"].Value <-
                            if isNull f then box DBNull.Value else box f
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
                    "SELECT worksheet, row, col, value, formula FROM Cell ORDER BY worksheet, row, col;",
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
            for (_, row, col, value, formula) in rows do
                let cell = ws.Cells.[row, col] :?> Range
                if not (isNull formula) then
                    // 有公式则写公式，由 Excel 重新计算值
                    cell.Formula <- formula
                elif not (isNull value) then
                    match tryParseIsoDate value with
                    | Some d ->
                        cell.Value2 <- box d
                        cell.NumberFormat <-
                            if d.TimeOfDay <> TimeSpan.Zero then
                                "yyyy-mm-dd hh:mm:ss"
                            else
                                "yyyy-mm-dd"
                    | None ->
                        let mutable f = 0.0
                        if Double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, &f) then
                            cell.Value2 <- box f
                        else
                            cell.Value2 <- box value)
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
