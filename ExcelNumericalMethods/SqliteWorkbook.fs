///工作簿与 SQLite 数据库之间的相互转换
module ExcelNumericalMethods.SqliteWorkbook

open System
open System.Data.SQLite
open System.Globalization
open Microsoft.Office.Interop.Excel

/// 字符串转换为合法的 SQLite 标识符（表名、列名）
let private sanitizeIdentifier (fallback: string) (name: string) =
    let sb = System.Text.StringBuilder()
    for ch in name do
        if Char.IsLetterOrDigit ch || ch = '_' then
            sb.Append ch |> ignore
        else
            sb.Append '_' |> ignore
    let s = sb.ToString().Trim('_')
    if String.IsNullOrEmpty s then fallback else s

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

/// 列类型
type private ColumnType =
    | Integer
    | Real
    | Text

/// 判断 double 是否为整数
let private isIntegral (d: double) = d = Math.Floor d && not (Double.IsInfinity d)

/// 根据数据行推断每一列的类型（第一行是表头，不参与推断）
let private inferColumnTypes (data: obj[,]) (rows: int) (cols: int) : ColumnType[] =
    [|
        for c in 1..cols do
            let mutable t = Integer
            for r in 2..rows do
                match data.[r, c] with
                | null -> ()
                | :? string -> t <- Text
                | :? DateTime -> t <- Text
                | :? double as d when not (isIntegral d) ->
                    if t = Integer then t <- Real
                | _ -> ()
            yield t
    |]

/// 单元格值转换为数据库参数值
let private toDbValue (colType: ColumnType) (v: obj) : obj =
    match v with
    | null -> box DBNull.Value
    | :? DateTime as d ->
        // 日期统一存为 ISO 文本，保证可逆
        box (d.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture))
    | :? bool as b ->
        match colType with
        | Text -> box (if b then "1" else "0")
        | _ -> box (if b then 1 else 0)
    | :? double as d ->
        match colType with
        | Integer when isIntegral d && d >= -9.2233720368547758E+18 && d <= 9.2233720368547758E+18 ->
            box (int64 d)
        | _ -> box d
    | :? float32 as f -> box (float f)
    | :? int as i -> box i
    | :? int64 as i -> box i
    | :? decimal as m -> box (double m)
    | :? string as s -> box s
    | other -> box (other.ToString())

/// 将当前工作簿另存为 SQLite 数据库：每个工作表一张表，第一行为列名
let saveWorkbookAs (path: string) (wb: Workbook) =
    use conn = new SQLiteConnection("Data Source=" + path + ";Version=3;")
    conn.Open()
    use tran = conn.BeginTransaction()
    let usedTableNames = Collections.Generic.HashSet<string>()
    for ws in Traversal.getWorksheets wb do
        let used = ws.UsedRange
        let rows = used.Rows.Count
        let cols = used.Columns.Count
        if rows > 0 && cols > 0 then
            let data = readValues used rows cols
            // 整个区域为空则跳过
            let mutable any = false
            for r in 1..rows do
                for c in 1..cols do
                    if not (isNull data.[r, c]) then any <- true
            if any then
                let tableName = uniqueName usedTableNames (sanitizeIdentifier "Sheet" ws.Name)
                // 列名：第一行为表头，空表头生成 col{序号}
                let usedColNames = Collections.Generic.HashSet<string>()
                let colNames =
                    [|
                        for c in 1..cols do
                            let h = data.[1, c]
                            let baseName =
                                match h with
                                | :? string as s when not (String.IsNullOrWhiteSpace s) -> s
                                | _ -> sprintf "col%d" c
                            yield uniqueName usedColNames (sanitizeIdentifier (sprintf "col%d" c) baseName)
                    |]
                let colTypes = inferColumnTypes data rows cols
                // 建表
                let quote (s: string) = "\"" + s.Replace("\"", "\"\"") + "\""
                use cmd = new SQLiteCommand("", conn, tran)
                cmd.CommandText <- sprintf "DROP TABLE IF EXISTS %s" (quote tableName)
                cmd.ExecuteNonQuery() |> ignore
                cmd.CommandText <-
                    let defs =
                        Array.map2
                            (fun name t ->
                                let ty =
                                    match t with
                                    | Integer -> "INTEGER"
                                    | Real -> "REAL"
                                    | Text -> "TEXT"
                                sprintf "%s %s" (quote name) ty)
                            colNames
                            colTypes
                        |> String.concat ", "
                    sprintf "CREATE TABLE %s (%s)" (quote tableName) defs
                cmd.ExecuteNonQuery() |> ignore
                // 插入数据行：命名参数，预置后按行更新值复用
                let pnames = [| for c in 1..cols -> sprintf "@c%s" colNames.[c - 1] |]
                cmd.CommandText <-
                    sprintf "INSERT INTO %s (%s) VALUES (%s)"
                        (quote tableName)
                        (colNames |> Array.map quote |> String.concat ", ")
                        (String.concat ", " pnames)
                for c in 1..cols do
                    cmd.Parameters.AddWithValue(pnames.[c - 1], box DBNull.Value) |> ignore
                for r in 2..rows do
                    for c in 1..cols do
                        cmd.Parameters.[pnames.[c - 1]].Value <-
                            toDbValue colTypes.[c - 1] data.[r, c]
                    cmd.ExecuteNonQuery() |> ignore
    tran.Commit()

/// 读取数据库中的表名（排除 SQLite 内部表）
let private getTableNames (conn: SQLiteConnection) : string[] =
    use cmd =
        new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name", conn)
    use reader = cmd.ExecuteReader()
    [| while reader.Read() do
           yield reader.GetString 0 |]

/// 读取一张表的列名和声明的类型
let private getColumns (conn: SQLiteConnection) (table: string) : (string * string)[] =
    let t = table.Replace("\"", "\"\"")
    use cmd = new SQLiteCommand(sprintf "PRAGMA table_info(\"%s\")" t, conn)
    use reader = cmd.ExecuteReader()
    [| while reader.Read() do
           yield reader.GetString 1, reader.GetString 2 |]

/// 声明的列类型是否为数值类型
let private isNumericType (declared: string) =
    let u = declared.ToUpperInvariant()
    u.Contains "INT" || u.Contains "REAL" || u.Contains "FLOA" || u.Contains "DOUB" || u.Contains "NUMERIC"
    || u.Contains "DEC"

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

/// 读取一张表的全部数据（不含表头）
let private readRows (conn: SQLiteConnection) (table: string) (columns: (string * string)[]) : obj[][] =
    let t = table.Replace("\"", "\"\"")
    use cmd = new SQLiteCommand(sprintf "SELECT * FROM \"%s\"" t, conn)
    use reader = cmd.ExecuteReader()
    [|
        while reader.Read() do
            yield
                [|
                    for i in 0..reader.FieldCount - 1 ->
                        match reader.GetValue i with
                        | :? DBNull -> null
                        | :? int64 as l -> box (double l)
                        | :? double as d -> box d
                        | :? string as s ->
                            let declared = if i < columns.Length then snd columns.[i] else ""
                            if isNumericType declared then
                                let mutable f = 0.0
                                if Double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, &f) then
                                    box f
                                else
                                    box s
                            else
                                match tryParseIsoDate s with
                                | Some d -> box d
                                | None -> box s
                        | :? (byte[]) as b ->
                            // BLOB 以十六进制文本表示
                            box (BitConverter.ToString(b).Replace("-", ""))
                        | other -> box (other.ToString())
                |]
    |]

/// 从 SQLite 数据库创建新的 Excel 工作簿：每张表一个工作表
let createWorkbookFrom (app: Application) (path: string) : Workbook =
    // 先在数据库中读取所有表的内容
    let tables =
        use conn = new SQLiteConnection("Data Source=" + path + ";Version=3;")
        conn.Open()
        let names = getTableNames conn
        if names.Length = 0 then
            invalidOp "数据库中没有任何数据表！"
        [| for t in names ->
               let cols = getColumns conn t
               t, cols, readRows conn t cols |]

    // 创建新的工作簿
    let wb = app.Workbooks.Add(Type.Missing)
    let sheets = wb.Worksheets
    let usedSheetNames = Collections.Generic.HashSet<string>()
    tables
    |> Array.iteri (fun i (table, cols, rows) ->
        let ws =
            if i = 0 then
                sheets.[1] :?> Worksheet
            else
                sheets.Add(Type.Missing, sheets.[sheets.Count], Type.Missing, Type.Missing) :?> Worksheet
        ws.Name <- toExcelSheetName usedSheetNames table
        let nCols = cols.Length
        let nRows = rows.Length
        // 组装 obj[,]：第一行为列名
        let arr = Array2D.create (nRows + 1) nCols null
        for c in 0..nCols - 1 do
            arr.[0, c] <- box (fst cols.[c])
        for r in 0..nRows - 1 do
            for c in 0..nCols - 1 do
                arr.[r + 1, c] <- rows.[r].[c]
        let target = ws.Range(ws.Cells.[1, 1], ws.Cells.[nRows + 1, nCols])
        target.Value2 <- arr
        // 日期列设置数字格式
        for c in 0..nCols - 1 do
            let isDateCol =
                rows
                |> Array.exists (fun row -> match row.[c] with :? DateTime -> true | _ -> false)
            if isDateCol then
                let hasTime =
                    rows
                    |> Array.exists (fun row ->
                        match row.[c] with
                        | :? DateTime as d -> d.TimeOfDay <> TimeSpan.Zero
                        | _ -> false)
                let fmt = if hasTime then "yyyy-mm-dd hh:mm:ss" else "yyyy-mm-dd"
                ws.Range(ws.Cells.[1, c + 1], ws.Cells.[nRows + 1, c + 1]).NumberFormat <- fmt
        target.Columns.AutoFit() |> ignore)
    // 删除多余的空白工作表
    let oldCount = sheets.Count
    if oldCount > tables.Length then
        let opt = app.DisplayAlerts
        app.DisplayAlerts <- false
        try
            for i in oldCount .. -1 .. (tables.Length + 1) do
                (sheets.[i] :?> Worksheet).Delete()
        finally
            app.DisplayAlerts <- opt
    wb
