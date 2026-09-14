///工作簿 SQLite 数据库访问包装库（netstandard2.0）
///数据库结构见本项目的 create_excel_db.sql，共三张表：
///  Workbook  工作簿的名称
///  Worksheet 工作表的顺序和名称
///  Cell      单元格：所在工作表、行地址、列地址、公式、格式
///本库只封装数据库访问，不依赖 Excel 互操作，可被任意 netstandard2.0 应用引用。
namespace ExcelWorkbookDb

open System
open System.Data.SQLite

/// 工作表记录：对应 Worksheet 表
type WorksheetRow =
    { Position: int // 工作表顺序，从 1 开始
      Name: string } // 工作表名称

/// 单元格记录：对应 Cell 表
type CellRow =
    { Worksheet: string // 所在工作表名称
      Row: int // 行地址，从 1 开始
      Col: int // 列地址，从 1 开始
      Formula: string option // 单元格内容：公式或文本形式的值
      Format: string option } // 单元格格式（数字格式）

/// 工作簿数据库的完整内容：名称、工作表、单元格
type WorkbookData =
    { Name: string option
      Worksheets: WorksheetRow[]
      Cells: CellRow[] }

/// 数据库访问包装：封装 Workbook / Worksheet / Cell 三张表的参数化读写
module WorkbookDb =

    /// 内嵌建表 SQL 的资源名（唯一事实来源：本项目的 create_excel_db.sql）
    let [<Literal>] private SchemaResourceName = "ExcelWorkbookDb.create_excel_db.sql"

    /// 建表 SQL：编译期内嵌为程序集资源，运行时从资源读取
    let createSchemaSql: string =
        let asm = System.Reflection.Assembly.GetExecutingAssembly()
        use stream = asm.GetManifestResourceStream SchemaResourceName
        if isNull stream then
            failwithf
                "未找到内嵌资源 %s：请确认 ExcelWorkbookDb.fsproj 已包含 create_excel_db.sql 的 EmbeddedResource"
                SchemaResourceName
        use reader = new System.IO.StreamReader(stream)
        reader.ReadToEnd()

    /// 生成连接字符串
    let connectionString (path: string) : string =
        "Data Source=" + path + ";Version=3;"

    /// 打开数据库连接；文件不存在时 SQLite 自动创建空文件（还需先建表才能读写）
    /// 开启外键强制：Cell.worksheet 引用 Worksheet.name 的约束生效
    let openConnection (path: string) : SQLiteConnection =
        let conn = new SQLiteConnection(connectionString path)
        conn.Open()
        use cmd = new SQLiteCommand("PRAGMA foreign_keys = ON;", conn)
        cmd.ExecuteNonQuery() |> ignore
        conn

    /// 打开数据库执行读取操作，结束后释放连接
    let withConnection (path: string) (f: SQLiteConnection -> 'T) : 'T =
        use conn = openConnection path
        f conn

    /// 在指定连接上执行建表 SQL
    let private createSchema (conn: SQLiteConnection) : unit =
        use cmd = new SQLiteCommand(createSchemaSql, conn)
        cmd.ExecuteNonQuery() |> ignore

    /// 新建数据库文件：直接覆盖目标文件（不合并原有数据）并执行建表 SQL
    let createDatabase (path: string) : unit =
        if System.IO.File.Exists path then
            System.IO.File.Delete path
        use conn = openConnection path
        createSchema conn

    /// 在单个事务中执行写入操作：成功提交，异常时回滚并释放连接
    let withTransaction (path: string) (action: SQLiteTransaction -> 'T) : 'T =
        use conn = openConnection path
        use tran = conn.BeginTransaction()
        let result = action tran
        tran.Commit()
        result

    // ---------- 读取 ----------

    /// 读取工作簿名称（Workbook 表为单行表）
    let getWorkbookName (conn: SQLiteConnection) : string option =
        use cmd = new SQLiteCommand("SELECT name FROM Workbook;", conn)
        use r = cmd.ExecuteReader()
        if r.Read() then Some(r.GetString 0) else None

    /// 按 position 升序读取全部工作表
    let getWorksheets (conn: SQLiteConnection) : WorksheetRow[] =
        use cmd = new SQLiteCommand("SELECT position, name FROM Worksheet ORDER BY position;", conn)
        use r = cmd.ExecuteReader()
        [|
            while r.Read() do
                yield
                    { Position = r.GetInt32 0
                      Name = r.GetString 1 }
        |]

    /// 按 (worksheet, row, col) 升序读取全部单元格
    let getCells (conn: SQLiteConnection) : CellRow[] =
        use cmd =
            new SQLiteCommand(
                "SELECT worksheet, row, col, formula, format FROM Cell ORDER BY worksheet, row, col;",
                conn)
        use r = cmd.ExecuteReader()
        [|
            while r.Read() do
                yield
                    { Worksheet = r.GetString 0
                      Row = r.GetInt32 1
                      Col = r.GetInt32 2
                      Formula = if r.IsDBNull 3 then None else Some(r.GetString 3)
                      Format = if r.IsDBNull 4 then None else Some(r.GetString 4) }
        |]

    /// 读取指定工作表的单元格
    let getCellsOf (conn: SQLiteConnection) (worksheet: string) : CellRow[] =
        use cmd =
            new SQLiteCommand(
                "SELECT row, col, formula, format FROM Cell WHERE worksheet = @worksheet ORDER BY row, col;",
                conn)
        cmd.Parameters.AddWithValue("@worksheet", worksheet) |> ignore
        use r = cmd.ExecuteReader()
        [|
            while r.Read() do
                yield
                    { Worksheet = worksheet
                      Row = r.GetInt32 0
                      Col = r.GetInt32 1
                      Formula = if r.IsDBNull 2 then None else Some(r.GetString 2)
                      Format = if r.IsDBNull 3 then None else Some(r.GetString 3) }
        |]

    /// 读取数据库的完整内容
    let load (path: string) : WorkbookData =
        withConnection path (fun conn ->
            { Name = getWorkbookName conn
              Worksheets = getWorksheets conn
              Cells = getCells conn })

    // ---------- 写入 ----------

    /// string option -> obj（None 转为 DBNull.Value）
    let private nullable (v: string option) : obj =
        match v with
        | Some s -> box s
        | None -> box DBNull.Value

    /// 写入工作簿名称（先清空 Workbook 表再插入）
    let setWorkbookName (tran: SQLiteTransaction) (name: string) : unit =
        use cmd = new SQLiteCommand("DELETE FROM Workbook;", tran.Connection, tran)
        cmd.ExecuteNonQuery() |> ignore
        use cmd = new SQLiteCommand("INSERT INTO Workbook (name) VALUES (@name);", tran.Connection, tran)
        cmd.Parameters.AddWithValue("@name", name) |> ignore
        cmd.ExecuteNonQuery() |> ignore

    /// 插入一条工作表记录
    let insertWorksheet (tran: SQLiteTransaction) (ws: WorksheetRow) : unit =
        use cmd =
            new SQLiteCommand(
                "INSERT INTO Worksheet (position, name) VALUES (@position, @name);",
                tran.Connection,
                tran)
        cmd.Parameters.AddWithValue("@position", ws.Position) |> ignore
        cmd.Parameters.AddWithValue("@name", ws.Name) |> ignore
        cmd.ExecuteNonQuery() |> ignore

    /// 插入一条单元格记录
    let insertCell (tran: SQLiteTransaction) (cell: CellRow) : unit =
        use cmd =
            new SQLiteCommand(
                "INSERT INTO Cell (worksheet, row, col, formula, format) VALUES (@worksheet, @row, @col, @formula, @format);",
                tran.Connection,
                tran)
        cmd.Parameters.AddWithValue("@worksheet", cell.Worksheet) |> ignore
        cmd.Parameters.AddWithValue("@row", cell.Row) |> ignore
        cmd.Parameters.AddWithValue("@col", cell.Col) |> ignore
        cmd.Parameters.AddWithValue("@formula", nullable cell.Formula) |> ignore
        cmd.Parameters.AddWithValue("@format", nullable cell.Format) |> ignore
        cmd.ExecuteNonQuery() |> ignore

    /// 清空三张表（保留表结构），用于整体覆盖写入而不重建文件
    let clear (tran: SQLiteTransaction) : unit =
        for table in [ "Cell"; "Worksheet"; "Workbook" ] do
            use cmd = new SQLiteCommand("DELETE FROM " + table + ";", tran.Connection, tran)
            cmd.ExecuteNonQuery() |> ignore

    /// 将完整工作簿数据另存为数据库文件：直接覆盖目标文件，整体写入（不合并原有数据）
    let save (path: string) (data: WorkbookData) : unit =
        if System.IO.File.Exists path then
            System.IO.File.Delete path
        use conn = openConnection path
        createSchema conn
        use tran = conn.BeginTransaction()
        match data.Name with
        | Some name -> setWorkbookName tran name
        | None -> ()
        data.Worksheets |> Array.iter (insertWorksheet tran)
        data.Cells |> Array.iter (insertCell tran)
        tran.Commit()
