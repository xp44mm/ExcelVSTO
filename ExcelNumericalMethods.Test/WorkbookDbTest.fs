namespace ExcelNumericalMethods.Test

open ExcelWorkbookDb
open Xunit

type WorkbookDbTest(output: ITestOutputHelper) =

    /// 临时数据库路径
    let tempPath () =
        System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.Guid.NewGuid().ToString("N") + ".db")

    /// 从运行目录向上查找包含 ExcelWorkbookDb\create_excel_db.sql 的项目根目录
    let rec findProjectRoot (dir: string) =
        if System.IO.File.Exists(System.IO.Path.Combine(dir, "ExcelWorkbookDb", "create_excel_db.sql")) then
            dir
        else
            let parent = System.IO.Directory.GetParent(dir)
            if isNull parent then
                failwith "向上找不到 ExcelWorkbookDb\create_excel_db.sql"
            else
                findProjectRoot parent.FullName

    [<Fact>]
    member this.``内嵌建表SQL与create_excel_db.sql一致``() =
        let root = findProjectRoot System.AppContext.BaseDirectory
        let expected = System.IO.File.ReadAllText(System.IO.Path.Combine(root, "ExcelWorkbookDb", "create_excel_db.sql"))
        Assert.Equal(expected, WorkbookDb.createSchemaSql)

    [<Fact>]
    member this.``save后load往返一致``() =
        let path = tempPath()
        try
            let data =
                { Name = "测试工作簿"
                  Worksheets =
                    [| { Position = 1; Name = "Sheet1" }
                       { Position = 2; Name = "Sheet2" } |]
                  Cells =
                    [| { Worksheet = "Sheet1"
                         Row = 1
                         Col = 1
                         Formula = "=1+1"
                         NumberFormat = "0.00" }
                       { Worksheet = "Sheet1"
                         Row = 2
                         Col = 1
                         Formula = "abc"
                         NumberFormat = WorkbookDb.DefaultNumberFormat }
                       { Worksheet = "Sheet2"
                         Row = 1
                         Col = 1
                         Formula = ""
                         NumberFormat = WorkbookDb.DefaultNumberFormat } |] }
            WorkbookDb.save path data
            let actual = WorkbookDb.load path
            Assert.Equal(data.Name, actual.Name)
            Assert.Equal<WorksheetRow[]>(data.Worksheets, actual.Worksheets)
            Assert.Equal<CellRow[]>(data.Cells, actual.Cells)
        finally
            System.IO.File.Delete path

    [<Fact>]
    member this.``createDatabase后事务写入并读取``() =
        let path = tempPath()
        try
            WorkbookDb.createDatabase path "测试库"
            WorkbookDb.withTransaction path (fun tran ->
                WorkbookDb.insertWorksheet tran { Position = 1; Name = "S1" }
                WorkbookDb.insertCell tran
                    { Worksheet = "S1"
                      Row = 3
                      Col = 2
                      Formula = "=A1"
                      NumberFormat = WorkbookDb.DefaultNumberFormat })
            WorkbookDb.withConnection path (fun conn ->
                let name = WorkbookDb.getWorkbookName conn
                Assert.Equal("测试库", name)
                let sheets = WorkbookDb.getWorksheets conn
                Assert.Equal(1, sheets.Length)
                Assert.Equal({ Position = 1; Name = "S1" }, sheets.[0])
                let cells = WorkbookDb.getCellsOf conn "S1"
                Assert.Single(cells) |> ignore
                Assert.Equal(3, cells.[0].Row)
                Assert.Equal(2, cells.[0].Col)
                Assert.Equal("=A1", cells.[0].Formula)
                Assert.Equal(WorkbookDb.DefaultNumberFormat, cells.[0].NumberFormat))
        finally
            System.IO.File.Delete path

    [<Fact>]
    member this.``违反外键约束时写入失败``() =
        let path = tempPath()
        try
            WorkbookDb.createDatabase path "测试库"
            Assert.Throws<System.Data.SQLite.SQLiteException>(fun () ->
                WorkbookDb.withTransaction path (fun tran ->
                    WorkbookDb.insertCell tran
                        { Worksheet = "不存在的表"
                          Row = 1
                          Col = 1
                          Formula = "=1"
                          NumberFormat = WorkbookDb.DefaultNumberFormat }))
            |> ignore
        finally
            System.IO.File.Delete path
    [<Fact>]
    member this.``NumberFormat省略时默认General且formula不能为NULL``() =
        let path = tempPath()
        try
            WorkbookDb.createDatabase path "测试库"
            WorkbookDb.withTransaction path (fun tran ->
                WorkbookDb.insertWorksheet tran { Position = 1; Name = "S1" })
            WorkbookDb.withConnection path (fun conn ->
                // 省略 NumberFormat 列：应取默认值 General
                use cmd =
                    new System.Data.SQLite.SQLiteCommand(
                        "INSERT INTO Cell (worksheet, row, col, formula) VALUES ('S1', 1, 1, '=1');",
                        conn)
                cmd.ExecuteNonQuery() |> ignore
                // formula 为 NULL 违反 NOT NULL 约束
                Assert.Throws<System.Data.SQLite.SQLiteException>(fun () ->
                    use cmd2 =
                        new System.Data.SQLite.SQLiteCommand(
                            "INSERT INTO Cell (worksheet, row, col, formula) VALUES ('S1', 1, 2, NULL);",
                            conn)
                    cmd2.ExecuteNonQuery() |> ignore)
                |> ignore)
            let data = WorkbookDb.load path
            Assert.Equal(1, data.Cells.Length)
            Assert.Equal("=1", data.Cells.[0].Formula)
            Assert.Equal(WorkbookDb.DefaultNumberFormat, data.Cells.[0].NumberFormat)
        finally
            System.IO.File.Delete path
