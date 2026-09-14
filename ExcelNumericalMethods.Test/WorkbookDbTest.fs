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
                { Name = Some "测试工作簿"
                  Worksheets =
                    [| { Position = 1; Name = "Sheet1" }
                       { Position = 2; Name = "Sheet2" } |]
                  Cells =
                    [| { Worksheet = "Sheet1"
                         Row = 1
                         Col = 1
                         Formula = Some "=1+1"
                         Format = Some "0.00" }
                       { Worksheet = "Sheet1"
                         Row = 2
                         Col = 1
                         Formula = Some "abc"
                         Format = None }
                       { Worksheet = "Sheet2"
                         Row = 1
                         Col = 1
                         Formula = None
                         Format = None } |] }
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
            WorkbookDb.createDatabase path
            WorkbookDb.withTransaction path (fun tran ->
                WorkbookDb.insertWorksheet tran { Position = 1; Name = "S1" }
                WorkbookDb.insertCell tran
                    { Worksheet = "S1"
                      Row = 3
                      Col = 2
                      Formula = Some "=A1"
                      Format = None })
            WorkbookDb.withConnection path (fun conn ->
                let name = WorkbookDb.getWorkbookName conn
                Assert.Null(name)
                let sheets = WorkbookDb.getWorksheets conn
                Assert.Equal(1, sheets.Length)
                Assert.Equal({ Position = 1; Name = "S1" }, sheets.[0])
                let cells = WorkbookDb.getCellsOf conn "S1"
                Assert.Single(cells) |> ignore
                Assert.Equal(3, cells.[0].Row)
                Assert.Equal(2, cells.[0].Col)
                Assert.Equal(Some "=A1", cells.[0].Formula)
                Assert.Equal(None, cells.[0].Format))
        finally
            System.IO.File.Delete path

    [<Fact>]
    member this.``违反外键约束时写入失败``() =
        let path = tempPath()
        try
            WorkbookDb.createDatabase path
            Assert.Throws<System.Data.SQLite.SQLiteException>(fun () ->
                WorkbookDb.withTransaction path (fun tran ->
                    WorkbookDb.insertCell tran
                        { Worksheet = "不存在的表"
                          Row = 1
                          Col = 1
                          Formula = Some "=1"
                          Format = None }))
            |> ignore
        finally
            System.IO.File.Delete path
