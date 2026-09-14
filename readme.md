# ExcelWorkbookDb

Excel 工作簿 SQLite 数据库访问包装库（netstandard2.0）。

本库只封装数据库访问，不依赖 Excel 互操作，可被任意 netstandard2.0 应用引用。建表 SQL 以程序集资源 `ExcelWorkbookDb.create_excel_db.sql` 内嵌在包中，是数据库结构的唯一事实来源。

## 数据库结构

共三张表：

| 表 | 字段 | 说明 |
|---|---|---|
| Workbook | name | 工作簿名称（单行表） |
| Worksheet | position（主键）、name | 工作表的顺序与名称 |
| Cell | worksheet、row、col（复合主键）、formula、NumberFormat | 单元格内容与数字格式 |

- `Cell.worksheet` 外键引用 `Worksheet.name`，打开连接时默认开启外键强制（`PRAGMA foreign_keys = ON`）。
- `formula` 不允许为 NULL（公式为空则该单元格不写入）。
- `NumberFormat` 不允许为 NULL，默认 `'General'`（对应 Excel API `Range.NumberFormat`）。

## 用法

```fsharp
open ExcelWorkbookDb

// 新建数据库文件（覆盖已有文件）并执行建表 SQL
WorkbookDb.createDatabase path

// 整体保存：名称 + 工作表 + 单元格（覆盖写入，不合并原有数据）
let data =
    { Name = Some "工作簿1"
      Worksheets = [| { Position = 1; Name = "Sheet1" } |]
      Cells =
        [| { Worksheet = "Sheet1"
             Row = 1
             Col = 1
             Formula = "=1+2"
             NumberFormat = WorkbookDb.DefaultNumberFormat } |] }
WorkbookDb.save path data

// 读取完整内容
let loaded = WorkbookDb.load path

// 事务写入：成功提交，异常回滚
WorkbookDb.withTransaction path (fun tran ->
    WorkbookDb.insertWorksheet tran { Position = 2; Name = "Sheet2" }
    WorkbookDb.insertCell tran { Worksheet = "Sheet2"; Row = 1; Col = 1; Formula = "=A1"; NumberFormat = WorkbookDb.DefaultNumberFormat })

// 读取指定工作表的单元格
WorkbookDb.withConnection path (fun conn -> WorkbookDb.getCellsOf conn "Sheet1")
```

## API

### 记录类型

- `WorksheetRow { Position: int; Name: string }` —— 对应 Worksheet 表
- `CellRow { Worksheet: string; Row: int; Col: int; Formula: string; NumberFormat: string }` —— 对应 Cell 表
- `WorkbookData { Name: string option; Worksheets: WorksheetRow[]; Cells: CellRow[] }` —— 数据库完整内容

### 模块 `WorkbookDb`

| 函数 | 说明 |
|---|---|
| `createDatabase path` | 新建数据库文件（覆盖）并执行建表 SQL |
| `save path data` | 将完整工作簿数据整体写入文件（覆盖，不合并原有数据） |
| `load path` | 读取数据库完整内容 |
| `withConnection path f` | 打开连接执行读取操作，结束后释放（开启外键强制） |
| `withTransaction path f` | 在单个事务中执行写入操作，成功提交、异常回滚 |
| `getWorkbookName conn` | 读取工作簿名称 |
| `getWorksheets conn` | 读取全部工作表（按 position 升序） |
| `getCells conn` | 读取全部单元格（按 worksheet, row, col 升序） |
| `getCellsOf conn worksheet` | 读取指定工作表的单元格 |
| `setWorkbookName tran name` | 写入工作簿名称（先清空 Workbook 表） |
| `insertWorksheet tran ws` | 插入一条工作表记录 |
| `insertCell tran cell` | 插入一条单元格记录 |
| `clear tran` | 清空三张表（保留表结构） |
| `DefaultNumberFormat` | 数字格式默认值常量 `"General"` |

## 构建与打包

```powershell
dotnet build ExcelWorkbookDb\ExcelWorkbookDb.fsproj -c Release
dotnet pack ExcelWorkbookDb\ExcelWorkbookDb.fsproj -c Release
# 产物：ExcelWorkbookDb\bin\Release\ExcelWorkbookDb.1.0.0.nupkg
```
