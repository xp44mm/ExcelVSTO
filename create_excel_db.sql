-- Excel 数据库结构（ExcelVSTO 插件：工作簿另存为SQLite / 从SQLite创建工作簿）
-- 一个数据库文件对应一个工作簿；另存时整体覆盖文件，不合并原有数据。
-- 本文件与 ExcelNumericalMethods/SqliteWorkbook.fs 中的 createSchemaSql 保持一致。
-- 三张表：
--   Workbook  工作簿的名称
--   Worksheet 工作表的顺序和名称
--   Cell      单元格：所在工作表、行地址、列地址、值、公式

CREATE TABLE IF NOT EXISTS Workbook (
    name TEXT PRIMARY KEY
);

CREATE TABLE IF NOT EXISTS Worksheet (
    position INTEGER NOT NULL,      -- 工作表顺序，从 1 开始
    name     TEXT NOT NULL UNIQUE,  -- 工作表名称
    PRIMARY KEY (position)
);

CREATE TABLE IF NOT EXISTS Cell (
    worksheet TEXT NOT NULL REFERENCES Worksheet(name),
    row       INTEGER NOT NULL,     -- 行地址，从 1 开始
    col       INTEGER NOT NULL,     -- 列地址，从 1 开始
    value     TEXT,                 -- 单元格值；空为 NULL
    formula   TEXT,                 -- 公式；无公式为 NULL
    PRIMARY KEY (worksheet, row, col)
);
