-- Excel 数据库结构（ExcelVSTO 插件：工作簿另存为SQLite / 从SQLite创建工作簿）
-- 一个数据库文件对应一个工作簿；另存时整体覆盖文件，不合并原有数据。
-- 本文件与 ExcelNumericalMethods/SqliteWorkbook.fs 中的 createSchemaSql 保持一致。
-- 三张表：
--   Workbook  工作簿的名称
--   Worksheet 工作表的顺序和名称
--   Cell      单元格：所在工作表、行地址、列地址、公式、格式

CREATE TABLE Workbook (
    name TEXT PRIMARY KEY
);

CREATE TABLE Worksheet (
    position INTEGER NOT NULL,      -- 工作表顺序，从 1 开始
    name     TEXT NOT NULL UNIQUE,  -- 工作表名称
    PRIMARY KEY (position)
);

CREATE TABLE Cell (
    worksheet TEXT NOT NULL REFERENCES Worksheet(name),
    row       INTEGER NOT NULL,     -- 行地址，从 1 开始
    col       INTEGER NOT NULL,     -- 列地址，从 1 开始
    formula   TEXT,                 -- 单元格内容：公式或文本形式的值
    format    TEXT,                 -- 单元格格式（数字格式）
    PRIMARY KEY (worksheet, row, col)
);
