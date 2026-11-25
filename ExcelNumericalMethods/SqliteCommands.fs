///「SQLite」命令的交互层：弹出文件对话框、调用 SqliteWorkbook 完成读写、提示结果。
/// C# 回调只做链接；本模块是这些命令的功能实现。
module ExcelNumericalMethods.SqliteCommands

open System
open System.IO
open System.Windows.Forms
open Microsoft.Office.Interop.Excel

let private showError (message: string) =
    MessageBox.Show(message, "ExcelVSTO", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore

/// 把当前工作簿另存为 SQLite 数据库（三张表，直接覆盖目标文件）。
let saveToSqlite (app: Application) =
    let wb = app.ActiveWorkbook
    if isNull wb then
        showError "当前没有打开的工作簿！"
    else
        use dlg = new SaveFileDialog()
        dlg.Title <- "另存为 SQLite 数据库"
        dlg.Filter <- "SQLite 数据库 (*.db)|*.db|所有文件 (*.*)|*.*"
        dlg.DefaultExt <- ".db"
        dlg.AddExtension <- true
        dlg.OverwritePrompt <- true
        dlg.InitialDirectory <- wb.Path
        dlg.FileName <- Path.GetFileNameWithoutExtension(wb.Name) + ".db"
        if dlg.ShowDialog() = DialogResult.OK then
            try
                SqliteWorkbook.saveWorkbookAs dlg.FileName wb
                MessageBox.Show("工作簿已保存到 SQLite 数据库：\n" + dlg.FileName) |> ignore
            with ex -> showError ex.Message

/// 从 SQLite 数据库导入工作表，追加到当前工作簿尾部，不新建工作簿、不保存文件。
/// 表名与现有表冲突时自动改名。
let createFromSqlite (app: Application) =
    let wb = app.ActiveWorkbook
    if isNull wb then
        showError "当前没有打开的工作簿！"
    else
        use dlg = new OpenFileDialog()
        dlg.Title <- "从 SQLite 数据库导入工作表"
        dlg.Filter <- "SQLite 数据库 (*.db)|*.db|所有文件 (*.*)|*.*"
        dlg.CheckFileExists <- true
        dlg.InitialDirectory <- wb.Path
        if dlg.ShowDialog() = DialogResult.OK then
            try
                let names = SqliteWorkbook.appendSheetsFrom wb dlg.FileName
                app.StatusBar <- box ("已在当前工作簿尾部追加 " + string names.Length + " 张工作表：" + String.Join("、", names))
            with ex -> showError ex.Message
