///「复制选择」：把工作簿文件名和当前选择地址，以
/// "工作簿文件名"
/// '工作表名'!A1:C3
/// 格式写入系统剪贴板。不复制单元格当前值。
module ExcelNumericalMethods.SelectionCopy

open System
open System.IO
open Microsoft.Office.Interop.Excel

/// 选择区域地址：A1 样式、不带 $；多区域由 Excel 以逗号连接。
let addressOf (sel: Range) : string =
    string (sel.get_Address(false, false))

/// 拼装剪贴板文本：
/// "工作簿文件名"
/// '工作表名'!A1:C3
/// 工作簿参数取文件名（去掉路径）；工作表名内的单引号按 Excel 惯例加倍。
let format
    (workbookPath: string)
    (sheetName: string)
    (address: string)
    : string
    =
    let nl = Environment.NewLine
    let fileName = Path.GetFileName workbookPath
    let sheet = sheetName.Replace("'", "''")
    sprintf "\"%s\"%s'%s'!%s" fileName nl sheet address

/// 将文本写入系统剪贴板；剪贴板被其他程序占用时重试数次。
let setClipboardText (text: string) : unit =
    let rec retry attempt =
        if attempt <= 0 then
            raise(Runtime.InteropServices.ExternalException "无法写入剪贴板，可能被其他程序占用，请重试。")
        else
            try
                System.Windows.Forms.Clipboard.SetDataObject(text, true, 5, 100)
            with :? Runtime.InteropServices.ExternalException ->
                System.Threading.Thread.Sleep(100)
                retry(attempt - 1)
    retry 5

/// 把当前选择复制到剪贴板，返回生成的文本（供状态栏提示等使用）。
let copySelection (sel: Range) (workbookPath: string) (sheetName: string) : string =
    let text = format workbookPath sheetName (addressOf sel)
    setClipboardText text
    text
