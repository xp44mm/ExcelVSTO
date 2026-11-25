///「复制当前单元格」：把当前单元格的简单当前值作为单行纯文本写入剪贴板。
/// 多行文本只取第一段（首个 CR 或 LF 之前的单行），其余内容忽略；
/// 没有可复制内容时不改动剪贴板，返回空字符串供调用方提示。
module ExcelNumericalMethods.ActiveCellCopy

open System
open Microsoft.Office.Interop.Excel

/// 读取单元格当前值；错误单元格的 Value2 是负 Int32 错误码，改取显示文本（如 #N/A）。
let private valueTextOf (cell: Range) : string =
    match cell.Value2 with
    | null -> ""
    | :? string as s -> s
    | :? int as n when n < 0 -> string cell.Text
    | v -> string v

/// 只保留第一段单行文本：遇到第一个 CR 或 LF 即截断。
let private firstSingleLine (text: string) : string =
    if isNull text then ""
    else
        let endIndex = text.IndexOfAny([| '\r'; '\n' |])
        if endIndex < 0 then text else text.Substring(0, endIndex)

/// 复制当前单元格的单行纯文本到剪贴板；没有内容时返回空字符串且不改动剪贴板。
let copyActiveCell (cell: Range) : string =
    let text = cell |> valueTextOf |> firstSingleLine
    if String.IsNullOrWhiteSpace text then
        ""
    else
        SelectionCopy.setClipboardText text
        text
